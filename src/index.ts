import PostalMime from "postal-mime";
import type { Attachment } from "postal-mime";
import ExcelJS from "exceljs";
import {
  extractYardCheckSource,
  generateYardCheckWorkbookBuffer,
} from "./yardCheck";

const DEFAULT_MAX_ATTACHMENT_BYTES = 30 * 1024 * 1024;
const SITE_TITLE = "Minot Vehicle ETIC";

type SheetSummary = {
  name: string;
  state: "visible" | "hidden" | "veryHidden";
  rowCount: number;
  columnCountEstimate: number;
  sampleHeaders: string[];
  nonEmptyCellCountEstimate: number;
  melMentions: number;
};

type AnalysisResult = {
  workbookFileName: string;
  receivedAtIso: string;
  dateKey: string;
  from: string;
  to: string;
  subject: string;
  workbookBytes: number;
  sheetSummaries: SheetSummary[];
  totalVisibleSheets: number;
  totalHiddenSheets: number;
  totalRowsAcrossSheets: number;
  melMentionsBySheet: Record<string, number>;
  melMentionsTotal: number;
};

type HistoryEntryDiff = {
  previousDateKey: string | null;
  deltaTotalRows: number | null;
  deltaMelMentionsTotal: number | null;
  deltaSheetsVisible: number | null;
};

type HistoryEntry = {
  dateKey: string;
  receivedAtIso: string;
  workbookFileName: string;
  workbookKey: string;
  analysisKey: string;
  totalVisibleSheets: number;
  totalHiddenSheets: number;
  totalRowsAcrossSheets: number;
  melMentionsTotal: number;
  diff: HistoryEntryDiff;
};

type HistoryIndex = {
  updatedAtIso: string;
  entries: HistoryEntry[];
};

interface Env {
  ETIC_BUCKET: R2Bucket;
  ALLOWED_SENDERS?: string;
  EXPECTED_ATTACHMENT_NAME?: string;
  MAX_ATTACHMENT_BYTES?: string;
}

export default {
  async email(message: ForwardableEmailMessage, env: Env, _ctx: ExecutionContext): Promise<void> {
    if (!isAuthorizedSender(message.from, env)) {
      message.setReject("Unauthorized sender");
      return;
    }

    const parsed = await PostalMime.parse(message.raw, { attachmentEncoding: "arraybuffer" });
    const workbookAttachment = pickWorkbookAttachment(parsed.attachments, env.EXPECTED_ATTACHMENT_NAME);
    if (!workbookAttachment) {
      console.warn(
        JSON.stringify({
          level: "warn",
          message: "No .xlsx attachment found on inbound email",
          from: message.from,
          to: message.to,
          subject: parsed.subject ?? "",
        }),
      );
      return;
    }

    const workbookBytes = normalizeAttachmentBinary(workbookAttachment.content);
    if (workbookBytes.byteLength > parseMaxAttachmentBytes(env.MAX_ATTACHMENT_BYTES)) {
      message.setReject("Attachment too large");
      return;
    }

    const now = new Date();
    const subject = parsed.subject ?? "";
    const dateKey = resolveAnalysisDateKey(subject, now);
    const safeName = sanitizeFileName(workbookAttachment.filename ?? "vehicle-etic.xlsx");
    const workbookKey = `workbooks/${dateKey}/${safeName}`;
    const analysisKey = `analyses/${dateKey}.json`;

    await env.ETIC_BUCKET.put(workbookKey, workbookBytes, {
      httpMetadata: {
        contentType:
          workbookAttachment.mimeType ||
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      customMetadata: {
        from: message.from,
        to: message.to,
        subject,
      },
    });

    const analysis = await analyzeWorkbook({
      binary: workbookBytes,
      fileName: safeName,
      receivedAtIso: now.toISOString(),
      dateKey,
      from: message.from,
      to: message.to,
      subject,
    });

    await env.ETIC_BUCKET.put(analysisKey, JSON.stringify(analysis, null, 2), {
      httpMetadata: { contentType: "application/json; charset=utf-8" },
    });

    const history = await loadHistory(env);
    const upsertedHistory = upsertHistoryEntry(history, analysis, workbookKey, analysisKey);
    await env.ETIC_BUCKET.put(
      "history/index.json",
      JSON.stringify(upsertedHistory, null, 2),
      {
        httpMetadata: { contentType: "application/json; charset=utf-8" },
      },
    );
    await env.ETIC_BUCKET.put(
      "analyses/latest.json",
      JSON.stringify(analysis, null, 2),
      {
        httpMetadata: { contentType: "application/json; charset=utf-8" },
      },
    );
  },

  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);

    if (url.pathname === "/healthz") {
      return Response.json({ ok: true, service: "etic-email-automation" });
    }

    if (url.pathname === "/api/history") {
      const history = await loadHistory(env);
      return Response.json(history, { headers: cacheHeaders() });
    }

    if (url.pathname === "/api/latest") {
      const analysis = await readJson<AnalysisResult>(env, "analyses/latest.json");
      if (!analysis) return new Response("Not Found", { status: 404 });
      return Response.json(analysis, { headers: cacheHeaders() });
    }

    if (url.pathname.startsWith("/api/analysis/")) {
      const dateKey = url.pathname.slice("/api/analysis/".length);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) {
        return new Response("Invalid date key", { status: 400 });
      }
      const analysis = await readJson<AnalysisResult>(env, `analyses/${dateKey}.json`);
      if (!analysis) return new Response("Not Found", { status: 404 });
      return Response.json(analysis, { headers: cacheHeaders() });
    }

    if (url.pathname === "/api/yard-check.xlsx") {
      return handleYardCheckDownload(env);
    }

    if (url.pathname === "/api/yard-check-meta") {
      return handleYardCheckMeta(env);
    }

    if (url.pathname === "/" || url.pathname === "/index.html") {
      return new Response(renderDashboardHtml(), {
        headers: {
          "Content-Type": "text/html; charset=utf-8",
          "Cache-Control": "no-store",
        },
      });
    }

    return new Response("Not Found", { status: 404 });
  },
} satisfies ExportedHandler<Env>;

async function handleYardCheckDownload(env: Env): Promise<Response> {
  const history = await loadHistory(env);
  const latestEntry = history.entries.length
    ? [...history.entries].sort((a, b) => a.dateKey.localeCompare(b.dateKey)).pop() ?? null
    : null;

  const latest = await readJson<AnalysisResult>(env, "analyses/latest.json");

  const workbookKey = latestEntry?.workbookKey ?? null;
  const sourceFileName = latestEntry?.workbookFileName ?? latest?.workbookFileName ?? "vehicle-etic.xlsx";
  const sourceDateKey = latestEntry?.dateKey ?? latest?.dateKey ?? "latest";

  if (!workbookKey) {
    return Response.json(
      { ok: false, error: "No source workbook found yet. Send the ETIC email to ingest first." },
      { status: 404 },
    );
  }

  const object = await env.ETIC_BUCKET.get(workbookKey);
  if (!object) {
    return Response.json(
      { ok: false, error: `Source workbook not found in R2: ${workbookKey}` },
      { status: 404 },
    );
  }

  const workbookBytes = await object.arrayBuffer();
  const source = await extractYardCheckSource(workbookBytes);
  if (!source) {
    return Response.json(
      {
        ok: false,
        error: "Could not locate a Work Orders sheet in the latest workbook.",
      },
      { status: 422 },
    );
  }

  const generatedAtIso = new Date().toISOString();
  const buffer = await generateYardCheckWorkbookBuffer(source, {
    sourceWorkbookFileName: sourceFileName,
    sourceDateKey,
    generatedAtIso,
  });

  const downloadName = `yard-check_${sourceDateKey}_${generatedAtIso.replace(/[:.]/g, "-")}.xlsx`;
  return new Response(buffer, {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${downloadName}"`,
      "Cache-Control": "no-store",
      "X-Source-Workbook": workbookKey,
      "X-Source-Date": sourceDateKey,
      "X-Work-Order-Sheet": source.workOrderSheet ?? "",
      "X-Fleet-Sheet": source.fleetSheet ?? "",
      "X-Total-Assets": String(source.totalAssets),
      "X-Total-Work-Orders": String(source.totalWorkOrders),
    },
  });
}

type YardCheckMetaOk = {
  ok: true;
  sourceDateKey: string;
  sourceFileName: string;
  workbookKey: string;
  receivedAtIso: string | null;
  subject: string | null;
  totalAssets: number;
  totalWorkOrders: number;
  workOrderSheet: string | null;
  fleetSheet: string | null;
};

type YardCheckMetaErr = { ok: false; error: string };

async function handleYardCheckMeta(env: Env): Promise<Response> {
  const meta = await buildYardCheckMeta(env);
  const status = meta.ok ? 200 : meta.error.includes("No source") ? 404 : 422;
  return Response.json(meta, { status, headers: cacheHeaders() });
}

async function buildYardCheckMeta(env: Env): Promise<YardCheckMetaOk | YardCheckMetaErr> {
  const history = await loadHistory(env);
  const latestEntry = history.entries.length
    ? [...history.entries].sort((a, b) => a.dateKey.localeCompare(b.dateKey)).pop() ?? null
    : null;
  const latest = await readJson<AnalysisResult>(env, "analyses/latest.json");
  const workbookKey = latestEntry?.workbookKey ?? null;
  const sourceFileName = latestEntry?.workbookFileName ?? latest?.workbookFileName ?? "vehicle-etic.xlsx";
  const sourceDateKey = latestEntry?.dateKey ?? latest?.dateKey ?? "latest";

  if (!workbookKey) {
    return { ok: false, error: "No source workbook found yet. Send the ETIC email to ingest first." };
  }

  const object = await env.ETIC_BUCKET.get(workbookKey);
  if (!object) {
    return { ok: false, error: `Source workbook not found in R2: ${workbookKey}` };
  }

  const workbookBytes = await object.arrayBuffer();
  const source = await extractYardCheckSource(workbookBytes);
  if (!source) {
    return { ok: false, error: "Could not locate a Work Orders sheet in the latest workbook." };
  }

  return {
    ok: true,
    sourceDateKey,
    sourceFileName,
    workbookKey,
    receivedAtIso: latest?.receivedAtIso ?? null,
    subject: latest?.subject ?? null,
    totalAssets: source.totalAssets,
    totalWorkOrders: source.totalWorkOrders,
    workOrderSheet: source.workOrderSheet ?? null,
    fleetSheet: source.fleetSheet ?? null,
  };
}

function cacheHeaders(): HeadersInit {
  return {
    "Cache-Control": "no-store",
  };
}

async function readJson<T>(env: Env, key: string): Promise<T | null> {
  const object = await env.ETIC_BUCKET.get(key);
  if (!object) return null;
  try {
    const text = await object.text();
    return JSON.parse(text) as T;
  } catch (error) {
    console.error(
      JSON.stringify({
        level: "error",
        message: "Failed to parse R2 JSON object",
        key,
        error: error instanceof Error ? error.message : String(error),
      }),
    );
    return null;
  }
}

async function loadHistory(env: Env): Promise<HistoryIndex> {
  const existing = await readJson<HistoryIndex>(env, "history/index.json");
  if (existing && Array.isArray(existing.entries)) {
    return existing;
  }
  return { updatedAtIso: new Date().toISOString(), entries: [] };
}

function upsertHistoryEntry(
  history: HistoryIndex,
  analysis: AnalysisResult,
  workbookKey: string,
  analysisKey: string,
): HistoryIndex {
  const filtered = history.entries.filter((entry) => entry.dateKey !== analysis.dateKey);
  const sortedPrevious = [...filtered].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  const previous = sortedPrevious.length > 0 ? sortedPrevious[sortedPrevious.length - 1] : null;

  const diff: HistoryEntryDiff = {
    previousDateKey: previous ? previous.dateKey : null,
    deltaTotalRows: previous ? analysis.totalRowsAcrossSheets - previous.totalRowsAcrossSheets : null,
    deltaMelMentionsTotal: previous
      ? analysis.melMentionsTotal - previous.melMentionsTotal
      : null,
    deltaSheetsVisible: previous
      ? analysis.totalVisibleSheets - previous.totalVisibleSheets
      : null,
  };

  const entry: HistoryEntry = {
    dateKey: analysis.dateKey,
    receivedAtIso: analysis.receivedAtIso,
    workbookFileName: analysis.workbookFileName,
    workbookKey,
    analysisKey,
    totalVisibleSheets: analysis.totalVisibleSheets,
    totalHiddenSheets: analysis.totalHiddenSheets,
    totalRowsAcrossSheets: analysis.totalRowsAcrossSheets,
    melMentionsTotal: analysis.melMentionsTotal,
    diff,
  };

  const merged = [...filtered, entry].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  return {
    updatedAtIso: new Date().toISOString(),
    entries: merged,
  };
}

function isAuthorizedSender(sender: string, env: Env): boolean {
  const configured = env.ALLOWED_SENDERS?.trim();
  if (!configured) return true;
  if (configured === "*") return true;
  const normalizedSender = normalizeEmailAddress(sender);
  const allowed = configured
    .split(",")
    .map((part) => normalizeEmailAddress(part))
    .filter(Boolean);
  return allowed.includes(normalizedSender);
}

function pickWorkbookAttachment(
  attachments: Attachment[],
  expectedAttachmentName?: string,
): Attachment | undefined {
  const expected = expectedAttachmentName?.trim().toLowerCase();
  if (expected) {
    const exact = attachments.find(
      (attachment) => (attachment.filename?.trim().toLowerCase() ?? "") === expected,
    );
    if (exact) return exact;
  }

  return attachments.find((attachment) => {
    const filename = attachment.filename?.toLowerCase() ?? "";
    return filename.endsWith(".xlsx");
  });
}

function normalizeAttachmentBinary(content: Attachment["content"]): ArrayBuffer {
  if (content instanceof ArrayBuffer) return content;
  if (content instanceof Uint8Array) return cloneToArrayBuffer(content);
  const encoded = new TextEncoder().encode(content);
  return cloneToArrayBuffer(encoded);
}

function cloneToArrayBuffer(view: ArrayBufferView): ArrayBuffer {
  const clone = new Uint8Array(view.byteLength);
  clone.set(new Uint8Array(view.buffer, view.byteOffset, view.byteLength));
  return clone.buffer;
}

function sanitizeFileName(fileName: string): string {
  return fileName.replace(/[^A-Za-z0-9._-]/g, "_");
}

function normalizeEmailAddress(email: string): string {
  return email.replace(/[<>]/g, "").trim().toLowerCase();
}

function parseMaxAttachmentBytes(raw: string | undefined): number {
  if (!raw) return DEFAULT_MAX_ATTACHMENT_BYTES;
  const value = Number.parseInt(raw, 10);
  return Number.isFinite(value) && value > 0 ? value : DEFAULT_MAX_ATTACHMENT_BYTES;
}

function isoDateKey(date: Date): string {
  const p = (n: number) => String(n).padStart(2, "0");
  return `${date.getUTCFullYear()}-${p(date.getUTCMonth() + 1)}-${p(date.getUTCDate())}`;
}

const MONTH_TOKEN: Record<string, number> = {
  jan: 1,
  feb: 2,
  mar: 3,
  apr: 4,
  may: 5,
  jun: 6,
  jul: 7,
  aug: 8,
  sep: 9,
  oct: 10,
  nov: 11,
  dec: 12,
};

/** Last DD-MMM-YY (or DD-MMM-YYYY) token in the subject → YYYY-MM-DD for R2 paths. */
function parseReportDateKeyFromSubject(subject: string): string | null {
  const s = subject.trim();
  let reportDateKey: string | null = null;
  const dateRe = /\b(\d{1,2})-([A-Za-z]{3})\.?-(\d{2,4})\b/g;
  let dm: RegExpExecArray | null;
  while ((dm = dateRe.exec(s)) !== null) {
    const day = Number.parseInt(dm[1] ?? "", 10);
    const monToken = (dm[2] ?? "").toLowerCase().replace(/\.$/, "").slice(0, 3);
    const month = MONTH_TOKEN[monToken];
    if (!month || !Number.isFinite(day) || day < 1 || day > 31) continue;
    const yRaw = dm[3] ?? "";
    const yNum = Number.parseInt(yRaw, 10);
    if (!Number.isFinite(yNum)) continue;
    const year = yRaw.length <= 2 ? (yNum >= 70 ? 1900 + yNum : 2000 + yNum) : yNum;
    reportDateKey = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }
  return reportDateKey;
}

/** Prefer report date from subject (forwarded ETICs); else UTC day of receipt. */
function resolveAnalysisDateKey(subject: string, receivedAt: Date): string {
  return parseReportDateKeyFromSubject(subject) ?? isoDateKey(receivedAt);
}

async function analyzeWorkbook(input: {
  binary: ArrayBuffer;
  fileName: string;
  receivedAtIso: string;
  dateKey: string;
  from: string;
  to: string;
  subject: string;
}): Promise<AnalysisResult> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(input.binary);

  const sheetSummaries = workbook.worksheets.map((sheet): SheetSummary => {
    const state = normalizeSheetState(sheet.state);
    const rowCount = estimateRowCount(sheet);
    const columnCountEstimate = estimateColumnCount(sheet);
    const sampleHeaders = collectHeaders(sheet);
    const nonEmptyCellCountEstimate = countNonEmptyCells(sheet);
    const melMentions = countMelMentions(sheet);

    return {
      name: sheet.name,
      state,
      rowCount,
      columnCountEstimate,
      sampleHeaders,
      nonEmptyCellCountEstimate,
      melMentions,
    };
  });

  const totalVisibleSheets = sheetSummaries.filter((s) => s.state === "visible").length;
  const totalHiddenSheets = sheetSummaries.filter((s) => s.state !== "visible").length;
  const totalRowsAcrossSheets = sheetSummaries.reduce((sum, s) => sum + s.rowCount, 0);
  const melMentionsBySheet = Object.fromEntries(sheetSummaries.map((s) => [s.name, s.melMentions]));
  const melMentionsTotal = sheetSummaries.reduce((sum, s) => sum + s.melMentions, 0);

  return {
    workbookFileName: input.fileName,
    receivedAtIso: input.receivedAtIso,
    dateKey: input.dateKey,
    from: input.from,
    to: input.to,
    subject: input.subject,
    workbookBytes: input.binary.byteLength,
    sheetSummaries,
    totalVisibleSheets,
    totalHiddenSheets,
    totalRowsAcrossSheets,
    melMentionsBySheet,
    melMentionsTotal,
  };
}

function normalizeSheetState(state: ExcelJS.WorksheetState | undefined): SheetSummary["state"] {
  if (state === "hidden") return "hidden";
  if (state === "veryHidden") return "veryHidden";
  return "visible";
}

function estimateRowCount(sheet: ExcelJS.Worksheet): number {
  let rows = 0;
  sheet.eachRow({ includeEmpty: false }, () => {
    rows += 1;
  });
  return rows;
}

function estimateColumnCount(sheet: ExcelJS.Worksheet): number {
  let max = 0;
  sheet.eachRow({ includeEmpty: false }, (row) => {
    const actual = row.actualCellCount;
    if (actual > max) max = actual;
  });
  return max;
}

function collectHeaders(sheet: ExcelJS.Worksheet): string[] {
  const firstRow = sheet.getRow(1);
  const headers: string[] = [];
  firstRow.eachCell({ includeEmpty: false }, (cell) => {
    const value = readCellText(cell).trim();
    if (value) headers.push(value);
  });
  return headers.slice(0, 12);
}

function countNonEmptyCells(sheet: ExcelJS.Worksheet): number {
  let count = 0;
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, () => {
      count += 1;
    });
  });
  return count;
}

function countMelMentions(sheet: ExcelJS.Worksheet): number {
  let count = 0;
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const text = readCellText(cell).toLowerCase();
      if (text.includes("mel")) count += 1;
    });
  });
  return count;
}

function readCellText(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map((part) => String(part ?? "")).join("");

  if (typeof value === "object") {
    const anyValue = value as unknown as Record<string, unknown>;
    const richText = anyValue.richText;
    if (Array.isArray(richText)) {
      return richText
        .map((part) =>
          part && typeof part === "object" ? String((part as { text?: unknown }).text ?? "") : "",
        )
        .join("");
    }
    if (typeof anyValue.text === "string") return anyValue.text;
    if (anyValue.result !== null && anyValue.result !== undefined) return String(anyValue.result);
    if (typeof anyValue.formula === "string") return anyValue.formula;
    if (typeof anyValue.hyperlink === "string") return anyValue.hyperlink;
    if (typeof anyValue.error === "string") return anyValue.error;
  }

  try {
    return String(value);
  } catch {
    return "";
  }
}

function renderDashboardHtml(): string {
  const escapedTitle = escapeHtml(SITE_TITLE);
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta name="color-scheme" content="dark light" />
  <title>${escapedTitle}</title>
  <style>
    :root {
      color-scheme: dark light;
      --panel: #111a2e;
      --border: #1f2a44;
      --text: #e6ecf5;
      --muted: #96a2b8;
      --accent: #6aa9ff;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Ubuntu, sans-serif;
      background: radial-gradient(1200px 800px at 10% -10%, #1a2748 0%, #0b1220 55%) fixed;
      color: var(--text);
      min-height: 100vh;
    }
    main { padding: 32px 24px 48px; max-width: 560px; margin: 0 auto; }
    .panel {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 28px 24px;
    }
    h1 { margin: 0 0 8px; font-size: 22px; font-weight: 650; }
    .muted { color: var(--muted); font-size: 13px; line-height: 1.5; margin-bottom: 20px; }
    dl { margin: 0 0 24px; display: grid; grid-template-columns: auto 1fr; gap: 8px 16px; font-size: 14px; }
    dt { color: var(--muted); }
    dd { margin: 0; }
    .btn {
      display: block;
      width: 100%;
      background: linear-gradient(180deg, #2a4d8f, #1c3971);
      color: #e7efff;
      border: 1px solid #3a5da5;
      padding: 14px 18px;
      border-radius: 12px;
      font-size: 15px;
      font-weight: 600;
      cursor: pointer;
    }
    .btn:hover { filter: brightness(1.08); }
    .btn:active { transform: translateY(1px); }
    .btn:disabled { opacity: 0.55; cursor: not-allowed; }
    #status { margin-top: 14px; font-size: 13px; min-height: 20px; color: var(--muted); }
    .err { color: #ff9b9b; }
  </style>
</head>
<body>
  <main>
    <div class="panel">
      <h1>Yard check</h1>
      <p class="muted">Download is built from the latest Vehicle ETIC workbook in storage (same file as email ingest).</p>
      <dl id="meta">
        <dt>Source file</dt><dd id="m-file">…</dd>
        <dt>Report day</dt><dd id="m-day">…</dd>
        <dt>Rows</dt><dd id="m-rows">…</dd>
      </dl>
      <button type="button" class="btn" id="download">Download yard check (.xlsx)</button>
      <div id="status"></div>
    </div>
  </main>
  <script>
    const fmt = (n) => typeof n === "number" ? n.toLocaleString() : "—";

    async function loadMeta() {
      const dl = document.getElementById("meta");
      const st = document.getElementById("status");
      try {
        const res = await fetch("/api/yard-check-meta");
        const data = await res.json();
        if (!data.ok) {
          dl.style.display = "none";
          st.className = "err";
          st.textContent = data.error || "Could not load source workbook.";
          return;
        }
        document.getElementById("m-file").textContent = data.sourceFileName;
        document.getElementById("m-day").textContent = data.sourceDateKey;
        document.getElementById("m-rows").textContent =
          fmt(data.totalAssets) + " assets · " + fmt(data.totalWorkOrders) + " work orders";
        st.textContent = "";
      } catch (e) {
        dl.style.display = "none";
        st.className = "err";
        st.textContent = "Failed to load: " + (e && e.message ? e.message : String(e));
      }
    }

    async function download() {
      const btn = document.getElementById("download");
      const st = document.getElementById("status");
      const prev = btn.textContent;
      btn.disabled = true;
      btn.textContent = "Preparing…";
      st.className = "";
      st.textContent = "Generating…";
      try {
        const res = await fetch("/api/yard-check.xlsx");
        if (!res.ok) {
          let msg = "Failed (" + res.status + ")";
          try {
            const j = await res.json();
            if (j && j.error) msg = j.error;
          } catch (_) {}
          st.className = "err";
          st.textContent = msg;
          return;
        }
        const disp = res.headers.get("Content-Disposition") || "";
        const m = /filename="?([^";]+)"?/i.exec(disp);
        const name = m ? m[1] : "yard-check.xlsx";
        const blob = await res.blob();
        const u = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = u;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(u);
        st.textContent = "Saved " + name;
      } catch (e) {
        st.className = "err";
        st.textContent = String(e && e.message ? e.message : e);
      } finally {
        btn.disabled = false;
        btn.textContent = prev;
      }
    }

    document.getElementById("download").addEventListener("click", () => { download(); });
    loadMeta();
  </script>
</body>
</html>`;
}

function escapeHtml(input: string): string {
  return input
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export {
  analyzeWorkbook,
  countMelMentions,
  isoDateKey,
  parseMaxAttachmentBytes,
  parseReportDateKeyFromSubject,
  pickWorkbookAttachment,
  readCellText,
  resolveAnalysisDateKey,
  sanitizeFileName,
  upsertHistoryEntry,
};

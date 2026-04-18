import PostalMime from "postal-mime";
import type { Attachment } from "postal-mime";
import ExcelJS from "exceljs";
import {
  extractWorkOrderRows,
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
    const dateKey = isoDateKey(now);
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
        subject: parsed.subject ?? "",
      },
    });

    const analysis = await analyzeWorkbook({
      binary: workbookBytes,
      fileName: safeName,
      receivedAtIso: now.toISOString(),
      dateKey,
      from: message.from,
      to: message.to,
      subject: parsed.subject ?? "",
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
  const extraction = await extractWorkOrderRows(workbookBytes);
  if (!extraction) {
    return Response.json(
      {
        ok: false,
        error: "Could not locate a Work Orders sheet in the latest workbook.",
      },
      { status: 422 },
    );
  }

  const generatedAtIso = new Date().toISOString();
  const buffer = await generateYardCheckWorkbookBuffer(extraction, {
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
      "X-Total-Rows": String(extraction.totalDataRows),
      "X-Unmatched-Headers": String(extraction.unmatchedHeaders.length),
    },
  });
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
      --bg: #0b1220;
      --panel: #111a2e;
      --panel-alt: #16213d;
      --border: #1f2a44;
      --text: #e6ecf5;
      --muted: #96a2b8;
      --accent: #6aa9ff;
      --up: #ff8c8c;
      --down: #8cff9f;
      --flat: #c7d0df;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Ubuntu, sans-serif;
      background: radial-gradient(1200px 800px at 10% -10%, #1a2748 0%, #0b1220 55%) fixed;
      color: var(--text);
      min-height: 100vh;
    }
    header {
      padding: 24px 32px;
      border-bottom: 1px solid var(--border);
      display: flex;
      align-items: baseline;
      gap: 14px;
      flex-wrap: wrap;
    }
    header h1 { margin: 0; font-size: 22px; letter-spacing: 0.02em; }
    header .sub { color: var(--muted); font-size: 13px; }
    main { padding: 24px 32px 80px; max-width: 1100px; margin: 0 auto; }
    .cards {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 14px;
      margin-bottom: 24px;
    }
    .card {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 16px 18px;
    }
    .card h3 { margin: 0 0 6px; font-size: 12px; color: var(--muted); font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; }
    .card .v { font-size: 24px; font-weight: 600; }
    .card .delta { margin-top: 6px; font-size: 12px; color: var(--muted); }
    .up { color: var(--up); }
    .down { color: var(--down); }
    .flat { color: var(--flat); }
    section { margin-bottom: 28px; }
    h2 { font-size: 16px; margin: 0 0 10px; letter-spacing: 0.04em; }
    table { width: 100%; border-collapse: collapse; background: var(--panel); border: 1px solid var(--border); border-radius: 12px; overflow: hidden; }
    th, td { padding: 10px 12px; text-align: left; font-size: 13px; border-bottom: 1px solid var(--border); }
    th { background: var(--panel-alt); color: var(--muted); font-weight: 600; }
    tr:last-child td { border-bottom: 0; }
    .pill { display: inline-block; padding: 2px 8px; border-radius: 999px; background: #1c2a4a; color: #cdd9ef; font-size: 11px; border: 1px solid var(--border); }
    .muted { color: var(--muted); }
    footer { padding: 24px 32px; color: var(--muted); font-size: 12px; text-align: center; }
    a { color: var(--accent); text-decoration: none; }
    a:hover { text-decoration: underline; }
    .empty { background: var(--panel); border: 1px dashed var(--border); border-radius: 12px; padding: 24px; color: var(--muted); text-align: center; }
    code { background: #0e1729; padding: 2px 6px; border-radius: 6px; border: 1px solid var(--border); }
    .btn {
      background: linear-gradient(180deg, #2a4d8f, #1c3971);
      color: #e7efff;
      border: 1px solid #3a5da5;
      padding: 8px 14px;
      border-radius: 10px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      transition: transform 0.05s ease, filter 0.15s ease;
    }
    .btn:hover { filter: brightness(1.1); }
    .btn:active { transform: translateY(1px); }
    .btn:disabled { opacity: 0.6; cursor: progress; }
  </style>
</head>
<body>
  <header>
    <h1>${escapedTitle}</h1>
    <span class="sub" id="updated">Loading…</span>
    <span class="spacer" style="flex:1"></span>
    <button id="download-yard-check" class="btn">Download Yard Check (.xlsx)</button>
  </header>
  <main>
    <div id="yard-check-status" class="muted" style="margin-bottom:16px; min-height: 18px;"></div>
    <div class="cards" id="summary-cards"></div>
    <section>
      <h2>Tools</h2>
      <div class="card">
        <div style="display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:wrap;">
          <div>
            <div style="font-size:15px; font-weight:600;">Yard Check Generator</div>
            <div class="muted" style="font-size:12px; max-width:640px;">
              Produces a landscape, print-ready Excel with every work order (Asset ID, Work Order ID, VIN/Serial, Make, Model, Previous Location) plus blank columns for new location and discrepancies. Use it to walk the compound and record current vehicle positions.
            </div>
          </div>
          <button id="download-yard-check-2" class="btn">Download .xlsx</button>
        </div>
      </div>
    </section>
    <section>
      <h2>History</h2>
      <div id="history-wrap"><div class="empty">Loading history…</div></div>
    </section>
    <section>
      <h2>Latest per-sheet summary</h2>
      <div id="sheets-wrap"><div class="empty">Loading latest analysis…</div></div>
    </section>
  </main>
  <footer>
    Ingest flow: email <code>etic@2t3.app</code> → R2 <code>eticanalysis</code> → this dashboard.
  </footer>
  <script>
    const fmt = (n) => typeof n === "number" ? n.toLocaleString() : "—";
    const deltaClass = (n) => n === null || n === undefined ? "flat" : n > 0 ? "up" : n < 0 ? "down" : "flat";
    const deltaPrefix = (n) => n === null || n === undefined ? "" : n > 0 ? "+" : "";

    async function loadAll() {
      const [historyRes, latestRes] = await Promise.allSettled([
        fetch("/api/history").then((r) => r.ok ? r.json() : null),
        fetch("/api/latest").then((r) => r.ok ? r.json() : null),
      ]);
      const history = historyRes.status === "fulfilled" ? historyRes.value : null;
      const latest = latestRes.status === "fulfilled" ? latestRes.value : null;
      renderSummary(latest, history);
      renderHistory(history);
      renderSheets(latest);
    }

    function renderSummary(latest, history) {
      const updatedEl = document.getElementById("updated");
      if (!latest) {
        updatedEl.textContent = "No analysis yet — waiting for first inbound email.";
      } else {
        updatedEl.textContent = "Latest: " + new Date(latest.receivedAtIso).toLocaleString() + " · " + latest.workbookFileName;
      }

      const cards = document.getElementById("summary-cards");
      const lastEntry = history && history.entries && history.entries.length
        ? history.entries[history.entries.length - 1]
        : null;

      const items = latest ? [
        { label: "Visible sheets", value: latest.totalVisibleSheets, delta: lastEntry?.diff?.deltaSheetsVisible ?? null },
        { label: "Hidden sheets", value: latest.totalHiddenSheets, delta: null },
        { label: "Rows (all sheets)", value: latest.totalRowsAcrossSheets, delta: lastEntry?.diff?.deltaTotalRows ?? null },
        { label: "MEL mentions", value: latest.melMentionsTotal, delta: lastEntry?.diff?.deltaMelMentionsTotal ?? null },
      ] : [];

      cards.innerHTML = items.length === 0
        ? '<div class="empty">No metrics yet.</div>'
        : items.map((item) => {
            const deltaHtml = item.delta === null || item.delta === undefined
              ? '<span class="muted">no prior snapshot</span>'
              : '<span class="' + deltaClass(item.delta) + '">Δ ' + deltaPrefix(item.delta) + fmt(item.delta) + ' vs prior day</span>';
            return '<div class="card"><h3>' + item.label + '</h3><div class="v">' + fmt(item.value) + '</div><div class="delta">' + deltaHtml + '</div></div>';
          }).join("");
    }

    function renderHistory(history) {
      const wrap = document.getElementById("history-wrap");
      if (!history || !history.entries || history.entries.length === 0) {
        wrap.innerHTML = '<div class="empty">No history yet.</div>';
        return;
      }
      const rows = [...history.entries].sort((a, b) => b.dateKey.localeCompare(a.dateKey)).map((e) => {
        const d = e.diff || {};
        const cls = (v) => deltaClass(v ?? null);
        const pref = (v) => deltaPrefix(v ?? null);
        const row = [
          '<tr>',
          '<td>' + e.dateKey + '</td>',
          '<td>' + e.workbookFileName + '</td>',
          '<td>' + fmt(e.totalVisibleSheets) + '</td>',
          '<td>' + fmt(e.totalHiddenSheets) + '</td>',
          '<td>' + fmt(e.totalRowsAcrossSheets) + ' <span class="' + cls(d.deltaTotalRows) + '">' + pref(d.deltaTotalRows) + fmt(d.deltaTotalRows) + '</span></td>',
          '<td>' + fmt(e.melMentionsTotal) + ' <span class="' + cls(d.deltaMelMentionsTotal) + '">' + pref(d.deltaMelMentionsTotal) + fmt(d.deltaMelMentionsTotal) + '</span></td>',
          '<td><span class="pill">' + (d.previousDateKey || "initial") + '</span></td>',
          '</tr>'
        ].join("");
        return row;
      }).join("");
      wrap.innerHTML = [
        '<table>',
        '<thead><tr>',
          '<th>Date</th>',
          '<th>Workbook</th>',
          '<th>Visible</th>',
          '<th>Hidden</th>',
          '<th>Rows (Δ)</th>',
          '<th>MEL (Δ)</th>',
          '<th>Compared to</th>',
        '</tr></thead>',
        '<tbody>', rows, '</tbody>',
        '</table>'
      ].join("");
    }

    function renderSheets(latest) {
      const wrap = document.getElementById("sheets-wrap");
      if (!latest || !Array.isArray(latest.sheetSummaries) || latest.sheetSummaries.length === 0) {
        wrap.innerHTML = '<div class="empty">No sheet data yet.</div>';
        return;
      }
      const rows = latest.sheetSummaries.map((s) => [
        '<tr>',
        '<td>' + s.name + '</td>',
        '<td>' + s.state + '</td>',
        '<td>' + fmt(s.rowCount) + '</td>',
        '<td>' + fmt(s.columnCountEstimate) + '</td>',
        '<td>' + fmt(s.nonEmptyCellCountEstimate) + '</td>',
        '<td>' + fmt(s.melMentions) + '</td>',
        '<td class="muted">' + (s.sampleHeaders && s.sampleHeaders.length ? s.sampleHeaders.join(" | ") : "—") + '</td>',
        '</tr>'
      ].join("")).join("");
      wrap.innerHTML = [
        '<table>',
        '<thead><tr>',
          '<th>Sheet</th>',
          '<th>State</th>',
          '<th>Rows</th>',
          '<th>Cols</th>',
          '<th>Cells</th>',
          '<th>MEL</th>',
          '<th>Sample headers</th>',
        '</tr></thead>',
        '<tbody>', rows, '</tbody>',
        '</table>'
      ].join("");
    }

    async function downloadYardCheck(btn) {
      const statusEl = document.getElementById("yard-check-status");
      const prevLabel = btn.textContent;
      btn.disabled = true;
      btn.textContent = "Preparing…";
      statusEl.textContent = "Generating yard check from latest workbook…";
      try {
        const res = await fetch("/api/yard-check.xlsx");
        if (!res.ok) {
          let message = "Failed to generate (" + res.status + ")";
          try {
            const body = await res.json();
            if (body && body.error) message = body.error;
          } catch (_) {}
          statusEl.textContent = message;
          return;
        }
        const disposition = res.headers.get("Content-Disposition") || "";
        const nameMatch = /filename=\"?([^\";]+)\"?/i.exec(disposition);
        const fileName = nameMatch ? nameMatch[1] : "yard-check.xlsx";
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        const rowCount = res.headers.get("X-Total-Rows") || "0";
        statusEl.textContent = "Downloaded " + fileName + " (" + rowCount + " rows).";
      } catch (err) {
        statusEl.textContent = "Error: " + (err && err.message ? err.message : String(err));
      } finally {
        btn.disabled = false;
        btn.textContent = prevLabel;
      }
    }

    document.addEventListener("click", (ev) => {
      const target = ev.target;
      if (target && target instanceof HTMLElement && (target.id === "download-yard-check" || target.id === "download-yard-check-2")) {
        ev.preventDefault();
        downloadYardCheck(target);
      }
    });

    loadAll().catch((err) => {
      document.getElementById("updated").textContent = "Failed to load: " + (err && err.message ? err.message : String(err));
    });
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
  pickWorkbookAttachment,
  readCellText,
  sanitizeFileName,
  upsertHistoryEntry,
};

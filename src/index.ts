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
      return handleYardCheckDownload(env, request);
    }

    if (url.pathname === "/api/yard-check-meta") {
      return handleYardCheckMeta(env, request);
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

function parseYardCheckDateParam(request: Request): string | null {
  const url = new URL(request.url);
  const d = url.searchParams.get("date");
  return d && /^\d{4}-\d{2}-\d{2}$/.test(d) ? d : null;
}

async function resolveYardCheckWorkbook(
  env: Env,
  dateKey: string | null,
): Promise<
  | { ok: true; workbookKey: string; sourceFileName: string; sourceDateKey: string }
  | { ok: false; error: string }
> {
  const history = await loadHistory(env);
  const latest = await readJson<AnalysisResult>(env, "analyses/latest.json");

  if (dateKey) {
    const entry = history.entries.find((e) => e.dateKey === dateKey);
    if (entry) {
      return {
        ok: true,
        workbookKey: entry.workbookKey,
        sourceFileName: entry.workbookFileName,
        sourceDateKey: entry.dateKey,
      };
    }
    return {
      ok: false,
      error: `No ETIC snapshot for ${dateKey}. Pick another date or ingest that workbook.`,
    };
  }

  const latestEntry = history.entries.length
    ? [...history.entries].sort((a, b) => a.dateKey.localeCompare(b.dateKey)).pop() ?? null
    : null;
  const workbookKey = latestEntry?.workbookKey ?? null;
  const sourceFileName = latestEntry?.workbookFileName ?? latest?.workbookFileName ?? "vehicle-etic.xlsx";
  const sourceDateKey = latestEntry?.dateKey ?? latest?.dateKey ?? "latest";

  if (!workbookKey) {
    return { ok: false, error: "No source workbook found yet. Send the ETIC email to ingest first." };
  }

  return { ok: true, workbookKey, sourceFileName, sourceDateKey };
}

async function handleYardCheckDownload(env: Env, request: Request): Promise<Response> {
  const paramDate = parseYardCheckDateParam(request);
  const resolved = await resolveYardCheckWorkbook(env, paramDate);
  if (!resolved.ok) {
    return Response.json({ ok: false, error: resolved.error }, { status: 404 });
  }
  const { workbookKey, sourceFileName, sourceDateKey } = resolved;

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
};

type YardCheckMetaErr = { ok: false; error: string };

async function handleYardCheckMeta(env: Env, request: Request): Promise<Response> {
  const meta = await buildYardCheckMeta(env, request);
  const status = meta.ok ? 200 : 404;
  return Response.json(meta, { status, headers: cacheHeaders() });
}

async function buildYardCheckMeta(env: Env, request: Request): Promise<YardCheckMetaOk | YardCheckMetaErr> {
  const paramDate = parseYardCheckDateParam(request);
  const resolved = await resolveYardCheckWorkbook(env, paramDate);
  if (!resolved.ok) {
    return { ok: false, error: resolved.error };
  }
  const { workbookKey, sourceFileName, sourceDateKey } = resolved;

  const object = await env.ETIC_BUCKET.get(workbookKey);
  if (!object) {
    return { ok: false, error: `Source workbook not found in R2: ${workbookKey}` };
  }

  return {
    ok: true,
    sourceDateKey,
    sourceFileName,
    workbookKey,
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
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&display=swap" rel="stylesheet" />
  <style>
    :root {
      color-scheme: dark;
      --bg0: #070b14;
      --bg1: #0d1526;
      --surface: rgba(17, 26, 46, 0.72);
      --surface-solid: #121c32;
      --border: rgba(106, 169, 255, 0.14);
      --border-strong: rgba(106, 169, 255, 0.28);
      --text: #eef3fb;
      --muted: #8b9ab5;
      --accent: #6aa9ff;
      --accent-dim: #4d87d9;
      --glow: rgba(106, 169, 255, 0.12);
      --radius: 16px;
      --font: "DM Sans", ui-sans-serif, system-ui, -apple-system, sans-serif;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: var(--font);
      color: var(--text);
      background:
        radial-gradient(ellipse 900px 500px at 15% -20%, var(--glow), transparent 55%),
        radial-gradient(ellipse 700px 400px at 95% 10%, rgba(139, 92, 246, 0.08), transparent 50%),
        linear-gradient(165deg, var(--bg1) 0%, var(--bg0) 45%, #050810 100%);
      background-attachment: fixed;
    }
    .app {
      max-width: 1040px;
      margin: 0 auto;
      padding: 28px 22px 56px;
    }
    .top {
      display: flex;
      flex-wrap: wrap;
      align-items: flex-end;
      justify-content: space-between;
      gap: 20px;
      margin-bottom: 28px;
    }
    .brand h1 {
      margin: 0 0 6px;
      font-size: clamp(1.55rem, 3vw, 1.85rem);
      font-weight: 700;
      letter-spacing: -0.03em;
      line-height: 1.15;
    }
    .brand p {
      margin: 0;
      color: var(--muted);
      font-size: 0.92rem;
      max-width: 42ch;
      line-height: 1.45;
    }
    .picker-wrap {
      display: flex;
      flex-direction: column;
      gap: 6px;
      min-width: min(100%, 280px);
    }
    .picker-wrap label {
      font-size: 0.72rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.11em;
      color: var(--muted);
    }
    select#etic-date {
      appearance: none;
      width: 100%;
      padding: 14px 44px 14px 16px;
      font-family: var(--font);
      font-size: 1rem;
      font-weight: 600;
      color: var(--text);
      background: var(--surface-solid) url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='20' height='20' viewBox='0 0 24 24' fill='none' stroke='%238b9ab5' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E") no-repeat right 14px center;
      border: 1px solid var(--border-strong);
      border-radius: 12px;
      cursor: pointer;
      box-shadow: 0 8px 32px rgba(0,0,0,0.25);
    }
    select#etic-date:focus {
      outline: none;
      border-color: var(--accent);
      box-shadow: 0 0 0 3px var(--glow), 0 8px 32px rgba(0,0,0,0.25);
    }
    .hero {
      position: relative;
      border-radius: var(--radius);
      padding: 28px 28px 26px;
      margin-bottom: 22px;
      background: var(--surface);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border: 1px solid var(--border);
      box-shadow: 0 24px 48px rgba(0,0,0,0.35);
      overflow: hidden;
    }
    .hero::before {
      content: "";
      position: absolute;
      inset: 0;
      background: linear-gradient(125deg, rgba(106,169,255,0.07) 0%, transparent 45%);
      pointer-events: none;
    }
    .hero-inner { position: relative; z-index: 1; }
    .hero-date {
      display: inline-flex;
      align-items: center;
      gap: 10px;
      font-size: 0.8rem;
      font-weight: 600;
      color: var(--accent);
      letter-spacing: 0.04em;
      text-transform: uppercase;
      margin-bottom: 10px;
    }
    .hero-date .dot {
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: var(--accent);
      box-shadow: 0 0 12px var(--accent);
    }
    .hero h2 {
      margin: 0 0 8px;
      font-size: clamp(1.35rem, 2.5vw, 1.65rem);
      font-weight: 700;
      letter-spacing: -0.02em;
    }
    .hero-file {
      margin: 0;
      color: var(--muted);
      font-size: 0.95rem;
    }
    .hero-file strong { color: #c5d4eb; font-weight: 500; }
    .card {
      border-radius: var(--radius);
      background: var(--surface);
      backdrop-filter: blur(12px);
      border: 1px solid var(--border);
      padding: 22px 22px 20px;
      box-shadow: 0 12px 40px rgba(0,0,0,0.22);
    }
    .card h3 {
      margin: 0 0 14px;
      font-size: 0.72rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      color: var(--muted);
    }
    .detail-rows { display: flex; flex-direction: column; gap: 12px; }
    .detail-row {
      display: grid;
      grid-template-columns: 120px 1fr;
      gap: 12px;
      font-size: 0.9rem;
      align-items: start;
    }
    @media (max-width: 520px) { .detail-row { grid-template-columns: 1fr; gap: 4px; } }
    .detail-row dt { color: var(--muted); font-weight: 500; margin: 0; }
    .detail-row dd { margin: 0; color: #d2ddf0; line-height: 1.45; word-break: break-word; }
    .ingest-card { margin-bottom: 22px; }
    .yard-card { margin-bottom: 8px; }
    .yard-card .yard-lead {
      margin: 0 0 18px;
      font-size: 0.92rem;
      color: var(--muted);
      line-height: 1.5;
      max-width: 52ch;
    }
    .yard-export-btn {
      display: flex;
      align-items: center;
      gap: 18px;
      width: 100%;
      margin: 0;
      padding: 20px 22px;
      border-radius: 14px;
      border: 1px solid var(--border-strong);
      background: linear-gradient(145deg, rgba(106, 169, 255, 0.12) 0%, rgba(7, 11, 20, 0.65) 55%);
      box-shadow: 0 8px 32px rgba(0, 0, 0, 0.28), inset 0 1px 0 rgba(255, 255, 255, 0.06);
      cursor: pointer;
      font-family: var(--font);
      text-align: left;
      transition: border-color 0.15s ease, box-shadow 0.15s ease, transform 0.12s ease;
    }
    .yard-export-btn:hover:not(:disabled) {
      border-color: rgba(106, 169, 255, 0.45);
      box-shadow: 0 12px 40px rgba(0, 0, 0, 0.32), 0 0 0 1px rgba(106, 169, 255, 0.12);
    }
    .yard-export-btn:active:not(:disabled) { transform: scale(0.992); }
    .yard-export-btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    .yard-export-btn .icon-wrap {
      flex-shrink: 0;
      width: 52px;
      height: 52px;
      border-radius: 14px;
      display: flex;
      align-items: center;
      justify-content: center;
      background: rgba(106, 169, 255, 0.18);
      border: 1px solid rgba(106, 169, 255, 0.28);
      color: var(--accent);
    }
    .yard-export-btn .icon-wrap svg { width: 26px; height: 26px; }
    .yard-export-btn .copy { flex: 1; min-width: 0; }
    .yard-export-btn .title {
      display: block;
      font-size: 1.08rem;
      font-weight: 700;
      letter-spacing: -0.02em;
      color: var(--text);
      margin-bottom: 4px;
    }
    .yard-export-btn .sub {
      display: block;
      font-size: 0.82rem;
      color: var(--muted);
      font-weight: 500;
    }
    .yard-export-btn .chev {
      flex-shrink: 0;
      color: var(--accent);
      opacity: 0.85;
    }
    .yard-export-btn .chev svg { width: 22px; height: 22px; display: block; }
    .yard-status-wrap { margin-top: 14px; }
    .status {
      font-size: 0.88rem;
      color: var(--muted);
      min-height: 22px;
      margin: 0;
    }
    .status.err { color: #ff9e9e; }
    .status.ok { color: #8fd4a8; }
    .empty-state {
      text-align: center;
      padding: 48px 24px;
      color: var(--muted);
      font-size: 0.95rem;
      line-height: 1.55;
    }
    .empty-state strong { color: var(--text); }
    .hidden { display: none !important; }
  </style>
</head>
<body>
  <div class="app">
    <header class="top">
      <div class="brand">
        <h1>${escapedTitle}</h1>
        <p>Pick a report date, then export the yard check for that file.</p>
      </div>
      <div class="picker-wrap">
        <label for="etic-date">ETIC report date</label>
        <select id="etic-date" aria-label="Select ETIC report date"></select>
      </div>
    </header>

    <div id="view-empty" class="empty-state hidden">
      <p><strong>No ETIC files yet.</strong><br />Email the Vehicle ETIC workbook to your ingest address to get dates here.</p>
    </div>

    <div id="view-main" class="hidden">
      <section class="hero">
        <div class="hero-inner">
          <div class="hero-date"><span class="dot" aria-hidden="true"></span> Selected snapshot</div>
          <h2 id="hero-title">—</h2>
          <p class="hero-file" id="hero-file"></p>
        </div>
      </section>

      <div class="card ingest-card">
        <h3>This snapshot</h3>
        <dl class="detail-rows" id="ingest-details"></dl>
      </div>

      <div class="card yard-card">
        <h3>Yard check</h3>
        <p class="yard-lead">Work orders rolled up for a walkaround: asset, shops, VIN, locations, plus blank columns for what you find in the yard.</p>
        <button type="button" class="yard-export-btn" id="btn-yard-check" aria-describedby="yard-status">
          <span class="icon-wrap" aria-hidden="true">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="1.75"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
          </span>
          <span class="copy">
            <span class="title">Export for compound walkaround</span>
            <span class="sub">Excel workbook · landscape · ready to print</span>
          </span>
          <span class="chev" aria-hidden="true">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" /></svg>
          </span>
        </button>
        <div class="yard-status-wrap">
          <p class="status" id="yard-status" role="status"></p>
        </div>
      </div>
    </div>
  </div>
  <script>
    function esc(s) {
      if (s == null || s === undefined) return "";
      return String(s)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    }

    function sortDesc(entries) {
      return [...entries].sort((a, b) => b.dateKey.localeCompare(a.dateKey));
    }

    function readHashDate() {
      const h = (location.hash || "").replace(/^#/, "");
      return /^\\d{4}-\\d{2}-\\d{2}$/.test(h) ? h : null;
    }

    let historyEntries = [];
    let selectedDate = null;

    function setHash(dateKey) {
      if (dateKey) location.hash = "#" + dateKey;
      else location.hash = "";
    }

    async function loadHistory() {
      const res = await fetch("/api/history");
      if (!res.ok) throw new Error("Could not load history");
      const data = await res.json();
      return Array.isArray(data.entries) ? data.entries : [];
    }

    function fillSelect(entries) {
      const sel = document.getElementById("etic-date");
      sel.innerHTML = "";
      const sorted = sortDesc(entries);
      for (const e of sorted) {
        const opt = document.createElement("option");
        opt.value = e.dateKey;
        opt.textContent = e.dateKey;
        sel.appendChild(opt);
      }
    }

    async function loadAnalysis(dateKey) {
      const res = await fetch("/api/analysis/" + encodeURIComponent(dateKey));
      if (!res.ok) throw new Error("No analysis for " + dateKey);
      return res.json();
    }

    function renderDetails(analysis) {
      document.getElementById("hero-title").textContent = analysis.dateKey;
      document.getElementById("hero-file").innerHTML =
        "File <strong>" + esc(analysis.workbookFileName) + "</strong>";

      const ing = document.getElementById("ingest-details");
      const recv = analysis.receivedAtIso
        ? new Date(analysis.receivedAtIso).toLocaleString(undefined, {
            dateStyle: "medium",
            timeStyle: "short",
          })
        : "—";
      ing.innerHTML = [
        "<div class='detail-row'><dt>Received</dt><dd>" + esc(recv) + "</dd></div>",
        "<div class='detail-row'><dt>From</dt><dd>" + esc(analysis.from) + "</dd></div>",
        "<div class='detail-row'><dt>To</dt><dd>" + esc(analysis.to) + "</dd></div>",
        "<div class='detail-row'><dt>Workbook size</dt><dd>" +
          (typeof analysis.workbookBytes === "number"
            ? (analysis.workbookBytes / 1024 / 1024).toFixed(2) + " MB"
            : "—") +
          "</dd></div>",
      ].join("");
    }

    async function selectDate(dateKey, pushHash) {
      selectedDate = dateKey;
      if (pushHash && dateKey) setHash(dateKey);
      document.getElementById("yard-status").textContent = "";
      document.getElementById("yard-status").className = "status";

      if (!dateKey) return;

      try {
        const analysis = await loadAnalysis(dateKey);
        renderDetails(analysis);
      } catch (e) {
        document.getElementById("ingest-details").innerHTML =
          "<div class='detail-row'><dt>Error</dt><dd>" + esc(e.message || String(e)) + "</dd></div>";
      }
    }

    async function downloadYardCheck() {
      const btn = document.getElementById("btn-yard-check");
      const st = document.getElementById("yard-status");
      if (!selectedDate) return;
      btn.disabled = true;
      st.className = "status";
      st.textContent = "Building your file…";
      try {
        const url = "/api/yard-check.xlsx?date=" + encodeURIComponent(selectedDate);
        const res = await fetch(url);
        if (!res.ok) {
          let msg = "Could not generate file.";
          try {
            const j = await res.json();
            if (j && j.error) msg = j.error;
          } catch (_) {}
          st.className = "status err";
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
        st.className = "status ok";
        st.textContent = "Saved as " + name;
      } catch (e) {
        st.className = "status err";
        st.textContent = String(e && e.message ? e.message : e);
      } finally {
        btn.disabled = false;
      }
    }

    async function init() {
      try {
        historyEntries = await loadHistory();
      } catch (e) {
        document.getElementById("view-empty").classList.remove("hidden");
        document.getElementById("view-empty").querySelector("p").innerHTML =
          "<strong>Could not load data.</strong><br />" + esc(e.message || String(e));
        return;
      }

      if (!historyEntries.length) {
        document.getElementById("view-empty").classList.remove("hidden");
        return;
      }

      document.getElementById("view-main").classList.remove("hidden");
      fillSelect(historyEntries);

      const sorted = sortDesc(historyEntries);
      const hashDate = readHashDate();
      const start =
        hashDate && sorted.some((e) => e.dateKey === hashDate)
          ? hashDate
          : sorted[0].dateKey;
      const sel = document.getElementById("etic-date");
      sel.value = start;
      await selectDate(start, true);

      sel.addEventListener("change", async () => {
        await selectDate(sel.value, true);
      });

      window.addEventListener("hashchange", async () => {
        const d = readHashDate();
        if (d && historyEntries.some((e) => e.dateKey === d) && d !== selectedDate) {
          sel.value = d;
          await selectDate(d, false);
        }
      });

      document.getElementById("btn-yard-check").addEventListener("click", downloadYardCheck);
    }

    init();
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

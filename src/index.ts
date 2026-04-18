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
  /** When the email subject includes a report date (e.g. 15-APR-26), mirrors dateKey; else null. */
  reportDateKey: string | null;
  from: string;
  to: string;
  subject: string;
  /** Mission-capable rate from subject, e.g. 75.49 */
  mcRatePercent: number | null;
  /** Count before the report date token when subject uses "… Report: &lt;n&gt; - DD-MMM-YY"; else null. */
  belowMelCriticalCount: number | null;
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
  reportDateKey: string | null;
  receivedAtIso: string;
  workbookFileName: string;
  workbookKey: string;
  analysisKey: string;
  mcRatePercent: number | null;
  belowMelCriticalCount: number | null;
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
    reportDateKey: analysis.reportDateKey,
    receivedAtIso: analysis.receivedAtIso,
    workbookFileName: analysis.workbookFileName,
    workbookKey,
    analysisKey,
    mcRatePercent: analysis.mcRatePercent,
    belowMelCriticalCount: analysis.belowMelCriticalCount,
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

/** Parse ETIC-style email subjects for report date + headline metrics. */
function parseEticEmailSubject(subject: string): {
  reportDateKey: string | null;
  mcRatePercent: number | null;
  belowMelCriticalCount: number | null;
} {
  const s = subject.trim();
  let mcRatePercent: number | null = null;
  const mcMatch = /MC\s*Rate:\s*([\d.]+)\s*%/i.exec(s);
  if (mcMatch) {
    const n = Number.parseFloat(mcMatch[1] ?? "");
    if (Number.isFinite(n)) mcRatePercent = n;
  }

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
    const candidate = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    reportDateKey = candidate;
  }

  let belowMelCriticalCount: number | null = null;
  // Count only when separated from the date (e.g. "Report: 3 – 15-APR-26"). "Report: 15-APR-26" is date-only.
  const belowWithCount =
    /Below\s+MEL\/Critical\s+Report:\s*(\d+)\s+[-–—]\s+\d{1,2}-[A-Za-z]{3}/i.exec(s);
  if (belowWithCount) {
    const c = Number.parseInt(belowWithCount[1] ?? "", 10);
    if (Number.isFinite(c)) belowMelCriticalCount = c;
  }

  return { reportDateKey, mcRatePercent, belowMelCriticalCount };
}

/** Prefer report date from subject (for forwarded historical ETICs); else UTC day of receipt. */
function resolveAnalysisDateKey(subject: string, receivedAt: Date): string {
  const parsed = parseEticEmailSubject(subject);
  return parsed.reportDateKey ?? isoDateKey(receivedAt);
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
  const subjectParsed = parseEticEmailSubject(input.subject);
  const reportDateKey = subjectParsed.reportDateKey;

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
    reportDateKey,
    from: input.from,
    to: input.to,
    subject: input.subject,
    mcRatePercent: subjectParsed.mcRatePercent,
    belowMelCriticalCount: subjectParsed.belowMelCriticalCount,
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
    .lead {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 20px 22px;
      margin-bottom: 22px;
    }
    .lead h2 { margin: 0 0 8px; font-size: 18px; font-weight: 650; letter-spacing: 0.02em; }
    .lead .meta { font-size: 12px; color: var(--muted); margin-bottom: 12px; }
    .lead .subject { font-size: 13px; line-height: 1.45; color: #c8d4ea; max-height: 4.4em; overflow: hidden; }
    .lead-metrics { display: flex; flex-wrap: wrap; gap: 20px 28px; margin-top: 14px; }
    .lead-metrics .big { font-size: 28px; font-weight: 700; letter-spacing: -0.02em; }
    .lead-metrics .lbl { font-size: 11px; text-transform: uppercase; color: var(--muted); letter-spacing: 0.07em; margin-bottom: 4px; }
    .chart-wrap {
      background: var(--panel);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 16px 18px 8px;
      margin-bottom: 22px;
    }
    .chart-wrap h2 { margin: 0 0 12px; font-size: 15px; }
    .chart-wrap svg { width: 100%; height: auto; display: block; }
    .chart-note { font-size: 12px; color: var(--muted); margin-top: 8px; }
    .cards {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
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
    .card .v { font-size: 22px; font-weight: 600; }
    .card .delta { margin-top: 6px; font-size: 12px; color: var(--muted); }
    .up { color: var(--up); }
    .down { color: var(--down); }
    .flat { color: var(--flat); }
    section { margin-bottom: 28px; }
    h2.sec { font-size: 16px; margin: 0 0 10px; letter-spacing: 0.04em; }
    .sec-hint { font-size: 12px; color: var(--muted); font-weight: 400; margin-left: 8px; }
    table { width: 100%; border-collapse: collapse; background: var(--panel); border: 1px solid var(--border); border-radius: 12px; overflow: hidden; }
    th, td { padding: 10px 12px; text-align: left; font-size: 13px; border-bottom: 1px solid var(--border); vertical-align: top; }
    th { background: var(--panel-alt); color: var(--muted); font-weight: 600; }
    tr:last-child td { border-bottom: 0; }
    td.num { text-align: right; font-variant-numeric: tabular-nums; }
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
    details.tools { background: var(--panel); border: 1px solid var(--border); border-radius: 12px; padding: 0 16px; }
    details.tools summary { cursor: pointer; padding: 14px 0; font-weight: 600; list-style: none; }
    details.tools summary::-webkit-details-marker { display: none; }
    details.tools .inner { padding-bottom: 16px; border-top: 1px solid var(--border); padding-top: 14px; }
  </style>
</head>
<body>
  <header>
    <h1>${escapedTitle}</h1>
    <span class="sub" id="updated">Loading…</span>
    <span class="spacer" style="flex:1"></span>
    <button id="download-yard-check" class="btn">Yard Check (.xlsx)</button>
  </header>
  <main>
    <div id="lead-wrap"><div class="empty">Loading overview…</div></div>
    <div class="chart-wrap" id="chart-wrap" style="display:none">
      <h2>Mission capable rate trend</h2>
      <svg id="mc-chart" viewBox="0 0 640 200" role="img" aria-label="MC rate over recent reports"></svg>
      <div class="chart-note" id="chart-note"></div>
    </div>
    <div class="cards" id="summary-cards"></div>
    <section>
      <h2 class="sec">Recent reports <span class="sec-hint" id="trend-hint"></span></h2>
      <div id="trends-wrap"><div class="empty">Loading…</div></div>
    </section>
    <section>
      <h2 class="sec">Latest workbook structure</h2>
      <div id="sheets-wrap"><div class="empty">Loading latest analysis…</div></div>
    </section>
    <section>
      <details class="tools">
        <summary>Tools — Yard Check generator</summary>
        <div class="inner">
          <div id="yard-check-status" class="muted" style="margin-bottom:12px; min-height: 18px;"></div>
          <div style="display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:wrap;">
            <div class="muted" style="font-size:13px; max-width:720px;">
              Print-ready Excel from the latest ingested workbook: work orders consolidated by asset, shops, VIN/serial, previous location, plus blank columns for new location and discrepancies.
            </div>
            <button id="download-yard-check-2" class="btn">Download .xlsx</button>
          </div>
        </div>
      </details>
    </section>
  </main>
  <footer>
    Ingest: forward saved ETIC emails (subject line preserved) to <code>etic@2t3.app</code> → R2 <code>eticanalysis</code> → this dashboard. Report date in the subject is used as the snapshot day.
  </footer>
  <script>
    const OVERVIEW_LIMIT = 60;
    const fmt = (n) => typeof n === "number" ? n.toLocaleString() : "—";
    const fmtPct = (n) => typeof n === "number" ? n.toFixed(2) + "%" : "—";
    const deltaClass = (n) => n === null || n === undefined ? "flat" : n > 0 ? "up" : n < 0 ? "down" : "flat";
    const deltaPrefix = (n) => n === null || n === undefined ? "" : n > 0 ? "+" : "";

    function sortHistoryDesc(entries) {
      return [...entries].sort((a, b) => b.dateKey.localeCompare(a.dateKey));
    }

    function escapeAttr(s) {
      return String(s || "")
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    }

    function renderMcChart(rowsOldestFirst) {
      const wrap = document.getElementById("chart-wrap");
      const svg = document.getElementById("mc-chart");
      const note = document.getElementById("chart-note");
      const withMc = rowsOldestFirst.filter((r) => typeof r.mcRatePercent === "number");
      if (withMc.length < 2) {
        wrap.style.display = "none";
        return;
      }
      wrap.style.display = "block";
      const W = 640, H = 200, padL = 44, padR = 12, padT = 16, padB = 28;
      const innerW = W - padL - padR, innerH = H - padT - padB;
      const vals = withMc.map((r) => r.mcRatePercent);
      let lo = Math.min.apply(null, vals);
      let hi = Math.max.apply(null, vals);
      if (hi - lo < 1) { lo -= 1; hi += 1; }
      lo = Math.max(0, lo - 1);
      hi = Math.min(100, hi + 1);
      const n = withMc.length;
      const pts = withMc.map((r, i) => {
        const x = padL + (n === 1 ? innerW / 2 : (i / (n - 1)) * innerW);
        const t = (r.mcRatePercent - lo) / (hi - lo || 1);
        const y = padT + innerH * (1 - t);
        return { x, y, r };
      });
      const d = pts.map((p, i) => (i === 0 ? "M" : "L") + p.x.toFixed(1) + " " + p.y.toFixed(1)).join(" ");
      const circles = pts.map((p) =>
        '<circle cx="' + p.x.toFixed(1) + '" cy="' + p.y.toFixed(1) + '" r="3.5" fill="#6aa9ff" stroke="#0b1220" stroke-width="1"><title>' +
        escapeAttr(p.r.dateKey + ": " + fmtPct(p.r.mcRatePercent)) + '</title></circle>'
      ).join("");
      const yTicks = [lo, (lo + hi) / 2, hi];
      const grid = yTicks.map((v) => {
        const t = (v - lo) / (hi - lo || 1);
        const y = padT + innerH * (1 - t);
        return '<line x1="' + padL + '" y1="' + y.toFixed(1) + '" x2="' + (W - padR) + '" y2="' + y.toFixed(1) + '" stroke="#1f2a44" stroke-dasharray="44" opacity="0.7"/>';
      }).join("");
      const yLabels = yTicks.map((v) => {
        const t = (v - lo) / (hi - lo || 1);
        const y = padT + innerH * (1 - t);
        return '<text x="' + (padL - 6) + '" y="' + (y + 4) + '" text-anchor="end" fill="#96a2b8" font-size="10">' + v.toFixed(0) + '%</text>';
      }).join("");
      svg.innerHTML =
        grid +
        '<path d="' + d + '" fill="none" stroke="#6aa9ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>' +
        circles +
        yLabels;
      note.textContent = "Oldest → newest across " + withMc.length + " reports with an MC rate in the email subject.";
    }

    async function loadAll() {
      const [historyRes, latestRes] = await Promise.allSettled([
        fetch("/api/history").then((r) => r.ok ? r.json() : null),
        fetch("/api/latest").then((r) => r.ok ? r.json() : null),
      ]);
      const history = historyRes.status === "fulfilled" ? historyRes.value : null;
      const latest = latestRes.status === "fulfilled" ? latestRes.value : null;
      renderLead(latest);
      renderSummary(latest, history);
      renderTrends(history);
      renderSheets(latest);
    }

    function renderLead(latest) {
      const wrap = document.getElementById("lead-wrap");
      const updatedEl = document.getElementById("updated");
      if (!latest) {
        updatedEl.textContent = "No analysis yet — forward an ETIC email with the workbook attached.";
        wrap.innerHTML = '<div class="empty">No ETIC has been ingested yet. Forward a saved email (keep the subject line) to <code>etic@2t3.app</code>.</div>';
        return;
      }
      const reportLabel = latest.reportDateKey || latest.dateKey;
      const subj = latest.subject || "";
      const shortSub = subj.length > 220 ? subj.slice(0, 217) + "…" : subj;
      updatedEl.textContent = "Last ingest: " + new Date(latest.receivedAtIso).toLocaleString() + " · " + latest.workbookFileName;
      const mc = latest.mcRatePercent;
      const below = latest.belowMelCriticalCount;
      wrap.innerHTML = [
        '<div class="lead">',
        '<h2>ETIC snapshot — ' + reportLabel + '</h2>',
        '<div class="meta">Report date from subject when present · otherwise UTC day of receipt</div>',
        '<div class="subject" title="' + escapeAttr(subj) + '">' + escapeAttr(shortSub) + '</div>',
        '<div class="lead-metrics">',
        '<div><div class="lbl">MC rate (from subject)</div><div class="big">' + (mc != null ? fmtPct(mc) : "—") + '</div></div>',
        '<div><div class="lbl">Below MEL / critical (from subject)</div><div class="big">' + (below != null ? fmt(below) : "—") + '</div></div>',
        '</div></div>'
      ].join("");
    }

    function renderSummary(latest, history) {
      const cards = document.getElementById("summary-cards");
      const sorted = history && history.entries ? sortHistoryDesc(history.entries) : [];
      const lastEntry = sorted.length ? sorted[0] : null;
      const prevEntry = sorted.length > 1 ? sorted[1] : null;

      let mcDelta = null;
      if (lastEntry && prevEntry && typeof lastEntry.mcRatePercent === "number" && typeof prevEntry.mcRatePercent === "number") {
        mcDelta = lastEntry.mcRatePercent - prevEntry.mcRatePercent;
      }
      let belowDelta = null;
      if (lastEntry && prevEntry && typeof lastEntry.belowMelCriticalCount === "number" && typeof prevEntry.belowMelCriticalCount === "number") {
        belowDelta = lastEntry.belowMelCriticalCount - prevEntry.belowMelCriticalCount;
      }

      const items = latest ? [
        { label: "MC rate (latest)", value: latest.mcRatePercent != null ? fmtPct(latest.mcRatePercent) : "—", delta: mcDelta, isPctDelta: true },
        { label: "Below MEL / critical", value: latest.belowMelCriticalCount != null ? fmt(latest.belowMelCriticalCount) : "—", delta: belowDelta, isPctDelta: false },
        { label: "Rows (all sheets)", value: latest.totalRowsAcrossSheets, delta: lastEntry?.diff?.deltaTotalRows ?? null, isPctDelta: false },
        { label: "MEL cell mentions", value: latest.melMentionsTotal, delta: lastEntry?.diff?.deltaMelMentionsTotal ?? null, isPctDelta: false },
      ] : [];

      cards.innerHTML = items.length === 0
        ? '<div class="empty">No metrics yet.</div>'
        : items.map((item) => {
            const deltaHtml = item.delta === null || item.delta === undefined
              ? '<span class="muted">no prior report</span>'
              : '<span class="' + deltaClass(item.isPctDelta ? -item.delta : item.delta) + '">Δ ' + deltaPrefix(item.delta) + (item.isPctDelta ? item.delta.toFixed(2) + " pp" : fmt(item.delta)) + ' vs prior</span>';
            const valStr = typeof item.value === "number" ? fmt(item.value) : item.value;
            return '<div class="card"><h3>' + item.label + '</h3><div class="v">' + valStr + '</div><div class="delta">' + deltaHtml + '</div></div>';
          }).join("");
    }

    function renderTrends(history) {
      const hint = document.getElementById("trend-hint");
      const wrap = document.getElementById("trends-wrap");
      if (!history || !history.entries || history.entries.length === 0) {
        hint.textContent = "";
        wrap.innerHTML = '<div class="empty">No history yet. Forward past ETIC emails (up to ' + OVERVIEW_LIMIT + ' days) to build this table.</div>';
        document.getElementById("chart-wrap").style.display = "none";
        return;
      }
      const desc = sortHistoryDesc(history.entries);
      const slice = desc.slice(0, OVERVIEW_LIMIT);
      hint.textContent = "(newest first, max " + OVERVIEW_LIMIT + " · " + slice.length + " loaded)";

      const oldestFirst = slice.slice().reverse();
      renderMcChart(oldestFirst);

      const rows = slice.map((e, idx) => {
        const next = slice[idx + 1];
        let dMc = null, dBelow = null;
        if (next && typeof e.mcRatePercent === "number" && typeof next.mcRatePercent === "number") dMc = e.mcRatePercent - next.mcRatePercent;
        if (next && typeof e.belowMelCriticalCount === "number" && typeof next.belowMelCriticalCount === "number") {
          dBelow = e.belowMelCriticalCount - next.belowMelCriticalCount;
        }
        const d = e.diff || {};
        const cls = (v) => deltaClass(v ?? null);
        const pref = (v) => deltaPrefix(v ?? null);
        const mcCell = typeof e.mcRatePercent === "number"
          ? fmtPct(e.mcRatePercent) + (dMc != null ? ' <span class="' + cls(-dMc) + '">' + pref(dMc) + dMc.toFixed(2) + '</span>' : "")
          : '<span class="muted">—</span>';
        const belowCell = typeof e.belowMelCriticalCount === "number"
          ? fmt(e.belowMelCriticalCount) + (dBelow != null ? ' <span class="' + cls(dBelow) + '">' + pref(dBelow) + fmt(dBelow) + '</span>' : "")
          : '<span class="muted">—</span>';
        return [
          '<tr>',
          '<td>' + e.dateKey + '</td>',
          '<td class="num">' + mcCell + '</td>',
          '<td class="num">' + belowCell + '</td>',
          '<td class="num">' + fmt(e.totalRowsAcrossSheets) + ' <span class="' + cls(d.deltaTotalRows) + '">' + pref(d.deltaTotalRows) + fmt(d.deltaTotalRows) + '</span></td>',
          '<td class="num">' + fmt(e.melMentionsTotal) + ' <span class="' + cls(d.deltaMelMentionsTotal) + '">' + pref(d.deltaMelMentionsTotal) + fmt(d.deltaMelMentionsTotal) + '</span></td>',
          '<td class="muted" style="font-size:12px; max-width:280px;">' + escapeAttr(e.workbookFileName) + '</td>',
          '</tr>'
        ].join("");
      }).join("");

      wrap.innerHTML = [
        '<table>',
        '<thead><tr>',
          '<th>Report day</th>',
          '<th class="num">MC rate</th>',
          '<th class="num">Below MEL / crit.</th>',
          '<th class="num">Rows (Δ)</th>',
          '<th class="num">MEL mentions (Δ)</th>',
          '<th>File</th>',
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
        const assets = res.headers.get("X-Total-Assets") || "0";
        const wo = res.headers.get("X-Total-Work-Orders") || "0";
        statusEl.textContent = "Downloaded " + fileName + " (" + assets + " assets across " + wo + " work orders).";
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
  parseEticEmailSubject,
  parseMaxAttachmentBytes,
  pickWorkbookAttachment,
  readCellText,
  resolveAnalysisDateKey,
  sanitizeFileName,
  upsertHistoryEntry,
};

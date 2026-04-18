import PostalMime from "postal-mime";
import type { Attachment } from "postal-mime";
import ExcelJS from "exceljs";

type AnalysisResult = {
  workbookFileName: string;
  receivedAtIso: string;
  from: string;
  to: string;
  subject: string;
  workbookBytes: number;
  sheetSummaries: SheetSummary[];
  totalVisibleSheets: number;
  totalHiddenSheets: number;
  totalRowsAcrossSheets: number;
  melMentionsBySheet: Record<string, number>;
};

type SheetSummary = {
  name: string;
  state: "visible" | "hidden" | "veryHidden";
  rowCount: number;
  columnCountEstimate: number;
  sampleHeaders: string[];
  nonEmptyCellCountEstimate: number;
  melMentions: number;
};

interface Env {
  ETIC_BUCKET: R2Bucket;
  REPORT_EMAIL: SendEmail;
  REPORT_TO: string;
  REPORT_FROM: string;
  EXPECTED_ATTACHMENT_NAME?: string;
  ALLOWED_SENDERS?: string;
  MAX_ATTACHMENT_BYTES?: string;
}

const DEFAULT_MAX_ATTACHMENT_BYTES = 30 * 1024 * 1024;

export default {
  async email(message: ForwardableEmailMessage, env: Env, ctx: ExecutionContext): Promise<void> {
    const replyRecipient = resolveReportRecipient(message.from, env.REPORT_TO);

    try {
      if (!isAuthorizedSender(message.from, env)) {
        message.setReject("Unauthorized sender");
        return;
      }

      const parsed = await PostalMime.parse(message.raw, {
        attachmentEncoding: "arraybuffer",
      });

      const workbookAttachment = pickWorkbookAttachment(parsed.attachments, env.EXPECTED_ATTACHMENT_NAME);
      if (!workbookAttachment) {
        await sendTextEmail(
          env,
          `[Vehicle ETIC] No workbook attachment`,
          [
            "An inbound email was received, but no .xlsx attachment was found.",
            `From: ${message.from}`,
            `To: ${message.to}`,
            `Subject: ${parsed.subject ?? "(no subject)"}`,
          ].join("\n"),
          replyRecipient,
        );
        return;
      }

      const workbookBytes = normalizeAttachmentBinary(workbookAttachment.content);
      if (workbookBytes.byteLength > parseMaxAttachmentBytes(env.MAX_ATTACHMENT_BYTES)) {
        message.setReject("Attachment too large");
        return;
      }

      const now = new Date();
      const ts = compactTimestamp(now);
      const safeName = sanitizeFileName(workbookAttachment.filename ?? "vehicle-etic.xlsx");
      const rawKey = `incoming/${ts}/${safeName}`;

      await env.ETIC_BUCKET.put(rawKey, workbookBytes, {
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
        from: message.from,
        to: message.to,
        subject: parsed.subject ?? "",
      });

      const reportText = renderReportText(analysis);
      const reportJsonKey = `reports/${ts}/analysis.json`;
      const reportTxtKey = `reports/${ts}/report.txt`;

      await Promise.all([
        env.ETIC_BUCKET.put(reportJsonKey, JSON.stringify(analysis, null, 2), {
          httpMetadata: { contentType: "application/json; charset=utf-8" },
        }),
        env.ETIC_BUCKET.put(reportTxtKey, reportText, {
          httpMetadata: { contentType: "text/plain; charset=utf-8" },
        }),
        env.ETIC_BUCKET.put("reports/latest.json", JSON.stringify(analysis, null, 2), {
          httpMetadata: { contentType: "application/json; charset=utf-8" },
        }),
        env.ETIC_BUCKET.put("reports/latest.txt", reportText, {
          httpMetadata: { contentType: "text/plain; charset=utf-8" },
        }),
      ]);

      ctx.waitUntil(
        sendTextEmail(
          env,
          `[Vehicle ETIC] Daily report - ${safeName}`,
          [
            "Vehicle ETIC analysis completed successfully.",
            "",
            `Workbook key: ${rawKey}`,
            `JSON report key: ${reportJsonKey}`,
            `Text report key: ${reportTxtKey}`,
            "",
            reportText,
          ].join("\n"),
          replyRecipient,
        ),
      );
    } catch (error) {
      const messageText = error instanceof Error ? error.message : String(error);
      ctx.waitUntil(
        sendTextEmail(
          env,
          `[Vehicle ETIC] Analysis failure`,
          `The ETIC analysis worker failed.\n\nError: ${messageText}`,
          replyRecipient,
        ),
      );
      throw error;
    }
  },

  async fetch(request: Request): Promise<Response> {
    const url = new URL(request.url);
    if (url.pathname === "/healthz") {
      return Response.json({ ok: true, service: "vehicle-etic-email-automation" });
    }
    return new Response("Not Found", { status: 404 });
  },
} satisfies ExportedHandler<Env>;

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
  if (content instanceof Uint8Array) {
    return cloneToArrayBuffer(content);
  }
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

function resolveReportRecipient(sender: string, fallback: string): string {
  const normalizedSender = normalizeEmailAddress(sender);
  if (normalizedSender) return normalizedSender;
  return normalizeEmailAddress(fallback);
}

function parseMaxAttachmentBytes(raw: string | undefined): number {
  if (!raw) return DEFAULT_MAX_ATTACHMENT_BYTES;
  const value = Number.parseInt(raw, 10);
  return Number.isFinite(value) && value > 0 ? value : DEFAULT_MAX_ATTACHMENT_BYTES;
}

function compactTimestamp(date: Date): string {
  const p = (n: number) => String(n).padStart(2, "0");
  return [
    date.getUTCFullYear(),
    p(date.getUTCMonth() + 1),
    p(date.getUTCDate()),
    "-",
    p(date.getUTCHours()),
    p(date.getUTCMinutes()),
    p(date.getUTCSeconds()),
  ].join("");
}

async function analyzeWorkbook(input: {
  binary: ArrayBuffer;
  fileName: string;
  receivedAtIso: string;
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

  return {
    workbookFileName: input.fileName,
    receivedAtIso: input.receivedAtIso,
    from: input.from,
    to: input.to,
    subject: input.subject,
    workbookBytes: input.binary.byteLength,
    sheetSummaries,
    totalVisibleSheets,
    totalHiddenSheets,
    totalRowsAcrossSheets,
    melMentionsBySheet,
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
    const value = cell.text?.trim();
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
      const text = cell.text?.toLowerCase() ?? "";
      if (text.includes("mel")) count += 1;
    });
  });
  return count;
}

function renderReportText(result: AnalysisResult): string {
  const lines: string[] = [];
  lines.push("Vehicle ETIC Daily Analysis");
  lines.push("===========================");
  lines.push(`Workbook: ${result.workbookFileName}`);
  lines.push(`Received: ${result.receivedAtIso}`);
  lines.push(`Sender: ${result.from}`);
  lines.push(`Recipient: ${result.to}`);
  lines.push(`Subject: ${result.subject || "(no subject)"}`);
  lines.push(`Size: ${result.workbookBytes} bytes`);
  lines.push("");
  lines.push(`Visible sheets: ${result.totalVisibleSheets}`);
  lines.push(`Hidden/veryHidden sheets: ${result.totalHiddenSheets}`);
  lines.push(`Total non-empty rows (all sheets): ${result.totalRowsAcrossSheets}`);
  lines.push("");
  lines.push("Per-sheet summary:");
  for (const sheet of result.sheetSummaries) {
    lines.push(`- ${sheet.name}`);
    lines.push(`  state: ${sheet.state}`);
    lines.push(`  rows: ${sheet.rowCount}`);
    lines.push(`  columns (max non-empty row width): ${sheet.columnCountEstimate}`);
    lines.push(`  non-empty cells: ${sheet.nonEmptyCellCountEstimate}`);
    lines.push(`  MEL mentions: ${sheet.melMentions}`);
    lines.push(
      `  sample headers: ${sheet.sampleHeaders.length ? sheet.sampleHeaders.join(" | ") : "(none detected)"}`,
    );
  }
  return lines.join("\n");
}

async function sendTextEmail(env: Env, subject: string, text: string, to: string): Promise<void> {
  if (!env.REPORT_FROM || !to) {
    console.error(
      JSON.stringify({
        level: "warn",
        message: "Report email skipped due to missing REPORT_FROM or recipient",
        subject,
      }),
    );
    return;
  }

  await env.REPORT_EMAIL.send({
    from: env.REPORT_FROM,
    to,
    subject,
    text,
  });
}

export {
  analyzeWorkbook,
  compactTimestamp,
  countMelMentions,
  parseMaxAttachmentBytes,
  pickWorkbookAttachment,
  resolveReportRecipient,
  renderReportText,
  sanitizeFileName,
};

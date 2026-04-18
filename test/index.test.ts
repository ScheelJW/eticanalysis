import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import {
  analyzeWorkbook,
  compactTimestamp,
  parseMaxAttachmentBytes,
  pickWorkbookAttachment,
  renderReportText,
  sanitizeFileName,
} from "../src/index";

describe("utility helpers", () => {
  it("sanitizes file names safely", () => {
    expect(sanitizeFileName("Vehicle ETIC (Daily).xlsx")).toBe("Vehicle_ETIC__Daily_.xlsx");
  });

  it("formats compact UTC timestamps", () => {
    const date = new Date("2026-04-17T12:34:56.000Z");
    expect(compactTimestamp(date)).toBe("20260417-123456");
  });

  it("parses max attachment bytes with fallback", () => {
    expect(parseMaxAttachmentBytes(undefined)).toBe(30 * 1024 * 1024);
    expect(parseMaxAttachmentBytes("2097152")).toBe(2097152);
    expect(parseMaxAttachmentBytes("0")).toBe(30 * 1024 * 1024);
    expect(parseMaxAttachmentBytes("abc")).toBe(30 * 1024 * 1024);
  });

  it("prefers exact expected attachment then xlsx fallback", () => {
    const attachments = [
      {
        filename: "other.txt",
        mimeType: "text/plain",
        disposition: "attachment" as const,
        content: new TextEncoder().encode("hello"),
      },
      {
        filename: "Vehicle ETIC.xlsx",
        mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disposition: "attachment" as const,
        content: new Uint8Array([1, 2, 3]),
      },
    ];

    const exact = pickWorkbookAttachment(attachments, "vehicle etic.xlsx");
    expect(exact?.filename).toBe("Vehicle ETIC.xlsx");

    const fallback = pickWorkbookAttachment(attachments, "does-not-exist.xlsx");
    expect(fallback?.filename).toBe("Vehicle ETIC.xlsx");
  });
});

describe("workbook analysis", () => {
  it("analyzes visible and hidden worksheets", async () => {
    const workbook = new ExcelJS.Workbook();
    const sheetA = workbook.addWorksheet("Vehicle Status");
    sheetA.addRow(["VIN", "MEL", "Status"]);
    sheetA.addRow(["123", "MEL 2", "Ready"]);
    sheetA.addRow(["456", "none", "Maintenance"]);

    const sheetB = workbook.addWorksheet("Raw Data", { state: "hidden" });
    sheetB.addRow(["id", "note"]);
    sheetB.addRow([1, "mel pending"]);

    const binary = await workbook.xlsx.writeBuffer();
    const result = await analyzeWorkbook({
      binary,
      fileName: "Vehicle ETIC.xlsx",
      receivedAtIso: "2026-04-17T00:00:00.000Z",
      from: "ops@example.com",
      to: "intake@example.com",
      subject: "Daily ETIC",
    });

    expect(result.totalVisibleSheets).toBe(1);
    expect(result.totalHiddenSheets).toBe(1);
    expect(result.sheetSummaries).toHaveLength(2);
    expect(result.melMentionsBySheet["Vehicle Status"]).toBeGreaterThanOrEqual(1);

    const report = renderReportText(result);
    expect(report).toContain("Vehicle ETIC Daily Analysis");
    expect(report).toContain("Vehicle Status");
    expect(report).toContain("Raw Data");
  });

  it("handles null and formula cells without crashing", async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Formula Edge Cases");
    sheet.addRow(["Header A", "Header B"]);
    sheet.getCell("A2").value = "";
    sheet.getCell("B2").value = { formula: "A2", result: "" };
    sheet.getCell("A3").value = { formula: "1+1", result: 2 };
    sheet.getCell("B3").value = "MEL READY";

    const binary = await workbook.xlsx.writeBuffer();
    const result = await analyzeWorkbook({
      binary,
      fileName: "Vehicle ETIC.xlsx",
      receivedAtIso: "2026-04-18T00:00:00.000Z",
      from: "ops@example.com",
      to: "intake@example.com",
      subject: "Daily ETIC edge case",
    });

    expect(result.sheetSummaries[0]?.name).toBe("Formula Edge Cases");
    expect(result.sheetSummaries[0]?.sampleHeaders).toContain("Header A");
    expect(result.melMentionsBySheet["Formula Edge Cases"]).toBeGreaterThanOrEqual(1);
  });
});

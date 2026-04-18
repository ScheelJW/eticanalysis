import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import {
  analyzeWorkbook,
  isoDateKey,
  parseMaxAttachmentBytes,
  pickWorkbookAttachment,
  sanitizeFileName,
  upsertHistoryEntry,
} from "../src/index";

describe("utility helpers", () => {
  it("sanitizes file names safely", () => {
    expect(sanitizeFileName("Vehicle ETIC (Daily).xlsx")).toBe("Vehicle_ETIC__Daily_.xlsx");
  });

  it("formats ISO date keys", () => {
    const date = new Date("2026-04-17T12:34:56.000Z");
    expect(isoDateKey(date)).toBe("2026-04-17");
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

async function makeAnalysis(
  overrides: Partial<{
    dateKey: string;
    totalRowsAcrossSheets: number;
    melMentionsTotal: number;
    totalVisibleSheets: number;
    totalHiddenSheets: number;
    workbookFileName: string;
  }>,
) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Vehicles");
  sheet.addRow(["VIN", "MEL", "Status"]);
  sheet.addRow(["123", "MEL 2", "Ready"]);
  const binary = await workbook.xlsx.writeBuffer();
  const analysis = await analyzeWorkbook({
    binary,
    fileName: overrides.workbookFileName ?? "Vehicle ETIC.xlsx",
    receivedAtIso: new Date("2026-04-17T00:00:00.000Z").toISOString(),
    dateKey: overrides.dateKey ?? "2026-04-17",
    from: "ops@example.com",
    to: "etic@2t3.app",
    subject: "Daily ETIC",
  });
  return {
    ...analysis,
    totalRowsAcrossSheets: overrides.totalRowsAcrossSheets ?? analysis.totalRowsAcrossSheets,
    melMentionsTotal: overrides.melMentionsTotal ?? analysis.melMentionsTotal,
    totalVisibleSheets: overrides.totalVisibleSheets ?? analysis.totalVisibleSheets,
    totalHiddenSheets: overrides.totalHiddenSheets ?? analysis.totalHiddenSheets,
  };
}

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
      dateKey: "2026-04-17",
      from: "ops@example.com",
      to: "etic@2t3.app",
      subject: "Daily ETIC",
    });

    expect(result.totalVisibleSheets).toBe(1);
    expect(result.totalHiddenSheets).toBe(1);
    expect(result.sheetSummaries).toHaveLength(2);
    expect(result.melMentionsBySheet["Vehicle Status"]).toBeGreaterThanOrEqual(1);
    expect(result.melMentionsTotal).toBeGreaterThanOrEqual(2);
  });

  it("handles formula cells without crashing", async () => {
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
      dateKey: "2026-04-18",
      from: "ops@example.com",
      to: "etic@2t3.app",
      subject: "Daily ETIC edge case",
    });

    expect(result.sheetSummaries[0]?.name).toBe("Formula Edge Cases");
    expect(result.sheetSummaries[0]?.sampleHeaders).toContain("Header A");
    expect(result.melMentionsBySheet["Formula Edge Cases"]).toBeGreaterThanOrEqual(1);
  });
});

describe("history upsert", () => {
  it("appends the first entry with no diff", async () => {
    const analysis = await makeAnalysis({
      dateKey: "2026-04-17",
      totalRowsAcrossSheets: 10,
      melMentionsTotal: 3,
    });
    const history = upsertHistoryEntry(
      { updatedAtIso: new Date().toISOString(), entries: [] },
      analysis,
      "workbooks/2026-04-17/x.xlsx",
      "analyses/2026-04-17.json",
    );
    expect(history.entries).toHaveLength(1);
    expect(history.entries[0]?.diff.previousDateKey).toBeNull();
    expect(history.entries[0]?.diff.deltaTotalRows).toBeNull();
  });

  it("computes delta vs previous day", async () => {
    const day1 = await makeAnalysis({
      dateKey: "2026-04-17",
      totalRowsAcrossSheets: 10,
      melMentionsTotal: 3,
      totalVisibleSheets: 5,
    });
    const day2 = await makeAnalysis({
      dateKey: "2026-04-18",
      totalRowsAcrossSheets: 12,
      melMentionsTotal: 5,
      totalVisibleSheets: 6,
    });
    let history = upsertHistoryEntry(
      { updatedAtIso: new Date().toISOString(), entries: [] },
      day1,
      "workbooks/2026-04-17/x.xlsx",
      "analyses/2026-04-17.json",
    );
    history = upsertHistoryEntry(
      history,
      day2,
      "workbooks/2026-04-18/y.xlsx",
      "analyses/2026-04-18.json",
    );
    expect(history.entries).toHaveLength(2);
    const latest = history.entries[1];
    expect(latest?.diff.previousDateKey).toBe("2026-04-17");
    expect(latest?.diff.deltaTotalRows).toBe(2);
    expect(latest?.diff.deltaMelMentionsTotal).toBe(2);
    expect(latest?.diff.deltaSheetsVisible).toBe(1);
  });

  it("replaces same-day entry on re-send", async () => {
    const initial = await makeAnalysis({
      dateKey: "2026-04-18",
      totalRowsAcrossSheets: 5,
      melMentionsTotal: 1,
    });
    let history = upsertHistoryEntry(
      { updatedAtIso: new Date().toISOString(), entries: [] },
      initial,
      "workbooks/2026-04-18/a.xlsx",
      "analyses/2026-04-18.json",
    );
    const corrected = await makeAnalysis({
      dateKey: "2026-04-18",
      totalRowsAcrossSheets: 7,
      melMentionsTotal: 2,
    });
    history = upsertHistoryEntry(
      history,
      corrected,
      "workbooks/2026-04-18/b.xlsx",
      "analyses/2026-04-18.json",
    );
    expect(history.entries).toHaveLength(1);
    expect(history.entries[0]?.totalRowsAcrossSheets).toBe(7);
    expect(history.entries[0]?.workbookKey).toBe("workbooks/2026-04-18/b.xlsx");
  });
});

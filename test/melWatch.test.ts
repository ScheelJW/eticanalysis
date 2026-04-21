import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { extractMelRowsFromBinary } from "../src/melWatch";

async function workbookToBuffer(wb: ExcelJS.Workbook): Promise<ArrayBuffer> {
  const buf = await wb.xlsx.writeBuffer();
  return buf instanceof ArrayBuffer ? buf : new Uint8Array(buf).buffer;
}

describe("extractMelRowsFromBinary", () => {
  it("finds MEL grid when sheet is not named MEL Calculator", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Daily MEL Rollup");
    ws.addRow(["MEL Key", "MEL assigned total", "NMC count", "FMC count", "MEL status"]);
    ws.addRow(["UVAN791", "19", "15", "4", "BELOW MEL"]);
    const bytes = await workbookToBuffer(wb);
    const rows = await extractMelRowsFromBinary(bytes);
    expect(rows.length).toBe(1);
    expect(rows[0].melKey).toBe("UVAN791");
    expect(rows[0].melAssignedTotal).toBe(19);
    expect(rows[0].nmcCount).toBe(15);
    expect(rows[0].fmcCount).toBe(4);
    expect(rows[0].melStatus).toBe("below");
  });
});

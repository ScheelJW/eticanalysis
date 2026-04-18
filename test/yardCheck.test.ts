import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import {
  extractWorkOrderRows,
  generateYardCheckWorkbookBuffer,
  matchHeaderToColumn,
} from "../src/yardCheck";

describe("matchHeaderToColumn", () => {
  it("matches canonical headers", () => {
    expect(matchHeaderToColumn("Asset ID")).toBe("assetId");
    expect(matchHeaderToColumn("Work Order ID")).toBe("workOrderId");
    expect(matchHeaderToColumn("Work Order Remarks")).toBe("workOrderRemarks");
    expect(matchHeaderToColumn("VIN/Serial")).toBe("vinSerial");
    expect(matchHeaderToColumn("Make")).toBe("make");
    expect(matchHeaderToColumn("Model")).toBe("model");
    expect(matchHeaderToColumn("Previous Location")).toBe("previousLocation");
  });

  it("matches header variants", () => {
    expect(matchHeaderToColumn("Asset #")).toBe("assetId");
    expect(matchHeaderToColumn("WO Number")).toBe("workOrderId");
    expect(matchHeaderToColumn("Remarks")).toBe("workOrderRemarks");
    expect(matchHeaderToColumn("Serial Number")).toBe("vinSerial");
    expect(matchHeaderToColumn("Manufacturer")).toBe("make");
    expect(matchHeaderToColumn("Last Known Location")).toBe("previousLocation");
  });

  it("returns null for unknowns", () => {
    expect(matchHeaderToColumn("Total Cost")).toBeNull();
    expect(matchHeaderToColumn("")).toBeNull();
  });
});

async function buildTestWorkbookWithWorkOrders(): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Work Orders");
  sheet.addRow(["Asset ID", "Work Order ID", "VIN/Serial", "Make", "Model", "Previous Location", "Work Order Remarks"]);
  sheet.addRow(["V001", "WO-1001", "1HGCM82633A004352", "Ford", "F-150", "Bay 3", "Oil change and tire rotation"]);
  sheet.addRow(["V002", "WO-1002", "JT8BD69S8Y0091234", "Chevy", "Silverado", "Lot A", "Brake inspection"]);
  sheet.addRow(["", "", "", "", "", "", ""]);
  sheet.addRow(["V003", "WO-1003", "SN-77-AA", "Kubota", "Tractor", "Shed 2", "MEL reset after hydraulic repair"]);

  const buffer = await workbook.xlsx.writeBuffer();
  return toArrayBufferLike(buffer);
}

function toArrayBufferLike(buf: unknown): ArrayBuffer {
  if (buf instanceof ArrayBuffer) return buf;
  const view = buf as unknown as { buffer?: ArrayBufferLike; byteOffset?: number; byteLength?: number };
  if (view && view.buffer && typeof view.byteOffset === "number" && typeof view.byteLength === "number") {
    const clone = new Uint8Array(view.byteLength);
    clone.set(new Uint8Array(view.buffer as ArrayBuffer, view.byteOffset, view.byteLength));
    return clone.buffer;
  }
  throw new Error("Unsupported buffer type in test helper");
}

describe("extractWorkOrderRows", () => {
  it("extracts rows from a Work Orders sheet", async () => {
    const binary = await buildTestWorkbookWithWorkOrders();
    const result = await extractWorkOrderRows(binary);
    expect(result).not.toBeNull();
    if (!result) return;
    expect(result.sheetName).toBe("Work Orders");
    expect(result.headerRowIndex).toBe(1);
    expect(result.totalDataRows).toBe(3);
    expect(result.rows[0]?.assetId).toBe("V001");
    expect(result.rows[1]?.model).toBe("Silverado");
    expect(result.rows[2]?.workOrderRemarks).toContain("MEL reset");
    expect(result.headerMap.assetId).toBe("Asset ID");
    expect(result.headerMap.previousLocation).toBe("Previous Location");
  });

  it("returns null when no work order sheet exists", async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet("Summary").addRow(["Hello", "World"]);
    const buffer = await workbook.xlsx.writeBuffer();
    const result = await extractWorkOrderRows(toArrayBufferLike(buffer));
    expect(result).toBeNull();
  });
});

describe("generateYardCheckWorkbookBuffer", () => {
  it("produces a readable workbook with required headers", async () => {
    const binary = await buildTestWorkbookWithWorkOrders();
    const extraction = await extractWorkOrderRows(binary);
    expect(extraction).not.toBeNull();
    if (!extraction) return;
    const output = await generateYardCheckWorkbookBuffer(extraction, {
      sourceWorkbookFileName: "Vehicle ETIC.xlsx",
      sourceDateKey: "2026-04-18",
      generatedAtIso: "2026-04-18T12:00:00.000Z",
    });
    expect(output.byteLength).toBeGreaterThan(1000);

    const reloaded = new ExcelJS.Workbook();
    await reloaded.xlsx.load(output);
    const sheet = reloaded.getWorksheet("Yard Check");
    expect(sheet).toBeDefined();
    if (!sheet) return;
    expect(sheet.pageSetup?.orientation).toBe("landscape");

    const header = sheet.getRow(1);
    const headerValues = header.values as (string | undefined)[];
    expect(headerValues).toContain("Asset ID");
    expect(headerValues).toContain("Work Order ID");
    expect(headerValues).toContain("VIN / Serial");
    expect(headerValues).toContain("Make");
    expect(headerValues).toContain("Model");
    expect(headerValues).toContain("Previous Location");
    expect(headerValues).toContain("New Location");
    expect(headerValues).toContain("Discrepancies");
    expect(headerValues).toContain("Work Order Remarks");

    expect(sheet.actualRowCount).toBeGreaterThanOrEqual(extraction.rows.length + 1);
  });
});

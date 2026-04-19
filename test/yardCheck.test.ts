import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import {
  extractYardCheckSource,
  generateYardCheckWorkbookBuffer,
  matchHeader,
} from "../src/yardCheck";

type WorkOrderField =
  | "assetId"
  | "workOrderId"
  | "remarks"
  | "shop"
  | "shop2"
  | "etiCLocation"
  | "makeModel";

type FleetField = "assetId" | "vinSerial" | "makeModel" | "etiCLocation";

const workOrderSynonyms: Record<WorkOrderField, string[]> = {
  assetId: ["asset id"],
  workOrderId: ["work order id", "wo id", "work order"],
  remarks: ["remarks", "work order remarks"],
  shop: ["shop", "primary shop"],
  shop2: ["shop2", "shop 2"],
  etiCLocation: ["etic location", "location"],
  makeModel: ["make model", "make/model"],
};

const fleetSynonyms: Record<FleetField, string[]> = {
  assetId: ["asset id"],
  vinSerial: ["serial nbr", "serial number", "vin"],
  makeModel: ["make/model"],
  etiCLocation: ["etic location"],
};

describe("matchHeader", () => {
  it("matches Work Orders aliases", () => {
    expect(matchHeader("Asset Id", workOrderSynonyms)).toBe("assetId");
    expect(matchHeader("Work Order Id", workOrderSynonyms)).toBe("workOrderId");
    expect(matchHeader("Remarks", workOrderSynonyms)).toBe("remarks");
    expect(matchHeader("Shop", workOrderSynonyms)).toBe("shop");
    expect(matchHeader("Shop2", workOrderSynonyms)).toBe("shop2");
    expect(matchHeader("ETIC Location", workOrderSynonyms)).toBe("etiCLocation");
    expect(matchHeader("Make Model", workOrderSynonyms)).toBe("makeModel");
  });

  it("matches Fleet (P&A) Serial Nbr", () => {
    expect(matchHeader("Serial Nbr", fleetSynonyms)).toBe("vinSerial");
    expect(matchHeader("Make/Model", fleetSynonyms)).toBe("makeModel");
  });
});

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

async function buildRealLikeWorkbook(): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  const wo = workbook.addWorksheet("WO INQUIRY");
  wo.addRow([
    "Work Order Id",
    "Open Dt",
    "Asset Id",
    "Remarks",
    "Make Model",
    "Shop",
    "Shop2",
    "ETIC Location",
  ]);
  wo.addRow([
    "WO-001",
    "2026-04-14",
    "AF17B00485",
    "Fuel gauge troubleshoot",
    "FORD F250",
    "GP",
    "",
    "Bay 3",
  ]);
  wo.addRow([
    "WO-002",
    "2026-04-16",
    "AF17B00485",
    "Brake service",
    "FORD F250",
    "FARM",
    "GP",
    "",
  ]);
  wo.addRow([
    "WO-003",
    "2026-04-17",
    "AF08B01113",
    "Engine diagnostics",
    "FORD F150",
    "FARM",
    "",
    "Lot A",
  ]);

  const fleet = workbook.addWorksheet("Fleet (P&A)");
  fleet.addRow(["Asset Id", "Make/Model", "Serial Nbr", "WO Inquiry.ETIC Location"]);
  fleet.addRow(["AF17B00485", "FORD F250", "1FDBF2A60HED60465", "Compound"]);
  fleet.addRow(["AF08B01113", "FORD F150", "1FTFW1E55HFB12345", "Shed 2"]);

  return toArrayBufferLike(await workbook.xlsx.writeBuffer());
}

describe("extractYardCheckSource", () => {
  it("groups work orders per asset and joins VIN from Fleet (P&A)", async () => {
    const binary = await buildRealLikeWorkbook();
    const source = await extractYardCheckSource(binary);
    expect(source).not.toBeNull();
    if (!source) return;

    expect(source.workOrderSheet).toBe("WO INQUIRY");
    expect(source.fleetSheet).toBe("Fleet (P&A)");
    expect(source.totalAssets).toBe(2);
    expect(source.totalWorkOrders).toBe(3);

    const row1 = source.rows.find((r) => r.assetId === "AF17B00485");
    expect(row1).toBeDefined();
    if (!row1) return;
    expect(row1.workOrderIds).toEqual(["WO-001", "WO-002"]);
    expect(row1.shops).toContain("GP");
    expect(row1.shops).toContain("FARM");
    expect(row1.vinSerial).toBe("1FDBF2A60HED60465");
    expect(row1.workOrderRemarks.length).toBe(2);

    const row2 = source.rows.find((r) => r.assetId === "AF08B01113");
    expect(row2).toBeDefined();
    if (!row2) return;
    expect(row2.workOrderIds).toEqual(["WO-003"]);
    expect(row2.vinSerial).toBe("1FTFW1E55HFB12345");
    expect(row2.previousLocations).toContain("Lot A");
  });

  it("returns null when no work order sheet exists", async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet("Summary").addRow(["Hello", "World"]);
    const buffer = await workbook.xlsx.writeBuffer();
    const result = await extractYardCheckSource(toArrayBufferLike(buffer));
    expect(result).toBeNull();
  });
});

describe("generateYardCheckWorkbookBuffer", () => {
  it("produces landscape workbook with expected headers", async () => {
    const binary = await buildRealLikeWorkbook();
    const source = await extractYardCheckSource(binary);
    expect(source).not.toBeNull();
    if (!source) return;
    const output = await generateYardCheckWorkbookBuffer(source, {
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
    expect(headerValues).toContain("Work Order ID(s)");
    expect(headerValues).toContain("VIN / Serial");
    expect(headerValues).toContain("Make / Model");
    expect(headerValues).toContain("Shop(s)");
    expect(headerValues).toContain("Previous Location");
    expect(headerValues).toContain("New Location");
    expect(headerValues).toContain("Discrepancies");
    expect(headerValues).toContain("Work Order Remarks");

    expect(sheet.actualRowCount).toBe(source.rows.length + 1);
  });
});

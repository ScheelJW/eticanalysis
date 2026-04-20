import ExcelJS from "exceljs";

export type YardCheckRow = {
  assetId: string;
  workOrderIds: string[];
  workOrderRemarks: string[];
  vinSerial: string;
  makeModel: string;
  shops: string[];
  previousLocations: string[];
};

export type YardCheckSource = {
  workOrderSheet: string | null;
  workOrderHeaderRowIndex: number | null;
  fleetSheet: string | null;
  fleetHeaderRowIndex: number | null;
  totalAssets: number;
  totalWorkOrders: number;
  rows: YardCheckRow[];
  unmatchedWorkOrderHeaders: string[];
  unmatchedFleetHeaders: string[];
};

export type YardCheckMeta = {
  sourceWorkbookFileName: string;
  sourceDateKey: string;
  generatedAtIso: string;
};

type WorkOrderField =
  | "assetId"
  | "workOrderId"
  | "remarks"
  | "shop"
  | "shop2"
  | "etiCLocation"
  | "makeModel"
  | "partsStatus"
  | "eticDue"
  | "currentMel";

type FleetField = "assetId" | "vinSerial" | "makeModel" | "etiCLocation";

const WORK_ORDER_SYNONYMS: Record<WorkOrderField, string[]> = {
  assetId: ["asset id", "asset", "asset number", "asset #", "equipment id", "equip id", "unit id"],
  workOrderId: [
    "work order id",
    "work order number",
    "work order no",
    "work order #",
    "wo id",
    "wo number",
    "wo no",
    "wo #",
    "work order",
  ],
  remarks: [
    "remarks",
    "work order remarks",
    "wo remarks",
    "comments",
    "notes",
    "description",
    "job description",
    "work description",
    "defect",
    "defect description",
  ],
  shop: ["shop", "primary shop", "assigned shop"],
  shop2: ["shop2", "secondary shop", "shop 2", "support shop"],
  etiCLocation: [
    "etic location",
    "current location",
    "last known location",
    "last location",
    "previous location",
    "location",
  ],
  makeModel: ["make model", "make/model", "make / model"],
  partsStatus: [
    "parts status",
    "part status",
    "parts",
    "parts availability",
    "material status",
  ],
  eticDue: [
    "etic",
    "etic date",
    "current etic",
    "projected etic",
    "etic due",
    "estimated completion",
  ],
  currentMel: [
    "mel",
    "current mel",
    "mel level",
    "mission essential",
    "mel status",
    "m e l",
  ],
};

const FLEET_SYNONYMS: Record<FleetField, string[]> = {
  assetId: ["asset id", "asset", "asset number", "asset #"],
  vinSerial: [
    "serial nbr",
    "serial number",
    "serial",
    "vin",
    "vin number",
    "vin/serial",
    "vin / serial",
    "vin or serial",
    "vin-serial",
  ],
  makeModel: ["make/model", "make / model", "make model", "make", "model"],
  etiCLocation: [
    "wo inquiry.etic location",
    "etic location",
    "current location",
    "location",
    "previous location",
  ],
};

function normalizeHeader(raw: string): string {
  return raw
    .replace(/[\u00a0]/g, " ")
    .replace(/\s+/g, " ")
    .replace(/[._]/g, " ")
    .trim()
    .toLowerCase();
}

function readCellString(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map((part) => String(part ?? "")).join("");
  if (typeof value === "object") {
    const v = value as unknown as Record<string, unknown>;
    if (Array.isArray(v.richText)) {
      return v.richText
        .map((p) => (p && typeof p === "object" ? String((p as { text?: unknown }).text ?? "") : ""))
        .join("");
    }
    if (typeof v.text === "string") return v.text;
    if (v.result !== null && v.result !== undefined) return String(v.result);
    if (typeof v.formula === "string") return v.formula;
    if (typeof v.hyperlink === "string") return v.hyperlink;
    if (typeof v.error === "string") return v.error;
  }
  try {
    return String(value);
  } catch {
    return "";
  }
}

function cleanText(raw: string): string {
  return raw.replace(/\s+/g, " ").trim();
}

export function matchHeader<T extends string>(
  header: string,
  synonymMap: Record<T, string[]>,
): T | null {
  const normalized = normalizeHeader(header);
  if (!normalized) return null;

  let bestField: T | null = null;
  let bestScore = 0;

  (Object.keys(synonymMap) as T[]).forEach((field) => {
    const synonyms = synonymMap[field];
    for (const candidate of synonyms) {
      if (normalized === candidate) {
        const s = candidate.length + 1000;
        if (s > bestScore) {
          bestField = field;
          bestScore = s;
        }
        continue;
      }
      if (normalized.includes(candidate) && candidate.length > bestScore) {
        bestField = field;
        bestScore = candidate.length;
      }
    }
  });

  return bestField;
}

type HeaderScan<T extends string> = {
  rowIndex: number;
  mapping: Map<number, T>;
  rawHeaders: Map<number, string>;
  unmatched: string[];
};

function scanHeaderRow<T extends string>(
  sheet: ExcelJS.Worksheet,
  synonymMap: Record<T, string[]>,
  maxRows = 25,
): HeaderScan<T> | null {
  let best: HeaderScan<T> | null = null;

  for (let rowIndex = 1; rowIndex <= maxRows; rowIndex += 1) {
    const row = sheet.getRow(rowIndex);
    if (!row || row.cellCount === 0) continue;

    const mapping = new Map<number, T>();
    const rawHeaders = new Map<number, string>();
    const unmatched: string[] = [];

    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const raw = readCellString(cell).trim();
      if (!raw) return;
      rawHeaders.set(colNumber, raw);
      const matched = matchHeader<T>(raw, synonymMap);
      if (matched) {
        if (!mapping.has(colNumber) && !hasValue(mapping, matched)) {
          mapping.set(colNumber, matched);
        }
      } else {
        unmatched.push(raw);
      }
    });

    if (rawHeaders.size === 0) continue;
    if (!best || mapping.size > best.mapping.size) {
      best = { rowIndex, mapping, rawHeaders, unmatched };
    }
    if (best.mapping.size >= 4) break;
  }

  if (!best || best.mapping.size === 0) return null;
  return best;
}

function hasValue<T>(map: Map<number, T>, value: T): boolean {
  for (const v of map.values()) {
    if (v === value) return true;
  }
  return false;
}

function findWorkOrderSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet | null {
  const all = workbook.worksheets;
  return (
    all.find((s) => /^work\s*orders?$/i.test(s.name)) ??
    all.find((s) => /^wo\s*inquiry/i.test(s.name)) ??
    all.find((s) => /work\s*orders?/i.test(s.name)) ??
    all.find((s) => /\bwo\b/i.test(s.name)) ??
    null
  );
}

function findFleetSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet | null {
  const all = workbook.worksheets;
  return (
    all.find((s) => /^fleet\b/i.test(s.name)) ??
    all.find((s) => /fleet/i.test(s.name)) ??
    null
  );
}

export type RawWorkOrder = {
  assetId: string;
  workOrderId: string;
  remarks: string;
  shop: string;
  shop2: string;
  etiCLocation: string;
  makeModel: string;
  partsStatus: string;
  eticDue: string;
  currentMel: string;
};

type FleetRecord = {
  assetId: string;
  vinSerial: string;
  makeModel: string;
  etiCLocation: string;
};

function extractWorkOrders(sheet: ExcelJS.Worksheet): {
  rows: RawWorkOrder[];
  scan: HeaderScan<WorkOrderField> | null;
} {
  const scan = scanHeaderRow<WorkOrderField>(sheet, WORK_ORDER_SYNONYMS);
  if (!scan) return { rows: [], scan: null };

  const rows: RawWorkOrder[] = [];
  const lastRow = sheet.actualRowCount || sheet.rowCount;
  for (let r = scan.rowIndex + 1; r <= lastRow; r += 1) {
    const row = sheet.getRow(r);
    if (!row || row.cellCount === 0) continue;
    const record: RawWorkOrder = {
      assetId: "",
      workOrderId: "",
      remarks: "",
      shop: "",
      shop2: "",
      etiCLocation: "",
      makeModel: "",
      partsStatus: "",
      eticDue: "",
      currentMel: "",
    };
    let hasAny = false;
    scan.mapping.forEach((field, colNumber) => {
      const text = cleanText(readCellString(row.getCell(colNumber)));
      if (text) hasAny = true;
      (record as Record<string, string>)[field] = text;
    });
    if (hasAny && record.assetId) rows.push(record);
  }
  return { rows, scan };
}

function extractFleet(sheet: ExcelJS.Worksheet): {
  byAsset: Map<string, FleetRecord>;
  scan: HeaderScan<FleetField> | null;
} {
  const scan = scanHeaderRow<FleetField>(sheet, FLEET_SYNONYMS);
  if (!scan) return { byAsset: new Map(), scan: null };

  const byAsset = new Map<string, FleetRecord>();
  const lastRow = sheet.actualRowCount || sheet.rowCount;
  for (let r = scan.rowIndex + 1; r <= lastRow; r += 1) {
    const row = sheet.getRow(r);
    if (!row || row.cellCount === 0) continue;
    const record: FleetRecord = {
      assetId: "",
      vinSerial: "",
      makeModel: "",
      etiCLocation: "",
    };
    let hasAny = false;
    scan.mapping.forEach((field, colNumber) => {
      const text = cleanText(readCellString(row.getCell(colNumber)));
      if (text) hasAny = true;
      (record as Record<string, string>)[field] = text;
    });
    if (hasAny && record.assetId) {
      byAsset.set(record.assetId, record);
    }
  }
  return { byAsset, scan };
}

/** Raw WO rows from the workbook (one row per work order line). */
export async function extractRawWorkOrdersFromBinary(binary: ArrayBuffer): Promise<RawWorkOrder[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(binary);
  const woSheet = findWorkOrderSheet(workbook);
  if (!woSheet) return [];
  return extractWorkOrders(woSheet).rows;
}

export async function extractYardCheckSource(binary: ArrayBuffer): Promise<YardCheckSource | null> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(binary);

  const woSheet = findWorkOrderSheet(workbook);
  if (!woSheet) return null;

  const woResult = extractWorkOrders(woSheet);
  const fleetSheet = findFleetSheet(workbook);
  const fleetResult = fleetSheet ? extractFleet(fleetSheet) : { byAsset: new Map<string, FleetRecord>(), scan: null };

  const grouped = new Map<string, YardCheckRow>();
  for (const wo of woResult.rows) {
    if (!wo.assetId) continue;
    const existing = grouped.get(wo.assetId);
    const fleet = fleetResult.byAsset.get(wo.assetId);
    if (!existing) {
      const row: YardCheckRow = {
        assetId: wo.assetId,
        workOrderIds: wo.workOrderId ? [wo.workOrderId] : [],
        workOrderRemarks: formatRemark(wo) ? [formatRemark(wo)] : [],
        vinSerial: fleet?.vinSerial ?? "",
        makeModel: (fleet?.makeModel ?? "") || wo.makeModel,
        shops: uniqNonEmpty([wo.shop, wo.shop2]),
        previousLocations: uniqNonEmpty([wo.etiCLocation, fleet?.etiCLocation ?? ""]),
      };
      grouped.set(wo.assetId, row);
    } else {
      if (wo.workOrderId && !existing.workOrderIds.includes(wo.workOrderId)) {
        existing.workOrderIds.push(wo.workOrderId);
      }
      const remark = formatRemark(wo);
      if (remark && !existing.workOrderRemarks.includes(remark)) {
        existing.workOrderRemarks.push(remark);
      }
      for (const shop of [wo.shop, wo.shop2]) {
        if (shop && !existing.shops.includes(shop)) existing.shops.push(shop);
      }
      if (wo.etiCLocation && !existing.previousLocations.includes(wo.etiCLocation)) {
        existing.previousLocations.push(wo.etiCLocation);
      }
      if (!existing.vinSerial && fleet?.vinSerial) existing.vinSerial = fleet.vinSerial;
      if (!existing.makeModel) existing.makeModel = (fleet?.makeModel ?? "") || wo.makeModel;
      if (fleet?.etiCLocation && !existing.previousLocations.includes(fleet.etiCLocation)) {
        existing.previousLocations.push(fleet.etiCLocation);
      }
    }
  }

  const rows = [...grouped.values()].sort((a, b) => a.assetId.localeCompare(b.assetId));

  return {
    workOrderSheet: woSheet.name,
    workOrderHeaderRowIndex: woResult.scan?.rowIndex ?? null,
    fleetSheet: fleetSheet?.name ?? null,
    fleetHeaderRowIndex: fleetResult.scan?.rowIndex ?? null,
    totalAssets: rows.length,
    totalWorkOrders: woResult.rows.length,
    rows,
    unmatchedWorkOrderHeaders: woResult.scan?.unmatched ?? [],
    unmatchedFleetHeaders: fleetResult.scan?.unmatched ?? [],
  };
}

function formatRemark(wo: RawWorkOrder): string {
  if (!wo.remarks) return "";
  if (wo.workOrderId) return `${wo.workOrderId}: ${wo.remarks}`;
  return wo.remarks;
}

function uniqNonEmpty(values: Array<string | undefined>): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const value of values) {
    if (!value) continue;
    if (seen.has(value)) continue;
    seen.add(value);
    out.push(value);
  }
  return out;
}

export async function generateYardCheckWorkbookBuffer(
  source: YardCheckSource,
  meta: YardCheckMeta,
): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Minot Vehicle ETIC Dashboard";
  workbook.created = new Date();

  const sheet = workbook.addWorksheet("Yard Check", {
    pageSetup: {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalCentered: true,
      margins: {
        left: 0.4,
        right: 0.4,
        top: 0.6,
        bottom: 0.6,
        header: 0.3,
        footer: 0.3,
      },
      printTitlesRow: "1:1",
    },
    headerFooter: {
      oddHeader: `&L&"Calibri,Bold"&12Minot Vehicle ETIC — Yard Check&R&10Source: ${meta.sourceWorkbookFileName} (${meta.sourceDateKey})`,
      oddFooter: `&LGenerated ${meta.generatedAtIso}&RPage &P of &N`,
      evenHeader: `&L&"Calibri,Bold"&12Minot Vehicle ETIC — Yard Check&R&10Source: ${meta.sourceWorkbookFileName} (${meta.sourceDateKey})`,
      evenFooter: `&LGenerated ${meta.generatedAtIso}&RPage &P of &N`,
    },
    views: [{ state: "frozen", ySplit: 1 }],
    properties: { defaultRowHeight: 22 },
  });

  sheet.columns = [
    { header: "Asset ID", key: "assetId", width: 13 },
    { header: "Work Order ID(s)", key: "workOrderIds", width: 20 },
    { header: "VIN / Serial", key: "vinSerial", width: 22 },
    { header: "Make / Model", key: "makeModel", width: 18 },
    { header: "Shop(s)", key: "shops", width: 12 },
    { header: "Previous Location", key: "previousLocations", width: 18 },
    { header: "New Location", key: "newLocation", width: 22 },
    { header: "Discrepancies", key: "discrepancies", width: 28 },
    { header: "Work Order Remarks", key: "workOrderRemarks", width: 48 },
  ];

  const headerRow = sheet.getRow(1);
  headerRow.height = 26;
  headerRow.font = { bold: true, size: 11 };
  headerRow.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  headerRow.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEEF2F8" } };
    cell.border = {
      top: { style: "thin", color: { argb: "FF9FB1C8" } },
      left: { style: "thin", color: { argb: "FF9FB1C8" } },
      bottom: { style: "thin", color: { argb: "FF9FB1C8" } },
      right: { style: "thin", color: { argb: "FF9FB1C8" } },
    };
  });

  for (const record of source.rows) {
    const added = sheet.addRow({
      assetId: record.assetId,
      workOrderIds: record.workOrderIds.join("\n"),
      vinSerial: record.vinSerial,
      makeModel: record.makeModel,
      shops: record.shops.join(", "),
      previousLocations: record.previousLocations.join("\n"),
      newLocation: "",
      discrepancies: "",
      workOrderRemarks: record.workOrderRemarks.join("\n\n"),
    });
    added.font = { size: 10 };
    added.alignment = { vertical: "top", wrapText: true };
    added.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = {
        top: { style: "hair", color: { argb: "FFBCC6D6" } },
        left: { style: "hair", color: { argb: "FFBCC6D6" } },
        bottom: { style: "hair", color: { argb: "FFBCC6D6" } },
        right: { style: "hair", color: { argb: "FFBCC6D6" } },
      };
    });
  }

  if (source.rows.length === 0) {
    const empty = sheet.addRow({ assetId: "No work orders found in source workbook." });
    empty.font = { italic: true };
  }

  const bufferLike = await workbook.xlsx.writeBuffer();
  return toArrayBuffer(bufferLike);
}

function toArrayBuffer(buffer: ExcelJS.Buffer): ArrayBuffer {
  if (buffer instanceof ArrayBuffer) return buffer;
  const view = buffer as unknown as { buffer?: ArrayBufferLike; byteOffset?: number; byteLength?: number };
  if (view && view.buffer && typeof view.byteOffset === "number" && typeof view.byteLength === "number") {
    const clone = new Uint8Array(view.byteLength);
    clone.set(new Uint8Array(view.buffer as ArrayBuffer, view.byteOffset, view.byteLength));
    return clone.buffer;
  }
  throw new Error("Unsupported buffer type returned by ExcelJS");
}

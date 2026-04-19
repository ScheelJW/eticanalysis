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

export type WorkOrderField =
  | "assetId"
  | "workOrderId"
  | "remarks"
  | "shop"
  | "shop2"
  | "etiCLocation"
  | "makeModel"
  | "partsStatus"
  | "eticDue"
  | "currentMel"
  | "owningUnit"
  | "melKey"
  | "mgmtCd";

export type FleetField = "assetId" | "vinSerial" | "makeModel" | "etiCLocation" | "mgmtCd" | "vehNomen";

export const WORK_ORDER_SYNONYMS: Record<WorkOrderField, string[]> = {
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
  shop: ["shop assigned", "assigned shop", "primary shop", "shop", "work center"],
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
    "parts stat",
  ],
  eticDue: [
    "est service end dt",
    "est service end date",
    "est service end",
    "estimated service end",
    "service end dt",
    "service end date",
    "etic due",
    "etic date",
    "current etic",
    "projected etic",
    "next etic",
    "new etic",
    "proj etic",
    "estimated completion",
    "estimated time in commission",
    "est time in commission",
    "est completion",
    "proj complete",
    "projected completion",
    "projected complete date",
    "completion date",
  ],
  currentMel: [
    "current mel",
    "mel",
    "mel level",
    "mel tier",
    "mel status",
    "mission essential",
    "mission essential level",
    "m e l",
    "above mel",
    "at mel",
    "below mel",
    "mel position",
    "mission capability",
  ],
  owningUnit: [
    "organization",
    "owning unit",
    "owning org",
    "owning organization",
    "user unit",
    "user/unit",
    "assigned unit",
    "unit",
    "org",
    "squadron",
    "customer",
    "customer unit",
  ],
  melKey: [
    "mel key",
    "mel code",
    "mel id",
    "priority key",
    "priority code",
    "lin code",
    "lin",
  ],
  mgmtCd: [
    "mgmt cd",
    "mgmt code",
    "management code",
    "management cd",
    "master mgmt cd",
    "mstr mgmt cd",
    "lin/tamcn",
  ],
};

export const FLEET_SYNONYMS: Record<FleetField, string[]> = {
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
  mgmtCd: [
    "mgmt cd",
    "mgmt code",
    "management code",
    "master mgmt cd",
    "mstr mgmt cd",
    "lin/tamcn",
  ],
  vehNomen: [
    "veh nomen",
    "vehicle nomenclature",
    "vehicle nomen",
    "nomenclature",
    "nomen",
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

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Score a (header, field) match. 0 = no match.
 * Exact match beats word-boundary beats substring.
 */
export function scoreHeaderMatch<T extends string>(
  header: string,
  synonymMap: Record<T, string[]>,
): { field: T; score: number } | null {
  const normalized = normalizeHeader(header);
  if (!normalized) return null;

  let bestField: T | null = null;
  let bestScore = 0;

  (Object.keys(synonymMap) as T[]).forEach((field) => {
    const synonyms = synonymMap[field];
    for (const candidate of synonyms) {
      let score = 0;
      if (normalized === candidate) {
        score = 10000 + candidate.length;
      } else {
        const wholeWord = new RegExp(`\\b${escapeRegex(candidate)}\\b`);
        const wordHit = wholeWord.test(normalized);
        if (wordHit) {
          score = 100 + candidate.length;
        } else if (normalized.includes(candidate)) {
          score = candidate.length;
        } else if (candidate.length >= 6 && candidate.includes(normalized) && normalized.length >= 3) {
          score = normalized.length;
        }
      }
      if (score > bestScore) {
        bestField = field;
        bestScore = score;
      }
    }
  });

  return bestField !== null ? { field: bestField, score: bestScore } : null;
}

/** Legacy wrapper used by tests — returns just the best-scoring field. */
export function matchHeader<T extends string>(
  header: string,
  synonymMap: Record<T, string[]>,
): T | null {
  return scoreHeaderMatch(header, synonymMap)?.field ?? null;
}

type HeaderScan<T extends string> = {
  rowIndex: number;
  mapping: Map<number, T>;
  rawHeaders: Map<number, string>;
  unmatched: string[];
};

function buildMappingFromScores<T extends string>(
  scored: Array<{ col: number; field: T; score: number }>,
): Map<number, T> {
  scored.sort((a, b) => b.score - a.score);
  const mapping = new Map<number, T>();
  const usedFields = new Set<T>();
  for (const s of scored) {
    if (mapping.has(s.col)) continue;
    if (usedFields.has(s.field)) continue;
    mapping.set(s.col, s.field);
    usedFields.add(s.field);
  }
  return mapping;
}

function scanHeaderRow<T extends string>(
  sheet: ExcelJS.Worksheet,
  synonymMap: Record<T, string[]>,
  maxRows = 25,
): HeaderScan<T> | null {
  let best: HeaderScan<T> | null = null;

  for (let rowIndex = 1; rowIndex <= maxRows; rowIndex += 1) {
    const row = sheet.getRow(rowIndex);
    if (!row || row.cellCount === 0) continue;

    const rawHeaders = new Map<number, string>();
    const scored: Array<{ col: number; field: T; score: number }> = [];
    const matchedCols = new Set<number>();

    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const raw = readCellString(cell).trim();
      if (!raw) return;
      rawHeaders.set(colNumber, raw);
      const m = scoreHeaderMatch<T>(raw, synonymMap);
      if (m) {
        scored.push({ col: colNumber, field: m.field, score: m.score });
        matchedCols.add(colNumber);
      }
    });

    if (rawHeaders.size === 0) continue;

    const mapping = buildMappingFromScores(scored);
    const unmatched: string[] = [];
    for (const [col, raw] of rawHeaders) {
      if (!matchedCols.has(col) || !mapping.has(col)) unmatched.push(raw);
    }
    if (!best || mapping.size > best.mapping.size) {
      best = { rowIndex, mapping, rawHeaders, unmatched };
    }
  }

  if (!best || best.mapping.size === 0) return null;
  return best;
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
  owningUnit: string;
  melKey: string;
  mgmtCd: string;
  vehNomen: string;
  /**
   * Every cell we read for this row, keyed by the *normalized* header name
   * (lowercased, whitespace-collapsed). Includes mapped fields, unmapped
   * columns from the WO sheet, and any merged-in fleet columns. Persisted
   * verbatim to D1 so future field additions can be backfilled from JSON
   * without re-parsing the workbook.
   */
  rawColumns: Record<string, string>;
};

/** Lower-case + collapse whitespace so JSON keys are stable across slight header drift. */
function normalizeHeaderKey(raw: string): string {
  return raw.toLowerCase().replace(/\s+/g, " ").trim();
}

type FleetRecord = {
  assetId: string;
  vinSerial: string;
  makeModel: string;
  etiCLocation: string;
  mgmtCd: string;
  vehNomen: string;
};

/**
 * Read a cell, falling back to the master cell when part of a merged range.
 * Some workbooks only store the value on the master; ExcelJS returns null for
 * the other cells in the merge. This restores the expected value.
 */
function readCellWithMerge(cell: ExcelJS.Cell): string {
  const direct = readCellString(cell);
  if (direct) return direct;
  const master = (cell as unknown as { master?: ExcelJS.Cell }).master;
  if (master && master !== cell) return readCellString(master);
  return "";
}

function extractWorkOrders(sheet: ExcelJS.Worksheet): {
  rows: RawWorkOrder[];
  scan: HeaderScan<WorkOrderField> | null;
  debug: {
    sheetRowCount: number;
    rowsScanned: number;
    rowsKept: number;
    rowsWithAssetIdRaw: number;
    rowsWithWorkOrderId: number;
    rowsWithCurrentMel: number;
    rowsWithEticDue: number;
    rowsFilledFromPriorAsset: number;
    samples: Array<{ row: number } & Partial<RawWorkOrder>>;
  };
} {
  const scan = scanHeaderRow<WorkOrderField>(sheet, WORK_ORDER_SYNONYMS);
  const sheetRowCount = sheet.actualRowCount || sheet.rowCount;
  const debug = {
    sheetRowCount,
    rowsScanned: 0,
    rowsKept: 0,
    rowsWithAssetIdRaw: 0,
    rowsWithWorkOrderId: 0,
    rowsWithCurrentMel: 0,
    rowsWithEticDue: 0,
    rowsFilledFromPriorAsset: 0,
    samples: [] as Array<{ row: number } & Partial<RawWorkOrder>>,
  };
  if (!scan) return { rows: [], scan: null, debug };

  const rows: RawWorkOrder[] = [];
  const lastRow = sheetRowCount;
  let fillDownAsset = "";
  for (let r = scan.rowIndex + 1; r <= lastRow; r += 1) {
    const row = sheet.getRow(r);
    if (!row || row.cellCount === 0) continue;
    debug.rowsScanned += 1;
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
      owningUnit: "",
      melKey: "",
      mgmtCd: "",
      vehNomen: "",
      rawColumns: {},
    };
    let hasAny = false;
    scan.mapping.forEach((field, colNumber) => {
      const text = cleanText(readCellWithMerge(row.getCell(colNumber)));
      if (text) hasAny = true;
      (record as unknown as Record<string, string>)[field] = text;
    });
    // Capture EVERY header (mapped + unmapped) so we can backfill new typed
    // columns from D1 later without re-reading R2.
    for (const [colNumber, header] of scan.rawHeaders) {
      const text = cleanText(readCellWithMerge(row.getCell(colNumber)));
      if (!text) continue;
      record.rawColumns[normalizeHeaderKey(header)] = text;
      hasAny = true;
    }

    if (record.assetId) {
      debug.rowsWithAssetIdRaw += 1;
      fillDownAsset = record.assetId;
    } else if (fillDownAsset && (record.workOrderId || record.remarks || record.eticDue || record.currentMel || record.partsStatus)) {
      record.assetId = fillDownAsset;
      debug.rowsFilledFromPriorAsset += 1;
    }
    if (record.workOrderId) debug.rowsWithWorkOrderId += 1;
    if (record.currentMel) debug.rowsWithCurrentMel += 1;
    if (record.eticDue) debug.rowsWithEticDue += 1;

    const keep = hasAny && (record.workOrderId || record.assetId);
    if (keep) {
      rows.push(record);
      debug.rowsKept += 1;
      if (debug.samples.length < 10) {
        debug.samples.push({ row: r, ...record });
      }
    }
  }
  return { rows, scan, debug };
}

function extractFleet(sheet: ExcelJS.Worksheet): {
  byAsset: Map<string, FleetRecord>;
  rawByAsset: Map<string, Record<string, string>>;
  scan: HeaderScan<FleetField> | null;
} {
  const scan = scanHeaderRow<FleetField>(sheet, FLEET_SYNONYMS);
  if (!scan) return { byAsset: new Map(), rawByAsset: new Map(), scan: null };

  const byAsset = new Map<string, FleetRecord>();
  const rawByAsset = new Map<string, Record<string, string>>();
  const lastRow = sheet.actualRowCount || sheet.rowCount;
  for (let r = scan.rowIndex + 1; r <= lastRow; r += 1) {
    const row = sheet.getRow(r);
    if (!row || row.cellCount === 0) continue;
    const record: FleetRecord = {
      assetId: "",
      vinSerial: "",
      makeModel: "",
      etiCLocation: "",
      mgmtCd: "",
      vehNomen: "",
    };
    const rawCols: Record<string, string> = {};
    let hasAny = false;
    scan.mapping.forEach((field, colNumber) => {
      const text = cleanText(readCellString(row.getCell(colNumber)));
      if (text) hasAny = true;
      (record as Record<string, string>)[field] = text;
    });
    for (const [colNumber, header] of scan.rawHeaders) {
      const text = cleanText(readCellString(row.getCell(colNumber)));
      if (!text) continue;
      rawCols[`fleet.${normalizeHeaderKey(header)}`] = text;
      hasAny = true;
    }
    if (hasAny && record.assetId) {
      byAsset.set(record.assetId, record);
      rawByAsset.set(record.assetId, rawCols);
    }
  }
  return { byAsset, rawByAsset, scan };
}

/**
 * Raw WO rows from the workbook (one row per work order line). Fleet sheet is
 * also consulted to fill in `vehNomen`, plus to backfill `mgmtCd` and
 * `makeModel` when those columns are missing on the WO row.
 */
export async function extractRawWorkOrdersFromBinary(binary: ArrayBuffer): Promise<RawWorkOrder[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(binary);
  const woSheet = findWorkOrderSheet(workbook);
  if (!woSheet) return [];
  const rows = extractWorkOrders(woSheet).rows;
  const fleetSheet = findFleetSheet(workbook);
  if (!fleetSheet) return rows;
  const fleetOut = extractFleet(fleetSheet);
  const fleet = fleetOut.byAsset;
  const fleetRaw = fleetOut.rawByAsset;
  if (fleet.size === 0) return rows;
  for (const row of rows) {
    if (!row.assetId) continue;
    const f = fleet.get(row.assetId);
    if (!f) continue;
    if (!row.vehNomen) row.vehNomen = f.vehNomen ?? "";
    if (!row.mgmtCd) row.mgmtCd = f.mgmtCd ?? "";
    if (!row.makeModel) row.makeModel = f.makeModel ?? "";
    const fr = fleetRaw.get(row.assetId);
    if (fr) Object.assign(row.rawColumns, fr);
  }
  return rows;
}

/** Detailed scrape report for debugging — what headers did we see, what got mapped, what rows came back. */
export async function debugScrapeWorkbook(binary: ArrayBuffer): Promise<Record<string, unknown>> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(binary);
  const sheets = workbook.worksheets.map((s) => ({
    name: s.name,
    rowCount: s.rowCount,
    actualRowCount: s.actualRowCount,
  }));
  const woSheet = findWorkOrderSheet(workbook);
  if (!woSheet) {
    return { sheets, error: "No work order sheet found" };
  }
  const out = extractWorkOrders(woSheet);
  const scan = out.scan;
  const scanInfo = scan
    ? {
        headerRowIndex: scan.rowIndex,
        headers: [...scan.rawHeaders.entries()]
          .sort((a, b) => a[0] - b[0])
          .map(([col, raw]) => ({ col, header: raw, mappedTo: scan.mapping.get(col) ?? null })),
        unmatched: scan.unmatched,
        mappedFieldCount: scan.mapping.size,
      }
    : null;

  const fleetSheet = findFleetSheet(workbook);
  const fleetScan = fleetSheet ? scanHeaderRow<FleetField>(fleetSheet, FLEET_SYNONYMS) : null;
  const fleetInfo = fleetScan
    ? {
        sheet: fleetSheet?.name,
        headerRowIndex: fleetScan.rowIndex,
        headers: [...fleetScan.rawHeaders.entries()]
          .sort((a, b) => a[0] - b[0])
          .map(([col, raw]) => ({ col, header: raw, mappedTo: fleetScan.mapping.get(col) ?? null })),
        unmatched: fleetScan.unmatched,
        mappedFieldCount: fleetScan.mapping.size,
      }
    : { sheet: fleetSheet?.name ?? null, note: "No header row detected" };

  return {
    sheets,
    workOrderSheet: woSheet.name,
    workOrderScan: scanInfo,
    fleetScan: fleetInfo,
    extractionStats: out.debug,
    uniqueWorkOrderIds: new Set(out.rows.map((r) => r.workOrderId).filter(Boolean)).size,
    uniqueAssetIds: new Set(out.rows.map((r) => r.assetId).filter(Boolean)).size,
    totalRowsExtracted: out.rows.length,
  };
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

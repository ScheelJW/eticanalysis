import ExcelJS from "exceljs";

export type YardCheckColumn =
  | "assetId"
  | "workOrderId"
  | "workOrderRemarks"
  | "vinSerial"
  | "make"
  | "model"
  | "previousLocation";

export type YardCheckRow = {
  assetId: string;
  workOrderId: string;
  workOrderRemarks: string;
  vinSerial: string;
  make: string;
  model: string;
  previousLocation: string;
};

export type YardCheckExtractionResult = {
  sheetName: string;
  headerRowIndex: number;
  totalDataRows: number;
  rows: YardCheckRow[];
  headerMap: Record<YardCheckColumn, string | null>;
  unmatchedHeaders: string[];
};

const HEADER_SYNONYMS: Record<YardCheckColumn, string[]> = {
  assetId: [
    "asset id",
    "asset",
    "asset #",
    "asset no",
    "asset number",
    "asset num",
    "equip id",
    "equipment id",
    "equipment #",
    "equipment number",
    "vehicle id",
    "unit id",
    "reg no",
    "registration number",
  ],
  workOrderId: [
    "work order id",
    "work order",
    "wo id",
    "wo #",
    "wo no",
    "wo num",
    "wo number",
    "work order #",
    "work order no",
    "work order number",
    "work order num",
  ],
  workOrderRemarks: [
    "work order remarks",
    "remarks",
    "wo remarks",
    "comments",
    "notes",
    "description",
    "job description",
    "work description",
    "defect",
    "defect description",
  ],
  vinSerial: [
    "vin/serial",
    "vin / serial",
    "vin or serial",
    "vin-serial",
    "vin",
    "vin number",
    "serial",
    "serial number",
    "serial #",
    "serial no",
  ],
  make: ["make", "manufacturer", "mfg"],
  model: ["model"],
  previousLocation: [
    "previous location",
    "prev location",
    "last known location",
    "last location",
    "current location",
    "location",
    "parking location",
    "yard location",
    "where",
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

export function matchHeaderToColumn(header: string): YardCheckColumn | null {
  const normalized = normalizeHeader(header);
  if (!normalized) return null;

  let bestColumn: YardCheckColumn | null = null;
  let bestScore = 0;

  (Object.keys(HEADER_SYNONYMS) as YardCheckColumn[]).forEach((column) => {
    const synonyms = HEADER_SYNONYMS[column];
    for (const candidate of synonyms) {
      if (normalized === candidate) {
        if (candidate.length > bestScore) {
          bestColumn = column;
          bestScore = candidate.length + 1000;
        }
        continue;
      }
      if (normalized.includes(candidate) && candidate.length > bestScore) {
        bestColumn = column;
        bestScore = candidate.length;
      }
    }
  });

  return bestColumn;
}

function findWorkOrderSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet | null {
  const candidates = workbook.worksheets;
  const preferred = candidates.find((s) => /work\s*orders?/i.test(s.name));
  if (preferred) return preferred;
  const loose = candidates.find((s) => /wo\b/i.test(s.name));
  return loose ?? null;
}

type HeaderScan = {
  rowIndex: number;
  mapping: Map<number, YardCheckColumn>;
  rawHeaders: Map<number, string>;
  unmatched: string[];
};

function scanHeaderRow(sheet: ExcelJS.Worksheet, maxRows = 25): HeaderScan | null {
  let best: HeaderScan | null = null;

  for (let rowIndex = 1; rowIndex <= maxRows; rowIndex += 1) {
    const row = sheet.getRow(rowIndex);
    if (!row || row.cellCount === 0) continue;

    const mapping = new Map<number, YardCheckColumn>();
    const rawHeaders = new Map<number, string>();
    const unmatched: string[] = [];

    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const raw = readCellString(cell).trim();
      if (!raw) return;
      rawHeaders.set(colNumber, raw);
      const column = matchHeaderToColumn(raw);
      if (column && !mapping.has(colNumber) && !isAlreadyAssigned(mapping, column)) {
        mapping.set(colNumber, column);
      } else if (!column) {
        unmatched.push(raw);
      }
    });

    if (!mapping.has(0) && mapping.size === 0 && rawHeaders.size === 0) continue;

    const score = mapping.size;
    if (!best || score > best.mapping.size) {
      best = { rowIndex, mapping, rawHeaders, unmatched };
    }
    if (best && best.mapping.size >= 4) {
      // Good enough match; stop scanning to avoid picking deeper rows
      break;
    }
  }

  if (!best || best.mapping.size === 0) return null;
  return best;
}

function isAlreadyAssigned(mapping: Map<number, YardCheckColumn>, column: YardCheckColumn): boolean {
  for (const existing of mapping.values()) {
    if (existing === column) return true;
  }
  return false;
}

export async function extractWorkOrderRows(
  binary: ArrayBuffer,
): Promise<YardCheckExtractionResult | null> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(binary);

  const sheet = findWorkOrderSheet(workbook);
  if (!sheet) return null;

  const headerScan = scanHeaderRow(sheet);
  if (!headerScan) {
    return {
      sheetName: sheet.name,
      headerRowIndex: 1,
      totalDataRows: 0,
      rows: [],
      headerMap: {
        assetId: null,
        workOrderId: null,
        workOrderRemarks: null,
        vinSerial: null,
        make: null,
        model: null,
        previousLocation: null,
      },
      unmatchedHeaders: [],
    };
  }

  const rows: YardCheckRow[] = [];
  const lastRow = sheet.actualRowCount || sheet.rowCount;
  for (let rowIndex = headerScan.rowIndex + 1; rowIndex <= lastRow; rowIndex += 1) {
    const row = sheet.getRow(rowIndex);
    if (!row || row.cellCount === 0) continue;

    const record: YardCheckRow = {
      assetId: "",
      workOrderId: "",
      workOrderRemarks: "",
      vinSerial: "",
      make: "",
      model: "",
      previousLocation: "",
    };
    let hasAnyValue = false;

    headerScan.mapping.forEach((column, colNumber) => {
      const value = readCellString(row.getCell(colNumber)).trim();
      if (value) hasAnyValue = true;
      (record as Record<string, string>)[column] = value;
    });

    if (hasAnyValue) rows.push(record);
  }

  const headerMap: Record<YardCheckColumn, string | null> = {
    assetId: null,
    workOrderId: null,
    workOrderRemarks: null,
    vinSerial: null,
    make: null,
    model: null,
    previousLocation: null,
  };
  headerScan.mapping.forEach((column, colNumber) => {
    headerMap[column] = headerScan.rawHeaders.get(colNumber) ?? null;
  });

  return {
    sheetName: sheet.name,
    headerRowIndex: headerScan.rowIndex,
    totalDataRows: rows.length,
    rows,
    headerMap,
    unmatchedHeaders: headerScan.unmatched,
  };
}

export type YardCheckMeta = {
  sourceWorkbookFileName: string;
  sourceDateKey: string;
  generatedAtIso: string;
};

export async function generateYardCheckWorkbookBuffer(
  extraction: YardCheckExtractionResult,
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
    { header: "Asset ID", key: "assetId", width: 12 },
    { header: "Work Order ID", key: "workOrderId", width: 14 },
    { header: "VIN / Serial", key: "vinSerial", width: 20 },
    { header: "Make", key: "make", width: 12 },
    { header: "Model", key: "model", width: 16 },
    { header: "Previous Location", key: "previousLocation", width: 18 },
    { header: "New Location", key: "newLocation", width: 22 },
    { header: "Discrepancies", key: "discrepancies", width: 30 },
    { header: "Work Order Remarks", key: "workOrderRemarks", width: 40 },
  ];

  const headerRow = sheet.getRow(1);
  headerRow.height = 24;
  headerRow.font = { bold: true, size: 11 };
  headerRow.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  headerRow.eachCell((cell) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFEEF2F8" },
    };
    cell.border = {
      top: { style: "thin", color: { argb: "FF9FB1C8" } },
      left: { style: "thin", color: { argb: "FF9FB1C8" } },
      bottom: { style: "thin", color: { argb: "FF9FB1C8" } },
      right: { style: "thin", color: { argb: "FF9FB1C8" } },
    };
  });

  for (const record of extraction.rows) {
    const added = sheet.addRow({
      assetId: record.assetId,
      workOrderId: record.workOrderId,
      vinSerial: record.vinSerial,
      make: record.make,
      model: record.model,
      previousLocation: record.previousLocation,
      newLocation: "",
      discrepancies: "",
      workOrderRemarks: record.workOrderRemarks,
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

  if (extraction.rows.length === 0) {
    const empty = sheet.addRow({ assetId: "No work orders found in source workbook." });
    empty.font = { italic: true };
  }

  const bufferLike = await workbook.xlsx.writeBuffer();
  return toArrayBuffer(bufferLike);
}

function toArrayBuffer(buffer: ExcelJS.Buffer): ArrayBuffer {
  if (buffer instanceof ArrayBuffer) return buffer;
  if (ArrayBuffer.isView(buffer)) {
    const view = buffer as ArrayBufferView;
    const clone = new Uint8Array(view.byteLength);
    clone.set(new Uint8Array(view.buffer, view.byteOffset, view.byteLength));
    return clone.buffer;
  }
  const anyBuf = buffer as unknown as { buffer?: ArrayBuffer; byteOffset?: number; byteLength?: number };
  if (anyBuf.buffer && typeof anyBuf.byteOffset === "number" && typeof anyBuf.byteLength === "number") {
    const clone = new Uint8Array(anyBuf.byteLength);
    clone.set(new Uint8Array(anyBuf.buffer, anyBuf.byteOffset, anyBuf.byteLength));
    return clone.buffer;
  }
  throw new Error("Unsupported buffer type returned by ExcelJS");
}

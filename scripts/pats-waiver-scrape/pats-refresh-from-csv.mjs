/**
 * Refresh already-imported PATS waivers in D1 from an existing CSV — no scrape.
 *
 * Reads the same format as scrape.mjs (pats_waivers.csv): reg_no, waiver_row_id,
 * waiver_text, status, etc. Emits UPDATE statements keyed by import_key PATS:<id>.
 *
 * Updates for each Outstanding row:
 *   asset_id  (AF + reg_no, same rules as pats-to-d1-sql.mjs)
 *   title     (truncated waiver text)
 *   description = waiver_text only (no boilerplate)
 * Plus one statement:
 *   reviewed_note = 'Imported from NEI PATS.' for all import_key LIKE 'PATS:%'
 *
 * Usage:
 *   node pats-refresh-from-csv.mjs path/to/pats_waivers.csv path/to/refresh.sql
 *
 * Apply:
 *   npx wrangler d1 execute etic-snapshots --remote --file=path/to/refresh.sql
 */

import fs from "node:fs";
import { isNoiseWaiverText } from "./is-noise-waiver.mjs";

const DEFAULT_CARD_STATUS = "Outstanding";

function sqlStr(s) {
  return `'${String(s ?? "").replace(/'/g, "''")}'`;
}

function assetIdFromReg(regNo) {
  const r = (regNo ?? "").trim().replace(/\s+/g, "");
  if (!r) return "";
  if (/^AF/i.test(r)) return r;
  return `AF${r}`;
}

function parseCsvLine(line) {
  const out = [];
  let cur = "";
  let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQ) {
      if (c === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i++;
        } else inQ = false;
      } else cur += c;
    } else {
      if (c === ",") {
        out.push(cur);
        cur = "";
      } else if (c === '"') inQ = true;
      else cur += c;
    }
  }
  out.push(cur);
  return out;
}

function main() {
  const inPath = process.argv[2] || "pats_waivers.csv";
  const outPath = process.argv[3] || "pats_waiver_refresh.sql";
  const body = fs.readFileSync(inPath, "utf8");
  const lines = body.split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) {
    console.error("CSV needs a header and at least one row");
    process.exit(1);
  }
  const header = parseCsvLine(lines[0]).map((h) => h.trim().toLowerCase());
  const idx = (name) => header.indexOf(name);

  const iReg = idx("reg_no");
  const iRow = idx("waiver_row_id");
  const iText = idx("waiver_text");
  const iStatus = idx("status");
  if (iReg < 0 || iRow < 0 || iText < 0) {
    console.error("CSV must include columns: reg_no, waiver_row_id, waiver_text");
    process.exit(1);
  }

  const chunks = [
    "-- Refresh PATS-backed waivers from CSV (no scrape). import_key PATS:* must already exist.\n",
    "-- (D1 execute disallows BEGIN TRANSACTION in batch SQL; statements run in one batch.)\n",
    "UPDATE waiver SET reviewed_note = 'Imported from NEI PATS.' WHERE import_key LIKE 'PATS:%';\n",
  ];

  let n = 0;
  let skipped = 0;
  for (let li = 1; li < lines.length; li++) {
    const cols = parseCsvLine(lines[li]);
    const get = (i) => (i >= 0 ? (cols[i] ?? "").trim() : "");

    const regNo = get(iReg);
    const waiverRowId = get(iRow);
    const waiverText = get(iText);
    const legacyStatus = get(iStatus) || DEFAULT_CARD_STATUS;

    if (!waiverRowId || !waiverText) continue;
    if (isNoiseWaiverText(waiverText)) continue;
    if (legacyStatus.trim().toLowerCase() !== "outstanding") {
      skipped++;
      continue;
    }

    const importKey = `PATS:${waiverRowId}`;
    const assetId = assetIdFromReg(regNo);
    if (!assetId) continue;

    const title =
      waiverText.length <= 120 ? waiverText : `${waiverText.slice(0, 117)}...`;

    chunks.push(
      `UPDATE waiver SET asset_id = ${sqlStr(assetId)}, title = ${sqlStr(title)}, description = ${sqlStr(waiverText)} WHERE import_key = ${sqlStr(importKey)};\n`,
    );
    n++;
  }

  fs.writeFileSync(outPath, chunks.join(""), "utf8");
  console.error(
    `Wrote ${n} UPDATE statements to ${outPath}. Skipped ${skipped} non-Outstanding CSV rows.`,
  );
}

main();

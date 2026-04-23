import ExcelJS from "exceljs";

// One row per MEL key as it exists in the workbook's MEL Calculator sheet.
// MEL = the minimum number of mission-capable assets a unit needs to fulfill
// its mission. NMC + FMC = currently assigned. recall_delta is the +/- of
// loaner ("recall") vehicles brought in as a band-aid.
export type MelRow = {
  melKey: string;
  unit: string;
  userUnit: string;
  priorityTier: string;
  mgmtCodeName: string;
  detailDocNumber: string;
  melAssignedTotal: number;
  nmcCount: number;
  fmcCount: number;
  accAbus: number;
  melRequired: number;
  recallDelta: number;
  melDelta: number;
  /** Normalized: "below" | "at" | "above" | "unknown". */
  melStatus: MelStatus;
  /** Original cell text e.g. "BELOW MEL". */
  rawStatus: string;
};

export type MelStatus = "below" | "at" | "above" | "unknown";

export type MelStateRow = {
  mel_key: string;
  last_snapshot_date: string;
  unit: string;
  user_unit: string;
  priority_tier: string;
  mgmt_code_name: string;
  detail_doc_number: string;
  mel_assigned_total: number;
  nmc_count: number;
  fmc_count: number;
  acc_abus: number;
  mel_required: number;
  recall_delta: number;
  mel_delta: number;
  mel_status: string;
  raw_status: string;
  updated_at_iso: string;
};

export type MelHistoryPoint = {
  snapshot_date_key: string;
  mel_key: string;
  unit: string;
  user_unit: string;
  priority_tier: string;
  mgmt_code_name: string;
  detail_doc_number: string;
  mel_assigned_total: number;
  nmc_count: number;
  fmc_count: number;
  acc_abus: number;
  mel_required: number;
  recall_delta: number;
  mel_delta: number;
  mel_status: string;
  raw_status: string;
};

export type MelChangelogRow = {
  id: number;
  mel_key: string;
  snapshot_date_key: string;
  field: string;
  old_value: string | null;
  new_value: string | null;
  created_at_iso: string;
};

type Env = { ETIC_SNAPSHOTS: D1Database };

function readCellText(cell: ExcelJS.Cell): string {
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return String(v);
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "object") {
    const o = v as { text?: string; result?: unknown; richText?: Array<{ text: string }> };
    if (typeof o.text === "string") return o.text;
    if (Array.isArray(o.richText)) return o.richText.map((p) => p.text).join("");
    if (o.result !== undefined && o.result !== null) return String(o.result);
  }
  return "";
}

function toInt(s: string): number {
  if (!s) return 0;
  const cleaned = s.replace(/[, ]/g, "").trim();
  // "AT MEL" / "BELOW MEL" / "ABOVE MEL" sometimes appear in MEL +/- — those
  // aren't numbers, so just return 0 for those.
  const n = Number.parseInt(cleaned, 10);
  return Number.isFinite(n) ? n : 0;
}

function classifyStatus(raw: string): MelStatus {
  const s = raw.replace(/\s+/g, " ").trim().toLowerCase();
  if (!s) return "unknown";
  if (s.includes("below")) return "below";
  if (s.includes("above")) return "above";
  if (s.includes("at mel") || s === "at" || /\bat\b/.test(s)) return "at";
  return "unknown";
}

/** Pulls one MelRow per non-empty row out of the MEL Calculator sheet. */
export async function extractMelRowsFromBinary(bytes: ArrayBuffer): Promise<MelRow[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const sheet = wb.worksheets.find((ws) => /mel\s*calc/i.test(ws.name));
  if (!sheet) return [];

  // Find the header row (first row with >= 4 non-empty cells).
  let headerRow = -1;
  for (let r = 1; r <= 12; r++) {
    let count = 0;
    sheet.getRow(r).eachCell({ includeEmpty: false }, () => { count += 1; });
    if (count >= 4) { headerRow = r; break; }
  }
  if (headerRow < 1) return [];

  // Map header text → column index (case-insensitive, whitespace-collapsed).
  const norm = (s: string) => s.replace(/\s+/g, " ").trim().toLowerCase();
  const colByHeader = new Map<string, number>();
  sheet.getRow(headerRow).eachCell({ includeEmpty: false }, (cell, col) => {
    colByHeader.set(norm(readCellText(cell)), col);
  });
  const colOf = (...candidates: string[]): number => {
    for (const c of candidates) {
      const k = norm(c);
      if (colByHeader.has(k)) return colByHeader.get(k) as number;
    }
    // Fuzzy fallback: substring match.
    for (const [hdr, col] of colByHeader) {
      for (const c of candidates) {
        if (hdr.includes(norm(c))) return col;
      }
    }
    return -1;
  };

  const cMelKey = colOf("mel key");
  const cAssigned = colOf("mel assigned total", "assigned total");
  const cNmc = colOf("nmc count");
  const cFmc = colOf("fmc count");
  const cAcc = colOf("acc/abus", "acc abus");
  const cMel = colOf("mel");
  const cTier = colOf("prioritytier", "priority tier");
  const cUserUnit = colOf("user/unit", "user unit");
  const cUnit = colOf("unit");
  const cDoc = colOf("detail doc number");
  const cMgmt = colOf("afa4_mgmtcodename__c.1", "mgmt code name", "mgmtcodename");
  const cRecall = colOf("recall +/-", "recall");
  const cMelDelta = colOf("mel +/-");
  const cStatus = colOf("mel status", "status");

  if (cMelKey < 1 || cAssigned < 1 || cNmc < 1 || cFmc < 1) return [];

  const rows: MelRow[] = [];
  const last = sheet.actualRowCount || sheet.rowCount || 0;
  const seen = new Set<string>();
  for (let r = headerRow + 1; r <= last; r++) {
    const row = sheet.getRow(r);
    const melKeyRaw = readCellText(row.getCell(cMelKey)).trim();
    if (!melKeyRaw) continue;
    if (seen.has(melKeyRaw)) continue; // shouldn't happen, but be safe
    seen.add(melKeyRaw);
    const rawStatus = cStatus > 0 ? readCellText(row.getCell(cStatus)).trim() : "";
    rows.push({
      melKey: melKeyRaw,
      unit: cUnit > 0 ? readCellText(row.getCell(cUnit)).trim() : "",
      userUnit: cUserUnit > 0 ? readCellText(row.getCell(cUserUnit)).trim() : "",
      priorityTier: cTier > 0 ? readCellText(row.getCell(cTier)).trim() : "",
      mgmtCodeName: cMgmt > 0 ? readCellText(row.getCell(cMgmt)).trim() : "",
      detailDocNumber: cDoc > 0 ? readCellText(row.getCell(cDoc)).trim() : "",
      melAssignedTotal: toInt(readCellText(row.getCell(cAssigned))),
      nmcCount: toInt(readCellText(row.getCell(cNmc))),
      fmcCount: toInt(readCellText(row.getCell(cFmc))),
      accAbus: cAcc > 0 ? toInt(readCellText(row.getCell(cAcc))) : 0,
      melRequired: cMel > 0 ? toInt(readCellText(row.getCell(cMel))) : 0,
      recallDelta: cRecall > 0 ? toInt(readCellText(row.getCell(cRecall))) : 0,
      melDelta: cMelDelta > 0 ? toInt(readCellText(row.getCell(cMelDelta))) : 0,
      melStatus: classifyStatus(rawStatus),
      rawStatus,
    });
  }
  return rows;
}

/** Insert one MEL snapshot for a date, replace existing snapshot for that
 *  date if any, update mel_state, and emit per-key changelog entries vs. the
 *  prior snapshot for that key (including same-day re-ingest vs the earlier
 *  file for that date). Idempotent. */
export async function ingestMelSnapshot(
  env: Env,
  dateKey: string,
  rows: MelRow[],
  nowIso: string,
): Promise<{ rows: number; changes: number }> {
  if (!rows.length) return { rows: 0, changes: 0 };

  // Pull the most recent prior snapshot per MEL key so we can diff. We fetch
  // all of mel_state once (whose row is "the latest we've ever seen") and
  // only treat it as a baseline when its date < dateKey.
  const prior = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM mel_state`,
  ).all<MelStateRow>();
  const priorByKey = new Map<string, MelStateRow>();
  for (const p of prior.results ?? []) priorByKey.set(p.mel_key, p);

  // Same-day re-ingest: keep this date's changelog rows from the first pass,
  // then diff the new workbook against the prior MEL snapshot rows we are
  // about to replace (morning vs afternoon).
  const priorSnap = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM mel_snapshot WHERE snapshot_date_key = ?`,
  )
    .bind(dateKey)
    .all<MelHistoryPoint>();
  const priorSnapByKey = new Map<string, MelHistoryPoint>();
  for (const row of priorSnap.results ?? []) priorSnapByKey.set(row.mel_key, row);
  const sameDayReingest = priorSnapByKey.size > 0;

  if (!sameDayReingest) {
    await env.ETIC_SNAPSHOTS.prepare(
      `DELETE FROM mel_changelog WHERE snapshot_date_key = ?`,
    )
      .bind(dateKey)
      .run();
  }

  await env.ETIC_SNAPSHOTS.prepare(
    `DELETE FROM mel_snapshot WHERE snapshot_date_key = ?`,
  ).bind(dateKey).run();

  // Insert mel_snapshot rows in batches.
  const insertSnap = env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO mel_snapshot
       (snapshot_date_key, mel_key, unit, user_unit, priority_tier, mgmt_code_name,
        detail_doc_number, mel_assigned_total, nmc_count, fmc_count, acc_abus,
        mel_required, recall_delta, mel_delta, mel_status, raw_status)
     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
  );
  const batchSize = 50;
  for (let i = 0; i < rows.length; i += batchSize) {
    const slice = rows.slice(i, i + batchSize);
    await env.ETIC_SNAPSHOTS.batch(
      slice.map((r) =>
        insertSnap.bind(
          dateKey, r.melKey, r.unit, r.userUnit, r.priorityTier, r.mgmtCodeName,
          r.detailDocNumber, r.melAssignedTotal, r.nmcCount, r.fmcCount, r.accAbus,
          r.melRequired, r.recallDelta, r.melDelta, r.melStatus, r.rawStatus,
        ),
      ),
    );
  }

  // Diff vs. the prior baseline (only if prior date < this date) and emit
  // changelog entries for the fields that matter.
  const tracked: Array<[keyof MelStateRow, keyof MelRow, string]> = [
    ["nmc_count", "nmcCount", "NMC count"],
    ["fmc_count", "fmcCount", "FMC count"],
    ["mel_required", "melRequired", "MEL required"],
    ["mel_assigned_total", "melAssignedTotal", "Assigned total"],
    ["mel_status", "melStatus", "MEL status"],
    ["recall_delta", "recallDelta", "Recall +/-"],
    ["unit", "unit", "Unit"],
    ["user_unit", "userUnit", "User unit"],
    ["detail_doc_number", "detailDocNumber", "Detail doc number"],
    ["priority_tier", "priorityTier", "Priority tier"],
  ];

  const changeStmts: D1PreparedStatement[] = [];
  // INSERT OR IGNORE pairs with the UNIQUE(mel_key, snapshot_date_key, field) index
  // added in migration 0017. Same rationale as the WO changelog: re-ingest of a
  // snapshot is now idempotent instead of producing a duplicate row per field per
  // day. We also keep ON CONFLICT semantics simple — first observation wins.
  const insertChange = env.ETIC_SNAPSHOTS.prepare(
    `INSERT OR REPLACE INTO mel_changelog
       (mel_key, snapshot_date_key, field, old_value, new_value, created_at_iso)
     VALUES (?,?,?,?,?,?)`,
  );
  let changes = 0;
  for (const r of rows) {
    const snapPrior = priorSnapByKey.get(r.melKey);
    if (snapPrior) {
      for (const [pCol, rField, label] of tracked) {
        const oldV = String((snapPrior as Record<string, unknown>)[pCol] ?? "");
        const newV = String((r as Record<string, unknown>)[rField] ?? "");
        if (oldV !== newV) {
          changeStmts.push(insertChange.bind(r.melKey, dateKey, label, oldV, newV, nowIso));
          changes += 1;
        }
      }
      continue;
    }

    const p = priorByKey.get(r.melKey);
    if (!p) {
      // First time we've seen this MEL key.
      changeStmts.push(
        insertChange.bind(r.melKey, dateKey, "initial", null, r.melStatus, nowIso),
      );
      changes += 1;
      continue;
    }
    if (p.last_snapshot_date >= dateKey) continue;
    for (const [pCol, rField, label] of tracked) {
      const oldV = String((p as Record<string, unknown>)[pCol] ?? "");
      const newV = String((r as Record<string, unknown>)[rField] ?? "");
      if (oldV !== newV) {
        changeStmts.push(insertChange.bind(r.melKey, dateKey, label, oldV, newV, nowIso));
        changes += 1;
      }
    }
  }
  if (changeStmts.length) {
    for (let i = 0; i < changeStmts.length; i += batchSize) {
      await env.ETIC_SNAPSHOTS.batch(changeStmts.slice(i, i + batchSize));
    }
  }

  // Upsert mel_state — last_snapshot_date wins-by-date.
  const upsertState = env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO mel_state
       (mel_key, last_snapshot_date, unit, user_unit, priority_tier, mgmt_code_name,
        detail_doc_number, mel_assigned_total, nmc_count, fmc_count, acc_abus,
        mel_required, recall_delta, mel_delta, mel_status, raw_status, updated_at_iso)
     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
     ON CONFLICT(mel_key) DO UPDATE SET
       last_snapshot_date = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                                 THEN excluded.last_snapshot_date ELSE mel_state.last_snapshot_date END,
       unit = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                   THEN excluded.unit ELSE mel_state.unit END,
       user_unit = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                        THEN excluded.user_unit ELSE mel_state.user_unit END,
       priority_tier = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                            THEN excluded.priority_tier ELSE mel_state.priority_tier END,
       mgmt_code_name = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                             THEN excluded.mgmt_code_name ELSE mel_state.mgmt_code_name END,
       detail_doc_number = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                                THEN excluded.detail_doc_number ELSE mel_state.detail_doc_number END,
       mel_assigned_total = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                                 THEN excluded.mel_assigned_total ELSE mel_state.mel_assigned_total END,
       nmc_count = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                        THEN excluded.nmc_count ELSE mel_state.nmc_count END,
       fmc_count = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                        THEN excluded.fmc_count ELSE mel_state.fmc_count END,
       acc_abus = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                       THEN excluded.acc_abus ELSE mel_state.acc_abus END,
       mel_required = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                           THEN excluded.mel_required ELSE mel_state.mel_required END,
       recall_delta = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                           THEN excluded.recall_delta ELSE mel_state.recall_delta END,
       mel_delta = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                        THEN excluded.mel_delta ELSE mel_state.mel_delta END,
       mel_status = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                         THEN excluded.mel_status ELSE mel_state.mel_status END,
       raw_status = CASE WHEN excluded.last_snapshot_date >= mel_state.last_snapshot_date
                         THEN excluded.raw_status ELSE mel_state.raw_status END,
       updated_at_iso = excluded.updated_at_iso`,
  );
  for (let i = 0; i < rows.length; i += batchSize) {
    const slice = rows.slice(i, i + batchSize);
    await env.ETIC_SNAPSHOTS.batch(
      slice.map((r) =>
        upsertState.bind(
          r.melKey, dateKey, r.unit, r.userUnit, r.priorityTier, r.mgmtCodeName,
          r.detailDocNumber, r.melAssignedTotal, r.nmcCount, r.fmcCount, r.accAbus,
          r.melRequired, r.recallDelta, r.melDelta, r.melStatus, r.rawStatus, nowIso,
        ),
      ),
    );
  }

  return { rows: rows.length, changes };
}

/** Latest state for every MEL key (one row each). */
export async function getMelLatest(env: Env): Promise<MelStateRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM mel_state ORDER BY mel_status, unit, mel_key`,
  ).all<MelStateRow>();
  return r.results ?? [];
}

/** Snapshot for an arbitrary date. Returns mel_snapshot rows. */
export async function getMelForDate(env: Env, dateKey: string): Promise<MelHistoryPoint[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM mel_snapshot WHERE snapshot_date_key = ? ORDER BY mel_status, unit, mel_key`,
  ).bind(dateKey).all<MelHistoryPoint>();
  return r.results ?? [];
}

/** Full per-MEL-key time series. */
export async function getMelHistoryForKey(
  env: Env,
  melKey: string,
): Promise<MelHistoryPoint[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM mel_snapshot WHERE mel_key = ? ORDER BY snapshot_date_key ASC`,
  ).bind(melKey).all<MelHistoryPoint>();
  return r.results ?? [];
}

/** Per-MEL-key changelog. */
export async function getMelChangelogForKey(
  env: Env,
  melKey: string,
): Promise<MelChangelogRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM mel_changelog WHERE mel_key = ? ORDER BY snapshot_date_key DESC, id DESC`,
  ).bind(melKey).all<MelChangelogRow>();
  return r.results ?? [];
}

/** Distinct snapshot dates we have MEL data for, newest first. */
export async function getMelSnapshotDates(env: Env): Promise<string[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT DISTINCT snapshot_date_key FROM mel_snapshot ORDER BY snapshot_date_key DESC`,
  ).all<{ snapshot_date_key: string }>();
  return (r.results ?? []).map((x) => x.snapshot_date_key);
}

/* -------------------- mel_config (editable per-key UI metadata) -------------------- */

export type MelConfigRow = {
  mel_key: string;
  is_critical: number; // 0 / 1
  type_label: string;
  display_order: number;
  notes: string;
  updated_at_iso: string;
};

export type MelSubdivisionRow = {
  id: number;
  mel_key: string;
  type_label: string;
  mgmt_code: string;
  mel_required: number;
  assigned: number;
  fmc: number;
  nmc: number;
  accidents: number;
  abuses: number;
  display_order: number;
  updated_at_iso: string;
};

export async function getAllMelConfig(env: Env): Promise<MelConfigRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM mel_config`).all<MelConfigRow>();
  return r.results ?? [];
}

export async function getMelConfig(env: Env, melKey: string): Promise<MelConfigRow | null> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM mel_config WHERE mel_key = ?`,
  ).bind(melKey).first<MelConfigRow>();
  return r ?? null;
}

export type MelConfigPatch = {
  isCritical?: boolean;
  typeLabel?: string;
  displayOrder?: number;
  notes?: string;
};

export async function upsertMelConfig(
  env: Env,
  melKey: string,
  patch: MelConfigPatch,
  nowIso: string,
): Promise<MelConfigRow> {
  const existing = await getMelConfig(env, melKey);
  const merged = {
    is_critical: patch.isCritical != null ? (patch.isCritical ? 1 : 0) : existing?.is_critical ?? 0,
    type_label: patch.typeLabel != null ? patch.typeLabel : existing?.type_label ?? "",
    display_order: patch.displayOrder != null ? patch.displayOrder : existing?.display_order ?? 0,
    notes: patch.notes != null ? patch.notes : existing?.notes ?? "",
  };
  await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO mel_config (mel_key, is_critical, type_label, display_order, notes, updated_at_iso)
     VALUES (?,?,?,?,?,?)
     ON CONFLICT(mel_key) DO UPDATE SET
       is_critical = excluded.is_critical,
       type_label = excluded.type_label,
       display_order = excluded.display_order,
       notes = excluded.notes,
       updated_at_iso = excluded.updated_at_iso`,
  ).bind(melKey, merged.is_critical, merged.type_label, merged.display_order, merged.notes, nowIso).run();
  const out = await getMelConfig(env, melKey);
  return out as MelConfigRow;
}

export async function deleteMelConfig(env: Env, melKey: string): Promise<void> {
  await env.ETIC_SNAPSHOTS.prepare(`DELETE FROM mel_config WHERE mel_key = ?`).bind(melKey).run();
}

/* -------------------- mel_subdivision -------------------- */

export async function getMelSubdivisions(env: Env, melKey?: string): Promise<MelSubdivisionRow[]> {
  if (melKey) {
    const r = await env.ETIC_SNAPSHOTS.prepare(
      `SELECT * FROM mel_subdivision WHERE mel_key = ? ORDER BY display_order, id`,
    ).bind(melKey).all<MelSubdivisionRow>();
    return r.results ?? [];
  }
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM mel_subdivision ORDER BY mel_key, display_order, id`,
  ).all<MelSubdivisionRow>();
  return r.results ?? [];
}

export type MelSubdivisionInput = {
  id?: number;
  melKey: string;
  typeLabel: string;
  mgmtCode?: string;
  melRequired?: number;
  assigned?: number;
  fmc?: number;
  nmc?: number;
  accidents?: number;
  abuses?: number;
  displayOrder?: number;
};

export async function upsertMelSubdivision(
  env: Env,
  input: MelSubdivisionInput,
  nowIso: string,
): Promise<MelSubdivisionRow> {
  const v = {
    mgmt_code: input.mgmtCode ?? "",
    mel_required: input.melRequired ?? 0,
    assigned: input.assigned ?? 0,
    fmc: input.fmc ?? 0,
    nmc: input.nmc ?? 0,
    accidents: input.accidents ?? 0,
    abuses: input.abuses ?? 0,
    display_order: input.displayOrder ?? 0,
  };
  if (input.id) {
    await env.ETIC_SNAPSHOTS.prepare(
      `UPDATE mel_subdivision SET
         mel_key = ?, type_label = ?, mgmt_code = ?, mel_required = ?, assigned = ?,
         fmc = ?, nmc = ?, accidents = ?, abuses = ?, display_order = ?, updated_at_iso = ?
       WHERE id = ?`,
    ).bind(
      input.melKey, input.typeLabel, v.mgmt_code, v.mel_required, v.assigned,
      v.fmc, v.nmc, v.accidents, v.abuses, v.display_order, nowIso, input.id,
    ).run();
    const r = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM mel_subdivision WHERE id = ?`)
      .bind(input.id).first<MelSubdivisionRow>();
    return r as MelSubdivisionRow;
  }
  await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO mel_subdivision
       (mel_key, type_label, mgmt_code, mel_required, assigned, fmc, nmc, accidents, abuses, display_order, updated_at_iso)
     VALUES (?,?,?,?,?,?,?,?,?,?,?)
     ON CONFLICT(mel_key, type_label) DO UPDATE SET
       mgmt_code = excluded.mgmt_code,
       mel_required = excluded.mel_required,
       assigned = excluded.assigned,
       fmc = excluded.fmc,
       nmc = excluded.nmc,
       accidents = excluded.accidents,
       abuses = excluded.abuses,
       display_order = excluded.display_order,
       updated_at_iso = excluded.updated_at_iso`,
  ).bind(
    input.melKey, input.typeLabel, v.mgmt_code, v.mel_required, v.assigned,
    v.fmc, v.nmc, v.accidents, v.abuses, v.display_order, nowIso,
  ).run();
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM mel_subdivision WHERE mel_key = ? AND type_label = ?`,
  ).bind(input.melKey, input.typeLabel).first<MelSubdivisionRow>();
  return r as MelSubdivisionRow;
}

export async function deleteMelSubdivision(env: Env, id: number): Promise<void> {
  await env.ETIC_SNAPSHOTS.prepare(`DELETE FROM mel_subdivision WHERE id = ?`).bind(id).run();
}

/* -------------------- app_config (k/v JSON) -------------------- */

export async function getAppConfig<T = unknown>(env: Env, key: string): Promise<T | null> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT value FROM app_config WHERE key = ?`,
  ).bind(key).first<{ value: string }>();
  if (!r?.value) return null;
  try { return JSON.parse(r.value) as T; } catch { return null; }
}

export async function setAppConfig(env: Env, key: string, value: unknown, nowIso: string): Promise<void> {
  await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO app_config (key, value, updated_at_iso) VALUES (?,?,?)
     ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at_iso = excluded.updated_at_iso`,
  ).bind(key, JSON.stringify(value), nowIso).run();
}

export async function getAllAppConfig(env: Env): Promise<Record<string, unknown>> {
  const r = await env.ETIC_SNAPSHOTS.prepare(`SELECT key, value FROM app_config`)
    .all<{ key: string; value: string }>();
  const out: Record<string, unknown> = {};
  for (const row of r.results ?? []) {
    try { out[row.key] = JSON.parse(row.value); } catch { out[row.key] = row.value; }
  }
  return out;
}

/** Default staleness thresholds (days between remarks). Below = below MEL,
 *  At = at MEL, Above = above MEL. UI exposes these in Settings. */
export const DEFAULT_STALENESS: Record<string, number> = {
  below: 3,
  at: 5,
  above: 10,
};

export async function getStalenessThresholds(env: Env): Promise<Record<string, number>> {
  const v = await getAppConfig<Record<string, number>>(env, "staleness_thresholds");
  return { ...DEFAULT_STALENESS, ...(v ?? {}) };
}

/** Daily roll-up across every MEL key: how many are below/at/above, total NMC,
 *  total FMC. One row per snapshot date for charting. */
export type MelRollupRow = {
  snapshot_date_key: string;
  total_keys: number;
  below: number;
  at: number;
  above: number;
  total_nmc: number;
  total_fmc: number;
  total_recall: number;
};
export async function getMelRollup(env: Env): Promise<MelRollupRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT snapshot_date_key,
            COUNT(*) AS total_keys,
            SUM(CASE WHEN mel_status='below' THEN 1 ELSE 0 END) AS below,
            SUM(CASE WHEN mel_status='at'    THEN 1 ELSE 0 END) AS at,
            SUM(CASE WHEN mel_status='above' THEN 1 ELSE 0 END) AS above,
            SUM(nmc_count) AS total_nmc,
            SUM(fmc_count) AS total_fmc,
            SUM(recall_delta) AS total_recall
       FROM mel_snapshot
       GROUP BY snapshot_date_key
       ORDER BY snapshot_date_key ASC`,
  ).all<MelRollupRow>();
  return r.results ?? [];
}

/** Authorization manager: latest MEL Calculator rows + unit / detail-doc move log (not work orders). */
export type AuthzManagerRow = {
  melKey: string;
  unit: string;
  userUnit: string;
  detailDocNumber: string;
  mgmtCodeName: string;
  priorityTier: string;
  melStatus: string;
};

export type AuthzManagerMoveRow = {
  id: number;
  melKey: string;
  snapshotDateKey: string;
  field: string;
  oldValue: string;
  newValue: string;
  createdAtIso: string;
};

export async function getAuthzManagerData(env: Env): Promise<{
  asOfDateKey: string;
  assets: AuthzManagerRow[];
  moves: AuthzManagerMoveRow[];
}> {
  const dateRow = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT MAX(snapshot_date_key) AS k FROM mel_snapshot`,
  ).first<{ k: string | null }>();
  const asOfDateKey = (dateRow?.k ?? "").trim();
  if (!asOfDateKey) {
    return { asOfDateKey: "", assets: [], moves: [] };
  }

  const snap = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT mel_key, unit, user_unit, detail_doc_number, mgmt_code_name, priority_tier, mel_status
       FROM mel_snapshot
      WHERE snapshot_date_key = ?
   ORDER BY mel_status, unit, mel_key`,
  )
    .bind(asOfDateKey)
    .all<{
      mel_key: string;
      unit: string;
      user_unit: string;
      detail_doc_number: string;
      mgmt_code_name: string;
      priority_tier: string;
      mel_status: string;
    }>();

  const assets: AuthzManagerRow[] = (snap.results ?? []).map((r) => ({
    melKey: (r.mel_key ?? "").trim(),
    unit: (r.unit ?? "").trim(),
    userUnit: (r.user_unit ?? "").trim(),
    detailDocNumber: (r.detail_doc_number ?? "").trim(),
    mgmtCodeName: (r.mgmt_code_name ?? "").trim(),
    priorityTier: (r.priority_tier ?? "").trim(),
    melStatus: (r.mel_status ?? "").trim(),
  }));

  const AUTHZ_MOVE_FIELDS = ["Unit", "User unit", "Detail doc number"];
  const placeholders = AUTHZ_MOVE_FIELDS.map(() => "?").join(",");
  const movesR = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, mel_key, snapshot_date_key, field, old_value, new_value, created_at_iso
       FROM mel_changelog
      WHERE field IN (${placeholders})
   ORDER BY id DESC
      LIMIT 500`,
  )
    .bind(...AUTHZ_MOVE_FIELDS)
    .all<{
      id: number;
      mel_key: string;
      snapshot_date_key: string;
      field: string;
      old_value: string | null;
      new_value: string | null;
      created_at_iso: string;
    }>();

  const moves: AuthzManagerMoveRow[] = (movesR.results ?? []).map((r) => ({
    id: r.id,
    melKey: (r.mel_key ?? "").trim(),
    snapshotDateKey: r.snapshot_date_key,
    field: r.field,
    oldValue: r.old_value ?? "",
    newValue: r.new_value ?? "",
    createdAtIso: r.created_at_iso,
  }));

  return { asOfDateKey, assets, moves };
}

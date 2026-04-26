// Yard check sessions: a fleet manager walks the lot with their phone and
// taps each vehicle's actual location plus any new discrepancies. Backed by
// D1 tables created in migrations/0014_yard_check.sql.
//
// The "asset roster" comes from a specific ETIC snapshot date. Rolling checks
// also persist the active asset row JSON so later ingests cannot change what
// the walker saw when they marked that vehicle present.

type Env = { ETIC_SNAPSHOTS: D1Database; ETIC_BUCKET: R2Bucket };

export type YardSessionStatus = "open" | "closed";

export type YardSessionRow = {
  id: number;
  name: string;
  createdBy: string;
  createdAtIso: string;
  closedAtIso: string | null;
  sourceDateKey: string;
  notes: string;
  status: YardSessionStatus;
};

export type YardEntryStatus = "present" | "missing" | "unknown" | "not_applicable";

export type YardEntryRow = {
  id: number;
  sessionId: number;
  assetId: string;
  location: string;
  discrepancies: string;
  status: YardEntryStatus;
  enteredBy: string;
  enteredAtIso: string;
  updatedAtIso: string;
};

export type YardAsset = {
  assetId: string;
  owningUnit: string;
  shop: string;
  mgmtCd: string;
  makeModel: string;
  vehNomen: string;
  melKey: string;
  melTier: string;
  vinSerial: string;
  // Best-known prior location across all open WOs for this asset on the source
  // snapshot. Helpful default text for the walker.
  previousLocation: string;
  openWoCount: number;
  isNce: boolean;
};

export type YardRoster = {
  assets: YardAsset[];
  /** Distinct non-empty location strings seen across the snapshot, sorted. */
  locations: string[];
};

export type YardSessionDetail = {
  session: YardSessionRow;
  entries: YardEntryRow[];
  assets: YardAsset[];
  /** Distinct locations from the source snapshot, for dropdown autocomplete. */
  locations: string[];
};

function rowToSession(r: SessionReadRow): YardSessionRow {
  return {
    id: r.id,
    name: r.name ?? "",
    createdBy: r.created_by ?? "",
    createdAtIso: r.created_at_iso,
    closedAtIso: r.closed_at_iso ?? null,
    sourceDateKey: r.source_date_key ?? "",
    notes: r.notes ?? "",
    status: r.closed_at_iso ? "closed" : "open",
  };
}

type SessionReadRow = {
  id: number;
  name: string | null;
  created_by: string | null;
  created_at_iso: string;
  closed_at_iso: string | null;
  source_date_key: string | null;
  notes: string | null;
};

function rowToEntry(r: EntryReadRow): YardEntryRow {
  return {
    id: r.id,
    sessionId: r.session_id,
    assetId: r.asset_id,
    location: r.location ?? "",
    discrepancies: r.discrepancies ?? "",
    status: normalizeEntryStatus(r.status),
    enteredBy: r.entered_by ?? "",
    enteredAtIso: r.entered_at_iso,
    updatedAtIso: r.updated_at_iso,
  };
}

type EntryReadRow = {
  id: number;
  session_id: number;
  asset_id: string;
  location: string | null;
  discrepancies: string | null;
  status: string | null;
  entered_by: string | null;
  entered_at_iso: string;
  updated_at_iso: string;
};

export function normalizeEntryStatus(s: string | null | undefined): YardEntryStatus {
  const v = (s ?? "").trim().toLowerCase();
  if (v === "missing" || v === "unknown" || v === "not_applicable") return v;
  return "present";
}

/**
 * Yard checks and ETIC rows may disagree on casing / leading spaces; use this
 * when matching a walker's asset_id to `work_order_snapshot` (same idea as
 * `UPPER(TRIM(...))` in SQL, but for in-memory maps and JS keys).
 */
export function canonicalYardAssetKey(assetId: string): string {
  return assetId.trim().toUpperCase();
}

type YardCheckMergeRow = { id: number; asset_id: string; checked_at_iso: string };

function mergeYardCheckRowsByCanonical<T extends YardCheckMergeRow>(rows: T[]): T[] {
  const by = new Map<string, T>();
  for (const row of rows) {
    if (!row.asset_id) continue;
    const c = canonicalYardAssetKey(row.asset_id);
    const existing = by.get(c);
    if (!existing) {
      by.set(c, row);
      continue;
    }
    if (row.checked_at_iso > existing.checked_at_iso) by.set(c, row);
    else if (row.checked_at_iso === existing.checked_at_iso && row.id > existing.id) by.set(c, row);
  }
  return [...by.values()];
}

async function hasWorkOrderSnapshotRowsForDate(env: Env, dateKey: string): Promise<boolean> {
  if (!dateKey) return false;
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT 1 AS ok FROM work_order_snapshot WHERE snapshot_date_key = ? LIMIT 1`,
  )
    .bind(dateKey)
    .first<{ ok: number }>();
  return !!r;
}

async function hasFleetSnapshotRowsForDate(env: Env, dateKey: string): Promise<boolean> {
  if (!dateKey) return false;
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT 1 AS ok FROM fleet_p_a_snapshot WHERE snapshot_date_key = ? LIMIT 1`,
  )
    .bind(dateKey)
    .first<{ ok: number }>();
  return !!r;
}

async function hasYardRosterRowsForDate(env: Env, dateKey: string): Promise<boolean> {
  return await hasFleetSnapshotRowsForDate(env, dateKey) || await hasWorkOrderSnapshotRowsForDate(env, dateKey);
}

/**
 * Yard follows the newest D1 snapshot that has extracted roster rows. R2 latest
 * can lag after rebuilds, but the walker needs the latest queryable D1 roster.
 */
async function getLatestSnapshotDateKey(env: Env): Promise<string> {
  let row = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT e.date_key
       FROM etic_snapshots e
      WHERE e.deleted_at_iso IS NULL
        AND (
          EXISTS (SELECT 1 FROM fleet_p_a_snapshot f WHERE f.snapshot_date_key = e.date_key)
          OR EXISTS (SELECT 1 FROM work_order_snapshot w WHERE w.snapshot_date_key = e.date_key)
        )
      ORDER BY e.date_key DESC
      LIMIT 1`,
  ).first<{ date_key: string }>();
  if (row?.date_key) return row.date_key;

  row = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT snapshot_date_key AS date_key
       FROM fleet_p_a_snapshot
      GROUP BY snapshot_date_key
      ORDER BY snapshot_date_key DESC
      LIMIT 1`,
  ).first<{ date_key: string }>();
  if (row?.date_key) return row.date_key;

  try {
    const latestObj = await env.ETIC_BUCKET.get("analyses/latest.json");
    if (latestObj) {
      const latest = JSON.parse(await latestObj.text()) as { dateKey?: string };
      const dateKey = String(latest?.dateKey ?? "").trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(dateKey) && await hasYardRosterRowsForDate(env, dateKey)) {
        return dateKey;
      }
    }
  } catch {
    // Fall through to D1-based fallback.
  }
  row = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT date_key FROM etic_snapshots WHERE deleted_at_iso IS NULL ORDER BY date_key DESC LIMIT 1`,
  ).first<{ date_key: string }>();
  if (!row?.date_key) {
    try {
      row = await env.ETIC_SNAPSHOTS.prepare(
        `SELECT date_key FROM etic_snapshots ORDER BY date_key DESC LIMIT 1`,
      ).first<{ date_key: string }>();
    } catch {
      row = null;
    }
  }
  if (row?.date_key) return row.date_key;
  const snap = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT snapshot_date_key
       FROM work_order_snapshot
      GROUP BY snapshot_date_key
      ORDER BY snapshot_date_key DESC
      LIMIT 1`,
  ).first<{ snapshot_date_key: string }>();
  return snap?.snapshot_date_key ?? "";
}

/**
 * Pull a VIN/serial from a raw work-order JSON blob. The workbook may store it
 * under "vin", "serial", "serial nbr", "fleet.vin", etc. (see FLEET_SYNONYMS).
 */
function extractVinSerial(raw: Record<string, unknown>): string {
  // Try a few common spellings first for a tight match, then fall back to
  // any key whose normalized name contains "vin" or "serial".
  const exact = [
    "fleet.serial nbr",
    "fleet.serial number",
    "fleet.serial",
    "fleet.vin",
    "fleet.vin/serial",
    "serial nbr",
    "serial number",
    "serial",
    "vin",
    "vin/serial",
  ];
  for (const key of exact) {
    const v = String(raw[key] ?? "").trim();
    if (v) return v;
  }
  for (const [k, val] of Object.entries(raw)) {
    const lk = k.toLowerCase();
    if (lk.includes("vin") || lk.includes("serial")) {
      const s = String(val ?? "").trim();
      if (s) return s;
    }
  }
  return "";
}

/**
 * Find the best "current location" string for an asset out of its raw row.
 * Prefers ETIC location > current/last known > generic location > fleet.location.
 */
function extractLocation(raw: Record<string, unknown>): string {
  const order = [
    "etic location",
    "wo inquiry.etic location",
    "current location",
    "last known location",
    "last location",
    "previous location",
    "location",
    "fleet.etic location",
    "fleet.current location",
    "fleet.location",
    "fleet.previous location",
  ];
  for (const key of order) {
    const v = String(raw[key] ?? "").trim();
    if (v) return v;
  }
  for (const [k, val] of Object.entries(raw)) {
    const lk = k.toLowerCase();
    if (lk.includes("location")) {
      const s = String(val ?? "").trim();
      if (s) return s;
    }
  }
  return "";
}

/**
 * Detect NCE (Nuclear Certified Equipment) flag.
 *
 * Earlier this was `lk.includes("nce")`, which matched "license", "experience",
 * "maintenance", and a dozen other words — flagging every vehicle as NCE. Now
 * we only accept:
 *   1. an exact / known column name from {EXACT_NCE_KEYS}, OR
 *   2. a column that has "nce" or "nuclear" as a whole word.
 *
 * In every case the value must be a "yes-ish" string (yes/y/true/1/x) or be
 * something like "NCE" / "Nuclear" itself. Empty/no/false/0 are rejected.
 */
const EXACT_NCE_KEYS = new Set([
  "nce",
  "nce status",
  "nce flag",
  "nce ind",
  "nce indicator",
  "is nce",
  "nuclear",
  "nuclear cert",
  "nuclear certified",
  "nuclear certified equipment",
  "fleet.nce",
  "fleet.nce status",
  "fleet.nuclear",
  "fleet.nuclear certified",
  "fleet.nuclear certified equipment",
]);
const NCE_FALSY = new Set(["", "no", "n", "false", "0", "none", "n/a", "na", "-"]);
const NCE_TRUTHY = new Set(["yes", "y", "true", "1", "x", "nce", "nuclear", "nuc"]);

function nceValueIsTruthy(raw: unknown): boolean {
  const s = String(raw ?? "").trim().toLowerCase();
  if (NCE_FALSY.has(s)) return false;
  if (NCE_TRUTHY.has(s)) return true;
  // Free-text values like "NCE - certified" or "Nuclear Cert" are also truthy.
  if (s.includes("nce") || s.includes("nuclear")) return true;
  return false;
}

function detectNce(raw: Record<string, unknown>): boolean {
  for (const [k, v] of Object.entries(raw)) {
    if (EXACT_NCE_KEYS.has(k.toLowerCase())) {
      if (nceValueIsTruthy(v)) return true;
    }
  }
  // Fall-back: column name has "nce" or "nuclear" as a whole word (so we don't
  // match "license", "experience", "fence", "maintenance", etc.).
  for (const [k, v] of Object.entries(raw)) {
    if (/(^|[^a-z])nce([^a-z]|$)/i.test(k) || /(^|[^a-z])nuclear([^a-z]|$)/i.test(k)) {
      if (nceValueIsTruthy(v)) return true;
    }
  }
  return false;
}

/** NCE flag + a display string from stored Fleet P&A or WO snapshot `raw_row_json`. */
export function nceInfoFromSnapshotRawJson(rawRowJson: string | null): { nce: boolean; nceStatus: string } {
  if (!rawRowJson) return { nce: false, nceStatus: "" };
  try {
    const raw = JSON.parse(rawRowJson) as Record<string, unknown>;
    if (!detectNce(raw)) return { nce: false, nceStatus: "" };
    let nceStatus = "";
    for (const [k, v] of Object.entries(raw)) {
      const kl = k.toLowerCase();
      if (
        EXACT_NCE_KEYS.has(kl) ||
        /(^|[^a-z])nce([^a-z]|$)/i.test(kl) ||
        /(^|[^a-z])nuclear([^a-z]|$)/i.test(kl)
      ) {
        const s = String(v ?? "").trim();
        if (s) {
          nceStatus = s;
          break;
        }
      }
    }
    return { nce: true, nceStatus };
  } catch {
    return { nce: false, nceStatus: "" };
  }
}

type YardRosterSourceRow = {
  asset_id: string;
  owning_unit: string | null;
  shop: string | null;
  mgmt_cd: string | null;
  make_model: string | null;
  veh_nomen: string | null;
  mel_key: string | null;
  mel_tier?: string | null;
  raw_row_json: string | null;
};

function readYardAssetRow(row: YardRosterSourceRow): { prevLoc: string; vin: string; nce: boolean } {
  let prevLoc = "";
  let vin = "";
  let nce = false;
  if (row.raw_row_json) {
    try {
      const raw = JSON.parse(row.raw_row_json) as Record<string, unknown>;
      prevLoc = extractLocation(raw);
      vin = extractVinSerial(raw);
      nce = detectNce(raw);
    } catch {
      // ignore JSON parse failures
    }
  }
  return { prevLoc, vin, nce };
}

function collectLocationValues(row: YardRosterSourceRow, allLocations: Set<string>): void {
  if (!row.raw_row_json) return;
  try {
    const raw = JSON.parse(row.raw_row_json) as Record<string, unknown>;
    for (const [k, v] of Object.entries(raw)) {
      const key = k.toLowerCase();
      const val = String(v ?? "").trim();
      if (val && key.includes("location")) allLocations.add(val);
    }
  } catch {
    // ignore JSON parse failures
  }
}

/**
 * Build the yard-check roster from the full Fleet P&A snapshot, then overlay
 * open-WO context from the same ETIC day. Checked assets outside Fleet P&A are
 * added later by getRollingRoster as unlisted/floor-to-book finds.
 */
export async function getYardRosterForDate(env: Env, dateKey: string): Promise<YardRoster> {
  if (!dateKey) return { assets: [], locations: [] };
  const [fleetRows, woRows] = await Promise.all([
    env.ETIC_SNAPSHOTS.prepare(
      `SELECT asset_id, owning_unit, shop, mgmt_cd, make_model, veh_nomen, mel_key, raw_row_json
       FROM fleet_p_a_snapshot
       WHERE snapshot_date_key = ? AND asset_id != ''`,
    )
      .bind(dateKey)
      .all<YardRosterSourceRow>(),
    env.ETIC_SNAPSHOTS.prepare(
      `SELECT asset_id, owning_unit, shop, mgmt_cd, make_model, veh_nomen, mel_key, mel_tier, raw_row_json
       FROM work_order_snapshot
       WHERE snapshot_date_key = ? AND asset_id != ''`,
    )
      .bind(dateKey)
      .all<YardRosterSourceRow>(),
  ]);
  const grouped = new Map<string, YardAsset>();
  const allLocations = new Set<string>();

  for (const row of fleetRows.results ?? []) {
    const id = (row.asset_id ?? "").trim();
    if (!id) continue;
    collectLocationValues(row, allLocations);
    const { prevLoc, vin, nce } = readYardAssetRow(row);
    grouped.set(canonicalYardAssetKey(id), {
      assetId: id,
      owningUnit: row.owning_unit ?? "",
      shop: row.shop ?? "",
      mgmtCd: row.mgmt_cd ?? "",
      makeModel: row.make_model ?? "",
      vehNomen: row.veh_nomen ?? "",
      melKey: row.mel_key ?? "",
      melTier: "",
      vinSerial: vin,
      previousLocation: prevLoc,
      openWoCount: 0,
      isNce: nce,
    });
  }

  for (const row of woRows.results ?? []) {
    const id = (row.asset_id ?? "").trim();
    if (!id) continue;
    collectLocationValues(row, allLocations);
    const canon = canonicalYardAssetKey(id);
    const { prevLoc, vin, nce } = readYardAssetRow(row);
    const existing = grouped.get(canon);
    if (!existing) {
      grouped.set(canon, {
        assetId: id,
        owningUnit: row.owning_unit ?? "",
        shop: row.shop ?? "",
        mgmtCd: row.mgmt_cd ?? "",
        makeModel: row.make_model ?? "",
        vehNomen: row.veh_nomen ?? "",
        melKey: row.mel_key ?? "",
        melTier: row.mel_tier ?? "",
        vinSerial: vin,
        previousLocation: prevLoc,
        openWoCount: 1,
        isNce: nce,
      });
      continue;
    }
    existing.openWoCount += 1;
    if (row.owning_unit) existing.owningUnit = row.owning_unit;
    if (row.shop) existing.shop = row.shop;
    if (row.mgmt_cd) existing.mgmtCd = row.mgmt_cd;
    if (row.make_model) existing.makeModel = row.make_model;
    if (row.veh_nomen) existing.vehNomen = row.veh_nomen;
    if (row.mel_key) existing.melKey = row.mel_key;
    if (row.mel_tier) existing.melTier = row.mel_tier;
    if (prevLoc) existing.previousLocation = prevLoc;
    if (vin) existing.vinSerial = vin;
    if (nce) existing.isNce = true;
  }

  const assets = [...grouped.values()];
  assets.sort((a, b) => a.assetId.localeCompare(b.assetId, undefined, { numeric: true }));
  const locations = [...allLocations].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
  return { assets, locations };
}

/**
 * Backwards-compatible alias for callers that only need the asset list.
 * Prefer getYardRosterForDate going forward.
 */
export async function getAssetRosterForDate(env: Env, dateKey: string): Promise<YardAsset[]> {
  const r = await getYardRosterForDate(env, dateKey);
  return r.assets;
}

export async function listSessions(env: Env, limit = 100): Promise<YardSessionRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, name, created_by, created_at_iso, closed_at_iso, source_date_key, notes
     FROM yard_check_session
     ORDER BY id DESC
     LIMIT ?`,
  )
    .bind(limit)
    .all<SessionReadRow>();
  return (r.results ?? []).map(rowToSession);
}

export async function getSession(env: Env, sessionId: number): Promise<YardSessionRow | null> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, name, created_by, created_at_iso, closed_at_iso, source_date_key, notes
     FROM yard_check_session WHERE id = ?`,
  )
    .bind(sessionId)
    .first<SessionReadRow>();
  return r ? rowToSession(r) : null;
}

export async function getEntries(env: Env, sessionId: number): Promise<YardEntryRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, session_id, asset_id, location, discrepancies, status, entered_by, entered_at_iso, updated_at_iso
     FROM yard_check_entry WHERE session_id = ?
     ORDER BY updated_at_iso DESC`,
  )
    .bind(sessionId)
    .all<EntryReadRow>();
  return (r.results ?? []).map(rowToEntry);
}

export async function getSessionDetail(
  env: Env,
  sessionId: number,
): Promise<YardSessionDetail | null> {
  const session = await getSession(env, sessionId);
  if (!session) return null;
  const dateKey = session.sourceDateKey || (await getLatestSnapshotDateKey(env));
  const [entries, roster] = await Promise.all([
    getEntries(env, sessionId),
    getYardRosterForDate(env, dateKey),
  ]);
  return { session, entries, assets: roster.assets, locations: roster.locations };
}

export type CreateSessionInput = {
  name?: string;
  createdBy?: string;
  notes?: string;
  /** If omitted, the latest known snapshot is used. */
  sourceDateKey?: string;
};

export async function createSession(env: Env, input: CreateSessionInput): Promise<YardSessionRow> {
  const sourceDateKey = input.sourceDateKey || (await getLatestSnapshotDateKey(env));
  const nowIso = new Date().toISOString();
  const name = (input.name ?? "").trim() || defaultSessionName(nowIso);
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO yard_check_session (name, created_by, created_at_iso, source_date_key, notes)
     VALUES (?, ?, ?, ?, ?)
     RETURNING id, name, created_by, created_at_iso, closed_at_iso, source_date_key, notes`,
  )
    .bind(name, input.createdBy ?? "", nowIso, sourceDateKey, input.notes ?? "")
    .first<SessionReadRow>();
  if (!r) throw new Error("failed to create yard check session");
  return rowToSession(r);
}

function defaultSessionName(nowIso: string): string {
  const d = new Date(nowIso);
  if (Number.isNaN(d.getTime())) return "Yard check";
  // e.g. "Yard check 18 APR 26 0814"
  const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = months[d.getUTCMonth()];
  const yy = String(d.getUTCFullYear()).slice(-2);
  const hh = String(d.getUTCHours()).padStart(2, "0");
  const mi = String(d.getUTCMinutes()).padStart(2, "0");
  return `Yard check ${dd} ${mm} ${yy} ${hh}${mi}Z`;
}

export type UpsertEntryInput = {
  assetId: string;
  location?: string;
  discrepancies?: string;
  status?: YardEntryStatus;
  enteredBy?: string;
};

export async function upsertEntry(
  env: Env,
  sessionId: number,
  input: UpsertEntryInput,
): Promise<YardEntryRow> {
  const session = await getSession(env, sessionId);
  if (!session) throw new Error("session not found");
  if (session.status === "closed") throw new Error("session is closed");
  const assetId = input.assetId.trim();
  if (!assetId) throw new Error("assetId required");
  const status = normalizeEntryStatus(input.status);
  const nowIso = new Date().toISOString();
  const location = (input.location ?? "").trim();
  const discrepancies = (input.discrepancies ?? "").trim();
  const enteredBy = (input.enteredBy ?? "").trim();
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO yard_check_entry
       (session_id, asset_id, location, discrepancies, status, entered_by, entered_at_iso, updated_at_iso)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)
     ON CONFLICT(session_id, asset_id) DO UPDATE SET
       location = excluded.location,
       discrepancies = excluded.discrepancies,
       status = excluded.status,
       entered_by = excluded.entered_by,
       updated_at_iso = excluded.updated_at_iso
     RETURNING id, session_id, asset_id, location, discrepancies, status, entered_by, entered_at_iso, updated_at_iso`,
  )
    .bind(sessionId, assetId, location, discrepancies, status, enteredBy, nowIso, nowIso)
    .first<EntryReadRow>();
  if (!r) throw new Error("failed to upsert yard check entry");
  return rowToEntry(r);
}

export async function deleteEntry(env: Env, sessionId: number, assetId: string): Promise<boolean> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `DELETE FROM yard_check_entry WHERE session_id = ? AND asset_id = ?`,
  )
    .bind(sessionId, assetId)
    .run();
  // D1 result shape: meta.changes
  const meta = (r as unknown as { meta?: { changes?: number } }).meta;
  return (meta?.changes ?? 0) > 0;
}

export async function closeSession(env: Env, sessionId: number): Promise<YardSessionRow | null> {
  const nowIso = new Date().toISOString();
  await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE yard_check_session SET closed_at_iso = ? WHERE id = ? AND closed_at_iso IS NULL`,
  )
    .bind(nowIso, sessionId)
    .run();
  return getSession(env, sessionId);
}

export async function reopenSession(env: Env, sessionId: number): Promise<YardSessionRow | null> {
  await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE yard_check_session SET closed_at_iso = NULL WHERE id = ?`,
  )
    .bind(sessionId)
    .run();
  return getSession(env, sessionId);
}

export async function deleteSession(env: Env, sessionId: number): Promise<boolean> {
  await env.ETIC_SNAPSHOTS.prepare(
    `DELETE FROM yard_check_entry WHERE session_id = ?`,
  )
    .bind(sessionId)
    .run();
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `DELETE FROM yard_check_session WHERE id = ?`,
  )
    .bind(sessionId)
    .run();
  const meta = (r as unknown as { meta?: { changes?: number } }).meta;
  return (meta?.changes ?? 0) > 0;
}

/* =============================================================================
   ROLLING YARD CHECK MODEL
   Each "check" is a row in yard_check (no session). The searchable roster is the
   current fleet, but the "due" queue is intentionally narrower: assets with an
   open WO on the latest ingest and no present check inside the cadence.
   ========================================================================== */

/** How many days a check is fresh for. Override via app_config["yardCheckIntervalDays"]. */
export const DEFAULT_YARD_CHECK_INTERVAL_DAYS = 7;

export type YardCheckRow = {
  id: number;
  assetId: string;
  location: string;
  discrepancies: string;
  status: YardEntryStatus;
  checkedBy: string;
  checkedAtIso: string;
  sourceDateKey: string;
  assetSnapshotJson: string;
  /** Populated on asset detail: photo URLs from this same yard visit (check_id), newest first. */
  checkPhotoUrls?: string[];
};

/** One correction applied to an existing `yard_check` row (see migrations/0022). */
export type YardCheckEditSnapshot = {
  location: string;
  discrepancies: string;
  status: YardEntryStatus;
};

export type YardCheckEditRow = {
  id: number;
  checkId: number;
  editedAtIso: string;
  editedBy: string;
  before: YardCheckEditSnapshot;
  after: YardCheckEditSnapshot;
};

export type YardPhotoRow = {
  id: number;
  assetId: string;
  checkId: number | null;
  r2Key: string;
  contentType: string;
  sizeBytes: number;
  uploadedBy: string;
  uploadedAtIso: string;
  caption: string;
  /** Convenience URL for the client. */
  url: string;
};

export type RollingAssetState = "fresh" | "due" | "overdue" | "never";

export type RollingAsset = YardAsset & {
  /** Most recent check timestamp, or null if never. */
  lastCheckedAtIso: string | null;
  /** Walker who recorded the last check, or "". */
  lastCheckedBy: string;
  /** Number of days since the last check (rounded down), or null if never. */
  daysSinceLastCheck: number | null;
  /** Bucketed staleness — drives badge color. */
  rollingState: RollingAssetState;
  /** True if this asset has never been checked. */
  isNeverChecked: boolean;
  /** True if this is the asset's first appearance vs the previous snapshot. */
  isNewAsset: boolean;
  /**
   * True if the asset_id has check history but does NOT appear in the latest
   * ETIC snapshot — i.e. a "floor-to-book" find. Surface a UI hint so book-
   * keepers can reconcile it.
   */
  isUnlisted: boolean;
  /** Number of photos uploaded against this asset. */
  photoCount: number;
  /**
   * Best-known current location: most-recent yard-check location, falling back
   * to the snapshot's "previous location" extracted from raw row JSON.
   */
  lastLocation: string;
  /** Notes / discrepancies from the most recent yard_check (may be empty). */
  lastNotes: string;
  /** True if any open WO for this asset has mel_tier === 'below'. */
  isBelowMel: boolean;
};

export type RollingRoster = {
  dateKey: string;
  intervalDays: number;
  assets: RollingAsset[];
  locations: string[];
  /** Roll-up counts so the UI doesn't need to walk the array. */
  totals: {
    total: number;
    due: number;
    overdue: number;
    fresh: number;
    never: number;
    checkedToday: number;
    checkedThisWeek: number;
  };
};

/** One row per exact asset_id: latest yard_check (ties broken by id). Merged by canonical in JS. */
type LatestCheckFullRow = {
  id: number;
  asset_id: string;
  checked_at_iso: string;
  checked_by: string | null;
  discrepancies: string | null;
};

type CheckReadRow = {
  id: number;
  asset_id: string;
  location: string | null;
  discrepancies: string | null;
  status: string | null;
  checked_by: string | null;
  checked_at_iso: string;
  source_date_key: string | null;
  snapshot_asset_json?: string | null;
  asset_snapshot_json?: string | null;
};

function rowToCheck(r: CheckReadRow): YardCheckRow {
  return {
    id: r.id,
    assetId: r.asset_id,
    location: r.location ?? "",
    discrepancies: r.discrepancies ?? "",
    status: normalizeEntryStatus(r.status),
    checkedBy: r.checked_by ?? "",
    checkedAtIso: r.checked_at_iso,
    sourceDateKey: r.source_date_key ?? "",
    assetSnapshotJson: r.snapshot_asset_json ?? r.asset_snapshot_json ?? "",
  };
}

/** openWoCount from recordCheck snapshot JSON; null if missing or not parseable. */
function openWoCountFromCheckSnapshot(check: YardCheckRow | null): number | null {
  if (!check?.assetSnapshotJson?.trim()) return null;
  try {
    const parsed = JSON.parse(check.assetSnapshotJson) as { asset?: { openWoCount?: number } };
    const n = parsed?.asset?.openWoCount;
    if (typeof n === "number" && Number.isFinite(n) && n >= 0) return Math.floor(n);
  } catch {
    /* ignore */
  }
  return null;
}

/** Fleet P&A row captured on the check (when the walker saved). Used when the ID is off-roster today. */
function yardAssetFromCheckSnapshotJson(json: string | null | undefined): YardAsset | null {
  if (!json?.trim()) return null;
  try {
    const parsed = JSON.parse(json) as { asset?: YardAsset };
    const a = parsed?.asset;
    if (!a || typeof a !== "object") return null;
    if (typeof a.assetId !== "string" || !a.assetId.trim()) return null;
    return a;
  } catch {
    return null;
  }
}

function rollingAssetFromTriggerCheck(
  check: YardCheckRow,
  photos: YardPhotoRow[],
  intervalDays: number,
): RollingAsset {
  const snap = yardAssetFromCheckSnapshotJson(check.assetSnapshotJson);
  const base: YardAsset = snap ?? {
    assetId: check.assetId,
    owningUnit: "",
    shop: "",
    mgmtCd: "",
    makeModel: "",
    vehNomen: "",
    melKey: "",
    melTier: "",
    vinSerial: "",
    previousLocation: "",
    openWoCount: 0,
    isNce: false,
  };
  const nowIso = new Date().toISOString();
  const days = daysBetween(check.checkedAtIso, nowIso);
  const rollingState = bucketState(days, intervalDays);
  return {
    ...base,
    assetId: check.assetId.trim() || base.assetId,
    lastCheckedAtIso: check.checkedAtIso,
    lastCheckedBy: check.checkedBy,
    daysSinceLastCheck: days,
    rollingState,
    isNeverChecked: false,
    isNewAsset: false,
    isUnlisted: true,
    photoCount: photos.length,
    lastLocation: (check.location ?? "").trim() || base.previousLocation || "",
    lastNotes: (check.discrepancies ?? "").trim(),
    isBelowMel: false,
  };
}

function mergeRollingAssetFromSnapshotJson(asset: RollingAsset, snap: YardAsset): RollingAsset {
  const pick = <K extends keyof YardAsset>(k: K) => {
    const cur = asset[k];
    const v = snap[k];
    if ((cur === "" || cur == null) && v != null && String(v).trim()) (asset as YardAsset)[k] = v;
  };
  pick("owningUnit");
  pick("shop");
  pick("mgmtCd");
  pick("makeModel");
  pick("vehNomen");
  pick("melKey");
  pick("melTier");
  pick("vinSerial");
  pick("previousLocation");
  if (!asset.lastLocation?.trim() && snap.previousLocation?.trim()) {
    asset.lastLocation = snap.previousLocation.trim();
  }
  if (
    asset.isUnlisted &&
    typeof snap.openWoCount === "number" &&
    Number.isFinite(snap.openWoCount) &&
    snap.openWoCount >= 0
  ) {
    asset.openWoCount = Math.floor(snap.openWoCount);
  }
  if (snap.isNce) asset.isNce = true;
  return asset;
}

function hydrateFindingRollingAsset(
  f: YardFinding,
  intervalDays_fallback: number,
  photoByCanon: Map<string, number>,
): void {
  if (!f.triggerCheck) return;
  if (f.asset) {
    const snap = yardAssetFromCheckSnapshotJson(f.triggerCheck.assetSnapshotJson);
    if (snap) mergeRollingAssetFromSnapshotJson(f.asset, snap);
    const c = canonicalYardAssetKey(f.assetId);
    const pc = photoByCanon.get(c);
    if (pc != null && pc > f.photoCount) f.photoCount = pc;
    return;
  }
  const c = canonicalYardAssetKey(f.assetId);
  const stubPhotos: YardPhotoRow[] = [];
  f.asset = rollingAssetFromTriggerCheck(
    f.triggerCheck,
    stubPhotos,
    intervalDays_fallback > 0 ? intervalDays_fallback : DEFAULT_YARD_CHECK_INTERVAL_DAYS,
  );
  const snap = yardAssetFromCheckSnapshotJson(f.triggerCheck.assetSnapshotJson);
  if (snap) mergeRollingAssetFromSnapshotJson(f.asset, snap);
  const pc = photoByCanon.get(c);
  if (pc != null) f.photoCount = pc;
}

async function batchPhotoCountsForCanonicalAssets(env: Env, canonSet: Set<string>): Promise<Map<string, number>> {
  const out = new Map<string, number>();
  const ids = [...canonSet].filter(Boolean);
  if (!ids.length) return out;
  const chunk = 80;
  for (let i = 0; i < ids.length; i += chunk) {
    const slice = ids.slice(i, i + chunk);
    const ph2 = slice.map(() => "?").join(",");
    const r = await env.ETIC_SNAPSHOTS.prepare(
      `SELECT UPPER(TRIM(asset_id)) AS aid, COUNT(*) AS c FROM yard_photo
       WHERE UPPER(TRIM(asset_id)) IN (${ph2}) GROUP BY UPPER(TRIM(asset_id))`,
    )
      .bind(...slice)
      .all<{ aid: string; c: number }>();
    for (const row of r.results ?? []) {
      if (row.aid) out.set(row.aid.toUpperCase(), Math.floor(row.c ?? 0));
    }
  }
  return out;
}

/**
 * Keys "snapshot_date_key|CANON_ASSET_ID" for assets that appear on work_order_snapshot
 * for that report day. Batched to avoid one D1 round-trip per yard_check row.
 */
async function batchWorkOrderPresenceOnSnapshots(
  env: Env,
  byDateKey: Map<string, Set<string>>,
): Promise<Set<string>> {
  const out = new Set<string>();
  const chunkSize = 80;
  for (const [dk, canons] of byDateKey) {
    const dateKey = dk.trim();
    if (!dateKey || !canons.size) continue;
    const arr = [...canons];
    for (let i = 0; i < arr.length; i += chunkSize) {
      const slice = arr.slice(i, i + chunkSize);
      const ph = slice.map(() => "?").join(",");
      const r = await env.ETIC_SNAPSHOTS.prepare(
        `SELECT DISTINCT UPPER(TRIM(asset_id)) AS a
         FROM work_order_snapshot
         WHERE snapshot_date_key = ? AND UPPER(TRIM(asset_id)) IN (${ph})`,
      )
        .bind(dateKey, ...slice)
        .all<{ a: string }>();
      const prefix = dateKey + "|";
      for (const row of r.results ?? []) {
        const a = (row.a ?? "").trim();
        if (a) out.add(prefix + a);
      }
    }
  }
  return out;
}

async function getYardCheckIntervalDays(env: Env): Promise<number> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT value FROM app_config WHERE key = ?`,
  )
    .bind("yardCheckIntervalDays")
    .first<{ value: string }>();
  if (!r?.value) return DEFAULT_YARD_CHECK_INTERVAL_DAYS;
  try {
    const v = JSON.parse(r.value);
    if (typeof v === "number" && v > 0 && v <= 365) return Math.floor(v);
  } catch {
    // ignore parse failure, fall through
  }
  const n = Number.parseInt(r.value, 10);
  if (Number.isFinite(n) && n > 0) return n;
  return DEFAULT_YARD_CHECK_INTERVAL_DAYS;
}

function daysBetween(aIso: string, bIso: string): number {
  const a = new Date(aIso).getTime();
  const b = new Date(bIso).getTime();
  if (!Number.isFinite(a) || !Number.isFinite(b)) return 0;
  return Math.floor((b - a) / (1000 * 60 * 60 * 24));
}

function bucketState(daysSince: number | null, intervalDays: number): RollingAssetState {
  if (daysSince === null) return "never";
  if (daysSince >= intervalDays * 2) return "overdue";
  if (daysSince >= intervalDays) return "due";
  return "fresh";
}

function isOpenWoDue(openWoCount: number, daysSince: number | null, intervalDays: number): boolean {
  return openWoCount > 0 && (daysSince === null || daysSince >= intervalDays);
}

/**
 * One-shot rolling roster: latest snapshot's assets + their last-check info +
 * photo counts + new-asset flag (first appearance vs the prior snapshot).
 */
export async function getRollingRoster(env: Env): Promise<RollingRoster> {
  const intervalDays = await getYardCheckIntervalDays(env);
  const dateKey = await getLatestSnapshotDateKey(env);
  if (!dateKey) {
    return {
      dateKey: "",
      intervalDays,
      assets: [],
      locations: [],
      totals: { total: 0, due: 0, overdue: 0, fresh: 0, never: 0, checkedToday: 0, checkedThisWeek: 0 },
    };
  }

  // Fetch in parallel: latest roster, last-checks per asset, photo counts,
  // the prior snapshot's asset set (for "NEW" detection), the most-recent
  // location each walker recorded, and the per-asset below-MEL flag.
  const [roster, latestChecks, photoCounts, priorAssets, latestLocs, belowMelRows] = await Promise.all([
    getYardRosterForDate(env, dateKey),
    env.ETIC_SNAPSHOTS.prepare(
      `SELECT id, asset_id, checked_at_iso, checked_by, discrepancies
       FROM (
         SELECT
           id,
           asset_id,
           checked_at_iso,
           checked_by,
           discrepancies,
           ROW_NUMBER() OVER (PARTITION BY asset_id ORDER BY checked_at_iso DESC, id DESC) AS rn
         FROM yard_check
       ) WHERE rn = 1`,
    ).all<LatestCheckFullRow>(),
    env.ETIC_SNAPSHOTS.prepare(
      `SELECT asset_id, COUNT(*) AS c FROM yard_photo GROUP BY asset_id`,
    ).all<{ asset_id: string; c: number }>(),
    env.ETIC_SNAPSHOTS.prepare(
      `SELECT date_key FROM etic_snapshots
        WHERE deleted_at_iso IS NULL AND date_key < ? ORDER BY date_key DESC LIMIT 1`,
    ).bind(dateKey).first<{ date_key: string }>(),
    env.ETIC_SNAPSHOTS.prepare(
      `SELECT yc.asset_id AS asset_id, yc.location AS location
       FROM yard_check yc
       JOIN (
         SELECT asset_id, MAX(checked_at_iso) AS m FROM yard_check
          WHERE COALESCE(location, '') != ''
          GROUP BY asset_id
       ) m ON m.asset_id = yc.asset_id AND m.m = yc.checked_at_iso`,
    ).all<{ asset_id: string; location: string }>(),
    env.ETIC_SNAPSHOTS.prepare(
      `SELECT DISTINCT asset_id FROM work_order_snapshot
        WHERE snapshot_date_key = ? AND mel_tier = 'below'`,
    ).bind(dateKey).all<{ asset_id: string }>(),
  ]);

  const latestMerged = mergeYardCheckRowsByCanonical(latestChecks.results ?? []);
  const lastByAsset = new Map<string, { at: string; by: string; notes: string; displayId: string }>();
  for (const row of latestMerged) {
    if (!row.asset_id) continue;
    const c = canonicalYardAssetKey(row.asset_id);
    lastByAsset.set(c, {
      at: row.checked_at_iso,
      by: row.checked_by ?? "",
      notes: (row.discrepancies ?? "").trim(),
      displayId: row.asset_id.trim(),
    });
  }
  const photoByCanon = new Map<string, number>();
  for (const row of photoCounts.results ?? []) {
    if (row.asset_id) {
      const c = canonicalYardAssetKey(row.asset_id);
      photoByCanon.set(c, (photoByCanon.get(c) ?? 0) + (row.c ?? 0));
    }
  }
  const lastLocByCanon = new Map<string, string>();
  for (const row of latestLocs.results ?? []) {
    if (row.asset_id && row.location) {
      const c = canonicalYardAssetKey(row.asset_id);
      const loc = row.location;
      const prev = lastLocByCanon.get(c);
      if (!prev || loc.length > prev.length) lastLocByCanon.set(c, loc);
    }
  }
  const belowMelAssets = new Set<string>();
  for (const row of belowMelRows.results ?? []) {
    if (row.asset_id) belowMelAssets.add(canonicalYardAssetKey(row.asset_id));
  }

  let priorAssetIds = new Set<string>();
  if (priorAssets?.date_key) {
    const r = await env.ETIC_SNAPSHOTS.prepare(
      `SELECT DISTINCT asset_id FROM work_order_snapshot WHERE snapshot_date_key = ?`,
    ).bind(priorAssets.date_key).all<{ asset_id: string }>();
    priorAssetIds = new Set(
      (r.results ?? [])
        .map((x) => (x.asset_id ? canonicalYardAssetKey(x.asset_id) : ""))
        .filter(Boolean),
    );
  }

  const rosterByCanon = new Map<string, YardAsset>();
  for (const a of roster.assets) {
    rosterByCanon.set(canonicalYardAssetKey(a.assetId), a);
  }

  const nowIso = new Date().toISOString();
  const todayPrefix = nowIso.slice(0, 10);
  const totals = { total: 0, due: 0, overdue: 0, fresh: 0, never: 0, checkedToday: 0, checkedThisWeek: 0 };
  const out: RollingAsset[] = [];
  const inLatestSnapshot = new Set<string>();
  for (const a of roster.assets) {
    const c = canonicalYardAssetKey(a.assetId);
    inLatestSnapshot.add(c);
    const last = lastByAsset.get(c) ?? null;
    const days = last ? daysBetween(last.at, nowIso) : null;
    const state = bucketState(days, intervalDays);
    const isNew = priorAssetIds.size > 0 && !priorAssetIds.has(c);
    out.push({
      ...a,
      lastCheckedAtIso: last?.at ?? null,
      lastCheckedBy: last?.by ?? "",
      daysSinceLastCheck: days,
      rollingState: state,
      isNeverChecked: !last,
      isNewAsset: isNew,
      isUnlisted: false,
      photoCount: photoByCanon.get(c) ?? 0,
      lastLocation: lastLocByCanon.get(c) || "",
      lastNotes: last?.notes ?? "",
      isBelowMel: belowMelAssets.has(c),
    });
    totals.total += 1;
    if (isOpenWoDue(a.openWoCount, days, intervalDays)) {
      if (state === "overdue" || state === "never") totals.overdue += 1;
      else totals.due += 1;
    } else {
      totals.fresh += 1;
    }
    if (last && last.at.slice(0, 10) === todayPrefix) totals.checkedToday += 1;
    if (last && days !== null && days < 7) totals.checkedThisWeek += 1;
  }

  // Do not append true floor-to-book IDs here. The Fleet list is strictly the
  // current Excel Fleet P&A roster (plus WO overlay). Found IDs that are not in
  // the workbook still surface in Needs Fix via listOpenFindings.
  for (const [c, last] of lastByAsset) {
    if (inLatestSnapshot.has(c)) continue;
    if (last.at.slice(0, 10) === todayPrefix) totals.checkedToday += 1;
    const days = daysBetween(last.at, nowIso);
    if (days < 7) totals.checkedThisWeek += 1;
  }

  // Sort: never-checked first, then overdue, then due, then fresh; within
  // each bucket, oldest check first so the walker tackles the worst.
  const order: Record<RollingAssetState, number> = { never: 0, overdue: 1, due: 2, fresh: 3 };
  out.sort((a, b) => {
    const oa = order[a.rollingState];
    const ob = order[b.rollingState];
    if (oa !== ob) return oa - ob;
    const da = a.daysSinceLastCheck ?? 1e9;
    const db = b.daysSinceLastCheck ?? 1e9;
    if (db !== da) return db - da;
    return a.assetId.localeCompare(b.assetId, undefined, { numeric: true });
  });

  return {
    dateKey,
    intervalDays,
    assets: out,
    locations: roster.locations,
    totals,
  };
}

/* ----- recording a check ------------------------------------------------- */

export type RecordCheckInput = {
  assetId: string;
  location?: string;
  discrepancies?: string;
  status?: YardEntryStatus;
  checkedBy?: string;
  sourceDateKey?: string;
  /** When true (Find-unlisted flow only), reject if asset already on latest ETIC WO snapshot. */
  fromFindUnlisted?: boolean;
};

async function assetOnWorkOrderSnapshot(env: Env, dateKey: string, assetId: string): Promise<boolean> {
  if (!dateKey || !assetId) return false;
  const row = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT 1 AS x FROM work_order_snapshot
     WHERE snapshot_date_key = ? AND UPPER(TRIM(asset_id)) = UPPER(TRIM(?)) LIMIT 1`,
  )
    .bind(dateKey, assetId)
    .first<{ x: number }>();
  return !!row;
}

async function assetSnapshotJsonForCheck(env: Env, dateKey: string, assetId: string): Promise<string> {
  if (!dateKey || !assetId) return "";
  const roster = await getYardRosterForDate(env, dateKey);
  const canon = canonicalYardAssetKey(assetId);
  const asset = roster.assets.find((a) => canonicalYardAssetKey(a.assetId) === canon);
  if (!asset) return "";
  return JSON.stringify({ sourceDateKey: dateKey, asset });
}

function isMissingColumnError(err: unknown, column: string): boolean {
  return String(err instanceof Error ? err.message : err).toLowerCase().includes(column.toLowerCase());
}

export async function recordCheck(env: Env, input: RecordCheckInput): Promise<YardCheckRow> {
  const assetId = input.assetId.trim();
  if (!assetId) throw new Error("assetId required");
  const status = normalizeEntryStatus(input.status);
  if (status === "present" && !(input.location ?? "").trim()) {
    throw new Error("Location is required when marking present (checked)");
  }
  const sourceDateKey = input.sourceDateKey || (await getLatestSnapshotDateKey(env));
  if (input.fromFindUnlisted && sourceDateKey) {
    const onWo = await assetOnWorkOrderSnapshot(env, sourceDateKey, assetId);
    if (onWo) {
      throw new Error(
        "This asset already appears on the latest ETIC work order list (open WO). Use the fleet list to log a yard check, not Find.",
      );
    }
  }
  const assetSnapshotJson = await assetSnapshotJsonForCheck(env, sourceDateKey, assetId);
  const nowIso = new Date().toISOString();
  const values = [
    assetId,
    (input.location ?? "").trim(),
    (input.discrepancies ?? "").trim(),
    status,
    (input.checkedBy ?? "").trim(),
    nowIso,
    sourceDateKey,
  ] as const;
  let r: CheckReadRow | null = null;
  try {
    r = await env.ETIC_SNAPSHOTS.prepare(
      `INSERT INTO yard_check
         (asset_id, location, discrepancies, status, checked_by, checked_at_iso, source_date_key, snapshot_asset_json)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)
       RETURNING id, asset_id, location, discrepancies, status, checked_by, checked_at_iso, source_date_key, snapshot_asset_json`,
    )
      .bind(...values, assetSnapshotJson)
      .first<CheckReadRow>();
  } catch (err) {
    if (!isMissingColumnError(err, "snapshot_asset_json")) throw err;
    r = await env.ETIC_SNAPSHOTS.prepare(
      `INSERT INTO yard_check
         (asset_id, location, discrepancies, status, checked_by, checked_at_iso, source_date_key)
       VALUES (?, ?, ?, ?, ?, ?, ?)
       RETURNING id, asset_id, location, discrepancies, status, checked_by, checked_at_iso, source_date_key`,
    )
      .bind(...values)
      .first<CheckReadRow>();
    if (r) r.snapshot_asset_json = assetSnapshotJson;
  }
  if (!r) throw new Error("failed to record check");
  return rowToCheck(r);
}

export type UpdateYardCheckInput = {
  checkId: number;
  editedBy: string;
  location: string;
  discrepancies: string;
  status?: YardEntryStatus;
};

function snapshotFromCheckRow(r: CheckReadRow): YardCheckEditSnapshot {
  return {
    location: (r.location ?? "").trim(),
    discrepancies: (r.discrepancies ?? "").trim(),
    status: normalizeEntryStatus(r.status),
  };
}

/**
 * Correct an existing check in place. Logs a row in `yard_check_edit`.
 * Original `checked_at_iso` / `checked_by` are preserved (walker's sighting).
 */
export async function updateYardCheck(env: Env, input: UpdateYardCheckInput): Promise<YardCheckRow> {
  const editedBy = (input.editedBy ?? "").trim();
  if (!editedBy) throw new Error("editedBy required");
  const checkId = input.checkId;
  if (!Number.isFinite(checkId) || checkId <= 0) throw new Error("checkId required");

  const r0 = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, asset_id, location, discrepancies, status, checked_by, checked_at_iso, source_date_key, snapshot_asset_json
     FROM yard_check WHERE id = ?`,
  )
    .bind(checkId)
    .first<CheckReadRow>();
  if (!r0) throw new Error("check not found");

  const before = snapshotFromCheckRow(r0);
  const after: YardCheckEditSnapshot = {
    location: (input.location ?? "").trim(),
    discrepancies: (input.discrepancies ?? "").trim(),
    status: input.status !== undefined ? normalizeEntryStatus(input.status) : before.status,
  };

  if (
    before.location === after.location &&
    before.discrepancies === after.discrepancies &&
    before.status === after.status
  ) {
    return rowToCheck(r0);
  }

  if (after.status === "present" && !after.location.trim()) {
    throw new Error("Location is required when status is Present / found");
  }

  const nowIso = new Date().toISOString();
  const assetId = r0.asset_id;

  await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE yard_check SET location = ?, discrepancies = ?, status = ? WHERE id = ?`,
  )
    .bind(after.location, after.discrepancies, after.status, checkId)
    .run();

  await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO yard_check_edit (asset_id, check_id, edited_at_iso, edited_by, before_json, after_json)
     VALUES (?, ?, ?, ?, ?, ?)`,
  )
    .bind(
      assetId,
      checkId,
      nowIso,
      editedBy,
      JSON.stringify(before),
      JSON.stringify(after),
    )
    .run();

  const r1 = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, asset_id, location, discrepancies, status, checked_by, checked_at_iso, source_date_key, snapshot_asset_json
     FROM yard_check WHERE id = ?`,
  )
    .bind(checkId)
    .first<CheckReadRow>();
  if (!r1) throw new Error("failed to reload check");
  return rowToCheck(r1);
}

type EditReadRow = {
  id: number;
  check_id: number;
  edited_at_iso: string;
  edited_by: string | null;
  before_json: string | null;
  after_json: string | null;
};

function parseEditSnapshotJson(json: string | null): YardCheckEditSnapshot {
  try {
    const o = JSON.parse(json || "{}") as Record<string, unknown>;
    return {
      location: String(o.location ?? ""),
      discrepancies: String(o.discrepancies ?? ""),
      status: normalizeEntryStatus(typeof o.status === "string" ? o.status : "present"),
    };
  } catch {
    return { location: "", discrepancies: "", status: "present" };
  }
}

export async function getCheckEditsForAsset(env: Env, assetId: string): Promise<YardCheckEditRow[]> {
  const id = assetId.trim();
  if (!id) return [];
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, check_id, edited_at_iso, edited_by, before_json, after_json
     FROM yard_check_edit
     WHERE UPPER(TRIM(asset_id)) = UPPER(TRIM(?))
     ORDER BY edited_at_iso DESC`,
  )
    .bind(id)
    .all<EditReadRow>();
  return (r.results ?? []).map((row) => ({
    id: row.id,
    checkId: row.check_id,
    editedAtIso: row.edited_at_iso,
    editedBy: row.edited_by ?? "",
    before: parseEditSnapshotJson(row.before_json),
    after: parseEditSnapshotJson(row.after_json),
  }));
}

export async function getChecksForAsset(
  env: Env,
  assetId: string,
  limit = 50,
): Promise<YardCheckRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, asset_id, location, discrepancies, status, checked_by, checked_at_iso, source_date_key, snapshot_asset_json
     FROM yard_check
     WHERE UPPER(TRIM(asset_id)) = UPPER(TRIM(?))
     ORDER BY checked_at_iso DESC LIMIT ?`,
  )
    .bind(assetId, limit)
    .all<CheckReadRow>();
  return (r.results ?? []).map(rowToCheck);
}

export async function getRecentChecks(env: Env, limit = 100): Promise<YardCheckRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, asset_id, location, discrepancies, status, checked_by, checked_at_iso, source_date_key, snapshot_asset_json
     FROM yard_check ORDER BY checked_at_iso DESC LIMIT ?`,
  )
    .bind(limit)
    .all<CheckReadRow>();
  return (r.results ?? []).map(rowToCheck);
}

/* =============================================================================
   LATEST SIGHTINGS LOOKUP
   For every asset_id that's ever been logged with status='present', return the
   most-recent (location, when, by). Used across the desktop UI (Work Orders,
   MEL, ETIC Meeting, Presenter) to surface "where the asset actually is right
   now and when somebody last laid eyes on it" next to the asset id.

   Only `present` checks count. A walker tagging an asset with another status
   (legacy 'missing'/'unknown' rows) does NOT update the sighting — the whole
   point of the badge is to answer "where is it physically".
   ========================================================================== */

export type YardSighting = {
  /** Best-known parking spot at the time of sighting. */
  location: string;
  /** ISO timestamp of the most-recent present check. */
  at: string;
  /** Walker who logged it. May be empty. */
  by: string;
};

type SightingRow = {
  asset_id: string;
  location: string | null;
  checked_by: string | null;
  checked_at_iso: string;
};

/**
 * Returns a Map<assetId, YardSighting> with the most-recent 'present' check
 * for every asset that has one. Cheap — one row per asset (D1 GROUP BY).
 */
export async function getLatestSightings(env: Env): Promise<Map<string, YardSighting>> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT yc.asset_id, yc.location, yc.checked_by, yc.checked_at_iso
       FROM yard_check yc
       JOIN (
         SELECT asset_id, MAX(checked_at_iso) AS m
           FROM yard_check
          WHERE LOWER(COALESCE(status, 'present')) = 'present'
          GROUP BY asset_id
       ) m ON m.asset_id = yc.asset_id AND m.m = yc.checked_at_iso`,
  ).all<SightingRow>();
  const out = new Map<string, YardSighting>();
  for (const row of r.results ?? []) {
    if (!row.asset_id) continue;
    out.set(row.asset_id, {
      location: row.location ?? "",
      at: row.checked_at_iso,
      by: row.checked_by ?? "",
    });
  }
  return out;
}

type LatestPhotoRow = { asset_id: string; id: number };

/**
 * One row per asset: id of the most recently uploaded yard_photo (any check).
 * Used for cross-cutting "latest yard photo" thumbnails on WO / MEL / meeting.
 */
export async function getLatestYardPhotoIdsByAsset(env: Env): Promise<Map<string, number>> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT asset_id, id FROM (
       SELECT asset_id, id,
         ROW_NUMBER() OVER (PARTITION BY asset_id ORDER BY uploaded_at_iso DESC, id DESC) AS rn
       FROM yard_photo
     ) WHERE rn = 1`,
  ).all<LatestPhotoRow>();
  const out = new Map<string, number>();
  for (const row of r.results ?? []) {
    if (row.asset_id && Number.isFinite(row.id)) out.set(row.asset_id, row.id);
  }
  return out;
}

export async function deleteCheck(env: Env, checkId: number): Promise<boolean> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `DELETE FROM yard_check WHERE id = ?`,
  )
    .bind(checkId)
    .run();
  const meta = (r as unknown as { meta?: { changes?: number } }).meta;
  return (meta?.changes ?? 0) > 0;
}

/* ----- photos ----------------------------------------------------------- */

export const YARD_PHOTO_PREFIX = "yard-photos/";

function photoUrlFor(id: number): string {
  return "/api/yard/photo/" + id;
}

function rowToPhoto(r: PhotoReadRow): YardPhotoRow {
  return {
    id: r.id,
    assetId: r.asset_id,
    checkId: r.check_id ?? null,
    r2Key: r.r2_key,
    contentType: r.content_type ?? "image/jpeg",
    sizeBytes: r.size_bytes ?? 0,
    uploadedBy: r.uploaded_by ?? "",
    uploadedAtIso: r.uploaded_at_iso,
    caption: r.caption ?? "",
    url: photoUrlFor(r.id),
  };
}

async function attachCheckPhotoUrlsToChecks(env: Env, checks: YardCheckRow[]): Promise<void> {
  const ids = checks.map((c) => c.id).filter((id) => Number.isFinite(id));
  if (!ids.length) return;
  const byCheck = new Map<number, string[]>();
  const chunk = 80;
  for (let i = 0; i < ids.length; i += chunk) {
    const slice = ids.slice(i, i + chunk);
    const ph = slice.map(() => "?").join(",");
    const r = await env.ETIC_SNAPSHOTS.prepare(
      `SELECT id, check_id FROM yard_photo
       WHERE check_id IN (${ph})
       ORDER BY check_id, uploaded_at_iso DESC, id DESC`,
    )
      .bind(...slice)
      .all<{ id: number; check_id: number | null }>();
    for (const row of r.results ?? []) {
      const cid = row.check_id;
      if (!cid) continue;
      let arr = byCheck.get(cid);
      if (!arr) {
        arr = [];
        byCheck.set(cid, arr);
      }
      arr.push(photoUrlFor(row.id));
    }
  }
  for (const c of checks) {
    c.checkPhotoUrls = byCheck.get(c.id) ?? [];
  }
}

type PhotoReadRow = {
  id: number;
  asset_id: string;
  check_id: number | null;
  r2_key: string;
  content_type: string | null;
  size_bytes: number | null;
  uploaded_by: string | null;
  uploaded_at_iso: string;
  caption: string | null;
};

export type AddPhotoInput = {
  assetId: string;
  checkId?: number | null;
  body: ArrayBuffer | ReadableStream;
  contentType: string;
  sizeBytes?: number;
  uploadedBy?: string;
  caption?: string;
};

/** Crockford-ish base32 random id; good enough to dedupe upload R2 keys. */
function randomId(): string {
  const bytes = new Uint8Array(12);
  crypto.getRandomValues(bytes);
  let s = "";
  for (const b of bytes) s += b.toString(16).padStart(2, "0");
  return s;
}

function extensionForContentType(ct: string): string {
  const c = ct.toLowerCase();
  if (c.includes("png")) return "png";
  if (c.includes("webp")) return "webp";
  if (c.includes("heic") || c.includes("heif")) return "heic";
  if (c.includes("gif")) return "gif";
  return "jpg";
}

export async function addPhoto(env: Env, input: AddPhotoInput): Promise<YardPhotoRow> {
  const assetId = input.assetId.trim();
  if (!assetId) throw new Error("assetId required");
  const ct = input.contentType || "image/jpeg";
  const ext = extensionForContentType(ct);
  const r2Key = YARD_PHOTO_PREFIX + assetId + "/" + Date.now() + "-" + randomId() + "." + ext;
  await env.ETIC_BUCKET.put(r2Key, input.body, {
    httpMetadata: { contentType: ct },
  });
  const nowIso = new Date().toISOString();
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO yard_photo
       (asset_id, check_id, r2_key, content_type, size_bytes, uploaded_by, uploaded_at_iso, caption)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)
     RETURNING id, asset_id, check_id, r2_key, content_type, size_bytes, uploaded_by, uploaded_at_iso, caption`,
  )
    .bind(
      assetId,
      input.checkId ?? null,
      r2Key,
      ct,
      input.sizeBytes ?? 0,
      (input.uploadedBy ?? "").trim(),
      nowIso,
      (input.caption ?? "").trim(),
    )
    .first<PhotoReadRow>();
  if (!r) throw new Error("failed to record photo metadata");
  return rowToPhoto(r);
}

export async function listPhotosForAsset(env: Env, assetId: string): Promise<YardPhotoRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, asset_id, check_id, r2_key, content_type, size_bytes, uploaded_by, uploaded_at_iso, caption
     FROM yard_photo
     WHERE UPPER(TRIM(asset_id)) = UPPER(TRIM(?))
     ORDER BY uploaded_at_iso DESC`,
  )
    .bind(assetId)
    .all<PhotoReadRow>();
  return (r.results ?? []).map(rowToPhoto);
}

export async function getPhoto(env: Env, photoId: number): Promise<{
  meta: YardPhotoRow;
  body: ReadableStream;
} | null> {
  const meta = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, asset_id, check_id, r2_key, content_type, size_bytes, uploaded_by, uploaded_at_iso, caption
     FROM yard_photo WHERE id = ?`,
  )
    .bind(photoId)
    .first<PhotoReadRow>();
  if (!meta) return null;
  const obj = await env.ETIC_BUCKET.get(meta.r2_key);
  if (!obj) return null;
  return { meta: rowToPhoto(meta), body: obj.body };
}

export async function deletePhoto(env: Env, photoId: number): Promise<boolean> {
  const meta = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT r2_key FROM yard_photo WHERE id = ?`,
  )
    .bind(photoId)
    .first<{ r2_key: string }>();
  if (!meta) return false;
  await env.ETIC_BUCKET.delete(meta.r2_key);
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `DELETE FROM yard_photo WHERE id = ?`,
  )
    .bind(photoId)
    .run();
  const m = (r as unknown as { meta?: { changes?: number } }).meta;
  return (m?.changes ?? 0) > 0;
}

/** When an asset has yard_check rows but getRollingRoster hasn't matched yet (race) or id casing edge case, synthesize a minimal row so the client can render. */
function rollingAssetSyntheticUnlisted(
  id: string,
  checks: YardCheckRow[],
  photos: YardPhotoRow[],
  intervalDays: number,
): RollingAsset {
  const last = checks[0];
  const nowIso = new Date().toISOString();
  const days = daysBetween(last.checkedAtIso, nowIso);
  const rollingState = bucketState(days, intervalDays);
  return {
    assetId: id,
    owningUnit: "",
    shop: "",
    mgmtCd: "",
    makeModel: "",
    vehNomen: "",
    melKey: "",
    melTier: "",
    vinSerial: "",
    previousLocation: "",
    openWoCount: 0,
    isNce: false,
    lastCheckedAtIso: last.checkedAtIso,
    lastCheckedBy: last.checkedBy,
    daysSinceLastCheck: days,
    rollingState,
    isNeverChecked: false,
    isNewAsset: false,
    isUnlisted: true,
    photoCount: photos.length,
    lastLocation: last.location || "",
    lastNotes: (last.discrepancies ?? "").trim(),
    isBelowMel: false,
  };
}

/** Detail view for one asset: roster info + check history + photos + WO context + FM&A actions. */
export type AssetDetail = {
  asset: RollingAsset | null;
  checks: YardCheckRow[];
  /** In-place corrections to checks, newest first (group client-side by checkId). */
  checkEdits: YardCheckEditRow[];
  photos: YardPhotoRow[];
  /** Open work orders against this asset on the latest snapshot, with remarks. */
  openWorkOrders: AssetWorkOrder[];
  /**
   * Open WOs for this asset keyed by the ETIC snapshot date that was "current"
   * when each check was saved (`yard_check.source_date_key`) — the book state
   * at the time the walker saw the unit.
   */
  workOrdersByCheckSourceDate: Record<string, AssetWorkOrder[]>;
  /** True if the asset_id has yard_check rows but isn't in the latest snapshot. */
  isUnlisted: boolean;
  /** FM&A actions taken on this asset, most recent first. */
  actions: YardFindingAction[];
};

export type AssetWorkOrder = {
  workOrderId: string;
  shop: string;
  remarks: string;
  partsStatus: string;
  eticDate: string;
  eticRaw: string;
  melTier: string;
  lastRemarkChangeDate: string;
  eticPushCount: number;
};

/** One row from work_order_snapshot (asset appeared on that day's ETIC). */
export type WorkOrderSnapshotHistoryRow = {
  snapshotDateKey: string;
  workOrderId: string;
  assetId: string;
  melTier: string;
  shop: string;
  eticDate: string;
  eticRaw: string;
  remarks: string;
  partsStatus: string;
  lastRemarkChangeDate: string;
  /** Fleet P&A FM&A notes column from raw_row_json for that snapshot row. */
  fleetFmaNotes: string;
};

/**
 * All stored snapshot rows for an asset (newest report date first), for FM&A
 * to see when this asset last had an open work order in the historical books.
 */
export async function listWorkOrderSnapshotsForAsset(
  env: Env,
  assetId: string,
  opts?: { limit?: number },
): Promise<WorkOrderSnapshotHistoryRow[]> {
  const lim = Math.min(500, Math.max(1, Math.floor(opts?.limit ?? 200)));
  const id = assetId.trim();
  if (!id) return [];
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT snapshot_date_key, work_order_id, asset_id, mel_tier, shop,
            etic_date, etic_raw, remarks, parts_status, last_remark_change_date, raw_row_json
     FROM work_order_snapshot
     WHERE UPPER(TRIM(asset_id)) = UPPER(TRIM(?))
     ORDER BY snapshot_date_key DESC, work_order_id ASC
     LIMIT ?`,
  )
    .bind(id, lim)
    .all<{
      snapshot_date_key: string;
      work_order_id: string;
      asset_id: string;
      mel_tier: string | null;
      shop: string | null;
      etic_date: string | null;
      etic_raw: string | null;
      remarks: string | null;
      parts_status: string | null;
      last_remark_change_date: string | null;
      raw_row_json: string | null;
    }>();
  return (r.results ?? []).map((row) => {
    let raw: Record<string, unknown> = {};
    try {
      raw = JSON.parse(row.raw_row_json || "{}") as Record<string, unknown>;
    } catch {
      raw = {};
    }
    return {
      snapshotDateKey: row.snapshot_date_key,
      workOrderId: row.work_order_id,
      assetId: row.asset_id ?? "",
      melTier: row.mel_tier ?? "",
      shop: row.shop ?? "",
      eticDate: row.etic_date ?? "",
      eticRaw: row.etic_raw ?? "",
      remarks: row.remarks ?? "",
      partsStatus: row.parts_status ?? "",
      lastRemarkChangeDate: row.last_remark_change_date ?? "",
      fleetFmaNotes: extractFleetFmaNotesFromRaw(raw),
    };
  });
}

type WoReadRow = {
  work_order_id: string;
  shop: string | null;
  remarks: string | null;
  parts_status: string | null;
  etic_date: string | null;
  etic_raw: string | null;
  mel_tier: string | null;
  last_remark_change_date: string | null;
  etic_push_count: number | null;
};

/** Latest-snapshot work orders for an asset (most-recent remarks first). */
async function getOpenWorkOrdersForAsset(
  env: Env,
  dateKey: string,
  assetId: string,
): Promise<AssetWorkOrder[]> {
  if (!dateKey || !assetId) return [];
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id, shop, remarks, parts_status, etic_date, etic_raw,
            mel_tier, last_remark_change_date, etic_push_count
     FROM work_order_snapshot
     WHERE snapshot_date_key = ? AND UPPER(TRIM(asset_id)) = UPPER(TRIM(?))
     ORDER BY last_remark_change_date DESC, work_order_id ASC`,
  )
    .bind(dateKey, assetId)
    .all<WoReadRow>();
  return (r.results ?? []).map((row) => ({
    workOrderId: row.work_order_id,
    shop: row.shop ?? "",
    remarks: row.remarks ?? "",
    partsStatus: row.parts_status ?? "",
    eticDate: row.etic_date ?? "",
    eticRaw: row.etic_raw ?? "",
    melTier: row.mel_tier ?? "",
    lastRemarkChangeDate: row.last_remark_change_date ?? "",
    eticPushCount: row.etic_push_count ?? 0,
  }));
}

export async function getAssetDetail(env: Env, assetId: string): Promise<AssetDetail> {
  const id = assetId.trim();
  const idCanon = canonicalYardAssetKey(id);
  const [roster, checks, checkEdits, photos, actions] = await Promise.all([
    getRollingRoster(env),
    getChecksForAsset(env, id),
    getCheckEditsForAsset(env, id),
    listPhotosForAsset(env, id),
    listFindingActionsForAsset(env, id),
  ]);
  const inRoster = roster.assets.find((a) => canonicalYardAssetKey(a.assetId) === idCanon) ?? null;
  let asset: RollingAsset | null = inRoster;
  if (!asset && checks.length > 0) {
    asset = rollingAssetSyntheticUnlisted(id, checks, photos, roster.intervalDays);
  }
  const openWorkOrders = await getOpenWorkOrdersForAsset(env, roster.dateKey, id);
  const sourceDateKeys = [...new Set(checks.map((c) => c.sourceDateKey).filter(Boolean))] as string[];
  const workOrdersByCheckSourceDate: Record<string, AssetWorkOrder[]> = {};
  await Promise.all(
    sourceDateKeys.map(async (dk) => {
      workOrdersByCheckSourceDate[dk] = await getOpenWorkOrdersForAsset(env, dk, id);
    }),
  );
  await attachCheckPhotoUrlsToChecks(env, checks);
  return {
    asset,
    checks,
    checkEdits,
    photos,
    openWorkOrders,
    workOrdersByCheckSourceDate,
    isUnlisted: inRoster ? inRoster.isUnlisted : checks.length > 0,
    actions,
  };
}

/* =============================================================================
   FM&A FOLLOW-UP QUEUE

   `listOpenFindings` returns only kinds that need explicit FM&A follow-up on
   walker-tagged evidence:

     - UNLISTED — walker found/logged an asset with no open WO on latest ETIC
     - DISCREPANCY — latest check has non-empty discrepancy text
     - UNKNOWN — legacy rows only

   Absence-derived "not checked in time" is NOT listed here: cadence / never
   checked is visible on the rolling fleet list (getRollingRoster). The
   MISSING kind remains in the schema for old yard_finding_action rows.

   FM&A records yard_finding_action rows; for DISCREPANCY/UNLISTED the action
   anchors to the triggering yard_check id so a newer check can re-open.
   ========================================================================== */

/** How long an asset can go un-confirmed before it counts as Missing. Defaults
 *  to 2× the freshness interval, i.e. "missed two full check cycles." */
export function getMissingThresholdDays(intervalDays: number): number {
  const v = Math.round(intervalDays * 2);
  return v >= 1 ? v : 14;
}

export type FindingKind = "missing" | "unlisted" | "discrepancy" | "unknown";

export type FindingResolution =
  | "resolved"
  | "in_progress"
  | "dismissed"
  | "wo_opened"
  | "retired"
  | "reassigned";

export type YardFindingAction = {
  id: number;
  assetId: string;
  kind: FindingKind;
  checkId: number | null;
  resolution: FindingResolution;
  woOpened: string;
  note: string;
  resolvedBy: string;
  resolvedAtIso: string;
};

export type YardFinding = {
  assetId: string;
  kind: FindingKind;
  /** The yard_check that triggered this finding (null for unlisted). */
  triggerCheck: YardCheckRow | null;
  /** Most-recent FM&A action for this asset+kind, or null if untouched. */
  lastAction: YardFindingAction | null;
  /** True if lastAction is still valid (matches the trigger check). */
  isAcknowledged: boolean;
  /** Snapshot of the asset's roster info at time of read (may be null for unlisted). */
  asset: RollingAsset | null;
  /** Photo count for quick badge rendering. */
  photoCount: number;
  /** Latest yard photo URL for list thumbnails (`/api/yard/photo/:id`), or null. */
  previewPhotoUrl: string | null;
  /** Fleet (P&A) merged columns on the WO row — FM&A / notes style headers when present. */
  fleetFmaNotes: string;
};

type FindingActionRow = {
  id: number;
  asset_id: string;
  kind: string;
  check_id: number | null;
  resolution: string;
  wo_opened: string | null;
  note: string | null;
  resolved_by: string | null;
  resolved_at_iso: string;
};

function normalizeFleetRawKey(k: string): string {
  return k.toLowerCase().replace(/\s+/g, " ").trim();
}

/**
 * Pull FM&A / fleet notes from merged workbook JSON on a WO row (Fleet P&A columns).
 */
export function extractFleetFmaNotesFromRaw(raw: Record<string, unknown> | null | undefined): string {
  if (!raw) return "";
  const preferred = [
    "fleet.fm&a notes",
    "fleet.fma notes",
    "fleet.fm and a notes",
    "fleet.f&m notes",
    "fleet.fma note",
    "fm&a notes",
    "fma notes",
  ];
  const byNorm = new Map<string, string>();
  for (const [k, v] of Object.entries(raw)) {
    const val = String(v ?? "").trim();
    if (!val) continue;
    byNorm.set(normalizeFleetRawKey(k), val);
  }
  for (const p of preferred) {
    const hit = byNorm.get(normalizeFleetRawKey(p));
    if (hit) return hit;
  }
  for (const [k, v] of Object.entries(raw)) {
    const nk = normalizeFleetRawKey(k);
    const val = String(v ?? "").trim();
    if (!val) continue;
    const hasFma = nk.includes("fm&a") || nk.includes("fma") || nk.includes("f&m");
    if (hasFma && nk.includes("note")) return val;
  }
  return "";
}

async function enrichFindingsPreviewPhotos(
  env: Env,
  findings: YardFinding[],
): Promise<void> {
  const ids = [...new Set(findings.map((f) => f.assetId).filter(Boolean))];
  if (!ids.length) return;
  const byAsset = new Map<string, number>();
  const chunk = 80;
  for (let i = 0; i < ids.length; i += chunk) {
    const slice = ids.slice(i, i + chunk);
    const u = slice.map((s) => canonicalYardAssetKey(s));
    const ph2 = u.map(() => "?").join(",");
    const r = await env.ETIC_SNAPSHOTS.prepare(
      `SELECT id, asset_id FROM (
         SELECT id, asset_id,
           ROW_NUMBER() OVER (PARTITION BY UPPER(TRIM(asset_id)) ORDER BY uploaded_at_iso DESC, id DESC) AS rn
         FROM yard_photo WHERE UPPER(TRIM(asset_id)) IN (${ph2})
       ) WHERE rn = 1`,
    )
      .bind(...u)
      .all<{ id: number; asset_id: string }>();
    for (const row of r.results ?? []) {
      if (row.asset_id) byAsset.set(canonicalYardAssetKey(row.asset_id), row.id);
    }
  }
  for (const f of findings) {
    const pid = byAsset.get(canonicalYardAssetKey(f.assetId));
    f.previewPhotoUrl = pid != null ? photoUrlFor(pid) : null;
  }
}

async function enrichFindingsFleetFmaNotes(
  env: Env,
  dateKey: string,
  findings: YardFinding[],
): Promise<void> {
  if (!findings.length) return;
  /** snapshot date key -> canonical asset ids to read FM&A columns for */
  const bySnapshot = new Map<string, Set<string>>();
  for (const f of findings) {
    const canon = canonicalYardAssetKey(f.assetId);
    if (!canon) continue;
    let dk = (dateKey || "").trim();
    if (!f.asset || f.asset.isUnlisted) {
      dk = (f.triggerCheck?.sourceDateKey ?? "").trim() || dk;
    }
    if (!dk) continue;
    let s = bySnapshot.get(dk);
    if (!s) {
      s = new Set();
      bySnapshot.set(dk, s);
    }
    s.add(canon);
  }
  const best = new Map<string, string>();
  const chunk = 80;
  for (const [dk, idSet] of bySnapshot) {
    const ids = [...idSet];
    for (let i = 0; i < ids.length; i += chunk) {
      const slice = ids.slice(i, i + chunk);
      const ph2 = slice.map(() => "?").join(",");
      const r = await env.ETIC_SNAPSHOTS.prepare(
        `SELECT asset_id, raw_row_json FROM work_order_snapshot
         WHERE snapshot_date_key = ? AND UPPER(TRIM(asset_id)) IN (${ph2})
         ORDER BY asset_id, work_order_id`,
      )
        .bind(dk, ...slice)
        .all<{ asset_id: string; raw_row_json: string | null }>();
      for (const row of r.results ?? []) {
        const aid = canonicalYardAssetKey(row.asset_id ?? "");
        if (!aid) continue;
        let raw: Record<string, unknown> = {};
        try {
          raw = JSON.parse(row.raw_row_json || "{}") as Record<string, unknown>;
        } catch {
          raw = {};
        }
        const note = extractFleetFmaNotesFromRaw(raw);
        if (!note) continue;
        const prev = best.get(aid) ?? "";
        if (!prev || note.length > prev.length) best.set(aid, note);
      }
    }
  }
  for (const f of findings) {
    f.fleetFmaNotes = best.get(canonicalYardAssetKey(f.assetId)) ?? "";
  }
}

async function hydrateFindingRollingAssetsBatch(
  env: Env,
  findings: YardFinding[],
  intervalDays: number,
): Promise<void> {
  const need = new Set<string>();
  for (const f of findings) {
    if (!f.triggerCheck) continue;
    const c = canonicalYardAssetKey(f.assetId);
    if (c) need.add(c);
  }
  const photoByCanon = await batchPhotoCountsForCanonicalAssets(env, need);
  const iv = intervalDays > 0 ? intervalDays : DEFAULT_YARD_CHECK_INTERVAL_DAYS;
  for (const f of findings) {
    hydrateFindingRollingAsset(f, iv, photoByCanon);
  }
}

const VALID_KINDS = new Set<FindingKind>(["missing", "unlisted", "discrepancy", "unknown"]);
const VALID_RESOLUTIONS = new Set<FindingResolution>([
  "resolved",
  "in_progress",
  "dismissed",
  "wo_opened",
  "retired",
  "reassigned",
]);

function rowToAction(r: FindingActionRow): YardFindingAction {
  const kind = (VALID_KINDS.has(r.kind as FindingKind) ? r.kind : "discrepancy") as FindingKind;
  const resolution = (
    VALID_RESOLUTIONS.has(r.resolution as FindingResolution) ? r.resolution : "resolved"
  ) as FindingResolution;
  return {
    id: r.id,
    assetId: r.asset_id,
    kind,
    checkId: r.check_id ?? null,
    resolution,
    woOpened: r.wo_opened ?? "",
    note: r.note ?? "",
    resolvedBy: r.resolved_by ?? "",
    resolvedAtIso: r.resolved_at_iso,
  };
}

export async function listFindingActionsForAsset(
  env: Env,
  assetId: string,
): Promise<YardFindingAction[]> {
  if (!assetId) return [];
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, asset_id, kind, check_id, resolution, wo_opened, note, resolved_by, resolved_at_iso
     FROM yard_finding_action
     WHERE UPPER(TRIM(asset_id)) = UPPER(TRIM(?))
     ORDER BY resolved_at_iso DESC`,
  )
    .bind(assetId)
    .all<FindingActionRow>();
  return (r.results ?? []).map(rowToAction);
}

/**
 * The full FM&A queue: every (asset, kind) pair that's currently "open"
 * according to the latest yard_check + the latest snapshot, plus its most-
 * recent action (if any).
 *
 * See the FM&A FOLLOW-UP QUEUE comment block above for the rules.
 */
export async function listOpenFindings(env: Env): Promise<{
  findings: YardFinding[];
  totals: Record<FindingKind | "total" | "acknowledged", number>;
}> {
  const dateKey = await getLatestSnapshotDateKey(env);

  // Fetch in parallel: rolling asset list (metadata + isUnlisted), latest
  // check of any status per asset (unlisted/discrepancy anchors), and the
  // latest FM&A action per (asset, kind).
  const [roster, latestAnyRows, latestActionRows] = await Promise.all([
    getRollingRoster(env),
    env.ETIC_SNAPSHOTS.prepare(
      `SELECT yc.* FROM yard_check yc
       JOIN (
         SELECT asset_id, MAX(checked_at_iso) AS m FROM yard_check GROUP BY asset_id
       ) m ON m.asset_id = yc.asset_id AND m.m = yc.checked_at_iso`,
    ).all<CheckReadRow>(),
    env.ETIC_SNAPSHOTS.prepare(
      `SELECT a.* FROM yard_finding_action a
       JOIN (
         SELECT asset_id, kind, MAX(resolved_at_iso) AS m FROM yard_finding_action
          GROUP BY asset_id, kind
       ) m ON m.asset_id = a.asset_id AND m.kind = a.kind AND m.m = a.resolved_at_iso`,
    ).all<FindingActionRow>(),
  ]);

  const rosterByCanon = new Map<string, RollingAsset>();
  for (const a of roster.assets) {
    rosterByCanon.set(canonicalYardAssetKey(a.assetId), a);
  }

  const latestAnyByAsset = new Map<string, YardCheckRow>();
  for (const row of mergeYardCheckRowsByCanonical(latestAnyRows.results ?? [])) {
    if (!row.asset_id) continue;
    const c = canonicalYardAssetKey(row.asset_id);
    const displayId = rosterByCanon.get(c)?.assetId ?? row.asset_id.trim();
    latestAnyByAsset.set(displayId, rowToCheck({ ...row, asset_id: displayId }));
  }

  const actionByKey = new Map<string, YardFindingAction>();
  for (const row of latestActionRows.results ?? []) {
    if (!row.asset_id) continue;
    const action = rowToAction(row);
    actionByKey.set(canonicalYardAssetKey(action.assetId) + "|" + action.kind, action);
  }

  const assetByKey = rosterByCanon;

  const byDateKeyForWoFallback = new Map<string, Set<string>>();
  for (const [assetId, latest] of latestAnyByAsset) {
    if (openWoCountFromCheckSnapshot(latest) !== null) continue;
    const dk = (latest.sourceDateKey ?? "").trim();
    if (!dk) continue;
    const c = canonicalYardAssetKey(assetId);
    let s = byDateKeyForWoFallback.get(dk);
    if (!s) {
      s = new Set();
      byDateKeyForWoFallback.set(dk, s);
    }
    s.add(c);
  }
  const woPresenceAtCheck = await batchWorkOrderPresenceOnSnapshots(env, byDateKeyForWoFallback);

  const findings: YardFinding[] = [];
  const totals: Record<FindingKind | "total" | "acknowledged", number> = {
    missing: 0,
    unlisted: 0,
    discrepancy: 0,
    unknown: 0,
    total: 0,
    acknowledged: 0,
  };

  function pushFinding(assetId: string, kind: FindingKind, triggerCheck: YardCheckRow | null) {
    const c = canonicalYardAssetKey(assetId);
    const asset = assetByKey.get(c) ?? null;
    const lastAction = actionByKey.get(c + "|" + kind) ?? null;
    // Acknowledgment rules per kind:
    //   - discrepancy / unknown:  action.check_id matches the triggering check
    //   - unlisted:               any action newer than the most-recent check
    //   - missing (absence):      any action newer than the most-recent
    //                             *present* check (or, if never seen, just any
    //                             action — there's nothing to be newer than)
    let acknowledged = false;
    if (lastAction) {
      if (kind === "unlisted") {
        const lastCheck = triggerCheck;
        acknowledged = !lastCheck || lastAction.resolvedAtIso >= lastCheck.checkedAtIso;
      } else if (kind === "missing") {
        // For absence-derived missing, "newer than the last sighting" is the
        // right test: if FM&A acted on it and then a walker logged a new
        // present check, the finding would have evaporated entirely; if the
        // action came AFTER the last sighting, FM&A has the latest info.
        const lastSeen = triggerCheck;
        acknowledged = !lastSeen || lastAction.resolvedAtIso >= lastSeen.checkedAtIso;
      } else {
        acknowledged = !!triggerCheck && lastAction.checkId === triggerCheck.id;
      }
    }
    findings.push({
      assetId,
      kind,
      triggerCheck,
      lastAction,
      isAcknowledged: acknowledged,
      asset,
      photoCount: asset?.photoCount ?? 0,
      previewPhotoUrl: null,
      fleetFmaNotes: "",
    });
    totals.total += 1;
    totals[kind] += 1;
    if (acknowledged) totals.acknowledged += 1;
  }

  // Absence-derived "missing" / not-seen items are intentionally NOT listed here:
  // the embedded fleet list (walker view) already shows due / overdue / never
  // checked. This queue is only for walker-tagged or floor-to-book issues.

  // ---- UNLISTED + DISCREPANCY (and legacy UNKNOWN) ------------------------
  // Walk every asset that has any check history.
  for (const [assetId, latest] of latestAnyByAsset) {
    const ra = assetByKey.get(canonicalYardAssetKey(assetId));
    const hasCurrentOpenWo = !!ra && !ra.isUnlisted && (ra.openWoCount ?? 0) > 0;
    const fromJson = openWoCountFromCheckSnapshot(latest);
    let hadOpenWoWhenWalkerSawIt = false;
    if (fromJson !== null) {
      hadOpenWoWhenWalkerSawIt = fromJson > 0;
    } else {
      const dk = (latest.sourceDateKey ?? "").trim();
      if (dk) {
        hadOpenWoWhenWalkerSawIt = woPresenceAtCheck.has(dk + "|" + canonicalYardAssetKey(assetId));
      }
    }
    if (!hasCurrentOpenWo && !hadOpenWoWhenWalkerSawIt) {
      pushFinding(assetId, "unlisted", latest);
    }
    if (latest.discrepancies && latest.discrepancies.trim()) {
      pushFinding(assetId, "discrepancy", latest);
    }
    // Legacy: surface old walker-tagged 'unknown' rows so they don't get
    // silently buried. Walker no longer produces these; once the backlog is
    // cleaned up this branch can be deleted.
    if (latest.status === "unknown") pushFinding(assetId, "unknown", latest);
  }

  await hydrateFindingRollingAssetsBatch(env, findings, roster.intervalDays);
  await Promise.all([enrichFindingsPreviewPhotos(env, findings), enrichFindingsFleetFmaNotes(env, dateKey, findings)]);

  // Sort: unacknowledged first, then by trigger check age (newest first, with
  // "never seen" sorting first within the missing kind), then assetId.
  findings.sort((a, b) => {
    if (a.isAcknowledged !== b.isAcknowledged) return a.isAcknowledged ? 1 : -1;
    const at = a.triggerCheck?.checkedAtIso ?? "";
    const bt = b.triggerCheck?.checkedAtIso ?? "";
    if (at !== bt) return at < bt ? 1 : -1;
    return a.assetId.localeCompare(b.assetId, undefined, { numeric: true });
  });

  return { findings, totals };
}

export type ResolveFindingInput = {
  assetId: string;
  kind: FindingKind;
  /** check_id this action is tied to. Required for non-unlisted kinds. */
  checkId?: number | null;
  resolution: FindingResolution;
  woOpened?: string;
  note?: string;
  resolvedBy?: string;
};

export async function resolveFinding(
  env: Env,
  input: ResolveFindingInput,
): Promise<YardFindingAction> {
  const assetId = input.assetId.trim();
  if (!assetId) throw new Error("assetId required");
  if (!VALID_KINDS.has(input.kind)) throw new Error("invalid kind");
  if (!VALID_RESOLUTIONS.has(input.resolution)) throw new Error("invalid resolution");
  const nowIso = new Date().toISOString();
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO yard_finding_action
       (asset_id, kind, check_id, resolution, wo_opened, note, resolved_by, resolved_at_iso)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)
     RETURNING id, asset_id, kind, check_id, resolution, wo_opened, note, resolved_by, resolved_at_iso`,
  )
    .bind(
      assetId,
      input.kind,
      input.checkId ?? null,
      input.resolution,
      (input.woOpened ?? "").trim(),
      (input.note ?? "").trim(),
      (input.resolvedBy ?? "").trim(),
      nowIso,
    )
    .first<FindingActionRow>();
  if (!r) throw new Error("failed to record action");
  return rowToAction(r);
}

/** Reopen a finding by deleting its most-recent action for (asset, kind). */
export async function reopenFinding(
  env: Env,
  assetId: string,
  kind: FindingKind,
): Promise<boolean> {
  if (!VALID_KINDS.has(kind)) throw new Error("invalid kind");
  const last = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id FROM yard_finding_action
     WHERE asset_id = ? AND kind = ? ORDER BY resolved_at_iso DESC LIMIT 1`,
  )
    .bind(assetId, kind)
    .first<{ id: number }>();
  if (!last) return false;
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `DELETE FROM yard_finding_action WHERE id = ?`,
  )
    .bind(last.id)
    .run();
  const meta = (r as unknown as { meta?: { changes?: number } }).meta;
  return (meta?.changes ?? 0) > 0;
}

/** Recent activity feed: last N checks + actions, interleaved by timestamp. */
export type YardActivityItem =
  | { kind: "check"; at: string; check: YardCheckRow }
  | { kind: "action"; at: string; action: YardFindingAction };

export async function getRecentActivity(
  env: Env,
  limit = 100,
  range?: { fromIso?: string; toIso?: string },
): Promise<YardActivityItem[]> {
  const fromIso = (range?.fromIso ?? "").trim();
  const toIso = (range?.toIso ?? "").trim();
  const checkWhere: string[] = [];
  const actionWhere: string[] = [];
  const checkParams: string[] = [];
  const actionParams: string[] = [];
  if (fromIso) {
    checkWhere.push("checked_at_iso >= ?");
    actionWhere.push("resolved_at_iso >= ?");
    checkParams.push(fromIso);
    actionParams.push(fromIso);
  }
  if (toIso) {
    checkWhere.push("checked_at_iso < ?");
    actionWhere.push("resolved_at_iso < ?");
    checkParams.push(toIso);
    actionParams.push(toIso);
  }
  const checkSql =
    `SELECT id, asset_id, location, discrepancies, status, checked_by, checked_at_iso, source_date_key, snapshot_asset_json
       FROM yard_check` +
    (checkWhere.length ? ` WHERE ${checkWhere.join(" AND ")}` : "") +
    ` ORDER BY checked_at_iso DESC LIMIT ?`;
  const actionSql =
    `SELECT id, asset_id, kind, check_id, resolution, wo_opened, note, resolved_by, resolved_at_iso
       FROM yard_finding_action` +
    (actionWhere.length ? ` WHERE ${actionWhere.join(" AND ")}` : "") +
    ` ORDER BY resolved_at_iso DESC LIMIT ?`;
  const [checks, actions] = await Promise.all([
    env.ETIC_SNAPSHOTS.prepare(checkSql).bind(...checkParams, limit).all<CheckReadRow>(),
    env.ETIC_SNAPSHOTS.prepare(actionSql).bind(...actionParams, limit).all<FindingActionRow>(),
  ]);
  const items: YardActivityItem[] = [
    ...(checks.results ?? []).map((r) => ({ kind: "check" as const, at: r.checked_at_iso, check: rowToCheck(r) })),
    ...(actions.results ?? []).map((r) => {
      const a = rowToAction(r);
      return { kind: "action" as const, at: a.resolvedAtIso, action: a };
    }),
  ];
  items.sort((a, b) => (a.at < b.at ? 1 : a.at > b.at ? -1 : 0));
  return items.slice(0, limit);
}

import type { FleetRecord, RawWorkOrder } from "./yardCheck";
import {
  extractFleetMapsFromBinary,
  extractRawWorkOrdersFromBinary,
  FLEET_SYNONYMS,
  WORK_ORDER_SYNONYMS,
  scoreHeaderMatch,
} from "./yardCheck";
import { getStalenessThresholds } from "./melWatch";
import { buildLineCompletions } from "./meeting";

export type MelTier = "below" | "at" | "above" | "unknown";

/** Default staleness in days between expected remarks. Mirrored in melWatch.ts
 *  so the Settings tab can override these without touching code. */
const REMARK_INTERVAL_DEFAULT: Record<MelTier, number | null> = {
  below: 3,
  at: 5,
  above: 10,
  unknown: null,
};

function resolveStaleness(thresholds?: Record<string, number>): Record<MelTier, number | null> {
  if (!thresholds) return REMARK_INTERVAL_DEFAULT;
  return {
    below: typeof thresholds.below === "number" ? thresholds.below : REMARK_INTERVAL_DEFAULT.below,
    at: typeof thresholds.at === "number" ? thresholds.at : REMARK_INTERVAL_DEFAULT.at,
    above: typeof thresholds.above === "number" ? thresholds.above : REMARK_INTERVAL_DEFAULT.above,
    unknown: null,
  };
}

export function classifyMelTier(raw: string): MelTier {
  const s = raw.replace(/\s+/g, " ").trim().toLowerCase();
  if (!s) return "unknown";
  if (/\bbelow\b/i.test(raw) || /\bunder\b/i.test(raw) || /<\s*mel/i.test(raw) || /mel\s*[<≤]/i.test(raw)) {
    return "below";
  }
  if (/\babove\b/i.test(raw) || />\s*mel/i.test(raw) || /mel\s*[>≥]/i.test(raw)) {
    return "above";
  }
  if (/\bat\b/i.test(raw) && /mel/i.test(raw)) return "at";
  if (/^mel\s*[-–]?\s*at\b/i.test(s) || /^at\s+mel\b/i.test(s)) return "at";
  if (/^mel\s*[-–]?\s*below\b/i.test(s) || /^below\s+mel\b/i.test(s)) return "below";
  if (/^mel\s*[-–]?\s*above\b/i.test(s) || /^above\s+mel\b/i.test(s)) return "above";
  const digitOnly = s.replace(/[^\d]/g, "");
  if (digitOnly.length === 1) {
    const n = Number.parseInt(digitOnly, 10);
    if (n <= 2) return "below";
    if (n >= 4) return "above";
    if (n === 3) return "at";
  }
  return "unknown";
}

const MONTH_TOKEN: Record<string, number> = {
  jan: 1,
  feb: 2,
  mar: 3,
  apr: 4,
  may: 5,
  jun: 6,
  jul: 7,
  aug: 8,
  sep: 9,
  oct: 10,
  nov: 11,
  dec: 12,
};

/** Parse ETIC / due date from workbook cell text or Excel date serialization. */
export function parseEticDate(raw: string): string | null {
  const t = raw.replace(/\s+/g, " ").trim();
  if (!t) return null;
  const iso = /^\d{4}-\d{2}-\d{2}/.exec(t);
  if (iso) return iso[0] ?? null;
  const dm = /\b(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})\b/.exec(t);
  if (dm) {
    const mm = Number.parseInt(dm[1] ?? "", 10);
    const dd = Number.parseInt(dm[2] ?? "", 10);
    let yy = Number.parseInt(dm[3] ?? "", 10);
    if (!Number.isFinite(mm) || !Number.isFinite(dd) || !Number.isFinite(yy)) return null;
    if (yy < 100) yy = yy >= 70 ? 1900 + yy : 2000 + yy;
    if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
      return `${yy}-${String(mm).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
    }
  }
  const dmy = /\b(\d{1,2})[-/]([A-Za-z]{3})[-/](\d{2,4})\b/i.exec(t);
  if (dmy) {
    const day = Number.parseInt(dmy[1] ?? "", 10);
    const mon = MONTH_TOKEN[(dmy[2] ?? "").toLowerCase().slice(0, 3)] ?? 0;
    let yr = Number.parseInt(dmy[3] ?? "", 10);
    if (!mon || !Number.isFinite(day) || !Number.isFinite(yr)) return null;
    if (yr < 100) yr = yr >= 70 ? 1900 + yr : 2000 + yr;
    return `${yr}-${String(mon).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }
  const tryDate = new Date(t);
  if (!Number.isNaN(tryDate.getTime())) {
    const y = tryDate.getFullYear();
    const m = tryDate.getMonth() + 1;
    const d = tryDate.getDate();
    return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  return null;
}

function dateKeyToUtcNoon(dateKey: string): Date {
  const [y, m, d] = dateKey.split("-").map((x) => Number.parseInt(x, 10));
  return new Date(Date.UTC(y, (m ?? 1) - 1, d ?? 1, 12, 0, 0));
}

/** Whole calendar days from a to b (a,b are YYYY-MM-DD). */
export function calendarDaysBetween(aKey: string, bKey: string): number {
  const a = dateKeyToUtcNoon(aKey).getTime();
  const b = dateKeyToUtcNoon(bKey).getTime();
  return Math.floor((b - a) / (86400 * 1000));
}

type WoStateRow = {
  work_order_id: string;
  asset_id: string;
  last_snapshot_date: string;
  remarks: string;
  parts_status: string;
  etic_raw: string;
  etic_date: string | null;
  mel_tier: string;
  last_remark_change_date: string;
  etic_push_count: number;
  first_etic_date: string | null;
  last_etic_date: string | null;
  cumulative_etic_slip_days: number;
  owning_unit: string;
  mel_key: string;
  shop: string;
  mgmt_cd: string;
  make_model: string;
  veh_nomen: string;
  updated_at_iso: string;
};

export async function ingestWorkOrderSnapshot(
  env: { ETIC_SNAPSHOTS: D1Database },
  dateKey: string,
  rows: RawWorkOrder[],
  updatedAtIso: string,
): Promise<void> {
  const cleaned = rows
    .map((wo) => ({
      wid: (wo.workOrderId ?? "").trim(),
      assetId: (wo.assetId ?? "").trim(),
      remarks: (wo.remarks ?? "").trim(),
      partsStatus: (wo.partsStatus ?? "").trim(),
      eticRaw: (wo.eticDue ?? "").trim(),
      melTier: classifyMelTier(wo.currentMel ?? ""),
      owningUnit: (wo.owningUnit ?? "").trim(),
      melKey: (wo.melKey ?? "").trim(),
      shop: (wo.shop ?? "").trim(),
      mgmtCd: (wo.mgmtCd ?? "").trim(),
      makeModel: (wo.makeModel ?? "").trim(),
      vehNomen: (wo.vehNomen ?? "").trim(),
      // Cache the entire parsed row (mapped + unmapped headers) so we can
      // later derive new typed columns from D1 without re-reading R2.
      rawRowJson: JSON.stringify(wo.rawColumns ?? {}),
    }))
    .filter((w) => w.wid.length > 0);
  if (cleaned.length === 0) return;

  const dedup = new Map<string, (typeof cleaned)[number]>();
  for (const c of cleaned) dedup.set(c.wid, c);
  const workOrders = [...dedup.values()];

  const priorByWid = new Map<string, WoStateRow>();
  const selectBatchSize = 90;
  for (let i = 0; i < workOrders.length; i += selectBatchSize) {
    const chunk = workOrders.slice(i, i + selectBatchSize);
    const placeholders = chunk.map(() => "?").join(",");
    const stmt = env.ETIC_SNAPSHOTS.prepare(
      `SELECT work_order_id, asset_id, remarks, parts_status, etic_raw, etic_date, mel_tier,
              last_remark_change_date, etic_push_count, first_etic_date, last_etic_date, cumulative_etic_slip_days,
              owning_unit, mel_key, shop, mgmt_cd, make_model, veh_nomen,
              snapshot_date_key AS last_snapshot_date
       FROM work_order_snapshot
       WHERE work_order_id IN (${placeholders})
         AND snapshot_date_key <= ?
       ORDER BY work_order_id ASC, snapshot_date_key DESC`,
    ).bind(...chunk.map((c) => c.wid), dateKey);
    const r = await stmt.all<WoStateRow>();
    // First row per WO wins (ORDER BY snapshot_date_key DESC): same-day
    // re-ingest uses this morning's row as prev, not yesterday's.
    for (const row of r.results ?? []) {
      if (priorByWid.has(row.work_order_id)) continue;
      priorByWid.set(row.work_order_id, row);
    }
  }

  // INSERT OR REPLACE pairs with the UNIQUE(work_order_id, snapshot_date_key, field)
  // index added in migration 0017. Re-ingesting the same snapshot now refreshes
  // that date's change row (instead of leaving stale values), while still
  // preventing duplicate timeline rows.
  const insertLog = env.ETIC_SNAPSHOTS.prepare(
    `INSERT OR REPLACE INTO work_order_changelog (work_order_id, snapshot_date_key, changed_at_iso, field, old_value, new_value)
     VALUES (?, ?, ?, ?, ?, ?)`,
  );
  const upsert = env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO work_order_state (
      work_order_id, asset_id, last_snapshot_date, remarks, parts_status, etic_raw, etic_date, mel_tier,
      last_remark_change_date, etic_push_count, first_etic_date, last_etic_date, cumulative_etic_slip_days,
      owning_unit, mel_key, shop, mgmt_cd, make_model, veh_nomen, raw_row_json, updated_at_iso )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(work_order_id) DO UPDATE SET
      asset_id = excluded.asset_id,
      last_snapshot_date = excluded.last_snapshot_date,
      remarks = excluded.remarks,
      parts_status = excluded.parts_status,
      etic_raw = excluded.etic_raw,
      etic_date = excluded.etic_date,
      mel_tier = excluded.mel_tier,
      last_remark_change_date = excluded.last_remark_change_date,
      etic_push_count = excluded.etic_push_count,
      first_etic_date = excluded.first_etic_date,
      last_etic_date = excluded.last_etic_date,
      cumulative_etic_slip_days = excluded.cumulative_etic_slip_days,
      owning_unit = excluded.owning_unit,
      mel_key = excluded.mel_key,
      shop = excluded.shop,
      mgmt_cd = excluded.mgmt_cd,
      make_model = excluded.make_model,
      veh_nomen = excluded.veh_nomen,
      raw_row_json = excluded.raw_row_json,
      updated_at_iso = excluded.updated_at_iso
    WHERE excluded.last_snapshot_date >= work_order_state.last_snapshot_date`,
  );
  const snapUpsert = env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO work_order_snapshot (
      snapshot_date_key, work_order_id, asset_id, remarks, parts_status, etic_raw, etic_date, mel_tier,
      last_remark_change_date, etic_push_count, first_etic_date, last_etic_date, cumulative_etic_slip_days,
      owning_unit, mel_key, shop, mgmt_cd, make_model, veh_nomen, raw_row_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(snapshot_date_key, work_order_id) DO UPDATE SET
      asset_id = excluded.asset_id,
      remarks = excluded.remarks,
      parts_status = excluded.parts_status,
      etic_raw = excluded.etic_raw,
      etic_date = excluded.etic_date,
      mel_tier = excluded.mel_tier,
      last_remark_change_date = excluded.last_remark_change_date,
      etic_push_count = excluded.etic_push_count,
      first_etic_date = excluded.first_etic_date,
      last_etic_date = excluded.last_etic_date,
      cumulative_etic_slip_days = excluded.cumulative_etic_slip_days,
      owning_unit = excluded.owning_unit,
      mel_key = excluded.mel_key,
      shop = excluded.shop,
      mgmt_cd = excluded.mgmt_cd,
      make_model = excluded.make_model,
      veh_nomen = excluded.veh_nomen,
      raw_row_json = excluded.raw_row_json`,
  );

  const statements: D1PreparedStatement[] = [];
  for (const wo of workOrders) {
    const { wid, assetId, remarks, partsStatus, eticRaw, melTier, owningUnit, melKey, shop, mgmtCd, makeModel, vehNomen, rawRowJson } = wo;
    const eticDate = parseEticDate(eticRaw);
    const prev = priorByWid.get(wid) ?? null;

    let lastRemarkChange = dateKey;
    let pushCount = 0;
    let firstEtic: string | null = eticDate;
    let lastEtic: string | null = eticDate;
    let slipSum = 0;

    if (prev) {
      lastRemarkChange = prev.last_remark_change_date;
      if (remarks !== (prev.remarks ?? "")) {
        lastRemarkChange = dateKey;
        statements.push(insertLog.bind(wid, dateKey, updatedAtIso, "remarks", prev.remarks ?? "", remarks));
      }
      if (partsStatus !== (prev.parts_status ?? "")) {
        statements.push(insertLog.bind(wid, dateKey, updatedAtIso, "parts_status", prev.parts_status ?? "", partsStatus));
      }
      if (eticRaw !== (prev.etic_raw ?? "")) {
        statements.push(insertLog.bind(wid, dateKey, updatedAtIso, "etic", prev.etic_raw ?? "", eticRaw));
      }
      if (melTier !== (prev.mel_tier as MelTier)) {
        statements.push(insertLog.bind(wid, dateKey, updatedAtIso, "mel_tier", prev.mel_tier ?? "", melTier));
      }
      if (shop !== (prev.shop ?? "")) {
        statements.push(insertLog.bind(wid, dateKey, updatedAtIso, "shop", prev.shop ?? "", shop));
      }
      pushCount = prev.etic_push_count ?? 0;
      firstEtic = prev.first_etic_date ?? eticDate;
      slipSum = prev.cumulative_etic_slip_days ?? 0;
      lastEtic = prev.last_etic_date ?? null;

      const oldD = prev.etic_date;
      const newD = eticDate;
      if (oldD && newD && newD > oldD) {
        pushCount += 1;
        slipSum += calendarDaysBetween(oldD, newD);
        statements.push(insertLog.bind(wid, dateKey, updatedAtIso, "etic_date_slip", oldD, newD));
      }
      if (!firstEtic && newD) firstEtic = newD;
      if (newD) lastEtic = newD;
      else if (lastEtic === null) lastEtic = prev.last_etic_date;
    } else {
      statements.push(insertLog.bind(wid, dateKey, updatedAtIso, "initial", "", "first_seen"));
      firstEtic = eticDate;
      lastEtic = eticDate;
    }

    statements.push(
      upsert.bind(
        wid,
        assetId,
        dateKey,
        remarks,
        partsStatus,
        eticRaw,
        eticDate,
        melTier,
        lastRemarkChange,
        pushCount,
        firstEtic,
        lastEtic,
        slipSum,
        owningUnit,
        melKey,
        shop,
        mgmtCd,
        makeModel,
        vehNomen,
        rawRowJson,
        updatedAtIso,
      ),
    );
    statements.push(
      snapUpsert.bind(
        dateKey,
        wid,
        assetId,
        remarks,
        partsStatus,
        eticRaw,
        eticDate,
        melTier,
        lastRemarkChange,
        pushCount,
        firstEtic,
        lastEtic,
        slipSum,
        owningUnit,
        melKey,
        shop,
        mgmtCd,
        makeModel,
        vehNomen,
        rawRowJson,
      ),
    );
  }

  const BATCH = 50;
  for (let i = 0; i < statements.length; i += BATCH) {
    await env.ETIC_SNAPSHOTS.batch(statements.slice(i, i + BATCH));
  }

  // Verify any pending FM&A actions against the changes we just wrote.
  // Soft-fails so a verifier bug never blocks an ingest.
  try {
    await verifyWorkOrderActionsForSnapshot(env, dateKey);
  } catch (err) {
    console.error("verifyWorkOrderActionsForSnapshot failed", err);
  }
}

/** Schedule maintenance (Fleet P&A Schedule Mx / due columns). */
export type ScheduleMxBucket = "missing" | "no_due" | "overdue" | "due_soon" | "ok";

export type WatchRow = {
  workOrderId: string;
  assetId: string;
  melTier: MelTier;
  partsStatus: string;
  eticRaw: string;
  eticDate: string | null;
  remarks: string;
  lastRemarkChangeDate: string;
  daysSinceRemarkChange: number;
  requiredIntervalDays: number | null;
  remarkStale: boolean;
  eticPushCount: number;
  cumulativeEticSlipDays: number;
  firstEticDate: string | null;
  lastEticDate: string | null;
  lastSnapshotDate: string;
  owningUnit: string;
  melKey: string;
  shop: string;
  mgmtCd: string;
  makeModel: string;
  vehNomen: string;
  firstSeenDate: string;
  historyBounded: boolean;
  establishedDate: string;
  establishedDateIso: string | null;
  woReason: string;
  nce: boolean;
  nceStatus: string;
  /** Raw Schedule Mx Status / due cell text from Fleet (P&A). */
  scheduleMxStatus: string;
  scheduleMxDueIso: string | null;
  /** Days until scheduled/call-in date vs snapshot; not used to assert overdue (slicer does). */
  scheduleMxDaysUntil: number | null;
  /** Reserved; overdue is slicer-driven — no calendar day count. */
  scheduleMxOverdueByDays: number | null;
  scheduleMxBucket: ScheduleMxBucket;
  /** No fleet schedule-mx / due data — should be added in the workbook. */
  scheduleMxNeedsEntry: boolean;
  /** Fleet / workbook "days down" style column when present in raw row JSON. */
  daysDown: number | null;
};

type WatchReadRow = {
  work_order_id: string;
  asset_id: string;
  mel_tier: string;
  parts_status: string;
  etic_raw: string;
  etic_date: string | null;
  remarks: string;
  last_remark_change_date: string;
  etic_push_count: number;
  cumulative_etic_slip_days: number;
  first_etic_date: string | null;
  last_etic_date: string | null;
  last_snapshot_date: string;
  owning_unit: string | null;
  mel_key: string | null;
  shop: string | null;
  mgmt_cd: string | null;
  make_model: string | null;
  veh_nomen: string | null;
  raw_row_json: string | null;
};

/**
 * Pull "extra" typed values out of the raw row JSON we already cache in D1.
 * Headers are normalized (lowercased, whitespace-collapsed) by the workbook
 * extractor, so we match against several known variants without rerunning
 * an ingest.
 */
function readRawColumns(json: string | null): Record<string, string> {
  if (!json) return {};
  try {
    const parsed = JSON.parse(json);
    if (parsed && typeof parsed === "object") return parsed as Record<string, string>;
  } catch (_e) {
    /* corrupt JSON — treat as empty */
  }
  return {};
}

function pickRawValue(raw: Record<string, string>, candidates: string[]): string {
  for (const c of candidates) {
    const v = raw[c];
    if (v && v.trim()) return v.trim();
  }
  // Fuzzy fallback: substring match on the candidate.
  const lowerKeys = Object.keys(raw);
  for (const c of candidates) {
    const found = lowerKeys.find((k) => k.includes(c));
    if (found && raw[found] && raw[found].trim()) return raw[found].trim();
  }
  return "";
}

function pickDaysDownFromRaw(raw: Record<string, string>): number | null {
  const s = pickRawValue(raw, [
    "days down",
    "fleet.days down",
    "day down",
    "fleet.day down",
  ]);
  if (!s) return null;
  const n = Number.parseFloat(String(s).replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? Math.round(n) : null;
}

function findScheduleMxStatus(raw: Record<string, string>): string {
  const direct = pickRawValue(raw, [
    "fleet.schedule mx status",
    "fleet.schedule maintenance status",
    "schedule mx status",
    "schedule maintenance status",
    "scheduled mx status",
  ]);
  if (direct) return direct;
  for (const [k, v] of Object.entries(raw)) {
    if (!v?.trim()) continue;
    const kn = k.toLowerCase();
    if (kn.includes("wo inquiry")) continue;
    if (kn.includes("schedule mx status")) return v.trim();
    if (kn.includes("schedule maintenance") && kn.includes("status")) return v.trim();
  }
  return "";
}

function findScheduleMxDueText(raw: Record<string, string>): string {
  const direct = pickRawValue(raw, [
    "fleet.next schedule mx",
    "fleet.schedule mx due",
    "fleet.next scheduled maintenance",
    "next schedule mx",
    "schedule mx due",
    "next sched mx",
  ]);
  if (direct) return direct;
  for (const [k, v] of Object.entries(raw)) {
    if (!v?.trim()) continue;
    const kn = k.toLowerCase();
    if (kn.includes("next schedule mx") || kn.includes("schedule mx due") || kn.includes("next sched mx")) {
      return v.trim();
    }
  }
  return "";
}

/** Excel "Schedule Mx Slicer" — authoritative for whether work is overdue (vs scheduled call-in dates). */
function findScheduleMxSlicer(raw: Record<string, string>): string {
  const direct = pickRawValue(raw, [
    "fleet.schedule mx slicer",
    "schedule mx slicer",
    "scheduled mx slicer",
  ]);
  if (direct) return direct;
  for (const [k, v] of Object.entries(raw)) {
    if (!v?.trim()) continue;
    const kn = k.toLowerCase();
    if (kn.includes("wo inquiry")) continue;
    if (kn.includes("schedule mx slicer") || (kn.includes("slicer") && kn.includes("schedule mx"))) {
      return v.trim();
    }
  }
  return "";
}

function slicerSaysOverdue(slicerText: string): boolean {
  const s = slicerText.replace(/\s+/g, " ").trim().toLowerCase();
  return /\boverdue\b|\bpast\s*due\b/i.test(s);
}

/**
 * Derive schedule-maintenance health from merged Fleet (P&A) columns vs snapshot date.
 * Matches headers like "Schedule Mx Status" → `fleet.schedule mx status` in raw JSON.
 * **Overdue** follows the workbook "Schedule Mx Slicer" only: unless that cell says overdue,
 * we do not mark overdue from scheduled/call-in dates alone.
 */
export function analyzeScheduleMxFromRaw(
  raw: Record<string, string>,
  asOfDateKey: string,
): {
  scheduleMxStatus: string;
  scheduleMxDueIso: string | null;
  scheduleMxDaysUntil: number | null;
  scheduleMxOverdueByDays: number | null;
  scheduleMxBucket: ScheduleMxBucket;
  scheduleMxNeedsEntry: boolean;
} {
  const DUE_SOON_DAYS = 30;
  const slicer = findScheduleMxSlicer(raw);
  const status = findScheduleMxStatus(raw);
  const dueText = findScheduleMxDueText(raw);
  let dueIso = parseEticDate(dueText);
  if (!dueIso && status) dueIso = parseEticDate(status);

  const statusLower = status.toLowerCase();
  const slicerOverdue = slicerSaysOverdue(slicer);

  let daysUntil: number | null = null;
  if (dueIso) {
    daysUntil = calendarDaysBetween(asOfDateKey, dueIso);
  }
  const overdueByDays: number | null = null;

  const hasAnyText = !!(slicer.trim() || status.trim() || dueText.trim());
  const needsScheduleMxEntry = !hasAnyText && !dueIso;

  let bucket: ScheduleMxBucket;
  if (needsScheduleMxEntry) {
    bucket = "missing";
  } else if (slicerOverdue) {
    bucket = "overdue";
  } else if (dueIso !== null && daysUntil !== null && daysUntil >= 0 && daysUntil <= DUE_SOON_DAYS) {
    bucket = "due_soon";
  } else if (!dueIso && hasAnyText) {
    if (/\bcurrent\b|\bgreen\b|\bok\b|\bcompliant\b|\bontime\b|\bon[- ]?track\b/i.test(statusLower)) bucket = "ok";
    else bucket = "no_due";
  } else {
    bucket = "ok";
  }

  const displayParts = [slicer.trim(), status.trim(), dueText.trim()].filter(Boolean);
  const displayStatus = displayParts.length ? displayParts.join(" · ") : "";

  return {
    scheduleMxStatus: displayStatus,
    scheduleMxDueIso: dueIso,
    scheduleMxDaysUntil: daysUntil,
    scheduleMxOverdueByDays: overdueByDays,
    scheduleMxBucket: bucket,
    scheduleMxNeedsEntry: needsScheduleMxEntry,
  };
}

/** Excel exports dates as ISO strings, "YYYY-MM-DD HH:MM" or "M/D/YY". */
function parseEstablishedDate(raw: string): string | null {
  if (!raw) return null;
  const s = raw.trim();
  // ISO with time: 2026-04-01T00:00:00Z or 2026-04-01 00:00:00
  const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  // US: 4/1/2026, 04/01/26, 4/1/26 8:30 AM
  const usMatch = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (usMatch) {
    const m = Number(usMatch[1]);
    const d = Number(usMatch[2]);
    const yRaw = Number(usMatch[3]);
    const y = yRaw < 70 ? 2000 + yRaw : yRaw < 100 ? 1900 + yRaw : yRaw;
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    }
  }
  // Fallback: ISO Date.parse
  const t = Date.parse(s);
  if (!Number.isNaN(t)) {
    const d = new Date(t);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
  }
  return null;
}

function rowToWatchRow(
  row: WatchReadRow,
  asOfDateKey: string,
  ctx?: {
    firstSeenDate?: string | null;
    earliestSnapshot?: string | null;
    intervals?: Record<MelTier, number | null>;
  },
): WatchRow {
  const tier = (row.mel_tier as MelTier) || "unknown";
  const intervals = ctx?.intervals ?? REMARK_INTERVAL_DEFAULT;
  const interval = intervals[tier];
  const daysSince = calendarDaysBetween(row.last_remark_change_date, asOfDateKey);
  const stale = interval !== null && daysSince > interval;
  const firstSeen = ctx?.firstSeenDate ?? "";
  const earliest = ctx?.earliestSnapshot ?? "";
  // History is "bounded" (we don't actually know the true age) when the WO was
  // already present in our very first ingested snapshot AND its remarks have
  // not changed since then — i.e., the count below is a lower bound only.
  const historyBounded =
    !!firstSeen && !!earliest && firstSeen === earliest && row.last_remark_change_date <= firstSeen;
  const rawCols = readRawColumns(row.raw_row_json);
  const daysDown = pickDaysDownFromRaw(rawCols);
  return {
    workOrderId: row.work_order_id,
    assetId: row.asset_id,
    melTier: tier,
    partsStatus: row.parts_status,
    eticRaw: row.etic_raw,
    eticDate: row.etic_date,
    remarks: row.remarks,
    lastRemarkChangeDate: row.last_remark_change_date,
    daysSinceRemarkChange: daysSince,
    requiredIntervalDays: interval,
    remarkStale: stale,
    eticPushCount: row.etic_push_count,
    cumulativeEticSlipDays: row.cumulative_etic_slip_days,
    firstEticDate: row.first_etic_date,
    lastEticDate: row.last_etic_date,
    lastSnapshotDate: row.last_snapshot_date,
    owningUnit: row.owning_unit ?? "",
    melKey: row.mel_key ?? "",
    shop: row.shop ?? "",
    mgmtCd: row.mgmt_cd ?? "",
    makeModel: row.make_model ?? "",
    vehNomen: row.veh_nomen ?? "",
    firstSeenDate: firstSeen,
    historyBounded,
    daysDown,
    ...extractRawExtrasFromColumns(rawCols),
    ...analyzeScheduleMxFromRaw(rawCols, asOfDateKey),
  };
}

function extractRawExtrasFromColumns(raw: Record<string, string>): {
  establishedDate: string;
  establishedDateIso: string | null;
  woReason: string;
  nce: boolean;
  nceStatus: string;
} {
  const establishedDate = pickRawValue(raw, [
    "estbd dt/time",
    "estbd dt time",
    "estbd date",
    "estbd dt",
    "established date",
    "established dt/time",
    "open date",
    "wo open date",
    "date opened",
  ]);
  const woReason = pickRawValue(raw, [
    "reason",
    "wo reason",
    "work order reason",
    "reason code",
    "reason for maintenance",
    "reason for work",
  ]);
  // NCE comes from the merged Fleet (P&A) columns (any non-empty value in the
  // "NCE Vehicle Listing.Status" column means the asset is NCE).
  const nceStatus = pickRawValue(raw, [
    "fleet.nce vehicle listing.status",
    "fleet.nce vehicle listing status",
    "fleet.nce status",
    "fleet.nce",
  ]);
  return {
    establishedDate,
    establishedDateIso: parseEstablishedDate(establishedDate),
    woReason,
    nce: !!nceStatus,
    nceStatus,
  };
}

export type ScheduleMxFleetRow = {
  assetId: string;
  makeModel: string;
  mgmtCd: string;
  workOrderCount: number;
  nce: boolean;
  nceStatus: string;
  /** Overdue schedule maintenance on an NCE asset — highest leadership priority. */
  scheduleMxNceCritical: boolean;
} & ReturnType<typeof analyzeScheduleMxFromRaw>;

async function loadWoCountsPerAsset(
  env: { ETIC_SNAPSHOTS: D1Database },
  dateKey: string,
): Promise<Map<string, number>> {
  const m = new Map<string, number>();
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT asset_id, COUNT(*) AS c FROM work_order_snapshot WHERE snapshot_date_key = ? GROUP BY asset_id`,
  )
    .bind(dateKey)
    .all<{ asset_id: string; c: number }>();
  for (const row of r.results ?? []) {
    const aid = (row.asset_id ?? "").trim();
    if (aid) m.set(aid, Number(row.c) || 0);
  }
  return m;
}

function sortScheduleMxRows(out: ScheduleMxFleetRow[]): ScheduleMxFleetRow[] {
  function rank(a: ScheduleMxFleetRow): number {
    if (a.scheduleMxNceCritical) return 0;
    if (a.scheduleMxBucket === "overdue") return 1;
    if (a.scheduleMxBucket === "due_soon") return 2;
    if (a.scheduleMxBucket === "missing") return 3;
    return 4;
  }
  out.sort(function (a, b) {
    const ra = rank(a);
    const rb = rank(b);
    if (ra !== rb) return ra - rb;
    return a.assetId.localeCompare(b.assetId);
  });
  return out;
}

function buildScheduleMxRowsFromFleetSheet(
  rawByAsset: Map<string, Record<string, string>>,
  typedByAsset: Map<string, FleetRecord>,
  woCounts: Map<string, number>,
  dateKey: string,
): ScheduleMxFleetRow[] {
  const out: ScheduleMxFleetRow[] = [];
  for (const [assetId, merged] of rawByAsset) {
    const typed = typedByAsset.get(assetId);
    const nceStatus = pickRawValue(merged, [
      "fleet.nce vehicle listing.status",
      "fleet.nce vehicle listing status",
      "fleet.nce status",
      "fleet.nce",
    ]);
    const nce = !!nceStatus;
    const smx = analyzeScheduleMxFromRaw(merged, dateKey);
    const scheduleMxNceCritical = nce && smx.scheduleMxBucket === "overdue";
    const makeModel =
      (typed?.makeModel ?? "").trim() ||
      pickRawValue(merged, ["fleet.make/model", "fleet.make model", "make model"]);
    const mgmtCd =
      (typed?.mgmtCd ?? "").trim() ||
      pickRawValue(merged, ["fleet.mgmt cd", "fleet.mgmt code", "mgmt cd"]);
    out.push({
      assetId,
      makeModel,
      mgmtCd,
      workOrderCount: woCounts.get(assetId) ?? 0,
      nce,
      nceStatus,
      scheduleMxNceCritical,
      ...smx,
    });
  }
  return sortScheduleMxRows(out);
}

/** Fallback when the workbook is missing from R2 or has no Fleet sheet. */
async function scheduleMxFleetRowsFromWorkOrderSnapshotsOnly(
  env: { ETIC_SNAPSHOTS: D1Database },
  dateKey: string,
): Promise<ScheduleMxFleetRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT asset_id, raw_row_json, make_model, mgmt_cd
       FROM work_order_snapshot
      WHERE snapshot_date_key = ?`,
  )
    .bind(dateKey)
    .all<{ asset_id: string; raw_row_json: string | null; make_model: string | null; mgmt_cd: string | null }>();

  type Agg = { raws: Record<string, string>[]; make: string; mgmt: string; woCount: number };
  const byAsset = new Map<string, Agg>();
  for (const row of r.results ?? []) {
    const aid = (row.asset_id ?? "").trim();
    if (!aid) continue;
    const parsed = readRawColumns(row.raw_row_json);
    let g = byAsset.get(aid);
    if (!g) {
      g = {
        raws: [],
        make: (row.make_model ?? "").trim(),
        mgmt: (row.mgmt_cd ?? "").trim(),
        woCount: 0,
      };
      byAsset.set(aid, g);
    }
    g.raws.push(parsed);
    g.woCount += 1;
    if (!g.make && row.make_model) g.make = row.make_model.trim();
    if (!g.mgmt && row.mgmt_cd) g.mgmt = row.mgmt_cd.trim();
  }

  const out: ScheduleMxFleetRow[] = [];
  for (const [assetId, g] of byAsset) {
    const merged: Record<string, string> = {};
    for (const rr of g.raws) {
      for (const [k, v] of Object.entries(rr)) {
        if (v && String(v).trim()) merged[k] = String(v).trim();
      }
    }
    const nceStatus = pickRawValue(merged, [
      "fleet.nce vehicle listing.status",
      "fleet.nce vehicle listing status",
      "fleet.nce status",
      "fleet.nce",
    ]);
    const nce = !!nceStatus;
    const smx = analyzeScheduleMxFromRaw(merged, dateKey);
    const scheduleMxNceCritical = nce && smx.scheduleMxBucket === "overdue";
    out.push({
      assetId,
      makeModel: g.make,
      mgmtCd: g.mgmt,
      workOrderCount: g.woCount,
      nce,
      nceStatus,
      scheduleMxNceCritical,
      ...smx,
    });
  }
  return sortScheduleMxRows(out);
}

/**
 * Full Fleet (P&A) fleet: reads the stored ETIC workbook from R2 and parses the Fleet sheet
 * so every asset row appears (not only assets with open work orders). WO counts are joined from
 * `work_order_snapshot` when present.
 */
export async function getScheduleMxFleetForDate(
  env: { ETIC_SNAPSHOTS: D1Database; ETIC_BUCKET: R2Bucket },
  dateKey: string,
): Promise<ScheduleMxFleetRow[]> {
  const woCounts = await loadWoCountsPerAsset(env, dateKey);

  const snap = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT workbook_key FROM etic_snapshots WHERE date_key = ?`,
  )
    .bind(dateKey)
    .first<{ workbook_key: string | null }>();

  if (snap?.workbook_key) {
    const obj = await env.ETIC_BUCKET.get(snap.workbook_key);
    if (obj) {
      try {
        const bytes = await obj.arrayBuffer();
        const ex = await extractFleetMapsFromBinary(bytes);
        if (ex && ex.rawByAsset.size > 0) {
          return buildScheduleMxRowsFromFleetSheet(ex.rawByAsset, ex.byAsset, woCounts, dateKey);
        }
      } catch (err) {
        console.error("getScheduleMxFleetForDate: fleet extract failed", err);
      }
    }
  }

  return scheduleMxFleetRowsFromWorkOrderSnapshotsOnly(env, dateKey);
}

async function getEarliestSnapshotDate(env: { ETIC_SNAPSHOTS: D1Database }): Promise<string> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT MIN(snapshot_date_key) AS k FROM work_order_snapshot`,
  ).first<{ k: string | null }>();
  return r?.k ?? "";
}

async function getFirstSeenDates(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderIds?: string[],
): Promise<Map<string, string>> {
  const out = new Map<string, string>();
  if (workOrderIds && workOrderIds.length > 0) {
    const batchSize = 90;
    for (let i = 0; i < workOrderIds.length; i += batchSize) {
      const chunk = workOrderIds.slice(i, i + batchSize);
      const placeholders = chunk.map(() => "?").join(",");
      const r = await env.ETIC_SNAPSHOTS.prepare(
        `SELECT work_order_id, MIN(snapshot_date_key) AS first_seen
         FROM work_order_snapshot
         WHERE work_order_id IN (${placeholders})
         GROUP BY work_order_id`,
      )
        .bind(...chunk)
        .all<{ work_order_id: string; first_seen: string }>();
      for (const row of r.results ?? []) out.set(row.work_order_id, row.first_seen);
    }
  } else {
    const r = await env.ETIC_SNAPSHOTS.prepare(
      `SELECT work_order_id, MIN(snapshot_date_key) AS first_seen
       FROM work_order_snapshot
       GROUP BY work_order_id`,
    ).all<{ work_order_id: string; first_seen: string }>();
    for (const row of r.results ?? []) out.set(row.work_order_id, row.first_seen);
  }
  return out;
}

async function getFirstSeenDate(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderId: string,
): Promise<string> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT MIN(snapshot_date_key) AS first_seen
     FROM work_order_snapshot WHERE work_order_id = ?`,
  )
    .bind(workOrderId)
    .first<{ first_seen: string | null }>();
  return r?.first_seen ?? "";
}

/** One work order’s latest state; staleness vs asOfDateKey. Returns null if not in index. */
export async function getWatchRowById(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderId: string,
  asOfDateKey: string,
): Promise<WatchRow | null> {
  const row = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id, asset_id, mel_tier, parts_status, etic_raw, etic_date, remarks,
            last_remark_change_date, etic_push_count, cumulative_etic_slip_days, first_etic_date, last_etic_date,
            last_snapshot_date, owning_unit, mel_key, shop, mgmt_cd, make_model, veh_nomen, raw_row_json
     FROM work_order_state WHERE work_order_id = ?`,
  )
    .bind(workOrderId)
    .first<WatchReadRow>();
  if (!row) return null;
  const [firstSeen, earliest, intervals] = await Promise.all([
    getFirstSeenDate(env, workOrderId),
    getEarliestSnapshotDate(env),
    getStalenessThresholds(env).then(resolveStaleness),
  ]);
  return rowToWatchRow(row, asOfDateKey, { firstSeenDate: firstSeen, earliestSnapshot: earliest, intervals });
}

/** Latest state per WO; remark staleness computed vs asOfDateKey (report date). */
export async function getWatchRowsLatest(env: { ETIC_SNAPSHOTS: D1Database }, asOfDateKey: string): Promise<WatchRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id, asset_id, mel_tier, parts_status, etic_raw, etic_date, remarks,
            last_remark_change_date, etic_push_count, cumulative_etic_slip_days, first_etic_date, last_etic_date,
            last_snapshot_date, owning_unit, mel_key, shop, mgmt_cd, make_model, veh_nomen, raw_row_json
     FROM work_order_state
     ORDER BY work_order_id`,
  ).all<WatchReadRow>();

  const [firstSeenMap, earliest, intervals] = await Promise.all([
    getFirstSeenDates(env),
    getEarliestSnapshotDate(env),
    getStalenessThresholds(env).then(resolveStaleness),
  ]);
  const out: WatchRow[] = [];
  for (const row of r.results ?? []) {
    out.push(
      rowToWatchRow(row, asOfDateKey, {
        firstSeenDate: firstSeenMap.get(row.work_order_id) ?? "",
        earliestSnapshot: earliest,
        intervals,
      }),
    );
  }
  return out;
}

/**
 * All WOs that were present in the given snapshot date (the file the user is viewing).
 * Scales to years of daily files — indexed on snapshot_date_key.
 */
export async function getWatchRowsForDate(
  env: { ETIC_SNAPSHOTS: D1Database },
  dateKey: string,
): Promise<WatchRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id, asset_id, mel_tier, parts_status, etic_raw, etic_date, remarks,
            last_remark_change_date, etic_push_count, cumulative_etic_slip_days, first_etic_date, last_etic_date,
            snapshot_date_key AS last_snapshot_date, owning_unit, mel_key, shop, mgmt_cd, make_model, veh_nomen, raw_row_json
     FROM work_order_snapshot
     WHERE snapshot_date_key = ?
     ORDER BY work_order_id`,
  )
    .bind(dateKey)
    .all<WatchReadRow>();

  const ids = (r.results ?? []).map((row) => row.work_order_id);
  const [firstSeenMap, earliest, intervals] = await Promise.all([
    getFirstSeenDates(env, ids),
    getEarliestSnapshotDate(env),
    getStalenessThresholds(env).then(resolveStaleness),
  ]);
  const out: WatchRow[] = [];
  for (const row of r.results ?? []) {
    out.push(
      rowToWatchRow(row, dateKey, {
        firstSeenDate: firstSeenMap.get(row.work_order_id) ?? "",
        earliestSnapshot: earliest,
        intervals,
      }),
    );
  }
  return out;
}

/** Entire timeline of a single WO across every snapshot it appeared in. */
export async function getWorkOrderTimeline(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderId: string,
): Promise<WatchRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id, asset_id, mel_tier, parts_status, etic_raw, etic_date, remarks,
            last_remark_change_date, etic_push_count, cumulative_etic_slip_days, first_etic_date, last_etic_date,
            snapshot_date_key AS last_snapshot_date, owning_unit, mel_key, shop, mgmt_cd, make_model, veh_nomen, raw_row_json
     FROM work_order_snapshot
     WHERE work_order_id = ?
     ORDER BY snapshot_date_key DESC`,
  )
    .bind(workOrderId)
    .all<WatchReadRow>();

  const [firstSeen, earliest, intervals] = await Promise.all([
    getFirstSeenDate(env, workOrderId),
    getEarliestSnapshotDate(env),
    getStalenessThresholds(env).then(resolveStaleness),
  ]);
  const out: WatchRow[] = [];
  for (const row of r.results ?? []) {
    out.push(
      rowToWatchRow(row, row.last_snapshot_date, {
        firstSeenDate: firstSeen,
        earliestSnapshot: earliest,
        intervals,
      }),
    );
  }
  return out;
}

/** One WO's state on a specific snapshot date; null if it wasn't in that snapshot. */
export async function getWatchRowByIdForDate(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderId: string,
  dateKey: string,
): Promise<WatchRow | null> {
  const row = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id, asset_id, mel_tier, parts_status, etic_raw, etic_date, remarks,
            last_remark_change_date, etic_push_count, cumulative_etic_slip_days, first_etic_date, last_etic_date,
            snapshot_date_key AS last_snapshot_date, owning_unit, mel_key, shop, mgmt_cd, make_model, veh_nomen, raw_row_json
     FROM work_order_snapshot
     WHERE work_order_id = ? AND snapshot_date_key = ?`,
  )
    .bind(workOrderId, dateKey)
    .first<WatchReadRow>();
  if (!row) return null;
  const [firstSeen, earliest, intervals] = await Promise.all([
    getFirstSeenDate(env, workOrderId),
    getEarliestSnapshotDate(env),
    getStalenessThresholds(env).then(resolveStaleness),
  ]);
  return rowToWatchRow(row, dateKey, { firstSeenDate: firstSeen, earliestSnapshot: earliest, intervals });
}

export type WorkOrderChangelogEntry = {
  id: number;
  snapshot_date_key: string;
  changed_at_iso: string;
  field: string;
  old_value: string | null;
  new_value: string | null;
};

export async function getChangelog(env: { ETIC_SNAPSHOTS: D1Database }, workOrderId: string, limit = 200) {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, snapshot_date_key, changed_at_iso, field, old_value, new_value
     FROM work_order_changelog WHERE work_order_id = ? ORDER BY id DESC LIMIT ?`,
  )
    .bind(workOrderId, limit)
    .all<{
      id: number;
      snapshot_date_key: string;
      changed_at_iso: string;
      field: string;
      old_value: string | null;
      new_value: string | null;
    }>();
  return r.results ?? [];
}

type SnapChangelogRow = {
  snapshot_date_key: string;
  remarks: string;
  parts_status: string;
  etic_raw: string;
  etic_date: string | null;
  mel_tier: string;
  shop: string | null;
};

/**
 * Reconstruct field changes by diffing consecutive work_order_snapshot rows (ingest
 * order). This is the correct timeline for "when did the workbook first show a
 * change" even when work_order_changelog is stale, rebuilt in one pass, or all rows
 * share one snapshot_date_key. Uses the same field rules as ingestWorkOrderSnapshot.
 */
export async function getChangelogFromSnapshots(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderId: string,
  limit = 500,
): Promise<WorkOrderChangelogEntry[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT snapshot_date_key, remarks, parts_status, etic_raw, etic_date, mel_tier, shop
     FROM work_order_snapshot
     WHERE work_order_id = ?
     ORDER BY snapshot_date_key ASC`,
  )
    .bind(workOrderId)
    .all<SnapChangelogRow>();
  const rows = r.results ?? [];
  const out: WorkOrderChangelogEntry[] = [];
  let id = 0;
  const nowIso = new Date().toISOString();
  let tick = 0;
  const stamp = (k: string) => {
    const t = Date.parse(k + "T00:00:00.000Z");
    if (Number.isNaN(t)) return nowIso;
    return new Date(t + tick++).toISOString();
  };
  for (let i = 0; i < rows.length; i++) {
    const cur = rows[i]!;
    if (i === 0) {
      out.push({
        id: ++id,
        snapshot_date_key: cur.snapshot_date_key,
        changed_at_iso: stamp(cur.snapshot_date_key),
        field: "initial",
        old_value: "",
        new_value: "first_seen",
      });
      continue;
    }
    const prev = rows[i - 1]!;
    const k = cur.snapshot_date_key;
    if (cur.remarks !== (prev.remarks ?? "")) {
      out.push({
        id: ++id,
        snapshot_date_key: k,
        changed_at_iso: stamp(k),
        field: "remarks",
        old_value: prev.remarks ?? "",
        new_value: cur.remarks,
      });
    }
    if (cur.parts_status !== (prev.parts_status ?? "")) {
      out.push({
        id: ++id,
        snapshot_date_key: k,
        changed_at_iso: stamp(k),
        field: "parts_status",
        old_value: prev.parts_status ?? "",
        new_value: cur.parts_status,
      });
    }
    if (cur.etic_raw !== (prev.etic_raw ?? "")) {
      out.push({
        id: ++id,
        snapshot_date_key: k,
        changed_at_iso: stamp(k),
        field: "etic",
        old_value: prev.etic_raw ?? "",
        new_value: cur.etic_raw,
      });
    }
    if ((cur.mel_tier || "unknown") !== (prev.mel_tier || "unknown")) {
      out.push({
        id: ++id,
        snapshot_date_key: k,
        changed_at_iso: stamp(k),
        field: "mel_tier",
        old_value: prev.mel_tier ?? "",
        new_value: cur.mel_tier,
      });
    }
    if ((cur.shop ?? "") !== (prev.shop ?? "")) {
      out.push({
        id: ++id,
        snapshot_date_key: k,
        changed_at_iso: stamp(k),
        field: "shop",
        old_value: prev.shop ?? "",
        new_value: cur.shop ?? "",
      });
    }
    const oldD = prev.etic_date;
    const newD = cur.etic_date;
    if (oldD && newD && newD > oldD) {
      out.push({
        id: ++id,
        snapshot_date_key: k,
        changed_at_iso: stamp(k),
        field: "etic_date_slip",
        old_value: oldD,
        new_value: newD,
      });
    }
  }
  if (out.length <= limit) return out;
  return out.slice(-limit);
}

/**
 * Changelog for the work-order screen: `work_order_changelog` in D1 (populated on every
 * ingest and by POST /api/watch?rebuild=1) first; if empty, derive from
 * `work_order_snapshot` rows (still D1).
 */
export async function getChangelogForDisplay(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderId: string,
  limit = 500,
): Promise<WorkOrderChangelogEntry[]> {
  const fromTable = await getChangelog(env, workOrderId, limit);
  if (fromTable.length > 0) return fromTable;
  return getChangelogFromSnapshots(env, workOrderId, limit);
}

/** ETIC meeting rows for this WO (notes / due-outs) for the work-order timeline. */
export type WoTimelineMeetingNote = {
  workOrderId: string;
  meetingId: number;
  meetingTitle: string;
  startedAtIso: string;
  updatedAtIso: string;
  notes: string;
  dueOuts: string;
  /** 0/1 per line, same order as notes.split(newline) */
  notesLineDone: number[];
  /** 0/1 per line, same order as dueOuts.split(newline) */
  dueOutsLineDone: number[];
};

/** Rolling yard_check rows for the asset tied to this WO. */
export type WoTimelineYardWalk = {
  id: number;
  location: string;
  discrepancies: string;
  status: string;
  checkedBy: string;
  checkedAtIso: string;
  sourceDateKey: string;
};

export async function getAssetIdForWorkOrder(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderId: string,
): Promise<string | null> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT asset_id FROM work_order_state WHERE work_order_id = ?`,
  )
    .bind(workOrderId)
    .first<{ asset_id: string }>();
  const a = (r?.asset_id ?? "").trim();
  return a || null;
}

export async function getMeetingNotesForWorkOrder(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderId: string,
): Promise<WoTimelineMeetingNote[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT n.work_order_id, n.notes, n.due_outs, n.line_completions, n.updated_at_iso, m.id AS meeting_id, m.title AS meeting_title, m.started_at_iso
       FROM meeting_wo_note n
      JOIN meeting m ON m.id = n.meeting_id
      WHERE n.work_order_id = ?
        AND (length(trim(ifnull(n.notes,''))) > 0 OR length(trim(ifnull(n.due_outs,''))) > 0)
      ORDER BY n.updated_at_iso DESC
      LIMIT 200`,
  )
    .bind(workOrderId)
    .all<{
      work_order_id: string;
      notes: string;
      due_outs: string;
      line_completions: string;
      updated_at_iso: string;
      meeting_id: number;
      meeting_title: string;
      started_at_iso: string;
    }>();
  return (r.results ?? []).map((row) => {
    const lc = buildLineCompletions(
      row.line_completions ?? "",
      row.notes ?? "",
      row.due_outs ?? "",
    );
    return {
      workOrderId: row.work_order_id ?? "",
      meetingId: row.meeting_id,
      meetingTitle: row.meeting_title ?? "",
      startedAtIso: row.started_at_iso ?? "",
      updatedAtIso: row.updated_at_iso ?? "",
      notes: row.notes ?? "",
      dueOuts: row.due_outs ?? "",
      notesLineDone: lc.n,
      dueOutsLineDone: lc.d,
    };
  });
}

export async function getYardWalksForAsset(
  env: { ETIC_SNAPSHOTS: D1Database },
  assetId: string,
  limit = 80,
): Promise<WoTimelineYardWalk[]> {
  const aid = assetId.trim();
  if (!aid) return [];
  const lim = Math.max(1, Math.min(limit, 200));
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, location, discrepancies, status, checked_by, checked_at_iso, source_date_key
       FROM yard_check WHERE asset_id = ?
       ORDER BY checked_at_iso DESC LIMIT ?`,
  )
    .bind(aid, lim)
    .all<{
      id: number;
      location: string;
      discrepancies: string;
      status: string;
      checked_by: string;
      checked_at_iso: string;
      source_date_key: string;
    }>();
  return (r.results ?? []).map((row) => ({
    id: row.id,
    location: row.location ?? "",
    discrepancies: row.discrepancies ?? "",
    status: row.status ?? "",
    checkedBy: row.checked_by ?? "",
    checkedAtIso: row.checked_at_iso ?? "",
    sourceDateKey: row.source_date_key ?? "",
  }));
}

// ---------------------------------------------------------------------------
// Backfill typed columns from `raw_row_json` already stored in D1.
//
// Lets us add a new typed column (e.g. `mgmt_cd`) and re-derive its value
// across all history without re-parsing any .xlsx out of R2 — as long as
// the raw row JSON has been captured at least once.
// ---------------------------------------------------------------------------

/**
 * Maps each typed column we know how to backfill to:
 *   - the WO-sheet field name (used to look up synonyms in WORK_ORDER_SYNONYMS)
 *   - optionally a fleet-sheet field (used to look up synonyms in FLEET_SYNONYMS)
 * The backfill scores each JSON key against those synonyms and picks the
 * best-matching cell value for that row.
 */
const BACKFILL_FIELDS: Array<{
  column: string;
  woField?: keyof typeof WORK_ORDER_SYNONYMS;
  fleetField?: keyof typeof FLEET_SYNONYMS;
}> = [
  { column: "asset_id",     woField: "assetId",     fleetField: "assetId"     },
  { column: "owning_unit",  woField: "owningUnit"                              },
  { column: "mel_key",      woField: "melKey"                                  },
  { column: "shop",         woField: "shop"                                    },
  { column: "mgmt_cd",      woField: "mgmtCd",      fleetField: "mgmtCd"      },
  { column: "make_model",   woField: "makeModel",   fleetField: "makeModel"   },
  { column: "veh_nomen",                            fleetField: "vehNomen"    },
];

/**
 * Find the best-matching cell value in a parsed raw_row_json object for the
 * given WO and/or fleet synonyms. JSON keys produced by the extractor are
 * already normalized (lowercase, single-spaced); fleet keys are prefixed
 * with `fleet.`. We strip the prefix before scoring so fleet headers are
 * eligible too.
 */
function pickFromRawRow(
  raw: Record<string, string>,
  woField: keyof typeof WORK_ORDER_SYNONYMS | undefined,
  fleetField: keyof typeof FLEET_SYNONYMS | undefined,
): string {
  let bestScore = -Infinity;
  let bestVal = "";
  for (const [key, val] of Object.entries(raw)) {
    if (!val) continue;
    const isFleet = key.startsWith("fleet.");
    const headerText = isFleet ? key.slice("fleet.".length) : key;
    if (!isFleet && woField) {
      const m = scoreHeaderMatch(headerText, WORK_ORDER_SYNONYMS);
      if (m && m.field === woField && m.score > bestScore) {
        bestScore = m.score;
        bestVal = val;
      }
    }
    if (isFleet && fleetField) {
      const m = scoreHeaderMatch(headerText, FLEET_SYNONYMS);
      if (m && m.field === fleetField && m.score > bestScore) {
        bestScore = m.score;
        bestVal = val;
      }
    }
  }
  return bestVal;
}

export type BackfillReport = {
  table: "work_order_state" | "work_order_snapshot";
  rowsScanned: number;
  rowsWithJson: number;
  perColumn: Record<string, { updated: number; sample?: { wid: string; value: string } }>;
};

/**
 * Re-derive typed columns from each row's `raw_row_json` and write them back
 * if the typed column is empty (or, when overwrite=true, regardless).
 * This is the fast path for "I just added a new column and want it populated
 * across history" — no R2, no ExcelJS, just SQL + JSON.
 */
export async function backfillTypedColumnsFromJson(
  env: { ETIC_SNAPSHOTS: D1Database },
  opts: { table: "work_order_state" | "work_order_snapshot"; overwrite?: boolean; limit?: number } = { table: "work_order_state" },
): Promise<BackfillReport> {
  const overwrite = opts.overwrite === true;
  const limit = Math.max(1, Math.min(opts.limit ?? 100000, 200000));
  const report: BackfillReport = {
    table: opts.table,
    rowsScanned: 0,
    rowsWithJson: 0,
    perColumn: Object.fromEntries(BACKFILL_FIELDS.map((f) => [f.column, { updated: 0 }])),
  };

  // Stream rows in pages so we don't blow CPU/memory on a single query.
  const PAGE = 500;
  let offset = 0;
  while (offset < limit) {
    let pageRows: Array<Record<string, unknown>>;
    if (opts.table === "work_order_state") {
      const r = await env.ETIC_SNAPSHOTS.prepare(
        `SELECT rowid AS rid, work_order_id AS wid, raw_row_json,
                asset_id, owning_unit, mel_key, shop, mgmt_cd, make_model, veh_nomen
         FROM work_order_state ORDER BY rowid LIMIT ? OFFSET ?`,
      ).bind(PAGE, offset).all();
      pageRows = (r.results ?? []) as Array<Record<string, unknown>>;
    } else {
      const r = await env.ETIC_SNAPSHOTS.prepare(
        `SELECT rowid AS rid, work_order_id AS wid, raw_row_json,
                asset_id, owning_unit, mel_key, shop, mgmt_cd, make_model, veh_nomen
         FROM work_order_snapshot ORDER BY rowid LIMIT ? OFFSET ?`,
      ).bind(PAGE, offset).all();
      pageRows = (r.results ?? []) as Array<Record<string, unknown>>;
    }
    if (pageRows.length === 0) break;

    const updates: D1PreparedStatement[] = [];
    for (const row of pageRows) {
      report.rowsScanned += 1;
      const rawJson = String(row.raw_row_json ?? "");
      if (!rawJson || rawJson === "{}") continue;
      let parsed: Record<string, string>;
      try {
        parsed = JSON.parse(rawJson) as Record<string, string>;
      } catch {
        continue;
      }
      report.rowsWithJson += 1;

      const setClauses: string[] = [];
      const binds: Array<string | number> = [];
      for (const f of BACKFILL_FIELDS) {
        const cur = String(row[f.column] ?? "").trim();
        if (cur && !overwrite) continue;
        const next = pickFromRawRow(parsed, f.woField, f.fleetField).trim();
        if (!next || next === cur) continue;
        setClauses.push(`${f.column} = ?`);
        binds.push(next);
        const cell = report.perColumn[f.column];
        cell.updated += 1;
        if (!cell.sample) cell.sample = { wid: String(row.wid ?? ""), value: next };
      }
      if (setClauses.length === 0) continue;
      const sql = `UPDATE ${opts.table} SET ${setClauses.join(", ")} WHERE rowid = ?`;
      binds.push(Number(row.rid));
      updates.push(env.ETIC_SNAPSHOTS.prepare(sql).bind(...binds));
    }

    if (updates.length > 0) {
      const BATCH = 50;
      for (let i = 0; i < updates.length; i += BATCH) {
        await env.ETIC_SNAPSHOTS.batch(updates.slice(i, i + BATCH));
      }
    }

    offset += pageRows.length;
    if (pageRows.length < PAGE) break;
  }
  return report;
}

// ---------------------------------------------------------------------------
// FM&A (Fleet Managers & Analysis) hand-logged actions.
//
// A fleet manager opens a WO in the UI and logs "I just updated the remarks
// in the source system at 14:37, signed: Sgt Smith". That row goes into
// work_order_action with status='pending'. When the next ETIC .xlsx is
// ingested, verifyWorkOrderActionsForSnapshot() compares the action's
// expected_field against the changelog entries written for that snapshot:
//   - field changed in the snapshot   -> status='confirmed'
//   - first snapshot after the action AND no change to that field
//                                     -> status='missed'
// ---------------------------------------------------------------------------

// Note: MEL tier is excluded — it's a calculated field derived from remarks/
// position vs the MEL key, not something a fleet manager updates directly.
export type WorkOrderActionType =
  | "remarks_update"
  | "etic_update"
  | "parts_update"
  | "shop_update"
  | "other";

export type WorkOrderActionStatus = "pending" | "confirmed" | "missed";

export type WorkOrderAction = {
  id: number;
  workOrderId: string;
  createdAtIso: string;
  actionType: WorkOrderActionType;
  expectedField: string;
  actorName: string;
  note: string;
  status: WorkOrderActionStatus;
  verifiedAtIso: string | null;
  verifiedInSnapshot: string | null;
  snapshotsChecked: number;
  /** When set, FM&A marked this follow-up note as done (shop email / internal). */
  followupDoneAtIso: string | null;
};

type WorkOrderActionRow = {
  id: number;
  work_order_id: string;
  created_at_iso: string;
  action_type: string;
  expected_field: string;
  actor_name: string;
  note: string;
  status: string;
  verified_at_iso: string | null;
  verified_in_snapshot: string | null;
  snapshots_checked: number;
  followup_done_at_iso: string | null;
};

function rowToAction(r: WorkOrderActionRow): WorkOrderAction {
  return {
    id: r.id,
    workOrderId: r.work_order_id,
    createdAtIso: r.created_at_iso,
    actionType: r.action_type as WorkOrderActionType,
    expectedField: r.expected_field,
    actorName: r.actor_name,
    note: r.note,
    status: r.status as WorkOrderActionStatus,
    verifiedAtIso: r.verified_at_iso,
    verifiedInSnapshot: r.verified_in_snapshot,
    snapshotsChecked: r.snapshots_checked,
    followupDoneAtIso: r.followup_done_at_iso,
  };
}

/** Map FM&A action type -> the field name we'll watch for in the changelog. */
export function expectedFieldForActionType(t: WorkOrderActionType): string {
  switch (t) {
    case "remarks_update": return "remarks";
    case "etic_update":    return "etic";
    case "parts_update":   return "parts_status";
    case "shop_update":    return "shop";
    default:               return "";
  }
}

export async function logWorkOrderAction(
  env: { ETIC_SNAPSHOTS: D1Database },
  input: {
    workOrderId: string;
    actionType: WorkOrderActionType;
    actorName?: string;
    note?: string;
  },
): Promise<WorkOrderAction> {
  const wid = input.workOrderId.trim();
  if (!wid) throw new Error("workOrderId required");
  const expected = expectedFieldForActionType(input.actionType);
  const nowIso = new Date().toISOString();
  const res = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO work_order_action
       (work_order_id, created_at_iso, action_type, expected_field, actor_name, note, status, snapshots_checked)
     VALUES (?, ?, ?, ?, ?, ?, 'pending', 0)
     RETURNING id, work_order_id, created_at_iso, action_type, expected_field, actor_name, note, status,
               verified_at_iso, verified_in_snapshot, snapshots_checked, followup_done_at_iso`,
  )
    .bind(
      wid,
      nowIso,
      input.actionType,
      expected,
      (input.actorName ?? "").trim(),
      (input.note ?? "").trim(),
    )
    .first<WorkOrderActionRow>();
  if (!res) throw new Error("Insert failed");
  return rowToAction(res);
}

export async function getWorkOrderActions(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderId: string,
  limit = 200,
): Promise<WorkOrderAction[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, work_order_id, created_at_iso, action_type, expected_field, actor_name, note, status,
            verified_at_iso, verified_in_snapshot, snapshots_checked, followup_done_at_iso
     FROM work_order_action
     WHERE work_order_id = ?
     ORDER BY id DESC
     LIMIT ?`,
  )
    .bind(workOrderId, limit)
    .all<WorkOrderActionRow>();
  return (r.results ?? []).map(rowToAction);
}

/** FM&A actions that include a note — for shop email / follow-up reports. Shop/asset from live state or last snapshot if WO closed. */
export type FmaFollowUpReportRow = {
  id: number;
  workOrderId: string;
  createdAtIso: string;
  actionType: string;
  actorName: string;
  note: string;
  status: string;
  shop: string;
  assetId: string;
};

export async function listFmaFollowUpActionsForReport(
  env: { ETIC_SNAPSHOTS: D1Database },
  opts?: {
    sinceIso?: string;
    limit?: number;
    /** Substring match on resolved shop (case-insensitive). Empty = no filter. */
    shopContains?: string;
    /** pending = only rows still awaiting verification; all = include confirmed/missed */
    status?: "pending" | "all";
  },
): Promise<FmaFollowUpReportRow[]> {
  const lim = Math.min(800, Math.max(1, opts?.limit ?? 400));
  let sinceIso = (opts?.sinceIso ?? "").trim();
  if (!sinceIso) {
    const cutoff = Date.now() - 14 * 86400000;
    sinceIso = new Date(cutoff).toISOString();
  }
  const shopQ = (opts?.shopContains ?? "").trim();
  const statusMode = opts?.status === "all" ? "all" : "pending";
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM (
        SELECT
          a.id AS id,
          a.work_order_id AS work_order_id,
          a.created_at_iso AS created_at_iso,
          a.action_type AS action_type,
          a.actor_name AS actor_name,
          a.note AS note,
          a.status AS status,
          COALESCE(
            NULLIF(trim(ws.shop), ''),
            (SELECT trim(s.shop)
               FROM work_order_snapshot s
              WHERE s.work_order_id = a.work_order_id
              ORDER BY s.snapshot_date_key DESC
              LIMIT 1),
            ''
          ) AS shop_resolved,
          COALESCE(
            NULLIF(trim(ws.asset_id), ''),
            (SELECT trim(s.asset_id)
               FROM work_order_snapshot s
              WHERE s.work_order_id = a.work_order_id
              ORDER BY s.snapshot_date_key DESC
              LIMIT 1),
            ''
          ) AS asset_resolved
         FROM work_order_action a
         LEFT JOIN work_order_state ws ON ws.work_order_id = a.work_order_id
        WHERE length(trim(ifnull(a.note, ''))) > 0
          AND (a.followup_done_at_iso IS NULL OR length(trim(a.followup_done_at_iso)) = 0)
          AND datetime(a.created_at_iso) >= datetime(?)
          AND (? = 'all' OR a.status = 'pending')
      ) AS t
      WHERE (length(trim(?)) = 0 OR t.shop_resolved LIKE '%' || trim(?) || '%' COLLATE NOCASE)
      ORDER BY t.shop_resolved COLLATE NOCASE, t.created_at_iso DESC
      LIMIT ?`,
  )
    .bind(sinceIso, statusMode, shopQ, shopQ, lim)
    .all<{
      id: number;
      work_order_id: string;
      created_at_iso: string;
      action_type: string;
      actor_name: string;
      note: string;
      status: string;
      shop_resolved: string;
      asset_resolved: string;
    }>();
  return (r.results ?? []).map((row) => ({
    id: row.id,
    workOrderId: row.work_order_id,
    createdAtIso: row.created_at_iso,
    actionType: row.action_type,
    actorName: row.actor_name ?? "",
    note: row.note ?? "",
    status: row.status,
    shop: row.shop_resolved ?? "",
    assetId: row.asset_resolved ?? "",
  }));
}

/** Mark an FM&A action as follow-up done (or clear). Independent of auto-verify. */
export async function setWorkOrderActionFollowupDone(
  env: { ETIC_SNAPSHOTS: D1Database },
  actionId: number,
  done: boolean,
): Promise<boolean> {
  if (!Number.isFinite(actionId) || actionId <= 0) return false;
  const at = done ? new Date().toISOString() : null;
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE work_order_action
        SET followup_done_at_iso = ?
      WHERE id = ?`,
  )
    .bind(at, actionId)
    .run();
  return (r.meta?.changes ?? 0) > 0;
}

export async function deleteWorkOrderAction(
  env: { ETIC_SNAPSHOTS: D1Database },
  id: number,
): Promise<boolean> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `DELETE FROM work_order_action WHERE id = ?`,
  ).bind(id).run();
  return (r.meta?.changes ?? 0) > 0;
}

/**
 * Walk every pending action that was logged BEFORE this snapshot's date and
 * decide whether the snapshot reflected the change.
 *   - changelog has an entry for (wid, snapshot_date_key, expected_field)
 *     -> mark confirmed
 *   - no entry, but this snapshot is the first one after the action
 *     -> mark missed (action didn't take)
 *   - if expected_field is empty (action_type='other'), only count the
 *     snapshot but never auto-confirm/miss — humans verify those.
 *
 * Designed to be cheap: pulls just the pending rows for this date, joins
 * against the changelog rows that were just written for the same date.
 */
export async function verifyWorkOrderActionsForSnapshot(
  env: { ETIC_SNAPSHOTS: D1Database },
  snapshotDateKey: string,
): Promise<{ checked: number; confirmed: number; missed: number }> {
  // Only act on actions logged strictly before this snapshot date (00:00 UTC
  // of the next day). An action logged on the same calendar day as the
  // snapshot is treated as "already in this snapshot" only if the field
  // changed; otherwise it remains pending until the next snapshot lands.
  const cutoffIso = snapshotDateKey + "T23:59:59.999Z";
  const pending = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, work_order_id, created_at_iso, action_type, expected_field, actor_name, note, status,
            verified_at_iso, verified_in_snapshot, snapshots_checked, followup_done_at_iso
     FROM work_order_action
     WHERE status = 'pending' AND created_at_iso <= ?`,
  )
    .bind(cutoffIso)
    .all<WorkOrderActionRow>();
  const rows = pending.results ?? [];
  if (rows.length === 0) return { checked: 0, confirmed: 0, missed: 0 };

  // Pull every changelog entry written for THIS snapshot date for the WOs
  // we care about. One query covers them all — typically <50 actions.
  const wids = Array.from(new Set(rows.map((r) => r.work_order_id)));
  const placeholders = wids.map(() => "?").join(",");
  const cl = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id, field
     FROM work_order_changelog
     WHERE snapshot_date_key = ? AND work_order_id IN (${placeholders})`,
  )
    .bind(snapshotDateKey, ...wids)
    .all<{ work_order_id: string; field: string }>();
  const changedSet = new Set<string>();
  for (const c of cl.results ?? []) {
    changedSet.add(c.work_order_id + "::" + c.field);
  }

  const nowIso = new Date().toISOString();
  const updates: D1PreparedStatement[] = [];
  const stmtConfirm = env.ETIC_SNAPSHOTS.prepare(
    `UPDATE work_order_action
        SET status = 'confirmed',
            verified_at_iso = ?,
            verified_in_snapshot = ?,
            snapshots_checked = snapshots_checked + 1
      WHERE id = ?`,
  );
  const stmtMissed = env.ETIC_SNAPSHOTS.prepare(
    `UPDATE work_order_action
        SET status = 'missed',
            verified_at_iso = ?,
            verified_in_snapshot = ?,
            snapshots_checked = snapshots_checked + 1
      WHERE id = ?`,
  );
  const stmtTouchOnly = env.ETIC_SNAPSHOTS.prepare(
    `UPDATE work_order_action
        SET snapshots_checked = snapshots_checked + 1
      WHERE id = ?`,
  );

  let confirmed = 0;
  let missed = 0;
  for (const a of rows) {
    const created = a.created_at_iso;
    // Compare just the date portion: action was logged on or before this snapshot's day
    const actionDate = created.slice(0, 10);
    const isAfter = snapshotDateKey > actionDate;
    if (a.expected_field) {
      const key = a.work_order_id + "::" + a.expected_field;
      if (changedSet.has(key)) {
        updates.push(stmtConfirm.bind(nowIso, snapshotDateKey, a.id));
        confirmed += 1;
      } else if (isAfter) {
        // First snapshot strictly AFTER the action with no matching change
        // -> the FM&A edit didn't make it in. Mark missed.
        updates.push(stmtMissed.bind(nowIso, snapshotDateKey, a.id));
        missed += 1;
      } else {
        // Same-day snapshot, field didn't change yet — leave pending.
        updates.push(stmtTouchOnly.bind(a.id));
      }
    } else {
      // 'other' actions: just bump the counter, no auto status change.
      updates.push(stmtTouchOnly.bind(a.id));
    }
  }

  if (updates.length) {
    const BATCH = 50;
    for (let i = 0; i < updates.length; i += BATCH) {
      await env.ETIC_SNAPSHOTS.batch(updates.slice(i, i + BATCH));
    }
  }
  return { checked: rows.length, confirmed, missed };
}

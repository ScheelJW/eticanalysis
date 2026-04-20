import type { RawWorkOrder } from "./yardCheck";

export type MelTier = "below" | "at" | "above" | "unknown";

const REMARK_INTERVAL: Record<MelTier, number | null> = {
  below: 3,
  at: 5,
  above: 10,
  unknown: null,
};

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
  updated_at_iso: string;
};

export async function ingestWorkOrderSnapshot(
  env: { ETIC_SNAPSHOTS: D1Database },
  dateKey: string,
  rows: RawWorkOrder[],
  updatedAtIso: string,
): Promise<void> {
  const stmtState = env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id, asset_id, remarks, parts_status, etic_raw, etic_date, mel_tier,
 last_remark_change_date, etic_push_count, first_etic_date, last_etic_date, cumulative_etic_slip_days
     FROM work_order_state WHERE work_order_id = ?`,
  );
  const insertLog = env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO work_order_changelog (work_order_id, snapshot_date_key, changed_at_iso, field, old_value, new_value)
     VALUES (?, ?, ?, ?, ?, ?)`,
  );
  const upsert = env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO work_order_state (
      work_order_id, asset_id, last_snapshot_date, remarks, parts_status, etic_raw, etic_date, mel_tier,
      last_remark_change_date, etic_push_count, first_etic_date, last_etic_date, cumulative_etic_slip_days, updated_at_iso ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
      updated_at_iso = excluded.updated_at_iso`,
  );

  for (const wo of rows) {
    const wid = (wo.workOrderId ?? "").trim();
    if (!wid) continue;

    const remarks = (wo.remarks ?? "").trim();
    const partsStatus = (wo.partsStatus ?? "").trim();
    const eticRaw = (wo.eticDue ?? "").trim();
    const eticDate = parseEticDate(eticRaw);
    const melTier = classifyMelTier(wo.currentMel ?? "");
    const assetId = (wo.assetId ?? "").trim();

    const prev = await stmtState.bind(wid).first<WoStateRow>();

    let lastRemarkChange = dateKey;
    let pushCount = 0;
    let firstEtic: string | null = eticDate;
    let lastEtic: string | null = eticDate;
    let slipSum = 0;

    if (prev) {
      lastRemarkChange = prev.last_remark_change_date;
      if (remarks !== (prev.remarks ?? "")) {
        lastRemarkChange = dateKey;
        await insertLog
          .bind(wid, dateKey, updatedAtIso, "remarks", prev.remarks ?? "", remarks)
          .run();
      }
      if (partsStatus !== (prev.parts_status ?? "")) {
        await insertLog
          .bind(wid, dateKey, updatedAtIso, "parts_status", prev.parts_status ?? "", partsStatus)
          .run();
      }
      if (eticRaw !== (prev.etic_raw ?? "")) {
        await insertLog
          .bind(wid, dateKey, updatedAtIso, "etic", prev.etic_raw ?? "", eticRaw)
          .run();
      }
      if (melTier !== (prev.mel_tier as MelTier)) {
        await insertLog
          .bind(wid, dateKey, updatedAtIso, "mel_tier", prev.mel_tier ?? "", melTier)
          .run();
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
        await insertLog
          .bind(wid, dateKey, updatedAtIso, "etic_date_slip", oldD, newD)
          .run();
      }
      if (!firstEtic && newD) firstEtic = newD;
      if (newD) lastEtic = newD;
      else if (lastEtic === null) lastEtic = prev.last_etic_date;
    } else {
      await insertLog.bind(wid, dateKey, updatedAtIso, "initial", "", "first_seen").run();
      firstEtic = eticDate;
      lastEtic = eticDate;
    }

    await upsert
      .bind(
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
        updatedAtIso,
      )
      .run();
  }
}

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
};

function rowToWatchRow(
  row: {
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
  },
  asOfDateKey: string,
): WatchRow {
  const tier = (row.mel_tier as MelTier) || "unknown";
  const interval = REMARK_INTERVAL[tier];
  const daysSince = calendarDaysBetween(row.last_remark_change_date, asOfDateKey);
  const stale = interval !== null && daysSince > interval;
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
  };
}

/** One work order’s latest state; staleness vs asOfDateKey. Returns null if not in index. */
export async function getWatchRowById(
  env: { ETIC_SNAPSHOTS: D1Database },
  workOrderId: string,
  asOfDateKey: string,
): Promise<WatchRow | null> {
  const row = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id, asset_id, mel_tier, parts_status, etic_raw, etic_date, remarks,
            last_remark_change_date, etic_push_count, cumulative_etic_slip_days, first_etic_date, last_etic_date, last_snapshot_date
     FROM work_order_state WHERE work_order_id = ?`,
  )
    .bind(workOrderId)
    .first<{
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
    }>();
  if (!row) return null;
  return rowToWatchRow(row, asOfDateKey);
}

/** Latest state per WO; remark staleness computed vs asOfDateKey (report date). */
export async function getWatchRowsLatest(env: { ETIC_SNAPSHOTS: D1Database }, asOfDateKey: string): Promise<WatchRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id, asset_id, mel_tier, parts_status, etic_raw, etic_date, remarks,
            last_remark_change_date, etic_push_count, cumulative_etic_slip_days, first_etic_date, last_etic_date, last_snapshot_date
     FROM work_order_state
     ORDER BY work_order_id`,
  ).all<{
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
  }>();

  const out: WatchRow[] = [];
  for (const row of r.results ?? []) {
    out.push(rowToWatchRow(row, asOfDateKey));
  }
  return out;
}

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

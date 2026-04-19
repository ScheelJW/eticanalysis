// Waiver card system.
//
// A "waiver" is a defect on a vehicle that management has formally accepted
// (doesn't affect safety/serviceability), so the shop should NOT chase it on
// future visits. The mechanic-facing mobile app reads approved waivers when
// they pull a truck into the bay; the desktop dashboard handles approval,
// per-vehicle lookup, annual re-verification, and printing the physical
// card that lives in the cab so the driver also knows.
//
// Lifecycle:
//   pending  → mechanic submits with photo + description
//   approved → management approves; appears on the card
//   rejected → management rejects with a reason; kept for audit
//
// Verification:
//   Every approved waiver must be re-verified at least once per year. The
//   approval seeds an 'initial' verification; mechanics + management can
//   add 'annual' / 'adhoc' verifications from either app. Name is required
//   on every action — there is no "system" actor here.
//
// Photo storage:
//   We mirror the yard-photo pattern — bytes in R2 under waiver-photos/<id>/...,
//   metadata columns on the waiver row itself. One photo per waiver is enough
//   for the documenting-the-defect use case the shop described.

type Env = { ETIC_SNAPSHOTS: D1Database; ETIC_BUCKET: R2Bucket };

export const WAIVER_PHOTO_PREFIX = "waiver-photos/";
export const WAIVER_VERIFY_PHOTO_PREFIX = "waiver-verify-photos/";

/** ms in a year, used to bucket "verification overdue" badges. */
export const WAIVER_VERIFY_INTERVAL_MS = 365 * 24 * 60 * 60 * 1000;
/** Soft warning threshold (1 month before annual). */
export const WAIVER_VERIFY_DUE_SOON_MS = (365 - 30) * 24 * 60 * 60 * 1000;

export type WaiverStatus = "pending" | "approved" | "rejected";

export type Waiver = {
  id: number;
  assetId: string;
  title: string;
  description: string;
  status: WaiverStatus;
  hasPhoto: boolean;
  photoUrl: string;
  submittedBy: string;
  submittedAtIso: string;
  reviewedBy: string;
  reviewedAtIso: string;
  reviewedNote: string;
  lastVerifiedBy: string;
  lastVerifiedAtIso: string;
  /** Bucketed staleness for annual re-verification — drives badge color. */
  verifyState: WaiverVerifyState;
  /** Whole days since the most-recent verification (or since approval if none). */
  daysSinceVerified: number | null;
};

/**
 * fresh   = verified within the last ~11 months (well under annual)
 * dueSoon = within the last ~12 months but past the 11-month nag threshold
 * overdue = more than 12 months since last verification (or approval, if no
 *           verifications yet) — gets the loud badge on the card
 * na      = pending/rejected; verification is not meaningful
 */
export type WaiverVerifyState = "fresh" | "dueSoon" | "overdue" | "na";

export type WaiverVerification = {
  id: number;
  waiverId: number;
  verifiedBy: string;
  verifiedAtIso: string;
  note: string;
  kind: "initial" | "annual" | "adhoc";
  /** Present when mechanic attached a photo during verify. */
  hasPhoto: boolean;
  photoUrl: string;
};

export type SubmitWaiverInput = {
  assetId: string;
  title: string;
  description?: string;
  submittedBy: string;
  photo?: {
    body: ArrayBuffer;
    contentType: string;
  };
};

type WaiverRow = {
  id: number;
  asset_id: string;
  title: string;
  description: string | null;
  photo_r2_key: string | null;
  photo_content_type: string | null;
  status: WaiverStatus;
  submitted_by: string;
  submitted_at_iso: string;
  reviewed_by: string | null;
  reviewed_at_iso: string | null;
  reviewed_note: string | null;
  last_verified_by: string | null;
  last_verified_at_iso: string | null;
};

type VerificationRow = {
  id: number;
  waiver_id: number;
  verified_by: string;
  verified_at_iso: string;
  note: string | null;
  kind: "initial" | "annual" | "adhoc";
  photo_r2_key: string | null;
  photo_content_type: string | null;
};

const SELECT_WAIVER_COLS =
  "id, asset_id, title, description, photo_r2_key, photo_content_type, status, " +
  "submitted_by, submitted_at_iso, reviewed_by, reviewed_at_iso, reviewed_note, " +
  "last_verified_by, last_verified_at_iso";

/** Remove legacy PATS bulk-import boilerplate from stored descriptions. */
function stripLegacyPatsDescription(raw: string | null | undefined): string {
  const t = (raw ?? "").trim();
  if (!t) return "";
  const kept: string[] = [];
  for (const line of t.split(/\r?\n/)) {
    const s = line.trim();
    if (!s) continue;
    if (/^legacy nei pats import\.?$/i.test(s)) continue;
    if (/^pats waiver_row_id:/i.test(s)) continue;
    if (/^pats vehicle_id:/i.test(s)) continue;
    if (/^original registration/i.test(s)) continue;
    if (/^recorded employee:/i.test(s)) continue;
    const fm = s.match(/^full note:\s*(.+)$/i);
    if (fm) {
      const inner = fm[1].trim();
      if (inner) kept.push(inner);
      continue;
    }
    kept.push(line);
  }
  return kept.join("\n").trim();
}

/** Exclude noise rows from SQL reads (lists, counts, by-id). Matches `scripts/pats-waiver-scrape/is-noise-waiver.mjs`. */
const WAIVER_ROW_IS_NOISE_SQL = `(
  lower(trim(ifnull(title,''))) LIKE '%no waiver information%'
  OR lower(trim(ifnull(description,''))) LIKE '%no waiver information%'
  OR lower(trim(ifnull(title,''))) LIKE '%no waivered items%'
  OR lower(trim(ifnull(description,''))) LIKE '%no waivered items%'
  OR lower(trim(ifnull(title,''))) LIKE '%no waiverable items%'
  OR lower(trim(ifnull(description,''))) LIKE '%no waiverable items%'
  OR (trim(ifnull(title,'')) = '' AND trim(ifnull(description,'')) = '')
)`;

function rowToWaiver(r: WaiverRow): Waiver {
  const lastVerifiedAt = r.last_verified_at_iso ?? "";
  // "Anchor" date for staleness: prefer the last explicit verification; fall
  // back to the approval date so a freshly approved waiver isn't immediately
  // overdue. Pending/rejected waivers don't track verification at all.
  const anchor =
    r.status === "approved"
      ? lastVerifiedAt || r.reviewed_at_iso || ""
      : "";
  let verifyState: WaiverVerifyState = "na";
  let daysSinceVerified: number | null = null;
  if (r.status === "approved" && anchor) {
    const ms = Date.now() - Date.parse(anchor);
    if (Number.isFinite(ms)) {
      daysSinceVerified = Math.max(0, Math.floor(ms / (24 * 60 * 60 * 1000)));
      if (ms >= WAIVER_VERIFY_INTERVAL_MS) verifyState = "overdue";
      else if (ms >= WAIVER_VERIFY_DUE_SOON_MS) verifyState = "dueSoon";
      else verifyState = "fresh";
    }
  }
  const descClean = stripLegacyPatsDescription(r.description);
  return {
    id: r.id,
    assetId: r.asset_id,
    title: r.title,
    description: descClean,
    status: r.status,
    hasPhoto: !!r.photo_r2_key,
    photoUrl: r.photo_r2_key ? `/api/waivers/${r.id}/photo` : "",
    submittedBy: r.submitted_by,
    submittedAtIso: r.submitted_at_iso,
    reviewedBy: r.reviewed_by ?? "",
    reviewedAtIso: r.reviewed_at_iso ?? "",
    reviewedNote: r.reviewed_note ?? "",
    lastVerifiedBy: r.last_verified_by ?? "",
    lastVerifiedAtIso: lastVerifiedAt,
    verifyState,
    daysSinceVerified,
  };
}

function rowToVerification(r: VerificationRow): WaiverVerification {
  const hasPhoto = !!r.photo_r2_key;
  return {
    id: r.id,
    waiverId: r.waiver_id,
    verifiedBy: r.verified_by,
    verifiedAtIso: r.verified_at_iso,
    note: r.note ?? "",
    kind: r.kind,
    hasPhoto,
    photoUrl: hasPhoto ? `/api/waivers/verification/${r.id}/photo` : "",
  };
}

function randomHex(len = 12): string {
  const b = new Uint8Array(len);
  crypto.getRandomValues(b);
  let s = "";
  for (const x of b) s += x.toString(16).padStart(2, "0");
  return s;
}

function extensionForContentType(ct: string): string {
  const c = (ct || "").toLowerCase();
  if (c.includes("png")) return "png";
  if (c.includes("webp")) return "webp";
  if (c.includes("heic") || c.includes("heif")) return "heic";
  if (c.includes("gif")) return "gif";
  return "jpg";
}

// ---------------------------------------------------------------------------
// Submit / read
// ---------------------------------------------------------------------------

export async function submitWaiver(env: Env, input: SubmitWaiverInput): Promise<Waiver> {
  const assetId = (input.assetId ?? "").trim();
  const title = (input.title ?? "").trim();
  const submittedBy = (input.submittedBy ?? "").trim();
  if (!assetId) throw new Error("assetId required");
  if (!title) throw new Error("title required");
  if (!submittedBy) throw new Error("submittedBy required");

  const description = (input.description ?? "").trim();
  const submittedAtIso = new Date().toISOString();

  // Insert the row first (without photo metadata) so we have an id to key the
  // R2 path on. Then upload + UPDATE in a follow-up. If the upload fails we
  // still have a clean pending row the mechanic can retry photo on.
  const inserted = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO waiver
       (asset_id, title, description, status, submitted_by, submitted_at_iso)
     VALUES (?, ?, ?, 'pending', ?, ?)
     RETURNING ${SELECT_WAIVER_COLS}`,
  )
    .bind(assetId, title, description, submittedBy, submittedAtIso)
    .first<WaiverRow>();
  if (!inserted) throw new Error("failed to insert waiver");

  if (input.photo && input.photo.body && input.photo.body.byteLength > 0) {
    const ct = input.photo.contentType || "image/jpeg";
    const ext = extensionForContentType(ct);
    const r2Key = `${WAIVER_PHOTO_PREFIX}${inserted.id}/${Date.now()}-${randomHex()}.${ext}`;
    await env.ETIC_BUCKET.put(r2Key, input.photo.body, {
      httpMetadata: { contentType: ct },
    });
    const updated = await env.ETIC_SNAPSHOTS.prepare(
      `UPDATE waiver SET photo_r2_key = ?, photo_content_type = ?
        WHERE id = ?
        RETURNING ${SELECT_WAIVER_COLS}`,
    )
      .bind(r2Key, ct, inserted.id)
      .first<WaiverRow>();
    if (updated) return rowToWaiver(updated);
  }
  return rowToWaiver(inserted);
}

/**
 * List all waivers for one asset. By default returns approved + pending so
 * mechanics see "in-flight" requests too; pass {approvedOnly: true} for the
 * printable card view.
 */
export async function listWaiversForAsset(
  env: Env,
  assetId: string,
  opts?: { approvedOnly?: boolean },
): Promise<Waiver[]> {
  const id = (assetId ?? "").trim();
  if (!id) return [];
  const where = opts?.approvedOnly
    ? "asset_id = ? AND status = 'approved'"
    : "asset_id = ? AND status IN ('pending','approved')";
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT ${SELECT_WAIVER_COLS} FROM waiver
      WHERE ${where} AND NOT ${WAIVER_ROW_IS_NOISE_SQL}
      ORDER BY status DESC, submitted_at_iso DESC`,
  )
    .bind(id)
    .all<WaiverRow>();
  return (r.results ?? []).map(rowToWaiver);
}

/**
 * All pending submissions across the fleet — drives the desktop "Pending
 * review" queue. Newest first.
 */
export async function listPendingWaivers(env: Env): Promise<Waiver[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT ${SELECT_WAIVER_COLS} FROM waiver
      WHERE status = 'pending' AND NOT ${WAIVER_ROW_IS_NOISE_SQL}
      ORDER BY submitted_at_iso DESC`,
  ).all<WaiverRow>();
  return (r.results ?? []).map(rowToWaiver);
}

export async function getWaiverById(env: Env, id: number): Promise<Waiver | null> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT ${SELECT_WAIVER_COLS} FROM waiver WHERE id = ? AND NOT ${WAIVER_ROW_IS_NOISE_SQL}`,
  )
    .bind(id)
    .first<WaiverRow>();
  return r ? rowToWaiver(r) : null;
}

// ---------------------------------------------------------------------------
// Approve / reject
// ---------------------------------------------------------------------------

export async function approveWaiver(
  env: Env,
  id: number,
  reviewedBy: string,
  note?: string,
): Promise<Waiver> {
  const by = (reviewedBy ?? "").trim();
  if (!by) throw new Error("reviewedBy required");
  const nowIso = new Date().toISOString();
  const updated = await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE waiver
        SET status = 'approved',
            reviewed_by = ?,
            reviewed_at_iso = ?,
            reviewed_note = ?,
            last_verified_by = ?,
            last_verified_at_iso = ?
      WHERE id = ? AND status = 'pending'
      RETURNING ${SELECT_WAIVER_COLS}`,
  )
    .bind(by, nowIso, (note ?? "").trim() || null, by, nowIso, id)
    .first<WaiverRow>();
  if (!updated) throw new Error("waiver not found or not pending");
  // Seed an 'initial' verification entry so the audit log has a starting
  // point — otherwise the very first annual reminder would show "never
  // verified" instead of "verified at approval".
  await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO waiver_verification (waiver_id, verified_by, verified_at_iso, note, kind)
     VALUES (?, ?, ?, ?, 'initial')`,
  )
    .bind(id, by, nowIso, (note ?? "").trim() || null)
    .run();
  return rowToWaiver(updated);
}

export async function rejectWaiver(
  env: Env,
  id: number,
  reviewedBy: string,
  reason: string,
): Promise<Waiver> {
  const by = (reviewedBy ?? "").trim();
  const r = (reason ?? "").trim();
  if (!by) throw new Error("reviewedBy required");
  if (!r) throw new Error("rejection reason required");
  const nowIso = new Date().toISOString();
  const updated = await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE waiver
        SET status = 'rejected',
            reviewed_by = ?,
            reviewed_at_iso = ?,
            reviewed_note = ?
      WHERE id = ? AND status = 'pending'
      RETURNING ${SELECT_WAIVER_COLS}`,
  )
    .bind(by, nowIso, r, id)
    .first<WaiverRow>();
  if (!updated) throw new Error("waiver not found or not pending");
  return rowToWaiver(updated);
}

// ---------------------------------------------------------------------------
// Verify
// ---------------------------------------------------------------------------

export async function verifyWaiver(
  env: Env,
  id: number,
  verifiedBy: string,
  note?: string,
  kind: "annual" | "adhoc" = "annual",
  photo?: { body: ArrayBuffer; contentType: string },
): Promise<Waiver> {
  const by = (verifiedBy ?? "").trim();
  if (!by) throw new Error("verifiedBy required");
  const nowIso = new Date().toISOString();
  // Only approved waivers can be verified — pending/rejected has no card.
  const cur = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT status FROM waiver WHERE id = ?`,
  )
    .bind(id)
    .first<{ status: WaiverStatus }>();
  if (!cur) throw new Error("waiver not found");
  if (cur.status !== "approved") throw new Error("only approved waivers can be verified");

  const inserted = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO waiver_verification (waiver_id, verified_by, verified_at_iso, note, kind, photo_r2_key, photo_content_type)
     VALUES (?, ?, ?, ?, ?, NULL, NULL)
     RETURNING id`,
  )
    .bind(id, by, nowIso, (note ?? "").trim() || null, kind)
    .first<{ id: number }>();
  if (!inserted?.id) throw new Error("failed to record verification");

  if (photo && photo.body.byteLength > 0) {
    const ct = photo.contentType || "image/jpeg";
    const ext = extensionForContentType(ct);
    const r2Key = `${WAIVER_VERIFY_PHOTO_PREFIX}${inserted.id}/${Date.now()}-${randomHex()}.${ext}`;
    await env.ETIC_BUCKET.put(r2Key, photo.body, {
      httpMetadata: { contentType: ct },
    });
    await env.ETIC_SNAPSHOTS.prepare(
      `UPDATE waiver_verification SET photo_r2_key = ?, photo_content_type = ? WHERE id = ?`,
    )
      .bind(r2Key, ct, inserted.id)
      .run();
  }

  const updated = await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE waiver
        SET last_verified_by = ?, last_verified_at_iso = ?
      WHERE id = ?
      RETURNING ${SELECT_WAIVER_COLS}`,
  )
    .bind(by, nowIso, id)
    .first<WaiverRow>();
  if (!updated) throw new Error("waiver disappeared mid-verify");
  return rowToWaiver(updated);
}

export async function listVerifications(env: Env, waiverId: number): Promise<WaiverVerification[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, waiver_id, verified_by, verified_at_iso, note, kind,
            photo_r2_key, photo_content_type
       FROM waiver_verification
      WHERE waiver_id = ?
      ORDER BY verified_at_iso DESC`,
  )
    .bind(waiverId)
    .all<VerificationRow>();
  return (r.results ?? []).map(rowToVerification);
}

export async function getVerificationPhoto(env: Env, verificationId: number): Promise<{
  body: ReadableStream;
  contentType: string;
} | null> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT photo_r2_key, photo_content_type FROM waiver_verification WHERE id = ?`,
  )
    .bind(verificationId)
    .first<{ photo_r2_key: string | null; photo_content_type: string | null }>();
  if (!r?.photo_r2_key) return null;
  const obj = await env.ETIC_BUCKET.get(r.photo_r2_key);
  if (!obj) return null;
  return { body: obj.body, contentType: r.photo_content_type ?? "image/jpeg" };
}

// ---------------------------------------------------------------------------
// Delete (admin)
// ---------------------------------------------------------------------------

export async function deleteWaiver(env: Env, id: number): Promise<boolean> {
  const meta = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT photo_r2_key FROM waiver WHERE id = ?`,
  )
    .bind(id)
    .first<{ photo_r2_key: string | null }>();
  if (!meta) return false;
  if (meta.photo_r2_key) {
    try {
      await env.ETIC_BUCKET.delete(meta.photo_r2_key);
    } catch {
      // Best effort — never block delete on R2 cleanup.
    }
  }
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `DELETE FROM waiver WHERE id = ?`,
  )
    .bind(id)
    .run();
  const m = (r as unknown as { meta?: { changes?: number } }).meta;
  return (m?.changes ?? 0) > 0;
}

// ---------------------------------------------------------------------------
// Photo serving
// ---------------------------------------------------------------------------

export async function getWaiverPhoto(env: Env, id: number): Promise<{
  body: ReadableStream;
  contentType: string;
} | null> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT photo_r2_key, photo_content_type FROM waiver WHERE id = ?`,
  )
    .bind(id)
    .first<{ photo_r2_key: string | null; photo_content_type: string | null }>();
  if (!r || !r.photo_r2_key) return null;
  const obj = await env.ETIC_BUCKET.get(r.photo_r2_key);
  if (!obj) return null;
  return { body: obj.body, contentType: r.photo_content_type ?? "image/jpeg" };
}

// ---------------------------------------------------------------------------
// Per-asset counts (for badges everywhere asset id is rendered)
// ---------------------------------------------------------------------------

export type WaiverCount = {
  /** Approved waivers currently on the card. */
  approved: number;
  /** Pending submissions awaiting management review. */
  pending: number;
  /** Approved waivers whose annual verification is past due. */
  overdueVerify: number;
};

/**
 * One DB pass that returns Map<assetId, WaiverCount> for every asset that
 * has at least one approved or pending waiver. Cheap — one row per asset.
 * The desktop dashboard fetches this on tab enter (same TTL as sightings)
 * so the asset-id badge can show "2 waivers" without per-row queries.
 */
export async function getWaiverCounts(env: Env): Promise<Map<string, WaiverCount>> {
  const intervalDays = Math.floor(WAIVER_VERIFY_INTERVAL_MS / (24 * 60 * 60 * 1000));
  // SQLite's date('now','-365 days') gives us a comparable YYYY-MM-DDT00:00:00
  // boundary. We compare against last_verified_at_iso (or reviewed_at_iso as
  // the implicit initial verification) so the count matches what rowToWaiver()
  // calculates per row.
  const rows = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT asset_id,
            SUM(CASE WHEN status = 'approved' THEN 1 ELSE 0 END) AS approved,
            SUM(CASE WHEN status = 'pending'  THEN 1 ELSE 0 END) AS pending,
            SUM(CASE
                  WHEN status = 'approved'
                   AND COALESCE(last_verified_at_iso, reviewed_at_iso) < datetime('now', ?)
                  THEN 1 ELSE 0 END) AS overdue_verify
       FROM waiver
      WHERE status IN ('approved','pending')
        AND NOT ${WAIVER_ROW_IS_NOISE_SQL}
      GROUP BY asset_id`,
  )
    .bind(`-${intervalDays} days`)
    .all<{ asset_id: string; approved: number; pending: number; overdue_verify: number }>();
  const out = new Map<string, WaiverCount>();
  for (const r of rows.results ?? []) {
    if (!r.asset_id) continue;
    out.set(r.asset_id, {
      approved: Number(r.approved) || 0,
      pending: Number(r.pending) || 0,
      overdueVerify: Number(r.overdue_verify) || 0,
    });
  }
  return out;
}

/**
 * Accident / abuse cost-recovery program tracking (VFM/VMS).
 * D1: abuse_tracker_case, abuse_tracker_note, abuse_tracker_attachment, abuse_tracker_seq.
 * R2 keys: abuse-tracker/<caseId>/<uuid>-<safeName>
 */

import { recordCheck } from "./yardSession";

type Env = { ETIC_SNAPSHOTS: D1Database; ETIC_BUCKET: R2Bucket };

export const ABUSE_TRACKER_R2_PREFIX = "abuse-tracker/";

export type AbuseCaseType = "accident" | "abuse";
export type AbuseCaseStage =
  | "initial"
  | "awaiting_estimates"
  | "pending_legal_release"
  | "repair_in_mx_contract"
  | "repair_downtown"
  | "repair_on_base"
  | "no_repair_tracking"
  | "closed";

/** Map legacy DB/UI stage values to current enum (migration already normalizes DB). */
export const LEGACY_ABUSE_STAGE_MAP: Record<string, AbuseCaseStage> = {
  intake: "initial",
  estimates: "awaiting_estimates",
  release_pending: "pending_legal_release",
  approved_work: "repair_in_mx_contract",
  initial: "initial",
  awaiting_estimates: "awaiting_estimates",
  pending_legal_release: "pending_legal_release",
  repair_in_mx_contract: "repair_in_mx_contract",
  repair_downtown: "repair_downtown",
  repair_on_base: "repair_on_base",
  no_repair_tracking: "no_repair_tracking",
  closed: "closed",
};

export function normalizeAbuseCaseStage(raw: string | null | undefined): AbuseCaseStage {
  const k = (raw ?? "").trim();
  return LEGACY_ABUSE_STAGE_MAP[k] ?? "initial";
}

export const ABUSE_STAGE_VALUES: readonly AbuseCaseStage[] = [
  "initial",
  "awaiting_estimates",
  "pending_legal_release",
  "repair_in_mx_contract",
  "repair_downtown",
  "repair_on_base",
  "no_repair_tracking",
  "closed",
] as const;

export function isValidAbuseStageInput(raw: string): boolean {
  const n = normalizeAbuseCaseStage(raw);
  return (ABUSE_STAGE_VALUES as readonly string[]).includes(n);
}
export type AbuseAttachmentKind = "damage_photo" | "release_letter" | "estimate" | "other";

export type AbuseEstimate = {
  vendor: string;
  amount: number | null;
  note: string;
};

/** While status is “awaiting package” — what is still outstanding. */
export type AbusePackageChecklist = {
  sf91?: boolean;
  photos?: boolean;
  /** Vehicle has arrived at the VM maintenance compound / lot for intake. */
  vehicleAtVmCompound?: boolean;
};

export type AbuseCaseRow = {
  id: number;
  control_number: string;
  case_type: AbuseCaseType;
  asset_id: string;
  owning_unit: string;
  shop: string;
  make_model: string;
  veh_nomen: string;
  mgmt_cd: string;
  determination: string;
  responsible_party: string;
  reimbursed_to_vm: number;
  reimbursed_at_iso: string | null;
  reimbursed_note: string;
  stage: AbuseCaseStage;
  vehicle_location: string;
  estimates_json: string;
  email_token: string;
  created_at_iso: string;
  updated_at_iso: string;
  created_by: string;
  closed_at_iso: string | null;
  work_order_id: string;
  tracking_active: number;
  package_checklist_json: string;
  estimates_runner: string;
  estimates_downtown_planned_date: string;
};

export type AbuseNoteRow = {
  id: number;
  case_id: number;
  body: string;
  author: string;
  at_iso: string;
};

export type AbuseAttachmentRow = {
  id: number;
  case_id: number;
  kind: AbuseAttachmentKind;
  r2_key: string;
  filename: string;
  content_type: string;
  size_bytes: number;
  uploaded_by: string;
  uploaded_at_iso: string;
  source: "web" | "email";
};

export type AbuseTimelineKind =
  | "case_opened"
  | "stage"
  | "location"
  | "responsible_unit"
  | "work_order"
  | "tracking_mode"
  | "note"
  | "attachment"
  | "reimbursement"
  | "closed"
  | "other";

export type AbuseTimelineRow = {
  id: number;
  case_id: number;
  at_iso: string;
  kind: AbuseTimelineKind;
  payload_json: string;
  created_by: string;
};

function randomToken(len: number): string {
  const bytes = new Uint8Array(len);
  crypto.getRandomValues(bytes);
  let s = "";
  for (const b of bytes) s += b.toString(16).padStart(2, "0");
  return s.slice(0, len);
}

function safeFileName(name: string): string {
  return name.replace(/[^A-Za-z0-9._-]/g, "_").slice(0, 120) || "file";
}

export async function allocateControlNumber(env: Env, caseType: AbuseCaseType): Promise<string> {
  const y = new Date().getUTCFullYear();
  await env.ETIC_SNAPSHOTS.prepare(`INSERT INTO abuse_tracker_seq (year, last_n) VALUES (?, 0) ON CONFLICT DO NOTHING`)
    .bind(y)
    .run();
  await env.ETIC_SNAPSHOTS.prepare(`UPDATE abuse_tracker_seq SET last_n = last_n + 1 WHERE year = ?`)
    .bind(y)
    .run();
  const r = await env.ETIC_SNAPSHOTS.prepare(`SELECT last_n FROM abuse_tracker_seq WHERE year = ?`)
    .bind(y)
    .first<{ last_n: number }>();
  const n = r?.last_n ?? 1;
  const prefix = caseType === "abuse" ? "ABU" : "ACC";
  return `${prefix}-${y}-${String(n).padStart(5, "0")}`;
}

export async function findCaseByEmailToken(env: Env, token: string): Promise<AbuseCaseRow | null> {
  const t = token.trim().toLowerCase();
  if (!t) return null;
  return env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_case WHERE lower(email_token) = ?`)
    .bind(t)
    .first<AbuseCaseRow>();
}

type FleetSnapshot = {
  owning_unit: string;
  shop: string;
  make_model: string;
  veh_nomen: string;
  mgmt_cd: string;
};

/** Latest Fleet P&A row for an asset (refreshed each workbook ingest). */
async function hydrateFromFleetAssetCurrent(env: Env, assetId: string): Promise<FleetSnapshot | null> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT owning_unit, shop, make_model, veh_nomen, mgmt_cd FROM fleet_asset_current WHERE asset_id = ?`,
  )
    .bind(assetId.trim())
    .first<FleetSnapshot>();
  if (!r) return null;
  return {
    owning_unit: r.owning_unit ?? "",
    shop: r.shop ?? "",
    make_model: r.make_model ?? "",
    veh_nomen: r.veh_nomen ?? "",
    mgmt_cd: r.mgmt_cd ?? "",
  };
}

async function hydrateFleetFields(env: Env, assetId: string): Promise<FleetSnapshot> {
  const fromFleet = await hydrateFromFleetAssetCurrent(env, assetId);
  if (fromFleet) return fromFleet;
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT owning_unit, shop, make_model, veh_nomen, mgmt_cd
     FROM work_order_state WHERE asset_id = ?
     ORDER BY updated_at_iso DESC LIMIT 1`,
  )
    .bind(assetId.trim())
    .first<{
      owning_unit: string;
      shop: string;
      make_model: string;
      veh_nomen: string;
      mgmt_cd: string;
    }>();
  return {
    owning_unit: r?.owning_unit ?? "",
    shop: r?.shop ?? "",
    make_model: r?.make_model ?? "",
    veh_nomen: r?.veh_nomen ?? "",
    mgmt_cd: r?.mgmt_cd ?? "",
  };
}

/** Latest open WO row for an asset (for linking A/A cases to the active work order). */
export async function resolveLatestWorkOrderIdForAsset(env: Env, assetId: string): Promise<string | null> {
  const aid = assetId.trim();
  if (!aid) return null;
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT work_order_id FROM work_order_state WHERE asset_id = ? ORDER BY datetime(updated_at_iso) DESC LIMIT 1`,
  )
    .bind(aid)
    .first<{ work_order_id: string }>();
  const wid = (r?.work_order_id ?? "").trim();
  return wid || null;
}

async function hydrateFromWorkOrderId(
  env: Env,
  workOrderId: string,
): Promise<{ asset_id: string; owning_unit: string; shop: string; make_model: string; veh_nomen: string; mgmt_cd: string } | null> {
  const wid = workOrderId.trim();
  if (!wid) return null;
  return env.ETIC_SNAPSHOTS.prepare(
    `SELECT asset_id, owning_unit, shop, make_model, veh_nomen, mgmt_cd
     FROM work_order_state WHERE work_order_id = ? LIMIT 1`,
  )
    .bind(wid)
    .first<{
      asset_id: string;
      owning_unit: string;
      shop: string;
      make_model: string;
      veh_nomen: string;
      mgmt_cd: string;
    }>();
}

export type CreateAbuseCaseInput = {
  caseType: AbuseCaseType;
  assetId: string;
  workOrderId?: string;
  determination?: string;
  responsibleParty?: string;
  vehicleLocation?: string;
  createdBy: string;
};

export async function createAbuseCase(env: Env, input: CreateAbuseCaseInput): Promise<AbuseCaseRow> {
  const woTrim = (input.workOrderId ?? "").trim();
  const woRow = woTrim ? await hydrateFromWorkOrderId(env, woTrim) : null;
  let assetId = (input.assetId ?? "").trim();
  if (!assetId && woRow) assetId = (woRow.asset_id ?? "").trim();
  if (!assetId) throw new Error("Asset id is required (WO may be added later and does not need to match ingest yet).");
  if (woRow && (input.assetId ?? "").trim() && (woRow.asset_id ?? "").trim() !== assetId) {
    throw new Error("Asset id does not match that work order.");
  }
  const caseType = input.caseType;
  if (caseType !== "accident" && caseType !== "abuse") throw new Error("caseType must be accident or abuse");

  const open = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id FROM abuse_tracker_case WHERE asset_id = ? AND case_type = ? AND closed_at_iso IS NULL`,
  )
    .bind(assetId, caseType)
    .first<{ id: number }>();
  if (open) throw new Error(`An open ${caseType} case already exists for this asset.`);

  const control = await allocateControlNumber(env, caseType);
  const token = randomToken(16);
  const now = new Date().toISOString();
  // Snapshot org fields at case creation only — never updated from later ingests.
  const fleet = woRow
    ? {
        owning_unit: woRow.owning_unit ?? "",
        shop: woRow.shop ?? "",
        make_model: woRow.make_model ?? "",
        veh_nomen: woRow.veh_nomen ?? "",
        mgmt_cd: woRow.mgmt_cd ?? "",
      }
    : await hydrateFleetFields(env, assetId);

  const ins = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO abuse_tracker_case (
       control_number, case_type, asset_id, work_order_id, owning_unit, shop, make_model, veh_nomen, mgmt_cd,
       determination, responsible_party, reimbursed_to_vm, reimbursed_at_iso, reimbursed_note,
       stage, vehicle_location, estimates_json, email_token, created_at_iso, updated_at_iso, created_by, tracking_active,
       package_checklist_json, estimates_runner, estimates_downtown_planned_date
     ) VALUES (?,?,?,?,?,?,?,?,?,?,?,0,NULL,'','initial',?,?,?,?,?,?,1,?,?,?)`,
  )
    .bind(
      control,
      caseType,
      assetId,
      woTrim,
      fleet.owning_unit,
      fleet.shop,
      fleet.make_model,
      fleet.veh_nomen,
      fleet.mgmt_cd,
      (input.determination ?? "").trim(),
      (input.responsibleParty ?? "").trim(),
      (input.vehicleLocation ?? "").trim(),
      "[]",
      token,
      now,
      now,
      (input.createdBy ?? "").trim(),
      "{}",
      "",
      "",
    )
    .run();
  const id = Number(ins.meta.last_row_id ?? 0);
  if (!id) throw new Error("failed to create case");
  if (!woTrim) {
    const linked = await resolveLatestWorkOrderIdForAsset(env, assetId);
    if (linked) {
      await env.ETIC_SNAPSHOTS.prepare(`UPDATE abuse_tracker_case SET work_order_id = ?, updated_at_iso = ? WHERE id = ?`)
        .bind(linked, now, id)
        .run();
    }
  }
  const row = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_case WHERE id = ?`)
    .bind(id)
    .first<AbuseCaseRow>();
  if (!row) throw new Error("case not found after insert");
  await insertAbuseTimeline(env, {
    caseId: id,
    kind: "case_opened",
    payload: { control_number: control, case_type: caseType, asset_id: assetId },
    createdBy: (input.createdBy ?? "").trim() || "system",
    atIso: now,
  });
  return row;
}

export type UpdateAbuseCaseInput = {
  workOrderId?: string;
  determination?: string;
  responsibleParty?: string;
  reimbursedToVm?: boolean;
  reimbursedAtIso?: string | null;
  reimbursedNote?: string;
  stage?: AbuseCaseStage;
  vehicleLocation?: string;
  estimates?: AbuseEstimate[];
  closed?: boolean;
  trackingActive?: boolean;
  timelineAuthor?: string;
  packageChecklist?: AbusePackageChecklist;
  estimatesRunner?: string;
  estimatesDowntownPlannedDate?: string;
};

function parsePackageChecklistJson(raw: string | null | undefined): AbusePackageChecklist {
  try {
    const o = JSON.parse(raw || "{}");
    if (!o || typeof o !== "object") return {};
    return {
      sf91: !!o.sf91,
      photos: !!o.photos,
      vehicleAtVmCompound: !!o.vehicleAtVmCompound,
    };
  } catch {
    return {};
  }
}

async function insertAbuseTimeline(
  env: Env,
  opts: { caseId: number; kind: AbuseTimelineKind; payload: Record<string, unknown>; createdBy: string; atIso?: string },
): Promise<void> {
  const at = opts.atIso ?? new Date().toISOString();
  const payload = JSON.stringify(opts.payload ?? {});
  const by = (opts.createdBy ?? "").trim() || "system";
  try {
    await env.ETIC_SNAPSHOTS.prepare(
      `INSERT INTO abuse_tracker_timeline (case_id, at_iso, kind, payload_json, created_by) VALUES (?,?,?,?,?)`,
    )
      .bind(opts.caseId, at, opts.kind, payload, by)
      .run();
  } catch (e) {
    console.warn(
      JSON.stringify({
        level: "warn",
        message: "abuse_tracker_timeline insert skipped (migration 0031 applied?)",
        err: e instanceof Error ? e.message : String(e),
      }),
    );
  }
}

export async function listAbuseTimeline(env: Env, caseId: number): Promise<AbuseTimelineRow[]> {
  try {
    const r = await env.ETIC_SNAPSHOTS.prepare(
      `SELECT * FROM abuse_tracker_timeline WHERE case_id = ? ORDER BY datetime(at_iso) ASC, id ASC`,
    )
      .bind(caseId)
      .all<AbuseTimelineRow>();
    return r.results ?? [];
  } catch {
    return [];
  }
}

/** Distinct owning_unit values from fleet_asset_current (Fleet P&A). */
export async function listFleetOwningUnits(env: Env, limit: number): Promise<string[]> {
  const lim = Math.min(2000, Math.max(1, limit || 500));
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT DISTINCT trim(owning_unit) AS u FROM fleet_asset_current
     WHERE trim(owning_unit) != ''
     ORDER BY u ASC LIMIT ?`,
  )
    .bind(lim)
    .all<{ u: string }>();
  const out: string[] = [];
  for (const row of r.results ?? []) {
    const u = (row.u ?? "").trim();
    if (u && !out.includes(u)) out.push(u);
  }
  return out;
}

export async function updateAbuseCase(
  env: Env,
  caseId: number,
  patch: UpdateAbuseCaseInput,
): Promise<AbuseCaseRow | null> {
  const cur = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_case WHERE id = ?`)
    .bind(caseId)
    .first<AbuseCaseRow>();
  if (!cur) return null;
  const now = new Date().toISOString();
  const tlAuthor = (patch.timelineAuthor ?? "").trim() || "Manager";
  let estimatesJson = cur.estimates_json;
  if (patch.estimates !== undefined) {
    estimatesJson = JSON.stringify(patch.estimates);
  }
  const reimbursedToVm =
    patch.reimbursedToVm !== undefined ? (patch.reimbursedToVm ? 1 : 0) : cur.reimbursed_to_vm;
  const reimbursedAtIso =
    patch.reimbursedAtIso !== undefined ? patch.reimbursedAtIso : cur.reimbursed_at_iso;
  const reimbursedNote =
    patch.reimbursedNote !== undefined ? patch.reimbursedNote.trim() : cur.reimbursed_note;
  const determination = patch.determination !== undefined ? patch.determination.trim() : cur.determination;
  const responsibleParty =
    patch.responsibleParty !== undefined ? patch.responsibleParty.trim() : cur.responsible_party;
  const stageRaw = patch.stage !== undefined ? patch.stage : normalizeAbuseCaseStage(String(cur.stage));
  const stage = normalizeAbuseCaseStage(stageRaw);
  const vehicleLocation =
    patch.vehicleLocation !== undefined ? patch.vehicleLocation.trim() : cur.vehicle_location;
  let workOrderId =
    patch.workOrderId !== undefined ? patch.workOrderId.trim() : cur.work_order_id ?? "";
  let closedAt = cur.closed_at_iso;
  if (patch.closed === true) closedAt = now;
  if (patch.closed === false) closedAt = null;

  const trackingActive =
    patch.trackingActive !== undefined ? (patch.trackingActive ? 1 : 0) : (cur.tracking_active ?? 1);

  if (!workOrderId && !closedAt && trackingActive) {
    const linked = await resolveLatestWorkOrderIdForAsset(env, cur.asset_id);
    if (linked) workOrderId = linked;
  }

  const curStage = normalizeAbuseCaseStage(String(cur.stage));
  const locChanged =
    patch.vehicleLocation !== undefined &&
    patch.vehicleLocation.trim() !== (cur.vehicle_location ?? "").trim();
  const newLoc = patch.vehicleLocation !== undefined ? patch.vehicleLocation.trim() : cur.vehicle_location;
  const stageChanged = patch.stage !== undefined && stage !== curStage;
  const respChanged =
    patch.responsibleParty !== undefined &&
    (patch.responsibleParty ?? "").trim() !== (cur.responsible_party ?? "").trim();
  const woChanged = patch.workOrderId !== undefined && workOrderId !== (cur.work_order_id ?? "").trim();
  const trackChanged = patch.trackingActive !== undefined && trackingActive !== (cur.tracking_active ?? 1);
  const reimbChanged =
    patch.reimbursedToVm !== undefined ||
    patch.reimbursedAtIso !== undefined ||
    patch.reimbursedNote !== undefined;

  const curPkg = parsePackageChecklistJson(cur.package_checklist_json);
  let packageJson = cur.package_checklist_json ?? "{}";
  if (patch.packageChecklist !== undefined) {
    packageJson = JSON.stringify({
      sf91: !!patch.packageChecklist.sf91,
      photos: !!patch.packageChecklist.photos,
      vehicleAtVmCompound: !!patch.packageChecklist.vehicleAtVmCompound,
    });
  }
  const estimatesRunner =
    patch.estimatesRunner !== undefined ? patch.estimatesRunner.trim() : cur.estimates_runner ?? "";
  const estimatesPlanned =
    patch.estimatesDowntownPlannedDate !== undefined
      ? patch.estimatesDowntownPlannedDate.trim()
      : cur.estimates_downtown_planned_date ?? "";

  const pkgChanged = patch.packageChecklist !== undefined && packageJson !== (cur.package_checklist_json ?? "{}");
  const runnerChanged =
    patch.estimatesRunner !== undefined && estimatesRunner !== (cur.estimates_runner ?? "").trim();
  const planChanged =
    patch.estimatesDowntownPlannedDate !== undefined &&
    estimatesPlanned !== (cur.estimates_downtown_planned_date ?? "").trim();

  await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE abuse_tracker_case SET
       work_order_id = ?, determination = ?, responsible_party = ?, reimbursed_to_vm = ?, reimbursed_at_iso = ?,
       reimbursed_note = ?, stage = ?, vehicle_location = ?, estimates_json = ?,
       updated_at_iso = ?, closed_at_iso = ?, tracking_active = ?,
       package_checklist_json = ?, estimates_runner = ?, estimates_downtown_planned_date = ?
     WHERE id = ?`,
  )
    .bind(
      workOrderId,
      determination,
      responsibleParty,
      reimbursedToVm,
      reimbursedAtIso,
      reimbursedNote,
      stage,
      vehicleLocation,
      estimatesJson,
      now,
      closedAt,
      trackingActive,
      packageJson,
      estimatesRunner,
      estimatesPlanned,
      caseId,
    )
    .run();

  if (stageChanged) {
    await insertAbuseTimeline(env, {
      caseId,
      kind: "stage",
      payload: { from: curStage, to: stage },
      createdBy: tlAuthor,
      atIso: now,
    });
  }
  if (locChanged && newLoc.trim()) {
    await insertAbuseTimeline(env, {
      caseId,
      kind: "location",
      payload: { from: (cur.vehicle_location ?? "").trim(), to: newLoc.trim() },
      createdBy: tlAuthor,
      atIso: now,
    });
  }
  if (respChanged) {
    await insertAbuseTimeline(env, {
      caseId,
      kind: "responsible_unit",
      payload: { from: (cur.responsible_party ?? "").trim(), to: responsibleParty },
      createdBy: tlAuthor,
      atIso: now,
    });
  }
  if (woChanged) {
    await insertAbuseTimeline(env, {
      caseId,
      kind: "work_order",
      payload: { from: (cur.work_order_id ?? "").trim(), to: workOrderId },
      createdBy: tlAuthor,
      atIso: now,
    });
  }
  if (trackChanged) {
    await insertAbuseTimeline(env, {
      caseId,
      kind: "tracking_mode",
      payload: { active: !!trackingActive },
      createdBy: tlAuthor,
      atIso: now,
    });
  }
  if (reimbChanged && (reimbursedToVm !== cur.reimbursed_to_vm || reimbursedNote !== cur.reimbursed_note || reimbursedAtIso !== cur.reimbursed_at_iso)) {
    await insertAbuseTimeline(env, {
      caseId,
      kind: "reimbursement",
      payload: {
        to_vm: !!reimbursedToVm,
        at: reimbursedAtIso,
        note: reimbursedNote,
      },
      createdBy: tlAuthor,
      atIso: now,
    });
  }
  if (patch.closed === true && !cur.closed_at_iso) {
    await insertAbuseTimeline(env, {
      caseId,
      kind: "closed",
      payload: {},
      createdBy: tlAuthor,
      atIso: now,
    });
  }
  if (pkgChanged) {
    await insertAbuseTimeline(env, {
      caseId,
      kind: "other",
      payload: { detail: "package_checklist", from: curPkg, to: parsePackageChecklistJson(packageJson) },
      createdBy: tlAuthor,
      atIso: now,
    });
  }
  if (runnerChanged || planChanged) {
    await insertAbuseTimeline(env, {
      caseId,
      kind: "other",
      payload: {
        detail: "estimates_handoff",
        runner_from: (cur.estimates_runner ?? "").trim(),
        runner_to: estimatesRunner,
        planned_from: (cur.estimates_downtown_planned_date ?? "").trim(),
        planned_to: estimatesPlanned,
      },
      createdBy: tlAuthor,
      atIso: now,
    });
  }

  if (locChanged && newLoc.trim() && !closedAt && trackingActive) {
    try {
      await recordCheck(env, {
        assetId: cur.asset_id,
        location: newLoc.trim(),
        discrepancies: "A/A tracker (manager location update)",
        status: "present",
        checkedBy: tlAuthor || "A/A program",
      });
    } catch (e) {
      console.warn(
        JSON.stringify({
          level: "warn",
          message: "abuse case vehicle_location yard_check sync failed",
          caseId,
          assetId: cur.asset_id,
          err: e instanceof Error ? e.message : String(e),
        }),
      );
    }
  }

  return env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_case WHERE id = ?`).bind(caseId).first<AbuseCaseRow>();
}

export async function addAbuseNote(env: Env, caseId: number, body: string, author: string): Promise<AbuseNoteRow | null> {
  const b = body.trim();
  if (!b) return null;
  const now = new Date().toISOString();
  const ins = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO abuse_tracker_note (case_id, body, author, at_iso) VALUES (?,?,?,?)`,
  )
    .bind(caseId, b, (author ?? "").trim(), now)
    .run();
  const nid = Number(ins.meta.last_row_id ?? 0);
  await env.ETIC_SNAPSHOTS.prepare(`UPDATE abuse_tracker_case SET updated_at_iso = ? WHERE id = ?`).bind(now, caseId).run();
  const noteRow = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_note WHERE id = ?`).bind(nid).first<AbuseNoteRow>();
  if (noteRow) {
    await insertAbuseTimeline(env, {
      caseId,
      kind: "note",
      payload: {
        note_id: nid,
        author: (author ?? "").trim(),
        snippet: b.slice(0, 200),
      },
      createdBy: (author ?? "").trim() || "system",
      atIso: now,
    });
  }
  return noteRow;
}

export async function addAbuseAttachment(
  env: Env,
  opts: {
    caseId: number;
    kind: AbuseAttachmentKind;
    body: ArrayBuffer;
    filename: string;
    contentType: string;
    uploadedBy: string;
    source: "web" | "email";
  },
): Promise<AbuseAttachmentRow> {
  const idPart = randomToken(8);
  const key = `${ABUSE_TRACKER_R2_PREFIX}${opts.caseId}/${idPart}-${safeFileName(opts.filename)}`;
  await env.ETIC_BUCKET.put(key, opts.body, {
    httpMetadata: { contentType: opts.contentType || "application/octet-stream" },
  });
  const now = new Date().toISOString();
  const ins = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO abuse_tracker_attachment
       (case_id, kind, r2_key, filename, content_type, size_bytes, uploaded_by, uploaded_at_iso, source)
     VALUES (?,?,?,?,?,?,?,?,?)`,
  )
    .bind(
      opts.caseId,
      opts.kind,
      key,
      opts.filename.slice(0, 240),
      opts.contentType || "application/octet-stream",
      opts.body.byteLength,
      (opts.uploadedBy ?? "").trim(),
      now,
      opts.source,
    )
    .run();
  const aid = Number(ins.meta.last_row_id ?? 0);
  await env.ETIC_SNAPSHOTS.prepare(`UPDATE abuse_tracker_case SET updated_at_iso = ? WHERE id = ?`).bind(now, opts.caseId).run();
  const row = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_attachment WHERE id = ?`)
    .bind(aid)
    .first<AbuseAttachmentRow>();
  if (!row) throw new Error("attachment row missing");
  await insertAbuseTimeline(env, {
    caseId: opts.caseId,
    kind: "attachment",
    payload: {
      kind: opts.kind,
      filename: opts.filename,
      source: opts.source,
      attachment_id: aid,
    },
    createdBy: (opts.uploadedBy ?? "").trim() || "system",
    atIso: now,
  });
  return row;
}

export async function getAbuseTrackerStats(env: Env): Promise<{
  open: number;
  byStage: Record<string, number>;
}> {
  const openR = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT COUNT(*) AS c FROM abuse_tracker_case WHERE closed_at_iso IS NULL`,
  ).first<{ c: number }>();
  const stages = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT stage, COUNT(*) AS c FROM abuse_tracker_case WHERE closed_at_iso IS NULL GROUP BY stage`,
  ).all<{ stage: string; c: number }>();
  const byStage: Record<string, number> = {};
  for (const s of stages.results ?? []) byStage[s.stage] = s.c ?? 0;
  return {
    open: openR?.c ?? 0,
    byStage,
  };
}

export async function listOpenAbuseCasesForWorkOrder(env: Env, workOrderId: string): Promise<AbuseCaseRow[]> {
  const wid = workOrderId.trim();
  if (!wid) return [];
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM abuse_tracker_case WHERE closed_at_iso IS NULL AND work_order_id = ? ORDER BY datetime(updated_at_iso) DESC`,
  )
    .bind(wid)
    .all<AbuseCaseRow>();
  return r.results ?? [];
}

export async function listAbuseCases(
  env: Env,
  opts: { openOnly?: boolean; limit?: number },
): Promise<AbuseCaseRow[]> {
  const limit = Math.min(5000, Math.max(1, opts.limit ?? 1000));
  let sql = `SELECT * FROM abuse_tracker_case`;
  const binds: unknown[] = [];
  if (opts.openOnly) {
    sql += ` WHERE closed_at_iso IS NULL`;
  }
  sql += ` ORDER BY datetime(created_at_iso) DESC LIMIT ?`;
  binds.push(limit);
  const r = await env.ETIC_SNAPSHOTS.prepare(sql)
    .bind(...binds)
    .all<AbuseCaseRow>();
  return r.results ?? [];
}

export async function getAbuseCaseDetail(env: Env, caseId: number): Promise<{
  case: AbuseCaseRow;
  notes: AbuseNoteRow[];
  attachments: AbuseAttachmentRow[];
  timeline: AbuseTimelineRow[];
} | null> {
  const c = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_case WHERE id = ?`).bind(caseId).first<AbuseCaseRow>();
  if (!c) return null;
  const notes = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM abuse_tracker_note WHERE case_id = ? ORDER BY id ASC`,
  )
    .bind(caseId)
    .all<AbuseNoteRow>();
  const atts = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM abuse_tracker_attachment WHERE case_id = ? ORDER BY id DESC`,
  )
    .bind(caseId)
    .all<AbuseAttachmentRow>();
  const timeline = await listAbuseTimeline(env, caseId);
  return { case: c, notes: notes.results ?? [], attachments: atts.results ?? [], timeline };
}

export async function getAbuseAttachmentMeta(
  env: Env,
  attachmentId: number,
): Promise<AbuseAttachmentRow | null> {
  return env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_attachment WHERE id = ?`)
    .bind(attachmentId)
    .first<AbuseAttachmentRow>();
}

/** Typeahead for new abuse cases — reads `fleet_asset_current` (Fleet P&A ingest). */
export async function searchFleetAssetsForPicker(
  env: Env,
  query: string,
  limit: number,
): Promise<Array<{ asset_id: string; owning_unit: string; shop: string; make_model: string }>> {
  const raw = query.trim();
  if (!raw) return [];
  const safe = raw.replace(/[%_]/g, "").slice(0, 48);
  if (!safe) return [];
  const lim = Math.min(50, Math.max(1, limit || 20));
  const like = `%${safe}%`;
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT asset_id, owning_unit, shop, make_model FROM fleet_asset_current
     WHERE asset_id LIKE ?
        OR owning_unit LIKE ?
        OR shop LIKE ?
        OR make_model LIKE ?
     ORDER BY asset_id ASC LIMIT ?`,
  )
    .bind(like, like, like, like, lim)
    .all<{ asset_id: string; owning_unit: string; shop: string; make_model: string }>();
  return r.results ?? [];
}

/** Legacy local-part: filename extension chooses damage vs release vs other. */
export function abuseDamEmailLocalPart(token: string): string {
  return `abuse-dam-${token.trim().toLowerCase()}`;
}

const ABUSE_INGEST_PREFIX_TO_KIND: Record<string, "auto" | AbuseAttachmentKind> = {
  dam: "auto",
  photo: "damage_photo",
  rel: "release_letter",
  est: "estimate",
  doc: "other",
};

/** Typed ingest addresses (one per attachment kind). Corporate browsers often block uploads. */
export function abuseIngestEmailLocalPart(kind: AbuseAttachmentKind, token: string): string {
  const t = token.trim().toLowerCase();
  if (kind === "damage_photo") return `abuse-photo-${t}`;
  if (kind === "release_letter") return `abuse-rel-${t}`;
  if (kind === "estimate") return `abuse-est-${t}`;
  return `abuse-doc-${t}`;
}

export function abuseIngestEmailAddresses(token: string): {
  auto: string;
  damage_photo: string;
  release_letter: string;
  estimate: string;
  other: string;
} {
  const t = token.trim().toLowerCase();
  return {
    auto: abuseDamEmailLocalPart(t),
    damage_photo: abuseIngestEmailLocalPart("damage_photo", t),
    release_letter: abuseIngestEmailLocalPart("release_letter", t),
    estimate: abuseIngestEmailLocalPart("estimate", t),
    other: abuseIngestEmailLocalPart("other", t),
  };
}

/** Parse To (or Delivered-To style) for abuse-{dam|photo|rel|est|doc}-<token>@host */
export function parseAbuseEmailIngestFromTo(toHeader: string): { token: string; routeKind: "auto" | AbuseAttachmentKind } | null {
  const raw = toHeader.replace(/[<>]/g, " ").trim().toLowerCase();
  const re = /\babuse-(dam|photo|rel|est|doc)-([a-f0-9]{16})@/g;
  let m: RegExpExecArray | null;
  let last: { token: string; routeKind: "auto" | AbuseAttachmentKind } | null = null;
  while ((m = re.exec(raw)) !== null) {
    const short = m[1]!;
    const token = m[2]!;
    const rk = ABUSE_INGEST_PREFIX_TO_KIND[short];
    if (rk) last = { token, routeKind: rk };
  }
  return last;
}

/** @deprecated use parseAbuseEmailIngestFromTo */
export function parseAbuseDamTokenFromEmailTo(toHeader: string): string | null {
  const p = parseAbuseEmailIngestFromTo(toHeader);
  return p && p.routeKind === "auto" ? p.token : null;
}

export type ParsedEmailAttachment = {
  filename?: string;
  mimeType?: string;
  content: ArrayBuffer;
};

/** Store non-workbook attachments on an abuse case (email ingest). */
export async function ingestAbuseDamEmailFiles(
  env: Env,
  caseRow: AbuseCaseRow,
  attachments: ParsedEmailAttachment[],
  from: string,
  routeKind: "auto" | AbuseAttachmentKind = "auto",
): Promise<number> {
  let n = 0;
  const fromTrim = (from ?? "").trim();
  for (const att of attachments) {
    const fn = (att.filename ?? "").trim().toLowerCase();
    if (fn.endsWith(".xlsx")) continue;
    const bytes = att.content.byteLength;
    if (bytes <= 0 || bytes > 25 * 1024 * 1024) continue;
    const lower = fn;
    let kind: AbuseAttachmentKind;
    if (routeKind !== "auto") {
      kind = routeKind;
    } else {
      kind = "other";
      if (lower.endsWith(".pdf")) kind = "release_letter";
      else if (/\.(jpe?g|png|gif|webp|heic|heif)$/i.test(lower)) kind = "damage_photo";
    }
    const base = (att.filename || "attachment").replace(/^.*[/\\]/, "") || "attachment";
    let outName = base;
    if (routeKind !== "auto") {
      const ext = fn.match(/(\.[^.]+)$/)?.[1] ?? "";
      const stem = base.replace(/\.[^.]+$/, "");
      const routeLabel =
        routeKind === "damage_photo"
          ? "email-damage"
          : routeKind === "release_letter"
            ? "email-release"
            : routeKind === "estimate"
              ? "email-estimate"
              : "email-doc";
      outName = stem + "__" + routeLabel + ext;
    }
    await addAbuseAttachment(env, {
      caseId: caseRow.id,
      kind,
      body: att.content,
      filename: outName.slice(0, 240),
      contentType: att.mimeType || "application/octet-stream",
      uploadedBy: fromTrim || "email",
      source: "email",
    });
    n += 1;
  }
  return n;
}

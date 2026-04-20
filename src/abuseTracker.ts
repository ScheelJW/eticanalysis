/**
 * Accident / abuse cost-recovery program tracking (VFM/VMS).
 * D1: abuse_tracker_case, abuse_tracker_note, abuse_tracker_attachment, abuse_tracker_seq.
 * R2 keys: abuse-tracker/<caseId>/<uuid>-<safeName>
 */

type Env = { ETIC_SNAPSHOTS: D1Database; ETIC_BUCKET: R2Bucket };

export const ABUSE_TRACKER_R2_PREFIX = "abuse-tracker/";

export type AbuseCaseType = "accident" | "abuse";
export type AbuseCaseStage = "intake" | "estimates" | "release_pending" | "approved_work" | "closed";
export type AbuseAttachmentKind = "damage_photo" | "release_letter" | "estimate" | "other";

export type AbuseEstimate = {
  vendor: string;
  amount: number | null;
  note: string;
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

export async function allocateControlNumber(env: Env): Promise<string> {
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
  return `AA-${y}-${String(n).padStart(5, "0")}`;
}

export async function findCaseByEmailToken(env: Env, token: string): Promise<AbuseCaseRow | null> {
  const t = token.trim().toLowerCase();
  if (!t) return null;
  return env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_case WHERE lower(email_token) = ?`)
    .bind(t)
    .first<AbuseCaseRow>();
}

async function hydrateFleetFields(env: Env, assetId: string): Promise<{
  owning_unit: string;
  shop: string;
  make_model: string;
  veh_nomen: string;
  mgmt_cd: string;
}> {
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
  if (woTrim && !woRow) throw new Error("Work order not found in fleet data (check WO id).");
  let assetId = (input.assetId ?? "").trim();
  if (!assetId && woRow) assetId = (woRow.asset_id ?? "").trim();
  if (!assetId) throw new Error("assetId or workOrderId required");
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

  const control = await allocateControlNumber(env);
  const token = randomToken(16);
  const now = new Date().toISOString();
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
       stage, vehicle_location, estimates_json, email_token, created_at_iso, updated_at_iso, created_by
     ) VALUES (?,?,?,?,?,?,?,?,?,?,?,0,NULL,'','intake',?,?,?,?,?,?)`,
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
    )
    .run();
  const id = Number(ins.meta.last_row_id ?? 0);
  if (!id) throw new Error("failed to create case");
  const row = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_case WHERE id = ?`)
    .bind(id)
    .first<AbuseCaseRow>();
  if (!row) throw new Error("case not found after insert");
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
};

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
  const stage = patch.stage !== undefined ? patch.stage : cur.stage;
  const vehicleLocation =
    patch.vehicleLocation !== undefined ? patch.vehicleLocation.trim() : cur.vehicle_location;
  const workOrderId =
    patch.workOrderId !== undefined ? patch.workOrderId.trim() : cur.work_order_id ?? "";
  let closedAt = cur.closed_at_iso;
  if (patch.closed === true) closedAt = now;
  if (patch.closed === false) closedAt = null;

  await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE abuse_tracker_case SET
       work_order_id = ?, determination = ?, responsible_party = ?, reimbursed_to_vm = ?, reimbursed_at_iso = ?,
       reimbursed_note = ?, stage = ?, vehicle_location = ?, estimates_json = ?,
       updated_at_iso = ?, closed_at_iso = ?
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
      caseId,
    )
    .run();

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
  return env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_note WHERE id = ?`).bind(nid).first<AbuseNoteRow>();
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
  return row;
}

export async function getAbuseTrackerStats(env: Env): Promise<{
  open: number;
  closed: number;
  byStage: Record<string, number>;
}> {
  const openR = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT COUNT(*) AS c FROM abuse_tracker_case WHERE closed_at_iso IS NULL`,
  ).first<{ c: number }>();
  const closedR = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT COUNT(*) AS c FROM abuse_tracker_case WHERE closed_at_iso IS NOT NULL`,
  ).first<{ c: number }>();
  const stages = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT stage, COUNT(*) AS c FROM abuse_tracker_case WHERE closed_at_iso IS NULL GROUP BY stage`,
  ).all<{ stage: string; c: number }>();
  const byStage: Record<string, number> = {};
  for (const s of stages.results ?? []) byStage[s.stage] = s.c ?? 0;
  return {
    open: openR?.c ?? 0,
    closed: closedR?.c ?? 0,
    byStage,
  };
}

export async function listAbuseCases(
  env: Env,
  opts: { openOnly?: boolean; limit?: number },
): Promise<AbuseCaseRow[]> {
  const limit = Math.min(500, Math.max(1, opts.limit ?? 200));
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
  return { case: c, notes: notes.results ?? [], attachments: atts.results ?? [] };
}

export async function getAbuseAttachmentMeta(
  env: Env,
  attachmentId: number,
): Promise<AbuseAttachmentRow | null> {
  return env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM abuse_tracker_attachment WHERE id = ?`)
    .bind(attachmentId)
    .first<AbuseAttachmentRow>();
}

export function abuseDamEmailLocalPart(token: string): string {
  return `abuse-dam-${token.trim().toLowerCase()}`;
}

/** Parse To header for abuse-dam-<token>@host */
export function parseAbuseDamTokenFromEmailTo(toHeader: string): string | null {
  const raw = toHeader.replace(/[<>]/g, "").trim().toLowerCase();
  const m = raw.match(/abuse-dam-([a-f0-9]{16})@/);
  return m ? m[1]! : null;
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
): Promise<number> {
  let n = 0;
  const fromTrim = (from ?? "").trim();
  for (const att of attachments) {
    const fn = (att.filename ?? "").trim().toLowerCase();
    if (fn.endsWith(".xlsx")) continue;
    const bytes = att.content.byteLength;
    if (bytes <= 0 || bytes > 25 * 1024 * 1024) continue;
    const lower = fn;
    let kind: AbuseAttachmentKind = "other";
    if (lower.endsWith(".pdf")) kind = "release_letter";
    else if (/\.(jpe?g|png|gif|webp|heic|heif)$/i.test(lower)) kind = "damage_photo";
    await addAbuseAttachment(env, {
      caseId: caseRow.id,
      kind,
      body: att.content,
      filename: att.filename || "attachment",
      contentType: att.mimeType || "application/octet-stream",
      uploadedBy: fromTrim || "email",
      source: "email",
    });
    n += 1;
  }
  return n;
}

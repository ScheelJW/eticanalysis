/**
 * ETIC live-meeting persistence.
 *
 * A meeting captures:
 *   - meta (attendees, title, target duration, snapshot date, filter summary)
 *   - a seeded list of work-order rows (from the watch list at meeting-start time)
 *   - per-WO status + notes + due-outs captured live
 *   - an optional final markdown rollup stored on end
 *
 * D1 tables: `meeting`, `meeting_wo_note` (see migration 0005).
 */

export type MeetingStatus = "active" | "ended";
export type MeetingNoteStatus = "pending" | "covered" | "skipped" | "deferred";

export type MeetingRow = {
  id: number;
  started_at_iso: string;
  ended_at_iso: string | null;
  title: string;
  attendees: string;
  filter_summary: string;
  target_minutes: number;
  snapshot_date_key: string;
  status: MeetingStatus;
  notes_md: string;
  /** Work order currently being shown on the conference-room screen. */
  current_wid: string;
  /** ISO timestamp the cursor was last moved (used for poll change-detection). */
  cursor_updated_at: string;
  /** Normalized scroll fraction (0..1) of the recent-changes timeline.
   *  Lets the controller scroll the TV's timeline section remotely. */
  timeline_scroll: number;
};

export type MeetingNoteRow = {
  id: number;
  meeting_id: number;
  work_order_id: string;
  asset_id: string;
  owning_unit: string;
  mel_key: string;
  mel_tier: string;
  shop: string;
  mgmt_cd: string;
  make_model: string;
  veh_nomen: string;
  etic_date: string | null;
  status: MeetingNoteStatus;
  notes: string;
  due_outs: string;
  sort_order: number;
  updated_at_iso: string;
};

export type SeedWorkOrder = {
  workOrderId: string;
  assetId?: string;
  owningUnit?: string;
  melKey?: string;
  melTier?: string;
  shop?: string;
  mgmtCd?: string;
  makeModel?: string;
  vehNomen?: string;
  eticDate?: string | null;
};

export type CreateMeetingInput = {
  title?: string;
  attendees?: string;
  filterSummary?: string;
  targetMinutes?: number;
  snapshotDateKey?: string;
  workOrders: SeedWorkOrder[];
};

export type UpsertNoteInput = {
  status?: MeetingNoteStatus;
  notes?: string;
  dueOuts?: string;
};

type Env = { ETIC_SNAPSHOTS: D1Database };

export async function createMeeting(env: Env, input: CreateMeetingInput): Promise<MeetingRow> {
  const nowIso = new Date().toISOString();
  const ins = await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO meeting (started_at_iso, title, attendees, filter_summary, target_minutes, snapshot_date_key, status)
     VALUES (?, ?, ?, ?, ?, ?, 'active')`,
  )
    .bind(
      nowIso,
      (input.title ?? "").trim(),
      (input.attendees ?? "").trim(),
      (input.filterSummary ?? "").trim(),
      Math.max(1, Math.min(600, Math.floor(input.targetMinutes ?? 30))),
      (input.snapshotDateKey ?? "").trim(),
    )
    .run();
  const id = Number(ins.meta.last_row_id ?? 0);
  if (!id) throw new Error("Failed to create meeting");

  const seed = input.workOrders ?? [];
  if (seed.length > 0) {
    const insertNote = env.ETIC_SNAPSHOTS.prepare(
      `INSERT OR IGNORE INTO meeting_wo_note
         (meeting_id, work_order_id, asset_id, owning_unit, mel_key, mel_tier, shop, mgmt_cd, make_model, veh_nomen, etic_date, status, notes, due_outs, sort_order, updated_at_iso)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', '', '', ?, ?)`,
    );
    const batch: D1PreparedStatement[] = [];
    let i = 0;
    for (const wo of seed) {
      const wid = (wo.workOrderId ?? "").trim();
      if (!wid) continue;
      batch.push(
        insertNote.bind(
          id,
          wid,
          (wo.assetId ?? "").trim(),
          (wo.owningUnit ?? "").trim(),
          (wo.melKey ?? "").trim(),
          (wo.melTier ?? "").trim(),
          (wo.shop ?? "").trim(),
          (wo.mgmtCd ?? "").trim(),
          (wo.makeModel ?? "").trim(),
          (wo.vehNomen ?? "").trim(),
          wo.eticDate ?? null,
          i,
          nowIso,
        ),
      );
      i += 1;
    }
    const CHUNK = 50;
    for (let k = 0; k < batch.length; k += CHUNK) {
      await env.ETIC_SNAPSHOTS.batch(batch.slice(k, k + CHUNK));
    }
  }

  const row = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM meeting WHERE id = ?`)
    .bind(id)
    .first<MeetingRow>();
  if (!row) throw new Error("Meeting created but could not be read back");
  return row;
}

export async function listMeetings(env: Env, limit = 100): Promise<MeetingRow[]> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT id, started_at_iso, ended_at_iso, title, attendees, filter_summary, target_minutes,
            snapshot_date_key, status, '' AS notes_md, current_wid, cursor_updated_at, timeline_scroll
     FROM meeting ORDER BY started_at_iso DESC LIMIT ?`,
  )
    .bind(Math.max(1, Math.min(500, limit)))
    .all<MeetingRow>();
  return r.results ?? [];
}

/**
 * Move the conference-room screen to a different work order. The presenter
 * (laptop) calls this; the big-screen view polls and follows.
 *
 * `workOrderId` is required (pass empty string to clear the cursor).
 * `timelineScroll` is optional — when present, also pushes the controller's
 * recent-changes timeline scroll position so the TV mirrors it.
 */
export async function setMeetingCursor(
  env: Env,
  meetingId: number,
  workOrderId: string,
  timelineScroll?: number | null,
): Promise<MeetingRow | null> {
  const nowIso = new Date().toISOString();
  if (typeof timelineScroll === "number" && Number.isFinite(timelineScroll)) {
    const clamped = Math.max(0, Math.min(1, timelineScroll));
    await env.ETIC_SNAPSHOTS.prepare(
      `UPDATE meeting SET current_wid = ?, cursor_updated_at = ?, timeline_scroll = ? WHERE id = ?`,
    )
      .bind((workOrderId ?? "").trim(), nowIso, clamped, meetingId)
      .run();
  } else {
    await env.ETIC_SNAPSHOTS.prepare(
      `UPDATE meeting SET current_wid = ?, cursor_updated_at = ? WHERE id = ?`,
    )
      .bind((workOrderId ?? "").trim(), nowIso, meetingId)
      .run();
  }
  return env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM meeting WHERE id = ?`)
    .bind(meetingId)
    .first<MeetingRow>();
}

/**
 * Update only the controller's recent-changes timeline scroll position.
 * (Cheap call — fires a lot while the controller is dragging the scrollbar.)
 */
export async function setMeetingTimelineScroll(
  env: Env,
  meetingId: number,
  scrollFraction: number,
): Promise<void> {
  if (!Number.isFinite(scrollFraction)) return;
  const clamped = Math.max(0, Math.min(1, scrollFraction));
  await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE meeting SET timeline_scroll = ?, cursor_updated_at = ? WHERE id = ?`,
  )
    .bind(clamped, new Date().toISOString(), meetingId)
    .run();
}

export async function getMeetingWithNotes(
  env: Env,
  id: number,
): Promise<{ meeting: MeetingRow; notes: MeetingNoteRow[] } | null> {
  const meeting = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM meeting WHERE id = ?`)
    .bind(id)
    .first<MeetingRow>();
  if (!meeting) return null;
  const notes = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM meeting_wo_note WHERE meeting_id = ? ORDER BY sort_order, work_order_id`,
  )
    .bind(id)
    .all<MeetingNoteRow>();
  return { meeting, notes: notes.results ?? [] };
}

export async function upsertMeetingNote(
  env: Env,
  meetingId: number,
  workOrderId: string,
  patch: UpsertNoteInput,
): Promise<MeetingNoteRow | null> {
  const nowIso = new Date().toISOString();
  const sets: string[] = [];
  const binds: unknown[] = [];
  if (patch.status !== undefined) {
    sets.push("status = ?");
    binds.push(patch.status);
  }
  if (patch.notes !== undefined) {
    sets.push("notes = ?");
    binds.push(patch.notes);
  }
  if (patch.dueOuts !== undefined) {
    sets.push("due_outs = ?");
    binds.push(patch.dueOuts);
  }
  sets.push("updated_at_iso = ?");
  binds.push(nowIso);
  binds.push(meetingId, workOrderId);

  const sql = `UPDATE meeting_wo_note SET ${sets.join(", ")} WHERE meeting_id = ? AND work_order_id = ?`;
  await env.ETIC_SNAPSHOTS.prepare(sql)
    .bind(...binds)
    .run();

  return env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM meeting_wo_note WHERE meeting_id = ? AND work_order_id = ?`,
  )
    .bind(meetingId, workOrderId)
    .first<MeetingNoteRow>();
}

/**
 * Append a new WO (not seeded at start) into an active meeting. Used when the
 * facilitator discovers a late-arrival WO mid-meeting.
 */
export async function addWorkOrdersToMeeting(
  env: Env,
  meetingId: number,
  workOrders: SeedWorkOrder[],
): Promise<number> {
  if (!workOrders.length) return 0;
  const nowIso = new Date().toISOString();
  const currentMax = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT COALESCE(MAX(sort_order), -1) AS m FROM meeting_wo_note WHERE meeting_id = ?`,
  )
    .bind(meetingId)
    .first<{ m: number }>();
  let nextOrder = (currentMax?.m ?? -1) + 1;
  const insertNote = env.ETIC_SNAPSHOTS.prepare(
    `INSERT OR IGNORE INTO meeting_wo_note
       (meeting_id, work_order_id, asset_id, owning_unit, mel_key, mel_tier, shop, mgmt_cd, make_model, veh_nomen, etic_date, status, notes, due_outs, sort_order, updated_at_iso)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', '', '', ?, ?)`,
  );
  const batch: D1PreparedStatement[] = [];
  for (const wo of workOrders) {
    const wid = (wo.workOrderId ?? "").trim();
    if (!wid) continue;
    batch.push(
      insertNote.bind(
        meetingId,
        wid,
        (wo.assetId ?? "").trim(),
        (wo.owningUnit ?? "").trim(),
        (wo.melKey ?? "").trim(),
        (wo.melTier ?? "").trim(),
        (wo.shop ?? "").trim(),
        (wo.mgmtCd ?? "").trim(),
        (wo.makeModel ?? "").trim(),
        (wo.vehNomen ?? "").trim(),
        wo.eticDate ?? null,
        nextOrder,
        nowIso,
      ),
    );
    nextOrder += 1;
  }
  const CHUNK = 50;
  let added = 0;
  for (let k = 0; k < batch.length; k += CHUNK) {
    const res = await env.ETIC_SNAPSHOTS.batch(batch.slice(k, k + CHUNK));
    for (const r of res) added += r.meta?.changes ?? 0;
  }
  return added;
}

export async function endMeeting(
  env: Env,
  id: number,
  notesMd: string,
): Promise<MeetingRow | null> {
  const endedAt = new Date().toISOString();
  await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE meeting SET status = 'ended', ended_at_iso = ?, notes_md = ? WHERE id = ?`,
  )
    .bind(endedAt, notesMd, id)
    .run();
  return env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM meeting WHERE id = ?`)
    .bind(id)
    .first<MeetingRow>();
}

export async function deleteMeeting(env: Env, id: number): Promise<void> {
  await env.ETIC_SNAPSHOTS.batch([
    env.ETIC_SNAPSHOTS.prepare(`DELETE FROM meeting_wo_note WHERE meeting_id = ?`).bind(id),
    env.ETIC_SNAPSHOTS.prepare(`DELETE FROM meeting WHERE id = ?`).bind(id),
  ]);
}

/** Render a meeting's minutes as markdown. Used both server- and client-side. */
export function renderMeetingMinutesMarkdown(
  meeting: MeetingRow,
  notes: MeetingNoteRow[],
): string {
  const started = meeting.started_at_iso ? new Date(meeting.started_at_iso) : null;
  const ended = meeting.ended_at_iso ? new Date(meeting.ended_at_iso) : null;
  const durationMin =
    started && ended ? Math.max(0, Math.round((ended.getTime() - started.getTime()) / 60000)) : null;

  const dateLine = started ? started.toISOString().slice(0, 10) : "";
  const startedTxt = started ? started.toLocaleString() : "";
  const endedTxt = ended ? ended.toLocaleString() : "(in progress)";

  const lines: string[] = [];
  const title = meeting.title ? meeting.title : "ETIC Meeting";
  lines.push(`# ${title} — ${dateLine}`);
  lines.push("");
  if (meeting.attendees) lines.push(`**Attendees:** ${meeting.attendees}`);
  if (meeting.filter_summary) lines.push(`**Scope:** ${meeting.filter_summary}`);
  lines.push(`**Snapshot reviewed:** ${meeting.snapshot_date_key || "—"}`);
  lines.push(
    `**Time:** ${startedTxt} → ${endedTxt}${durationMin !== null ? `  (${durationMin} min)` : ""}  ·  Target ${meeting.target_minutes} min`,
  );
  lines.push("");

  const covered = notes.filter((n) => n.status === "covered");
  const deferred = notes.filter((n) => n.status === "deferred");
  const skipped = notes.filter((n) => n.status === "skipped");
  const pending = notes.filter((n) => n.status === "pending");

  lines.push("## Summary");
  lines.push(
    `- Covered: **${covered.length}** · Deferred: **${deferred.length}** · Skipped: **${skipped.length}** · Not reached: **${pending.length}**`,
  );
  lines.push(`- Total work orders on agenda: **${notes.length}**`);
  lines.push("");

  const allDueOuts = notes.filter((n) => n.due_outs && n.due_outs.trim());
  if (allDueOuts.length > 0) {
    lines.push("## Due-outs");
    for (const n of allDueOuts) {
      const header = `**${n.work_order_id}**${n.asset_id ? ` · ${n.asset_id}` : ""}${n.shop ? ` · ${n.shop}` : ""}${n.mgmt_cd ? ` · Mgmt ${n.mgmt_cd}` : ""}${n.make_model ? ` · ${n.make_model}` : ""}`;
      lines.push(`- ${header}`);
      for (const due of n.due_outs.split(/\r?\n/)) {
        const t = due.trim();
        if (t) lines.push(`  - ${t}`);
      }
    }
    lines.push("");
  }

  const sections: Array<{ title: string; rows: MeetingNoteRow[] }> = [
    { title: "Covered", rows: covered },
    { title: "Deferred", rows: deferred },
    { title: "Skipped", rows: skipped },
    { title: "Not reached", rows: pending },
  ];

  for (const section of sections) {
    if (section.rows.length === 0) continue;
    lines.push(`## ${section.title}`);
    for (const n of section.rows) {
      const parts: string[] = [];
      if (n.asset_id) parts.push(n.asset_id);
      if (n.owning_unit) parts.push(n.owning_unit);
      if (n.shop) parts.push(n.shop);
      if (n.mel_key) parts.push(`MEL ${n.mel_key}`);
      if (n.mel_tier) parts.push(`${n.mel_tier}-MEL`);
      if (n.mgmt_cd) parts.push(`Mgmt ${n.mgmt_cd}`);
      if (n.make_model) parts.push(n.make_model);
      if (n.veh_nomen) parts.push(n.veh_nomen);
      if (n.etic_date) parts.push(`ETIC ${n.etic_date}`);
      const meta = parts.length ? ` — ${parts.join(" · ")}` : "";
      lines.push(`### ${n.work_order_id}${meta}`);
      if (n.notes && n.notes.trim()) {
        for (const noteLine of n.notes.split(/\r?\n/)) {
          const t = noteLine.trim();
          if (t) lines.push(`- ${t}`);
        }
      }
      if (n.due_outs && n.due_outs.trim()) {
        lines.push("- **Due-outs:**");
        for (const due of n.due_outs.split(/\r?\n/)) {
          const t = due.trim();
          if (t) lines.push(`  - ${t}`);
        }
      }
      lines.push("");
    }
  }

  return lines.join("\n");
}

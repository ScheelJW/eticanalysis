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
  /** When set, the meeting timer is frozen for everyone (presenter + controller). */
  paused_at_iso: string | null;
  /** Milliseconds of completed pause intervals (excludes the current open pause). */
  paused_accum_ms: number;
  /** Presenter TV UI scale (1.0 = 100%). Set from the live meeting controller. */
  presenter_scale?: number;
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
  /** JSON: {"n":[0,1],"d":[0,0]} = done flags per line (notes / due-outs). */
  line_completions: string;
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

function lineCount(s: string): number {
  if (!s || !s.trim()) return 0;
  return s.split(/\r?\n/).length;
}

/** @internal — exported for workOrderWatch timeline + tests */
export function buildLineCompletions(
  lineCompletionsJson: string,
  notes: string,
  dueOuts: string,
): { n: number[]; d: number[] } {
  const nLen = lineCount(notes);
  const dLen = lineCount(dueOuts);
  const n: number[] = Array.from({ length: nLen }, () => 0);
  const d: number[] = Array.from({ length: dLen }, () => 0);
  if (!lineCompletionsJson.trim()) return { n, d };
  try {
    const o = JSON.parse(lineCompletionsJson) as { n?: unknown; d?: unknown };
    if (Array.isArray(o.n)) for (let i = 0; i < nLen; i++) n[i] = o.n[i] ? 1 : 0;
    if (Array.isArray(o.d)) for (let i = 0; i < dLen; i++) d[i] = o.d[i] ? 1 : 0;
  } catch {
    /* keep zeros */
  }
  return { n, d };
}

function serializeLineCompletions(c: { n: number[]; d: number[] }): string {
  return JSON.stringify({ n: c.n, d: c.d });
}

function mergeCompletionsOnTextChange(
  prevJson: string,
  oldNotes: string,
  newNotes: string,
  oldDue: string,
  newDue: string,
): string {
  const c = buildLineCompletions(prevJson, oldNotes, oldDue);
  const nLen = lineCount(newNotes);
  const dLen = lineCount(newDue);
  const n: number[] = Array.from({ length: nLen }, () => 0);
  const d: number[] = Array.from({ length: dLen }, () => 0);
  for (let i = 0; i < nLen; i++) n[i] = i < c.n.length ? c.n[i] : 0;
  for (let i = 0; i < dLen; i++) d[i] = i < c.d.length ? c.d[i] : 0;
  return serializeLineCompletions({ n, d });
}

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
         (meeting_id, work_order_id, asset_id, owning_unit, mel_key, mel_tier, shop, mgmt_cd, make_model, veh_nomen, etic_date, status, notes, due_outs, line_completions, sort_order, updated_at_iso)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', '', '', '', ?, ?)`,
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
            snapshot_date_key, status, '' AS notes_md, current_wid, cursor_updated_at, timeline_scroll,
            paused_at_iso, paused_accum_ms, presenter_scale
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
/**
 * Pause or resume the meeting timer. Persists so the presenter view stays in sync.
 */
export async function setMeetingPauseState(
  env: Env,
  meetingId: number,
  paused: boolean,
): Promise<MeetingRow | null> {
  const row = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM meeting WHERE id = ?`)
    .bind(meetingId)
    .first<MeetingRow>();
  if (!row) return null;
  const nowMs = Date.now();
  const nowIso = new Date(nowMs).toISOString();
  let accum = Number(row.paused_accum_ms) || 0;
  const curPause = row.paused_at_iso ?? null;

  if (paused) {
    if (!curPause) {
      await env.ETIC_SNAPSHOTS.prepare(`UPDATE meeting SET paused_at_iso = ? WHERE id = ?`)
        .bind(nowIso, meetingId)
        .run();
    }
  } else {
    if (curPause) {
      const pStart = new Date(curPause).getTime();
      accum += Math.max(0, nowMs - pStart);
      await env.ETIC_SNAPSHOTS.prepare(
        `UPDATE meeting SET paused_at_iso = NULL, paused_accum_ms = ? WHERE id = ?`,
      )
        .bind(Math.round(accum), meetingId)
        .run();
    }
  }
  return env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM meeting WHERE id = ?`)
    .bind(meetingId)
    .first<MeetingRow>();
}

const PRESENTER_SCALE_MIN = 0.75;
const PRESENTER_SCALE_MAX = 1.5;

export async function setMeetingPresenterScale(
  env: Env,
  meetingId: number,
  scale: number,
): Promise<MeetingRow | null> {
  if (!Number.isFinite(scale)) return null;
  const s = Math.min(PRESENTER_SCALE_MAX, Math.max(PRESENTER_SCALE_MIN, scale));
  await env.ETIC_SNAPSHOTS.prepare(`UPDATE meeting SET presenter_scale = ? WHERE id = ?`)
    .bind(s, meetingId)
    .run();
  return env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM meeting WHERE id = ?`)
    .bind(meetingId)
    .first<MeetingRow>();
}

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

  let prev: MeetingNoteRow | null = null;
  if (patch.notes !== undefined || patch.dueOuts !== undefined) {
    prev = await env.ETIC_SNAPSHOTS.prepare(
      `SELECT notes, due_outs, line_completions FROM meeting_wo_note WHERE meeting_id = ? AND work_order_id = ?`,
    )
      .bind(meetingId, workOrderId)
      .first<MeetingNoteRow>();
  }

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
  if (prev && (patch.notes !== undefined || patch.dueOuts !== undefined)) {
    const merged = mergeCompletionsOnTextChange(
      prev.line_completions ?? "",
      prev.notes ?? "",
      patch.notes !== undefined ? patch.notes : (prev.notes ?? ""),
      prev.due_outs ?? "",
      patch.dueOuts !== undefined ? patch.dueOuts : (prev.due_outs ?? ""),
    );
    sets.push("line_completions = ?");
    binds.push(merged);
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
 * Toggle a single “done” line for meeting notes or due-outs. Persists in line_completions JSON.
 */
export async function toggleMeetingNoteLineDone(
  env: Env,
  meetingId: number,
  workOrderId: string,
  which: "notes" | "dueOuts",
  lineIndex: number,
  done: boolean,
): Promise<MeetingNoteRow | null> {
  if (!Number.isInteger(lineIndex) || lineIndex < 0) return null;
  const row = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT * FROM meeting_wo_note WHERE meeting_id = ? AND work_order_id = ?`,
  )
    .bind(meetingId, workOrderId)
    .first<MeetingNoteRow>();
  if (!row) return null;
  const notes = row.notes ?? "";
  const dueOuts = row.due_outs ?? "";
  const c = buildLineCompletions(row.line_completions ?? "", notes, dueOuts);
  if (which === "notes") {
    if (lineIndex >= c.n.length) return null;
    c.n[lineIndex] = done ? 1 : 0;
  } else {
    if (lineIndex >= c.d.length) return null;
    c.d[lineIndex] = done ? 1 : 0;
  }
  const nowIso = new Date().toISOString();
  const json = serializeLineCompletions(c);
  await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE meeting_wo_note SET line_completions = ?, updated_at_iso = ? WHERE meeting_id = ? AND work_order_id = ?`,
  )
    .bind(json, nowIso, meetingId, workOrderId)
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
       (meeting_id, work_order_id, asset_id, owning_unit, mel_key, mel_tier, shop, mgmt_cd, make_model, veh_nomen, etic_date, status, notes, due_outs, line_completions, sort_order, updated_at_iso)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', '', '', '', ?, ?)`,
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
  const row = await env.ETIC_SNAPSHOTS.prepare(`SELECT * FROM meeting WHERE id = ?`)
    .bind(id)
    .first<MeetingRow>();
  const endedAt = new Date().toISOString();
  let accum = row ? Number(row.paused_accum_ms) || 0 : 0;
  if (row?.paused_at_iso) {
    const pStart = new Date(row.paused_at_iso).getTime();
    accum += Math.max(0, Date.now() - pStart);
  }
  await env.ETIC_SNAPSHOTS.prepare(
    `UPDATE meeting SET status = 'ended', ended_at_iso = ?, notes_md = ?, paused_at_iso = NULL, paused_accum_ms = ? WHERE id = ?`,
  )
    .bind(endedAt, notesMd, Math.round(accum), id)
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

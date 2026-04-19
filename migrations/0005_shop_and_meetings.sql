-- Persist shop column for meeting-room filtering and richer watch-list filters.
ALTER TABLE work_order_state ADD COLUMN shop TEXT NOT NULL DEFAULT '';
ALTER TABLE work_order_snapshot ADD COLUMN shop TEXT NOT NULL DEFAULT '';

CREATE INDEX IF NOT EXISTS idx_wos_date_shop ON work_order_snapshot (snapshot_date_key, shop);
CREATE INDEX IF NOT EXISTS idx_wst_shop ON work_order_state (shop);

-- ETIC live-meeting tables.
CREATE TABLE IF NOT EXISTS meeting (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  started_at_iso TEXT NOT NULL,
  ended_at_iso TEXT,
  title TEXT NOT NULL DEFAULT '',
  attendees TEXT NOT NULL DEFAULT '',
  filter_summary TEXT NOT NULL DEFAULT '',
  target_minutes INTEGER NOT NULL DEFAULT 30,
  snapshot_date_key TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'active',
  notes_md TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_meeting_started ON meeting (started_at_iso DESC);
CREATE INDEX IF NOT EXISTS idx_meeting_status ON meeting (status);

CREATE TABLE IF NOT EXISTS meeting_wo_note (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  meeting_id INTEGER NOT NULL,
  work_order_id TEXT NOT NULL,
  asset_id TEXT NOT NULL DEFAULT '',
  owning_unit TEXT NOT NULL DEFAULT '',
  mel_key TEXT NOT NULL DEFAULT '',
  mel_tier TEXT NOT NULL DEFAULT '',
  shop TEXT NOT NULL DEFAULT '',
  etic_date TEXT,
  status TEXT NOT NULL DEFAULT 'pending',
  notes TEXT NOT NULL DEFAULT '',
  due_outs TEXT NOT NULL DEFAULT '',
  sort_order INTEGER NOT NULL DEFAULT 0,
  updated_at_iso TEXT NOT NULL,
  UNIQUE (meeting_id, work_order_id)
);
CREATE INDEX IF NOT EXISTS idx_meeting_wo_note_meeting ON meeting_wo_note (meeting_id);

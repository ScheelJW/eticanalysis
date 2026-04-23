-- Accident / abuse cost-recovery tracking (VFM/VMS). One open row per (asset_id, case_type).
-- Renumbered from 0026 to avoid collision with 0026_fma_meeting_line_done.sql.

CREATE TABLE IF NOT EXISTS abuse_tracker_seq (
  year INTEGER PRIMARY KEY,
  last_n INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS abuse_tracker_case (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  control_number TEXT NOT NULL UNIQUE,
  case_type TEXT NOT NULL,
  asset_id TEXT NOT NULL,
  owning_unit TEXT NOT NULL DEFAULT '',
  shop TEXT NOT NULL DEFAULT '',
  make_model TEXT NOT NULL DEFAULT '',
  veh_nomen TEXT NOT NULL DEFAULT '',
  mgmt_cd TEXT NOT NULL DEFAULT '',
  determination TEXT NOT NULL DEFAULT '',
  responsible_party TEXT NOT NULL DEFAULT '',
  reimbursed_to_vm INTEGER NOT NULL DEFAULT 0,
  reimbursed_at_iso TEXT,
  reimbursed_note TEXT NOT NULL DEFAULT '',
  stage TEXT NOT NULL DEFAULT 'intake',
  vehicle_location TEXT NOT NULL DEFAULT '',
  estimates_json TEXT NOT NULL DEFAULT '[]',
  email_token TEXT NOT NULL UNIQUE,
  created_at_iso TEXT NOT NULL,
  updated_at_iso TEXT NOT NULL,
  created_by TEXT NOT NULL DEFAULT '',
  closed_at_iso TEXT,
  CHECK (case_type IN ('accident', 'abuse')),
  CHECK (stage IN ('intake', 'estimates', 'release_pending', 'approved_work', 'closed'))
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_abuse_open_per_asset_type
  ON abuse_tracker_case (asset_id, case_type)
  WHERE closed_at_iso IS NULL;

CREATE INDEX IF NOT EXISTS idx_abuse_case_asset ON abuse_tracker_case (asset_id);
CREATE INDEX IF NOT EXISTS idx_abuse_case_stage ON abuse_tracker_case (stage);
CREATE INDEX IF NOT EXISTS idx_abuse_case_created ON abuse_tracker_case (created_at_iso DESC);

CREATE TABLE IF NOT EXISTS abuse_tracker_note (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  case_id INTEGER NOT NULL REFERENCES abuse_tracker_case(id) ON DELETE CASCADE,
  body TEXT NOT NULL,
  author TEXT NOT NULL DEFAULT '',
  at_iso TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_abuse_note_case ON abuse_tracker_note (case_id, id DESC);

CREATE TABLE IF NOT EXISTS abuse_tracker_attachment (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  case_id INTEGER NOT NULL REFERENCES abuse_tracker_case(id) ON DELETE CASCADE,
  kind TEXT NOT NULL DEFAULT 'other',
  r2_key TEXT NOT NULL,
  filename TEXT NOT NULL DEFAULT '',
  content_type TEXT NOT NULL DEFAULT 'application/octet-stream',
  size_bytes INTEGER NOT NULL DEFAULT 0,
  uploaded_by TEXT NOT NULL DEFAULT '',
  uploaded_at_iso TEXT NOT NULL,
  source TEXT NOT NULL DEFAULT 'web',
  CHECK (kind IN ('damage_photo', 'release_letter', 'estimate', 'other')),
  CHECK (source IN ('web', 'email'))
);

CREATE INDEX IF NOT EXISTS idx_abuse_att_case ON abuse_tracker_attachment (case_id, id DESC);

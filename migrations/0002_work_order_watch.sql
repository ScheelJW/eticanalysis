-- Per work order tracking: remark staleness, ETIC pushes, field changelog
CREATE TABLE IF NOT EXISTS work_order_state (
  work_order_id TEXT PRIMARY KEY NOT NULL,
  asset_id TEXT NOT NULL DEFAULT '',
  last_snapshot_date TEXT NOT NULL,
  remarks TEXT NOT NULL DEFAULT '',
  parts_status TEXT NOT NULL DEFAULT '',
  etic_raw TEXT NOT NULL DEFAULT '',
  etic_date TEXT,
  mel_tier TEXT NOT NULL DEFAULT 'unknown',
  last_remark_change_date TEXT NOT NULL,
  etic_push_count INTEGER NOT NULL DEFAULT 0,
  first_etic_date TEXT,
  last_etic_date TEXT,
  cumulative_etic_slip_days INTEGER NOT NULL DEFAULT 0,
  updated_at_iso TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_wo_state_asset ON work_order_state (asset_id);
CREATE INDEX IF NOT EXISTS idx_wo_state_mel ON work_order_state (mel_tier);
CREATE INDEX IF NOT EXISTS idx_wo_state_snapshot ON work_order_state (last_snapshot_date DESC);

CREATE TABLE IF NOT EXISTS work_order_changelog (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  work_order_id TEXT NOT NULL,
  snapshot_date_key TEXT NOT NULL,
  changed_at_iso TEXT NOT NULL,
  field TEXT NOT NULL,
  old_value TEXT,
  new_value TEXT
);

CREATE INDEX IF NOT EXISTS idx_wo_changelog_wo ON work_order_changelog (work_order_id, id DESC);

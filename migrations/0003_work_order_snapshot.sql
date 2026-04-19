-- Per work order PER snapshot-date fact table. Scales to years of daily files.
-- work_order_state is "latest known" (fast "now" view). This table is the history.
CREATE TABLE IF NOT EXISTS work_order_snapshot (
  snapshot_date_key TEXT NOT NULL,
  work_order_id TEXT NOT NULL,
  asset_id TEXT NOT NULL DEFAULT '',
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
  PRIMARY KEY (snapshot_date_key, work_order_id)
);

CREATE INDEX IF NOT EXISTS idx_wos_date ON work_order_snapshot (snapshot_date_key);
CREATE INDEX IF NOT EXISTS idx_wos_wo ON work_order_snapshot (work_order_id, snapshot_date_key DESC);
CREATE INDEX IF NOT EXISTS idx_wos_date_mel ON work_order_snapshot (snapshot_date_key, mel_tier);
CREATE INDEX IF NOT EXISTS idx_wos_date_etic ON work_order_snapshot (snapshot_date_key, etic_date);
CREATE INDEX IF NOT EXISTS idx_wos_asset_date ON work_order_snapshot (asset_id, snapshot_date_key DESC);

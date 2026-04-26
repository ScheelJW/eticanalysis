-- ELMS Schedule Mx: one row per (import date, maintenance plan row).
-- Safe if schedule_mx_extract_snapshot (0037) was never applied.
CREATE TABLE IF NOT EXISTS schedule_mx_plan_snapshot (
  snapshot_date_key TEXT NOT NULL,
  plan_row_key TEXT NOT NULL,
  asset_id TEXT NOT NULL,
  raw_row_json TEXT NOT NULL DEFAULT '{}',
  source_filename TEXT NOT NULL DEFAULT '',
  received_at_iso TEXT NOT NULL,
  PRIMARY KEY (snapshot_date_key, plan_row_key)
);
CREATE INDEX IF NOT EXISTS idx_smx_plan_date ON schedule_mx_plan_snapshot (snapshot_date_key);
CREATE INDEX IF NOT EXISTS idx_smx_plan_asset ON schedule_mx_plan_snapshot (asset_id, snapshot_date_key DESC);

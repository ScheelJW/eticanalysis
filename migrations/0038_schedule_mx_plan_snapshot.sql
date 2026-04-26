-- ELMS Schedule Mx: one row per (import date, maintenance plan row), not per asset.
-- Replaces schedule_mx_extract_snapshot (migrated rows use plan_row_key = former asset_id).
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

-- One-time migration from per-asset table (0037). Safe if empty or already migrated.
INSERT OR IGNORE INTO schedule_mx_plan_snapshot (
  snapshot_date_key, plan_row_key, asset_id, raw_row_json, source_filename, received_at_iso
)
SELECT snapshot_date_key, asset_id, asset_id, raw_row_json, source_filename, received_at_iso
FROM schedule_mx_extract_snapshot;

DROP TABLE IF EXISTS schedule_mx_extract_snapshot;

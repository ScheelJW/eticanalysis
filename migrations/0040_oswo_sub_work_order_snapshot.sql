-- Open sub work orders: rows from oswo@ email extract (one row per sub-WO or plan line).
-- Joined to parent work_order_id for WO detail. Keyed by import calendar day (see resolveAnalysisDateKey).
CREATE TABLE IF NOT EXISTS oswo_sub_work_order_snapshot (
  snapshot_date_key TEXT NOT NULL,
  row_key TEXT NOT NULL,
  parent_work_order_id TEXT NOT NULL,
  sub_work_order_id TEXT NOT NULL DEFAULT '',
  raw_row_json TEXT NOT NULL DEFAULT '{}',
  source_filename TEXT NOT NULL DEFAULT '',
  received_at_iso TEXT NOT NULL,
  PRIMARY KEY (snapshot_date_key, row_key)
);
CREATE INDEX IF NOT EXISTS idx_oswo_parent ON oswo_sub_work_order_snapshot (parent_work_order_id, snapshot_date_key DESC);
CREATE INDEX IF NOT EXISTS idx_oswo_snapshot ON oswo_sub_work_order_snapshot (snapshot_date_key);

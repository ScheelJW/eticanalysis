-- ELMS / Schedule Mx email extract (prevmx@): one row per asset per snapshot date.
-- Populated by email ingest; full history for Schedule Mx tab without ETIC workbook Fleet sheet.
CREATE TABLE IF NOT EXISTS schedule_mx_extract_snapshot (
  snapshot_date_key TEXT NOT NULL,
  asset_id TEXT NOT NULL,
  raw_row_json TEXT NOT NULL DEFAULT '{}',
  source_filename TEXT NOT NULL DEFAULT '',
  received_at_iso TEXT NOT NULL,
  PRIMARY KEY (snapshot_date_key, asset_id)
);
CREATE INDEX IF NOT EXISTS idx_schedule_mx_extract_date ON schedule_mx_extract_snapshot (snapshot_date_key);
CREATE INDEX IF NOT EXISTS idx_schedule_mx_extract_asset ON schedule_mx_extract_snapshot (asset_id, snapshot_date_key DESC);

-- Preserves superseded same-calendar-day ETIC ingests (morning vs afternoon file).
CREATE TABLE IF NOT EXISTS etic_snapshot_prior_ingest (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  report_date_key TEXT NOT NULL,
  superseded_received_at_iso TEXT NOT NULL,
  archived_at_iso TEXT NOT NULL,
  workbook_key TEXT NOT NULL,
  workbook_file_name TEXT NOT NULL,
  analysis_archive_key TEXT NOT NULL,
  mc_rate REAL,
  fleet_total INTEGER,
  fmc INTEGER,
  nmc INTEGER,
  surplus INTEGER,
  asset_manager_ok INTEGER NOT NULL DEFAULT 0,
  total_rows INTEGER,
  mel_total INTEGER,
  visible_sheets INTEGER,
  hidden_sheets INTEGER,
  asset_manager_breakdown TEXT NOT NULL DEFAULT '',
  UNIQUE(report_date_key, superseded_received_at_iso)
);

CREATE INDEX IF NOT EXISTS idx_prior_ingest_report_date
  ON etic_snapshot_prior_ingest (report_date_key DESC);

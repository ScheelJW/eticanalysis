-- Snapshot index for comparisons (KPIs + R2 keys). Applied via: wrangler d1 execute … --file=migrations/0001_etic_snapshots.sql
CREATE TABLE IF NOT EXISTS etic_snapshots (
  date_key TEXT PRIMARY KEY NOT NULL,
  workbook_key TEXT NOT NULL,
  workbook_file_name TEXT NOT NULL,
  received_at_iso TEXT NOT NULL,
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
  updated_at_iso TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_etic_snapshots_date ON etic_snapshots (date_key DESC);

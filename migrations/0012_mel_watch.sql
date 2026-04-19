-- MEL (Minimum Essential List) per-key history. The MEL Calculator sheet has
-- one row per MEL key (a unit + management code combination). We snapshot the
-- counts on every workbook ingest so we can chart "below MEL" trends, see which
-- units are recovering / regressing, and explain Recall +/- bandaid loans.

CREATE TABLE IF NOT EXISTS mel_state (
  mel_key TEXT PRIMARY KEY,
  last_snapshot_date TEXT NOT NULL,
  unit TEXT NOT NULL DEFAULT '',
  user_unit TEXT NOT NULL DEFAULT '',
  priority_tier TEXT NOT NULL DEFAULT '',
  mgmt_code_name TEXT NOT NULL DEFAULT '',
  detail_doc_number TEXT NOT NULL DEFAULT '',
  mel_assigned_total INTEGER NOT NULL DEFAULT 0,
  nmc_count INTEGER NOT NULL DEFAULT 0,
  fmc_count INTEGER NOT NULL DEFAULT 0,
  acc_abus INTEGER NOT NULL DEFAULT 0,
  mel_required INTEGER NOT NULL DEFAULT 0,
  recall_delta INTEGER NOT NULL DEFAULT 0,
  mel_delta INTEGER NOT NULL DEFAULT 0,
  mel_status TEXT NOT NULL DEFAULT '',
  raw_status TEXT NOT NULL DEFAULT '',
  updated_at_iso TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS mel_state_unit ON mel_state(unit);
CREATE INDEX IF NOT EXISTS mel_state_status ON mel_state(mel_status);
CREATE INDEX IF NOT EXISTS mel_state_tier ON mel_state(priority_tier);

CREATE TABLE IF NOT EXISTS mel_snapshot (
  snapshot_date_key TEXT NOT NULL,
  mel_key TEXT NOT NULL,
  unit TEXT NOT NULL DEFAULT '',
  user_unit TEXT NOT NULL DEFAULT '',
  priority_tier TEXT NOT NULL DEFAULT '',
  mgmt_code_name TEXT NOT NULL DEFAULT '',
  detail_doc_number TEXT NOT NULL DEFAULT '',
  mel_assigned_total INTEGER NOT NULL DEFAULT 0,
  nmc_count INTEGER NOT NULL DEFAULT 0,
  fmc_count INTEGER NOT NULL DEFAULT 0,
  acc_abus INTEGER NOT NULL DEFAULT 0,
  mel_required INTEGER NOT NULL DEFAULT 0,
  recall_delta INTEGER NOT NULL DEFAULT 0,
  mel_delta INTEGER NOT NULL DEFAULT 0,
  mel_status TEXT NOT NULL DEFAULT '',
  raw_status TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (snapshot_date_key, mel_key)
);
CREATE INDEX IF NOT EXISTS mel_snapshot_key_date ON mel_snapshot(mel_key, snapshot_date_key);
CREATE INDEX IF NOT EXISTS mel_snapshot_unit_date ON mel_snapshot(unit, snapshot_date_key);
CREATE INDEX IF NOT EXISTS mel_snapshot_date ON mel_snapshot(snapshot_date_key);

CREATE TABLE IF NOT EXISTS mel_changelog (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  mel_key TEXT NOT NULL,
  snapshot_date_key TEXT NOT NULL,
  field TEXT NOT NULL,
  old_value TEXT,
  new_value TEXT,
  created_at_iso TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS mel_changelog_key_date ON mel_changelog(mel_key, snapshot_date_key);
CREATE INDEX IF NOT EXISTS mel_changelog_field_date ON mel_changelog(field, snapshot_date_key);

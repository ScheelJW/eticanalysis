-- Editable per-MEL-key configuration: which keys are "critical / focus",
-- a friendly TYPE label override (since afa4 mgmt-code labels are cryptic),
-- a display ordering for the Critical Slide, and freeform notes.
CREATE TABLE IF NOT EXISTS mel_config (
  mel_key TEXT PRIMARY KEY,
  is_critical INTEGER NOT NULL DEFAULT 0,
  type_label TEXT NOT NULL DEFAULT '',
  display_order INTEGER NOT NULL DEFAULT 0,
  notes TEXT NOT NULL DEFAULT '',
  updated_at_iso TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS mel_config_critical ON mel_config(is_critical);
CREATE INDEX IF NOT EXISTS mel_config_order ON mel_config(display_order);

-- Optional subdivision of a single MEL key into TYPEs (e.g. 791 OSS missile
-- crew transport split into PT TRACTOR / PT TRAILER / PTR TRAILER, each with
-- their own MEL target). FMC/NMC per type are entered manually here because
-- the workbook does not break the MEL key down by type.
CREATE TABLE IF NOT EXISTS mel_subdivision (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  mel_key TEXT NOT NULL,
  type_label TEXT NOT NULL,
  mgmt_code TEXT NOT NULL DEFAULT '',
  mel_required INTEGER NOT NULL DEFAULT 0,
  assigned INTEGER NOT NULL DEFAULT 0,
  fmc INTEGER NOT NULL DEFAULT 0,
  nmc INTEGER NOT NULL DEFAULT 0,
  accidents INTEGER NOT NULL DEFAULT 0,
  abuses INTEGER NOT NULL DEFAULT 0,
  display_order INTEGER NOT NULL DEFAULT 0,
  updated_at_iso TEXT NOT NULL,
  UNIQUE(mel_key, type_label)
);
CREATE INDEX IF NOT EXISTS mel_subdivision_key ON mel_subdivision(mel_key);

-- Generic key-value store for app-wide settings (staleness thresholds, etc).
-- Values are JSON strings so callers can store whatever shape they need.
CREATE TABLE IF NOT EXISTS app_config (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL,
  updated_at_iso TEXT NOT NULL
);

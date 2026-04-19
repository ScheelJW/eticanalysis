-- Rolling yard checks: instead of explicit "sessions", every visit is just a
-- row in yard_check. The latest check per asset_id determines whether the
-- asset is "due" (no check, or last check older than the configured cadence,
-- default 7 days).
CREATE TABLE IF NOT EXISTS yard_check (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id TEXT NOT NULL,
  location TEXT NOT NULL DEFAULT '',
  discrepancies TEXT NOT NULL DEFAULT '',
  -- present | missing | unknown | not_applicable
  status TEXT NOT NULL DEFAULT 'present',
  checked_by TEXT NOT NULL DEFAULT '',
  checked_at_iso TEXT NOT NULL,
  -- Which ETIC snapshot was active when the walker recorded this check.
  -- Useful for forensic "what did they see when they tagged this".
  source_date_key TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS yard_check_asset_at
  ON yard_check(asset_id, checked_at_iso DESC);
CREATE INDEX IF NOT EXISTS yard_check_at
  ON yard_check(checked_at_iso DESC);

-- Photos taken during a check. Stored in R2 under yard-photos/<asset>/<uuid>;
-- D1 keeps the metadata for fast listing per asset and per-check linking.
CREATE TABLE IF NOT EXISTS yard_photo (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id TEXT NOT NULL,
  -- Optional: the yard_check this photo was taken alongside. Photos can also
  -- exist independently (e.g. uploaded later from the asset detail view), so
  -- we don't enforce a foreign key.
  check_id INTEGER,
  r2_key TEXT NOT NULL,
  content_type TEXT NOT NULL DEFAULT 'image/jpeg',
  size_bytes INTEGER NOT NULL DEFAULT 0,
  uploaded_by TEXT NOT NULL DEFAULT '',
  uploaded_at_iso TEXT NOT NULL,
  caption TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS yard_photo_asset_at
  ON yard_photo(asset_id, uploaded_at_iso DESC);
CREATE INDEX IF NOT EXISTS yard_photo_check
  ON yard_photo(check_id);

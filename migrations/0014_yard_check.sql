-- Yard checks: a fleet manager walks the lot with their phone, ticking off
-- where each vehicle actually is and noting any new discrepancies.
--
-- A "session" is one walk-through (e.g. "Tuesday morning yard check"). Entries
-- are upserted per (session_id, asset_id) so the same person can re-tap an
-- asset and overwrite their last note without creating duplicates.
CREATE TABLE IF NOT EXISTS yard_check_session (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL DEFAULT '',
  created_by TEXT NOT NULL DEFAULT '',
  created_at_iso TEXT NOT NULL,
  closed_at_iso TEXT,
  -- Which ETIC snapshot supplied the asset list when the session was created.
  -- Lets us re-render the same asset roster later even if a newer xlsx lands
  -- mid-session.
  source_date_key TEXT NOT NULL DEFAULT '',
  notes TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS yard_check_session_created
  ON yard_check_session(created_at_iso DESC);

CREATE TABLE IF NOT EXISTS yard_check_entry (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id INTEGER NOT NULL REFERENCES yard_check_session(id) ON DELETE CASCADE,
  asset_id TEXT NOT NULL,
  -- Where the vehicle actually is right now (free text — bays, lot letters,
  -- "in route to GP", etc).
  location TEXT NOT NULL DEFAULT '',
  -- Anything broken / missing / wrong that the walker notices on the spot.
  discrepancies TEXT NOT NULL DEFAULT '',
  -- present | missing | unknown | not_applicable
  status TEXT NOT NULL DEFAULT 'present',
  entered_by TEXT NOT NULL DEFAULT '',
  entered_at_iso TEXT NOT NULL,
  updated_at_iso TEXT NOT NULL,
  UNIQUE(session_id, asset_id)
);
CREATE INDEX IF NOT EXISTS yard_check_entry_session
  ON yard_check_entry(session_id);
CREATE INDEX IF NOT EXISTS yard_check_entry_session_asset
  ON yard_check_entry(session_id, asset_id);

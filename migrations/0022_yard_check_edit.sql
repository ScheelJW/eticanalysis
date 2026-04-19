-- Audit trail when someone corrects an existing yard_check row in place.
-- Original checked_at_iso / checked_by stay as the walker's sighting; edits
-- record who fixed the record and what changed.
CREATE TABLE IF NOT EXISTS yard_check_edit (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id TEXT NOT NULL,
  check_id INTEGER NOT NULL REFERENCES yard_check(id) ON DELETE CASCADE,
  edited_at_iso TEXT NOT NULL,
  edited_by TEXT NOT NULL DEFAULT '',
  before_json TEXT NOT NULL DEFAULT '{}',
  after_json TEXT NOT NULL DEFAULT '{}'
);
CREATE INDEX IF NOT EXISTS yard_check_edit_asset_at
  ON yard_check_edit(asset_id, edited_at_iso DESC);
CREATE INDEX IF NOT EXISTS yard_check_edit_check
  ON yard_check_edit(check_id);

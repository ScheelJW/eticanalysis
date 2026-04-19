-- FM&A (Fleet Managers & Analysis) hand-logged actions on work orders.
-- These appear in the WO change timeline alongside snapshot-derived events,
-- AND are auto-verified against the next snapshot to confirm whether the
-- expected change actually showed up in the next ETIC.
--
-- status lifecycle:
--   pending   -> action just logged, not yet verified
--   confirmed -> the expected field changed in a later snapshot
--   missed    -> at least one snapshot ingested after the action shows no
--                change to the expected field (fleet manager's edit didn't
--                make it into the next workbook)
CREATE TABLE IF NOT EXISTS work_order_action (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  work_order_id TEXT NOT NULL,
  created_at_iso TEXT NOT NULL,
  action_type TEXT NOT NULL,            -- 'remarks_update', 'etic_update', 'parts_update', 'mel_update', 'shop_update', 'other'
  expected_field TEXT NOT NULL DEFAULT '', -- 'remarks', 'etic', 'parts_status', 'mel_tier', 'shop' or '' for 'other'
  actor_name TEXT NOT NULL DEFAULT '',
  note TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'pending',
  verified_at_iso TEXT,
  verified_in_snapshot TEXT,
  snapshots_checked INTEGER NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_woa_wo ON work_order_action (work_order_id, created_at_iso DESC);
CREATE INDEX IF NOT EXISTS idx_woa_pending ON work_order_action (status, work_order_id);

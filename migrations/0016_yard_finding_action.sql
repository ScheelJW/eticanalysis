-- FM&A follow-up actions on yard-check findings.
--
-- Findings themselves are computed live from yard_check + the latest snapshot
-- (no point duplicating state). When FM&A acts on one we record the action
-- here, scoped to the specific yard_check that triggered it. If a NEWER check
-- comes in for the same asset_id+kind, the finding re-opens automatically —
-- so this table is "the last thing FM&A did about this issue", not "the
-- current state".
CREATE TABLE IF NOT EXISTS yard_finding_action (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id TEXT NOT NULL,
  -- 'missing' | 'unlisted' | 'discrepancy' | 'unknown'
  kind TEXT NOT NULL,
  -- The yard_check.id that the action was taken in response to. NULL is
  -- allowed for "unlisted" findings since those aren't tied to a specific
  -- check (the asset_id alone is the trigger).
  check_id INTEGER,
  -- 'resolved' | 'in_progress' | 'dismissed' | 'wo_opened' | 'retired' | 'reassigned'
  resolution TEXT NOT NULL DEFAULT 'resolved',
  -- WO number opened in response, if any (free text).
  wo_opened TEXT NOT NULL DEFAULT '',
  note TEXT NOT NULL DEFAULT '',
  resolved_by TEXT NOT NULL DEFAULT '',
  resolved_at_iso TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS yard_fa_asset_kind
  ON yard_finding_action(asset_id, kind, resolved_at_iso DESC);
CREATE INDEX IF NOT EXISTS yard_fa_at
  ON yard_finding_action(resolved_at_iso DESC);

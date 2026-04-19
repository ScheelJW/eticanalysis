-- Dedupe + protect work_order_changelog and mel_changelog.
--
-- The original schema had no uniqueness constraint on the (entity, snapshot_date, field)
-- triple, and the ingest path did a plain INSERT. Any time a snapshot was re-ingested
-- (rebuild history, replay, or an admin re-emailed the same workbook for the same day)
-- every change row got written again, which is why the WO change timeline was showing
-- the same MEL TIER / PARTS row twice on the same date.
--
-- Two things this migration does:
--   1. Collapse existing duplicates: keep the earliest row (lowest id) per
--      (work_order_id, snapshot_date_key, field) and (mel_key, snapshot_date_key, field).
--   2. Add UNIQUE indexes so the DB physically rejects future duplicates. Combined
--      with the matching `INSERT OR IGNORE` change in the Worker code, the ingest is
--      now idempotent — re-running it is a no-op for already-recorded changes.
--
-- We intentionally key the uniqueness on (entity, date, field) — not also on
-- old_value/new_value — because our snapshots are at most daily, so a single field
-- can only have one "change vs. yesterday" per snapshot day. If the value somehow
-- flapped twice on one day, we'd want to keep the first observation, not duplicate.

-- ---------- work_order_changelog ----------

DELETE FROM work_order_changelog
WHERE id NOT IN (
  SELECT MIN(id)
    FROM work_order_changelog
   GROUP BY work_order_id, snapshot_date_key, field
);

CREATE UNIQUE INDEX IF NOT EXISTS uq_wo_changelog_entry
  ON work_order_changelog (work_order_id, snapshot_date_key, field);

-- ---------- mel_changelog ----------

DELETE FROM mel_changelog
WHERE id NOT IN (
  SELECT MIN(id)
    FROM mel_changelog
   GROUP BY mel_key, snapshot_date_key, field
);

CREATE UNIQUE INDEX IF NOT EXISTS uq_mel_changelog_entry
  ON mel_changelog (mel_key, snapshot_date_key, field);

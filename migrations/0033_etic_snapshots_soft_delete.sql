-- Soft-delete for ETIC snapshot index rows (D1). Excluded from dashboard "latest"
-- and history merge until restored or re-ingested (upsert clears deleted_at_iso).

ALTER TABLE etic_snapshots ADD COLUMN deleted_at_iso TEXT;

-- User-requested: hide 21 Apr 2026 from the active snapshot list.
UPDATE etic_snapshots
SET deleted_at_iso = '2026-04-21T00:00:00.000Z'
WHERE date_key = '2026-04-21' AND deleted_at_iso IS NULL;

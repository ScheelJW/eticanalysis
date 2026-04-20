-- One-off: remove all ETIC meeting and yard-check data (e.g. test rows).
-- Safe order: children first, then parents. Remote D1 (requires API token in env):
--   CLOUDFLARE_API_TOKEN=... npx wrangler d1 execute etic-snapshots --remote --file=scripts/clear-meetings-and-yard-checks.sql
--
-- R2 objects under yard-photos/ are not deleted here; only D1 metadata for
-- check-linked photos is removed. Orphan R2 files are harmless but can be
-- purged from the eticanalysis bucket if you want zero leftover bytes.

-- ETIC meetings
DELETE FROM meeting_wo_note;
DELETE FROM meeting;

-- Yard (FM&A actions reference check ids; clear before checks)
DELETE FROM yard_finding_action;

-- Photos recorded during a check (unlink before deleting checks)
DELETE FROM yard_photo WHERE check_id IS NOT NULL;

-- Rolling yard checks (yard_check_edit CASCADE-deletes with yard_check rows)
DELETE FROM yard_check;

-- Legacy session-based yard UI (if any rows exist)
DELETE FROM yard_check_entry;
DELETE FROM yard_check_session;

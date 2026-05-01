-- One-time / ops: remove all ETIC live-meeting sessions and per-WO notes (includes due-outs).
-- Apply: npx wrangler d1 execute etic-snapshots --remote --file=scripts/clear-etic-meetings.sql
--        (omit --remote for local D1 only)

DELETE FROM meeting_wo_note;
DELETE FROM meeting;

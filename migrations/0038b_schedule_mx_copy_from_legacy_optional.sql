-- OPTIONAL one-time copy: run ONLY if you already have schedule_mx_extract_snapshot with data
-- (from migration 0037) and have NOT yet dropped it. Example:
--   npx wrangler d1 execute etic-snapshots --remote --file=migrations/0038_schedule_mx_plan_snapshot.sql
--   npx wrangler d1 execute etic-snapshots --remote --file=migrations/0038b_schedule_mx_copy_from_legacy_optional.sql
--   npx wrangler d1 execute etic-snapshots --remote --file=migrations/0039_drop_legacy_schedule_mx_extract.sql
INSERT OR IGNORE INTO schedule_mx_plan_snapshot (
  snapshot_date_key, plan_row_key, asset_id, raw_row_json, source_filename, received_at_iso
)
SELECT snapshot_date_key, asset_id, asset_id, raw_row_json, source_filename, received_at_iso
FROM schedule_mx_extract_snapshot;

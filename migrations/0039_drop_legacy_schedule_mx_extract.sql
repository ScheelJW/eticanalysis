-- Remove legacy per-asset table from 0037 if it still exists (no error if already gone).
DROP TABLE IF EXISTS schedule_mx_extract_snapshot;

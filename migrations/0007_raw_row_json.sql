-- Cache the entire parsed work-order row (mapped + unmapped headers) as JSON
-- so future field additions become a SQL backfill instead of re-parsing the
-- 50MB+ workbooks out of R2 every time.
ALTER TABLE work_order_state ADD COLUMN raw_row_json TEXT NOT NULL DEFAULT '{}';
ALTER TABLE work_order_snapshot ADD COLUMN raw_row_json TEXT NOT NULL DEFAULT '{}';

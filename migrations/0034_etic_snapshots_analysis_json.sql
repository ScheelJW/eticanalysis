-- Full parsed workbook analysis (same shape as R2 analyses/YYYY-MM-DD.json).
-- D1 is authoritative for API reads; R2 can still exist as cold backup.
ALTER TABLE etic_snapshots ADD COLUMN analysis_json TEXT;

-- Preserve the ETIC roster context that was active when a walker recorded a
-- rolling yard check. This makes the check auditable even after later ingests
-- change the current work-order and fleet tables.
ALTER TABLE yard_check ADD COLUMN snapshot_asset_json TEXT NOT NULL DEFAULT '{}';

-- Rolling fleet roster from ETIC workbook ingest (one row per asset_id).
-- Used for lookups that do not require an open work order. Updated on each
-- snapshot ingest; abuse tracker cases snapshot these fields only at create time.

CREATE TABLE IF NOT EXISTS fleet_asset_current (
  asset_id TEXT PRIMARY KEY NOT NULL,
  owning_unit TEXT NOT NULL DEFAULT '',
  shop TEXT NOT NULL DEFAULT '',
  make_model TEXT NOT NULL DEFAULT '',
  veh_nomen TEXT NOT NULL DEFAULT '',
  mgmt_cd TEXT NOT NULL DEFAULT '',
  mel_key TEXT NOT NULL DEFAULT '',
  last_seen_snapshot_date TEXT NOT NULL DEFAULT '',
  updated_at_iso TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_fleet_asset_unit ON fleet_asset_current (owning_unit);
CREATE INDEX IF NOT EXISTS idx_fleet_asset_shop ON fleet_asset_current (shop);

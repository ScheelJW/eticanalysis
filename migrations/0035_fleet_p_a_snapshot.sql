-- Historical Fleet (P&A) sheet: one row per asset per report date. Used as the
-- fact table for fleet-wide views (rollups, schedule Mx, NCE) without re-reading R2.
-- work_order_snapshot = Work Order Inquiry; mel_snapshot = MEL Calculator; this table = Fleet P&A.
CREATE TABLE IF NOT EXISTS fleet_p_a_snapshot (
  snapshot_date_key TEXT NOT NULL,
  asset_id TEXT NOT NULL,
  owning_unit TEXT NOT NULL DEFAULT '',
  shop TEXT NOT NULL DEFAULT '',
  make_model TEXT NOT NULL DEFAULT '',
  veh_nomen TEXT NOT NULL DEFAULT '',
  mgmt_cd TEXT NOT NULL DEFAULT '',
  mel_key TEXT NOT NULL DEFAULT '',
  raw_row_json TEXT NOT NULL DEFAULT '',
  updated_at_iso TEXT NOT NULL,
  PRIMARY KEY (snapshot_date_key, asset_id)
);
CREATE INDEX IF NOT EXISTS idx_fleet_pa_snapshot_date ON fleet_p_a_snapshot (snapshot_date_key);
CREATE INDEX IF NOT EXISTS idx_fleet_pa_snapshot_asset ON fleet_p_a_snapshot (asset_id, snapshot_date_key DESC);

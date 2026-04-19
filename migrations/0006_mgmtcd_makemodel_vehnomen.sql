-- Add Mgmt Cd, Make/Model, Veh Nomen to the watch tables for richer filtering.
ALTER TABLE work_order_state ADD COLUMN mgmt_cd TEXT NOT NULL DEFAULT '';
ALTER TABLE work_order_state ADD COLUMN make_model TEXT NOT NULL DEFAULT '';
ALTER TABLE work_order_state ADD COLUMN veh_nomen TEXT NOT NULL DEFAULT '';

ALTER TABLE work_order_snapshot ADD COLUMN mgmt_cd TEXT NOT NULL DEFAULT '';
ALTER TABLE work_order_snapshot ADD COLUMN make_model TEXT NOT NULL DEFAULT '';
ALTER TABLE work_order_snapshot ADD COLUMN veh_nomen TEXT NOT NULL DEFAULT '';

CREATE INDEX IF NOT EXISTS idx_wos_date_mgmtcd ON work_order_snapshot (snapshot_date_key, mgmt_cd);
CREATE INDEX IF NOT EXISTS idx_wst_mgmtcd ON work_order_state (mgmt_cd);

-- Carry the same fields onto live meeting note rows so they're stable in minutes.
ALTER TABLE meeting_wo_note ADD COLUMN mgmt_cd TEXT NOT NULL DEFAULT '';
ALTER TABLE meeting_wo_note ADD COLUMN make_model TEXT NOT NULL DEFAULT '';
ALTER TABLE meeting_wo_note ADD COLUMN veh_nomen TEXT NOT NULL DEFAULT '';

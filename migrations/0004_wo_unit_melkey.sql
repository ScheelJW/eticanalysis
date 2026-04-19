-- Add owning-unit and MEL-key columns for richer filtering on the watch list.
ALTER TABLE work_order_state ADD COLUMN owning_unit TEXT NOT NULL DEFAULT '';
ALTER TABLE work_order_state ADD COLUMN mel_key TEXT NOT NULL DEFAULT '';

ALTER TABLE work_order_snapshot ADD COLUMN owning_unit TEXT NOT NULL DEFAULT '';
ALTER TABLE work_order_snapshot ADD COLUMN mel_key TEXT NOT NULL DEFAULT '';

CREATE INDEX IF NOT EXISTS idx_wos_date_unit ON work_order_snapshot (snapshot_date_key, owning_unit);
CREATE INDEX IF NOT EXISTS idx_wos_date_melkey ON work_order_snapshot (snapshot_date_key, mel_key);
CREATE INDEX IF NOT EXISTS idx_wst_unit ON work_order_state (owning_unit);
CREATE INDEX IF NOT EXISTS idx_wst_melkey ON work_order_state (mel_key);

-- Accident/abuse: expanded stages, tracking-only flag, timeline events.
-- SQLite: replace case table while preserving ids (notes/attachments keep FK targets).

PRAGMA foreign_keys = OFF;

CREATE TABLE abuse_tracker_case_new (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  control_number TEXT NOT NULL UNIQUE,
  case_type TEXT NOT NULL,
  asset_id TEXT NOT NULL,
  work_order_id TEXT NOT NULL DEFAULT '',
  owning_unit TEXT NOT NULL DEFAULT '',
  shop TEXT NOT NULL DEFAULT '',
  make_model TEXT NOT NULL DEFAULT '',
  veh_nomen TEXT NOT NULL DEFAULT '',
  mgmt_cd TEXT NOT NULL DEFAULT '',
  determination TEXT NOT NULL DEFAULT '',
  responsible_party TEXT NOT NULL DEFAULT '',
  reimbursed_to_vm INTEGER NOT NULL DEFAULT 0,
  reimbursed_at_iso TEXT,
  reimbursed_note TEXT NOT NULL DEFAULT '',
  stage TEXT NOT NULL DEFAULT 'initial',
  vehicle_location TEXT NOT NULL DEFAULT '',
  estimates_json TEXT NOT NULL DEFAULT '[]',
  email_token TEXT NOT NULL UNIQUE,
  created_at_iso TEXT NOT NULL,
  updated_at_iso TEXT NOT NULL,
  created_by TEXT NOT NULL DEFAULT '',
  closed_at_iso TEXT,
  tracking_active INTEGER NOT NULL DEFAULT 1,
  CHECK (case_type IN ('accident', 'abuse')),
  CHECK (stage IN (
    'initial',
    'awaiting_estimates',
    'pending_legal_release',
    'repair_in_mx_contract',
    'repair_downtown',
    'repair_on_base',
    'no_repair_tracking',
    'closed'
  )),
  CHECK (tracking_active IN (0, 1))
);

INSERT INTO abuse_tracker_case_new (
  id, control_number, case_type, asset_id, work_order_id, owning_unit, shop, make_model, veh_nomen, mgmt_cd,
  determination, responsible_party, reimbursed_to_vm, reimbursed_at_iso, reimbursed_note,
  stage, vehicle_location, estimates_json, email_token, created_at_iso, updated_at_iso, created_by, closed_at_iso, tracking_active
)
SELECT
  id, control_number, case_type, asset_id, work_order_id, owning_unit, shop, make_model, veh_nomen, mgmt_cd,
  determination, responsible_party, reimbursed_to_vm, reimbursed_at_iso, reimbursed_note,
  CASE stage
    WHEN 'intake' THEN 'initial'
    WHEN 'estimates' THEN 'awaiting_estimates'
    WHEN 'release_pending' THEN 'pending_legal_release'
    WHEN 'approved_work' THEN 'repair_in_mx_contract'
    WHEN 'initial' THEN 'initial'
    WHEN 'awaiting_estimates' THEN 'awaiting_estimates'
    WHEN 'pending_legal_release' THEN 'pending_legal_release'
    WHEN 'repair_in_mx_contract' THEN 'repair_in_mx_contract'
    WHEN 'repair_downtown' THEN 'repair_downtown'
    WHEN 'repair_on_base' THEN 'repair_on_base'
    WHEN 'no_repair_tracking' THEN 'no_repair_tracking'
    WHEN 'closed' THEN 'closed'
    ELSE 'initial'
  END,
  vehicle_location, estimates_json, email_token, created_at_iso, updated_at_iso, created_by, closed_at_iso,
  1
FROM abuse_tracker_case;

DROP TABLE abuse_tracker_case;
ALTER TABLE abuse_tracker_case_new RENAME TO abuse_tracker_case;

CREATE UNIQUE INDEX IF NOT EXISTS idx_abuse_open_per_asset_type
  ON abuse_tracker_case (asset_id, case_type)
  WHERE closed_at_iso IS NULL;

CREATE INDEX IF NOT EXISTS idx_abuse_case_asset ON abuse_tracker_case (asset_id);
CREATE INDEX IF NOT EXISTS idx_abuse_case_stage ON abuse_tracker_case (stage);
CREATE INDEX IF NOT EXISTS idx_abuse_case_created ON abuse_tracker_case (created_at_iso DESC);
CREATE INDEX IF NOT EXISTS idx_abuse_case_wo ON abuse_tracker_case (work_order_id);

CREATE TABLE IF NOT EXISTS abuse_tracker_timeline (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  case_id INTEGER NOT NULL REFERENCES abuse_tracker_case(id) ON DELETE CASCADE,
  at_iso TEXT NOT NULL,
  kind TEXT NOT NULL,
  payload_json TEXT NOT NULL DEFAULT '{}',
  created_by TEXT NOT NULL DEFAULT '',
  CHECK (kind IN (
    'case_opened', 'stage', 'location', 'responsible_unit', 'work_order', 'tracking_mode',
    'note', 'attachment', 'reimbursement', 'closed', 'other'
  ))
);

CREATE INDEX IF NOT EXISTS idx_abuse_tl_case_at ON abuse_tracker_timeline (case_id, datetime(at_iso) ASC);

PRAGMA foreign_keys = ON;

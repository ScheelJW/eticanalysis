-- Link accident/abuse cases to an ETIC work order when known (optional).

ALTER TABLE abuse_tracker_case ADD COLUMN work_order_id TEXT NOT NULL DEFAULT '';

CREATE INDEX IF NOT EXISTS idx_abuse_case_wo ON abuse_tracker_case (work_order_id);

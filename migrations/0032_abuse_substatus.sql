-- A/A case: package checklist, estimates handoff + planned downtown date.

ALTER TABLE abuse_tracker_case ADD COLUMN package_checklist_json TEXT NOT NULL DEFAULT '{}';
ALTER TABLE abuse_tracker_case ADD COLUMN estimates_runner TEXT NOT NULL DEFAULT '';
ALTER TABLE abuse_tracker_case ADD COLUMN estimates_downtown_planned_date TEXT NOT NULL DEFAULT '';

-- Presenter / follow-along sync: the controller (laptop) sets the current
-- work order; the conference-room screen polls and renders it.
ALTER TABLE meeting ADD COLUMN current_wid TEXT NOT NULL DEFAULT '';
ALTER TABLE meeting ADD COLUMN cursor_updated_at TEXT NOT NULL DEFAULT '';

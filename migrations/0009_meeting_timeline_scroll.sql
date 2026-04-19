-- Presenter / follow-along sync: the controller can also scroll the recent-
-- changes timeline so the conference-room screen mirrors what they're reading.
-- Stored as a normalized 0..1 fraction (scrollTop / (scrollHeight - clientHeight)).
ALTER TABLE meeting ADD COLUMN timeline_scroll REAL NOT NULL DEFAULT 0;

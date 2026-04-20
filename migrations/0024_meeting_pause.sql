-- Server-synced pause so the conference-room presenter follows the controller.
ALTER TABLE meeting ADD COLUMN paused_at_iso TEXT;
ALTER TABLE meeting ADD COLUMN paused_accum_ms INTEGER NOT NULL DEFAULT 0;

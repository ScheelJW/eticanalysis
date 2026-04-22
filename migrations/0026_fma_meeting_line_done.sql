-- FM&A: optional "follow-up cleared" (separate from auto-verify status).
ALTER TABLE work_order_action ADD COLUMN followup_done_at_iso TEXT;

-- ETIC meeting: per-line completion for notes and due-out lines (JSON: {"n":[0,1],"d":[0]})
ALTER TABLE meeting_wo_note ADD COLUMN line_completions TEXT NOT NULL DEFAULT '';

-- Optional photo on each verification (mechanic annual/adhoc re-check).
ALTER TABLE waiver_verification ADD COLUMN photo_r2_key TEXT;
ALTER TABLE waiver_verification ADD COLUMN photo_content_type TEXT;

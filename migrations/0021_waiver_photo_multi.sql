-- Multiple defect photos per waiver (mechanic submissions + card display).
CREATE TABLE IF NOT EXISTS waiver_photo (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  waiver_id INTEGER NOT NULL,
  r2_key TEXT NOT NULL,
  content_type TEXT NOT NULL,
  sort_index INTEGER NOT NULL DEFAULT 0,
  created_at_iso TEXT NOT NULL,
  FOREIGN KEY (waiver_id) REFERENCES waiver(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS waiver_photo_waiver_sort
  ON waiver_photo (waiver_id, sort_index ASC, id ASC);

-- One row per legacy single-photo waiver.
INSERT INTO waiver_photo (waiver_id, r2_key, content_type, sort_index, created_at_iso)
SELECT id, photo_r2_key, COALESCE(photo_content_type, 'image/jpeg'), 0, submitted_at_iso
FROM waiver
WHERE photo_r2_key IS NOT NULL AND trim(photo_r2_key) != '';

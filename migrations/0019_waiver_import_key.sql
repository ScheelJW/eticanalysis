-- Idempotent bulk import of legacy waivers (e.g. PATS scrape) without duping rows
-- when the same import SQL is re-applied.
--
-- Normal UI-created waivers leave import_key NULL.

ALTER TABLE waiver ADD COLUMN import_key TEXT;

-- SQLite treats NULLs as distinct in UNIQUE; partial index keeps only
-- non-null import_key values unique (PATS:123, etc.).
CREATE UNIQUE INDEX IF NOT EXISTS waiver_import_key_uq
  ON waiver (import_key)
  WHERE import_key IS NOT NULL;

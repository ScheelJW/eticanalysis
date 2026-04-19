-- Per-unit (and NCE) MC-rate breakdown extracted from the Asset Manager sheet.
-- Stored as JSON: [{ label, mcRatePercent, fleetTotal, fmc, nmc, surplus, isNce }, ...]
-- Empty string means "not yet extracted" — we'll backfill from R2 on demand.
ALTER TABLE etic_snapshots ADD COLUMN asset_manager_breakdown TEXT NOT NULL DEFAULT '';

-- Soft delete: row stays for audit; photos remain in R2. Active lists and print
-- exclude rows where deleted_at_iso is set.
ALTER TABLE waiver ADD COLUMN deleted_at_iso TEXT;
ALTER TABLE waiver ADD COLUMN deleted_by TEXT;

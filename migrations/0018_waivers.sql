-- Waiver card system.
--
-- A "waiver" is a known repair item on a vehicle that management has formally
-- waived because it doesn't affect safety or serviceability. The mechanic
-- app reads the approved waivers when working on a truck so they don't chase
-- a defect that's already accepted; management approves new submissions and
-- prints a physical card that lives in the vehicle so drivers also know.
--
-- Lifecycle:
--   pending  → mechanic submits w/ photo + description
--   approved → management approves; from then on it counts as "live" on the
--              card and shows next to the asset id wherever rendered
--   rejected → management rejects with a reason; not on the card, but kept
--              for audit
--
-- Annual verification:
--   Every approved waiver must be re-verified at least once per year.
--   `last_verified_at_iso` tracks the most-recent verification (initial
--   approval also seeds one). `waiver_verification` is the immutable audit
--   log so we can show "Verified by Jane Doe on 2026-04-18" history on the
--   card and in the desktop UI.

CREATE TABLE IF NOT EXISTS waiver (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id TEXT NOT NULL,

  -- Short headline ("Front bumper crack < 6\"") + free-form details.
  title TEXT NOT NULL,
  description TEXT NOT NULL DEFAULT '',

  -- Optional inline R2 photo (one per waiver — submit captures the defect at
  -- the time the waiver was requested). Stored under waiver-photos/<id>/...
  -- See addPhoto() in src/waivers.ts.
  photo_r2_key TEXT,
  photo_content_type TEXT,

  status TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending','approved','rejected')),

  submitted_by TEXT NOT NULL,
  submitted_at_iso TEXT NOT NULL,

  -- Approval / rejection metadata. reviewed_note carries the rejection reason
  -- or the approver's optional comment.
  reviewed_by TEXT,
  reviewed_at_iso TEXT,
  reviewed_note TEXT,

  -- Convenience pointers — duplicated from waiver_verification rows so the
  -- card / list views don't need a JOIN to compute "is this overdue?".
  last_verified_by TEXT,
  last_verified_at_iso TEXT
);

CREATE INDEX IF NOT EXISTS waiver_asset_status ON waiver (asset_id, status);
CREATE INDEX IF NOT EXISTS waiver_status_submitted ON waiver (status, submitted_at_iso DESC);

CREATE TABLE IF NOT EXISTS waiver_verification (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  waiver_id INTEGER NOT NULL,
  verified_by TEXT NOT NULL,
  verified_at_iso TEXT NOT NULL,
  -- Optional context ("re-checked, still applies" / link to a WO closeout etc.)
  note TEXT,
  -- 'initial' = seeded by the approval; 'annual' = recurring re-verify; 'adhoc'
  -- = mechanic spot-check before annual was due. Lets reports tell apart the
  -- automatic verification from the manual ones.
  kind TEXT NOT NULL DEFAULT 'annual' CHECK (kind IN ('initial','annual','adhoc')),
  FOREIGN KEY (waiver_id) REFERENCES waiver(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS waiver_verification_wid
  ON waiver_verification (waiver_id, verified_at_iso DESC);

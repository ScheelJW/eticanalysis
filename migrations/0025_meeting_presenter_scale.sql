-- Controller-adjustable zoom for the conference-room presenter view (1.0 = 100%).
ALTER TABLE meeting ADD COLUMN presenter_scale REAL NOT NULL DEFAULT 1.0;

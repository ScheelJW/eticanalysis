# Waiver scrape & export brief

> **Hand this to the scraping agent.** Your job is to extract the existing
> waivers from the legacy system (paper binder, SharePoint list, Access DB,
> spreadsheet, web app — whatever it is) and produce a clean import bundle.
> The Vehicle ETIC dashboard agent will consume that bundle and load the
> waivers into the new D1-backed waiver card system.

You do **not** need to know the new system's internals. You only need to
produce the output described in §2.

---

## 1. Background (so you know what "good" looks like)

A "waiver" is a defect on a base vehicle that management has formally
accepted because it does not affect safety or serviceability. Each waiver
is tied to one vehicle (`asset_id`) and has at minimum:

- a short title (what the defect is)
- a longer description (where on the vehicle, why it's accepted)
- a defect photo
- who originally requested it + when
- who approved it + when
- the most-recent re-verification date (waivers must be re-verified
  annually) and who re-verified it
- a status (we only care about waivers that are currently **active** —
  i.e. on the card)

The existing system may not track all of these cleanly. Capture what's
there; leave what isn't blank. Don't invent data. See §5 for handling
gaps.

---

## 2. Required output — the import bundle

Produce a **single zip** named `waivers-import.zip` containing:

```
waivers-import.zip
├── waivers.csv            (UTF-8, with BOM, comma-delimited, RFC 4180-quoted)
└── photos/                (one image per waiver, optional but strongly preferred)
    ├── M-1234-001.jpg
    ├── M-1234-001.heic
    └── ...
```

That's it. No nested folders inside `photos/`. No xlsx — CSV is far easier
to validate and re-run.

### 2.1 `waivers.csv` schema

One row per waiver. Header row required, **column names exactly as below**
(snake_case, lowercase). Order doesn't matter; extra columns are ignored.

| Column | Required? | Type | Notes |
|---|---|---|---|
| `asset_id` | **yes** | string | The vehicle id as it appears on the ETIC workbook (e.g. `M-1234`, `B-009`). Trim whitespace, uppercase. If the legacy system uses a different id, map it back to the ETIC asset id. |
| `title` | **yes** | string | Short defect headline, ≤120 chars. Plain text. e.g. `"Driver-side mirror crack < 2\""`. |
| `description` | no | string | Longer details — where it is, why it's accepted. Multi-line OK; quote per RFC 4180. |
| `status` | **yes** | enum | One of `approved` / `pending` / `rejected`. **Almost every legacy row should be `approved`** — these are waivers that are currently live on the truck. Only export `pending`/`rejected` if the legacy system has a clear queue concept and you want that history preserved. |
| `submitted_by` | **yes** | string | Original submitter's name. If unknown, use `"legacy import"`. |
| `submitted_at_iso` | no | ISO 8601 | When the waiver was first requested. ISO 8601 (`2024-08-15` or `2024-08-15T00:00:00Z`). If unknown, leave blank — importer will default to today. |
| `reviewed_by` | yes if `status=approved` or `rejected` | string | Approver/rejecter name. If unknown but the waiver is clearly approved, use `"legacy import"`. |
| `reviewed_at_iso` | yes if `status=approved` or `rejected` | ISO 8601 | When approved/rejected. **Critical for `approved` rows** — this anchors the annual re-verification clock if no `last_verified_at_iso` is set. If unknown, use the submission date or today. |
| `reviewed_note` | no | string | Approval comment or rejection reason. |
| `last_verified_by` | no | string | Most-recent re-verification name, if the legacy system tracked it. |
| `last_verified_at_iso` | no | ISO 8601 | Most-recent re-verification date. |
| `photo_filename` | no | string | Filename inside `photos/` (no path). Case-sensitive. e.g. `M-1234-001.jpg`. Leave blank if no photo. |
| `legacy_id` | no | string | Whatever id the legacy system used for this row. Stash it here for traceability — the importer will preserve it in the `reviewed_note` if present. |

### 2.2 Photos

- **One photo per waiver** is enough — the new system records a single
  defect photo per waiver row.
- Acceptable formats: `jpg`, `jpeg`, `png`, `webp`, `heic`, `heif`, `gif`.
- Max **10 MB per file** (importer rejects larger). Resize anything above
  that — long edge ≤ 2400 px is plenty.
- Filename in the CSV must match the file in `photos/` exactly (extension
  and casing).
- If a waiver has multiple photos in the legacy system, pick the clearest
  one that shows the defect. Don't try to stuff multiple files into one
  row.
- If you can't get the photos out for whatever reason, ship the CSV
  without `photo_filename` populated and note it in the handoff. Photos
  can be backfilled later through the mobile app.

---

## 3. Example `waivers.csv`

```csv
asset_id,title,description,status,submitted_by,submitted_at_iso,reviewed_by,reviewed_at_iso,reviewed_note,last_verified_by,last_verified_at_iso,photo_filename,legacy_id
M-1234,"Driver mirror crack < 2""","Hairline crack lower-right corner. Mirror still fully usable; does not impair visibility.",approved,SSgt Jane Doe,2024-08-15,MSgt John Smith,2024-08-20,Approved per shop chief,SSgt Mike Reyes,2025-09-02,M-1234-001.jpg,LW-2024-0042
B-009,"Rear bumper paint chip","Paint chipped on rear bumper, ~3in. No structural impact.",approved,legacy import,,legacy import,2023-04-10,Backfilled from paper binder,,,B-009-001.heic,
M-1199,"Glove box latch missing","Driver reports glove box won't latch shut. Contents secured with strap. No safety impact.",approved,SrA Carlos Vega,2025-01-22,MSgt John Smith,2025-01-22,,,,M-1199-001.jpg,LW-2025-0007
```

Column count = 13. Header row mandatory. Quote any field that contains
`,`, `"`, or a newline. Escape inner double-quotes by doubling them
(RFC 4180): `Driver mirror crack < 2""` represents the value
`Driver mirror crack < 2"`.

---

## 4. Validation checklist before you hand the bundle over

Run through this yourself before declaring the scrape complete:

- [ ] CSV opens cleanly in Excel **and** parses cleanly with a stdlib
      CSV parser (Python `csv`, Node `csv-parse`, etc.) — no header
      mismatches, no row-length drift.
- [ ] `asset_id` matches the ETIC workbook's id format. Sample a few
      against the dashboard at <https://minot.2t3.app> if unsure.
- [ ] `status` is exactly one of `approved` / `pending` / `rejected`
      (lowercase). Anything else gets rejected by the importer.
- [ ] All `*_at_iso` values parse as ISO 8601. `2024-08-15` is fine;
      `8/15/2024` is **not**.
- [ ] Every `photo_filename` referenced in the CSV exists in `photos/`.
      No orphan files in `photos/` (warn but okay).
- [ ] No file in `photos/` exceeds 10 MB.
- [ ] No PII or secrets beyond names + defect descriptions.
- [ ] Hand back a brief summary in your final message:
      `N waivers exported, P with photos, Q rows missing reviewer info,
      R rows defaulted to "legacy import"`.

---

## 5. Edge cases & policy

- **Active vs historical**: Only export waivers that are **currently
  live on the truck**. Do not export waivers that were closed out,
  retracted, or where the underlying defect has been repaired. We're
  populating the active card system, not building a historical archive.
- **Same defect, multiple entries**: If the legacy system has duplicate
  rows for the same defect on the same vehicle, dedupe to the most
  recent / most complete one.
- **Asset id no longer in the ETIC workbook**: Still export it. The
  importer will load the row; the dashboard already handles "unlisted"
  asset ids gracefully.
- **Unknown approver but obviously approved (paper trail says so)**:
  Use `legacy import` for `reviewed_by` and best-guess the
  `reviewed_at_iso` (year is enough — pick `YYYY-01-01` if all you have
  is the year). Note this in the handoff summary.
- **No date at all**: Leave the date column blank. The importer will
  use today's date for `reviewed_at_iso` so the annual-verify clock
  starts now. Flag these in your summary so the fleet manager can plan
  to re-verify them deliberately rather than wait a year.
- **Rejected / cancelled history**: Skip unless explicitly asked. The
  active-card-only rule wins.
- **HEIC photos from iPhones**: Keep them as `.heic`. Importer handles
  the content type. No need to re-encode to JPEG.

---

## 6. What you don't need to worry about

- Database schemas, IDs, primary keys — the importer assigns its own.
- Verification audit log — the importer seeds an `initial` verification
  from `reviewed_at_iso` (or `last_verified_at_iso` if present)
  automatically.
- Idempotency — the importer is run once against an empty waiver table.
  If we need to re-run, the dashboard agent will truncate first.
- Authentication / API tokens / network access — you produce a file,
  the dashboard agent ingests it.
- Photo R2 keys, MIME sniffing, compression — importer handles all of
  it from the raw bytes.

---

## 7. Handoff

When done, drop `waivers-import.zip` somewhere the dashboard agent can
read it (a path on the user's machine is fine — they'll point us at it).
Include in your final message:

1. Total waiver count exported.
2. How many had photos.
3. Counts of any defaulted fields (e.g. "47 rows used `legacy import`
   for `reviewed_by`").
4. Anything weird you ran into that the dashboard agent should know
   before importing (e.g. "8 waivers reference an asset id that doesn't
   appear in the current ETIC workbook").

That's it. Ship the zip and the summary, you're done.

---

## 8. NEI PATS scrape → D1 (optional, repo tooling)

If the legacy system is **NEI PATS** and you already ran
`scripts/pats-waiver-scrape/scrape.mjs` to produce `pats_waivers.csv`:

1. Apply D1 migration `migrations/0019_waiver_import_key.sql` (adds nullable
   `waiver.import_key` + unique partial index).
2. Run `node scripts/pats-waiver-scrape/pats-to-d1-sql.mjs pats_waivers.csv pats_waiver_import.sql`.
   That emits SQL with **`AF` prefixed** to each `reg_no` for `asset_id`
   (e.g. `00C00488` → `AF00C00488`), **Outstanding-only** rows, approved
   status, `import_key = PATS:<waiver_row_id>` for idempotent re-run.
3. Execute the SQL against the production D1 binding (see `wrangler.jsonc`):

   `npx wrangler d1 execute etic-snapshots --remote --file=pats_waiver_import.sql`

PATS had no defect photos in the export; `photo_r2_key` stays NULL. Rows
like "No Waivered Items" are still imported as text if they were Outstanding
in PATS — filter the CSV before conversion if you want to drop those.

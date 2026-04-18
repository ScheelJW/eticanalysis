# eticanalysis

Cloudflare Worker that turns your **daily Vehicle ETIC** email into a live dashboard.

## Flow

1. Email with `.xlsx` attachment lands at any `@2t3.app` mailbox (Email Routing catch-all).
2. Worker parses the email, stores the workbook in R2 (`eticanalysis`) under `workbooks/<date>/<file>.xlsx`.
3. Worker runs workbook analysis (visible + hidden sheets, row/cell estimates, MEL mentions).
4. Analysis is saved in R2:
   - `analyses/<date>.json`
   - `analyses/latest.json`
5. A history index is maintained at `history/index.json` with day-over-day deltas.
6. Dashboard site serves live HTML + JSON APIs at `minot.2t3.app`.

## Architecture

- **Runtime:** Cloudflare Workers (TypeScript)
- **Inbound trigger:** Email Routing → Worker `email()` handler
- **Storage:** R2 bucket binding `ETIC_BUCKET`
- **Parser:** `postal-mime` for MIME + attachments
- **Workbook analysis:** `exceljs`
- **Frontend:** Worker-served HTML dashboard + fetch-backed charts/tables

## Routes

- `/` - HTML dashboard
- `/healthz` - Service health JSON
- `/api/latest` - Latest analysis JSON
- `/api/history` - Historical entries with per-day deltas
- `/api/analysis/<YYYY-MM-DD>` - Specific day's analysis

## R2 layout

```
workbooks/<YYYY-MM-DD>/<file>.xlsx   -- raw received workbooks
analyses/<YYYY-MM-DD>.json           -- analysis per ingest day
analyses/latest.json                 -- most recent analysis
history/index.json                   -- rolling index with day-over-day deltas
```

## Configuration (`wrangler.jsonc -> vars`)

- `EXPECTED_ATTACHMENT_NAME` - Preferred attachment filename (default: `Vehicle ETIC.xlsx`).
- `ALLOWED_SENDERS` - `*` allows any sender, otherwise comma-separated allowlist.
- `MAX_ATTACHMENT_BYTES` - Safety limit; anything larger is rejected.

## Cloudflare wiring

- Worker script: `etic-email-automation`
- Custom domain: `minot.2t3.app` (managed via Worker `routes` in `wrangler.jsonc`)
- R2 bucket: `eticanalysis`
- Email Routing catch-all on `2t3.app` → this Worker

## Development

```bash
npm install
npm run check     # types + tests
npm run dev       # local wrangler dev
npm run deploy    # production deploy
```

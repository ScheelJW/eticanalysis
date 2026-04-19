# eticanalysis

Internal tooling for the **Minot Vehicle ETIC** workflow. A daily Excel workbook
arrives by email; this Cloudflare Worker ingests it, persists structured rows
to D1 + R2, and serves a dashboard at `minot.2t3.app` covering:

- **Snapshot home** — fleet KPIs, day-over-day deltas, NCE / unit breakdowns.
- **Work Orders** — every open WO with ETIC date, push/slip tracking, remarks
  staleness, MEL tier, and inline change timeline.
- **MEL** — per-key fleet readiness against the Master Equipment List, with
  per-MEL "Open work orders" drill-down and configurable critical thresholds.
- **ETIC Meeting** — live presenter + controller views for the daily standup.
  TV (`/?present=<id>`) auto-syncs with whatever the controller has selected.
- **Yard Check** — mobile walker UI at `/yard` for tagging where each asset
  actually is. Drives the FM&A "things to fix" follow-up queue.
- **Ask AI** — chat over the current snapshot context.

> **AI agents and contributors: read [`AGENTS.md`](./AGENTS.md) before touching code.**
> The dashboard's HTML/CSS/JS lives inside one ~14k-line TypeScript template
> literal and has a few sharp edges (function hoisting, comment backticks,
> escapeHtml scope) that will bite you if you don't know they're there.

---

## Stack

| Layer | Choice |
|---|---|
| Runtime | Cloudflare Workers (TypeScript, ES2022) |
| Inbound trigger | Cloudflare Email Routing → Worker `email()` handler |
| Object storage | R2 bucket binding `ETIC_BUCKET` (raw workbooks + JSON analyses) |
| Relational | D1 binding `ETIC_SNAPSHOTS` (work orders, MEL, yard checks, meetings) |
| MIME parsing | `postal-mime` |
| Workbook parsing | `exceljs` |
| Frontend | Single Worker-served HTML page with embedded CSS/JS (no build step) |

## Quick start

```powershell
npm install
npm run check          # types + tests
npm run dev            # http://127.0.0.1:8787 with local R2/D1 (data in .wrangler/state)
```

Deploy to Cloudflare:

```powershell
npm run deploy         # uses .env.cloudflare if present (gitignored)
```

The first deploy needs `wrangler login` once, or a `CLOUDFLARE_API_TOKEN` in
`.env.cloudflare` (template at `.env.cloudflare.example`).

## Email ingest flow

1. `.xlsx` lands at any `@2t3.app` mailbox (Email Routing catch-all).
2. Worker stores the workbook at `r2://workbooks/<YYYY-MM-DD>/<file>.xlsx`.
3. Worker analyzes the workbook and writes JSON to:
   - `r2://analyses/<YYYY-MM-DD>.json`
   - `r2://analyses/latest.json`
4. Per-WO and per-MEL rows are upserted into D1 (`work_order_state`,
   `mel_state`, etc.) with day-over-day diffs feeding the change timelines.
5. `r2://history/index.json` is updated with the rolling KPI series.

## R2 layout

```
workbooks/<YYYY-MM-DD>/<file>.xlsx   raw received workbooks
analyses/<YYYY-MM-DD>.json           per-day analysis JSON
analyses/latest.json                 most recent analysis
history/index.json                   KPI series + day-over-day deltas
```

## Configuration (`wrangler.jsonc → vars`)

| Var | Purpose |
|---|---|
| `EXPECTED_ATTACHMENT_NAME` | Preferred attachment filename, default `Vehicle ETIC.xlsx` |
| `ALLOWED_SENDERS` | `*` or comma-separated allowlist for inbound email |
| `MAX_ATTACHMENT_BYTES` | Safety cap; oversized attachments are rejected |

Bindings (also in `wrangler.jsonc`):

| Binding | Resource |
|---|---|
| `ETIC_BUCKET` | R2 bucket `eticanalysis` |
| `ETIC_SNAPSHOTS` | D1 database `etic-snapshots` |

> **Forking:** `wrangler.jsonc` carries this deployment's `account_id` and the
> D1 `database_id`. They're not secrets, but a fork must replace them with
> values from your own Cloudflare account before `wrangler deploy` will work.

## Database migrations

D1 schema lives under [`migrations/`](./migrations) and is applied with
`wrangler d1 execute`. Apply a single migration:

```powershell
npx wrangler d1 execute etic-snapshots --remote --file=migrations/0016_yard_finding_action.sql
```

Drop `--remote` for the local dev DB.

## Routes

| Path | Purpose |
|---|---|
| `/` | Desktop dashboard (HTML + embedded JS) |
| `/yard` | Mobile-first walker UI (PWA-ish, separate HTML doc) |
| `/?present=<meetingId>` | Conference-room presenter view |
| `/healthz` | Health JSON |
| `/api/*` | JSON APIs consumed by the dashboards (full list in `AGENTS.md`) |

## Repo layout

```
src/
  index.ts            # Worker entry: email() + fetch() + ALL HTML/CSS/JS
  workOrderWatch.ts   # WO ingest, change timeline, MEL-keyed helpers
  melWatch.ts         # MEL ingest, rollups, config
  meeting.ts          # ETIC meeting sessions, notes, minutes
  yardCheck.ts        # Workbook → yard-check seed data
  yardSession.ts      # Yard-check checks, findings, sightings, photos
  ai.ts               # /api/ask handler
migrations/           # D1 schema, applied in order
test/                 # Vitest unit tests
.github/workflows/    # CI
AGENTS.md             # MUST-READ for any agent / contributor
```

## License

Private / unlicensed. Not for redistribution.

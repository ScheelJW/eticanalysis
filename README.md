# eticanalysis

Cloudflare Worker automation for your **daily Vehicle ETIC** process:

1. Receive inbound email containing the ETIC workbook (`.xlsx` attachment).
2. Save the workbook to R2 bucket `eticanalysis`.
3. Analyze visible + hidden sheets (structure and MEL mentions).
4. Save machine + human-readable reports to R2.
5. Email the report back automatically.

## Architecture

- **Runtime:** Cloudflare Workers (TypeScript)
- **Inbound trigger:** Email Routing -> Worker `email()` handler
- **Storage:** R2 bucket binding `ETIC_BUCKET`
- **Parser:** `postal-mime` for inbound email + attachments
- **Workbook analysis:** `exceljs`
- **Outbound report email:** `send_email` binding `REPORT_EMAIL`

## Project layout

- `src/index.ts` - Worker handlers and workbook analysis logic.
- `wrangler.jsonc` - Worker configuration (R2 + email bindings).
- `test/index.test.ts` - Unit tests for analysis and report rendering.

## Environment variables (`wrangler.jsonc -> vars`)

- `REPORT_TO` - Destination address for reports.
- `REPORT_FROM` - Sender address (must be on your Email Routing domain).
- `EXPECTED_ATTACHMENT_NAME` - Exact file name to prefer (default: `Vehicle ETIC.xlsx`).
- `ALLOWED_SENDERS` - Optional comma-separated allowlist for inbound sender envelope addresses.
- `MAX_ATTACHMENT_BYTES` - Max attachment size in bytes (default currently set to `20000000` in config).

## One-time Cloudflare setup

1. **Email Routing**
   - Enable Email Routing for your domain in Cloudflare.
   - Create verified destination address(es).
   - Add routing rule from an inbound address (example: `vehicle-etic@yourdomain.com`) to this Worker.

2. **R2 bucket**
   - Ensure bucket `eticanalysis` exists (already done in your case).

3. **Configure worker vars**
   - Edit `wrangler.jsonc` and set:
     - `REPORT_TO`
     - `REPORT_FROM`
     - optional `ALLOWED_SENDERS`
   - Keep sensitive values in secrets if needed.

4. **Deploy**
   ```bash
   npm install
   npm run check
   npm run deploy
   ```

## Daily flow

1. Send your ETIC workbook by email to the inbound address tied to this Worker.
2. Worker stores workbook at `incoming/<timestamp>/<filename>.xlsx`.
3. Worker writes reports to:
   - `reports/<timestamp>/analysis.json`
   - `reports/<timestamp>/report.txt`
   - `reports/latest.json`
   - `reports/latest.txt`
4. Worker emails the report text back to `REPORT_TO`.

## Local development

Run:

```bash
npm run dev
```

Wrangler exposes local email test endpoint at `/cdn-cgi/handler/email` in dev mode.

## Health endpoint

`GET /healthz` returns a simple JSON response for uptime checks.

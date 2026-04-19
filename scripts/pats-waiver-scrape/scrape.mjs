/**
 * NEI PATS — export repair waiver rows for every registration number to CSV.
 *
 * Login uses the public NEI bench form (same as clicking Login from nei.net),
 * then reads waiver tables from the Waiver screen (aID=S116).
 *
 * Usage:
 *   export PATS_USER='...' PATS_PASS='...'
 *   npm run scrape -- ../output/pats_waivers.csv
 *
 * Optional:
 *   PATS_CONCURRENCY=8   (default 4)
 *   PATS_LIMIT=200       (optional: only first N registrations — for dry runs)
 *   PATS_START_URL       (default https://www.nei.net/pats/site/)
 */

import { chromium } from "playwright";
import fs from "node:fs";
import path from "node:path";
import { isNoiseWaiverText } from "./is-noise-waiver.mjs";

const BENCH_LOGIN = "https://www.nei.net/bench/index.cfm";
const PATS_ENTRY = process.env.PATS_START_URL || "https://www.nei.net/pats/site/";
const WAIVER_PAGE = "https://www.nei.net/pats/sm/index.cfm?aID=S116";

function requireEnv(name) {
  const v = process.env[name];
  if (!v) {
    console.error(`Missing env ${name} (do not pass secrets on the command line)`);
    process.exit(1);
  }
  return v;
}

function csvEscape(value) {
  const s = String(value ?? "");
  if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function parseRegOptions(html) {
  /** @type {{ id: string, regNo: string }[]} */
  const out = [];
  const re = /<option\s+value="(\d+)"\s*>([^<\r\n]+)/gi;
  let m;
  while ((m = re.exec(html))) {
    const id = m[1];
    const regNo = m[2].trim();
    if (!regNo || regNo.startsWith("==")) continue;
    out.push({ id, regNo });
  }
  return out;
}

/**
 * Pull waiver rows from the iframe HTML for stePID=4.
 * @param {string} html
 * @param {string} vehicleId
 * @param {string} regNo
 */
function parseWaiverRows(html, vehicleId, regNo) {
  if (html.includes("No Waiver Information!")) return [];

  /** @type {Record<string,string>[]} */
  const rows = [];
  const formRe = /<form[^>]*action=["']\/pats\/sm\/index\.cfm\?aID=S116["'][^>]*>([\s\S]*?)<\/form>/gi;
  let formMatch;
  while ((formMatch = formRe.exec(html))) {
    const chunk = formMatch[1];
    if (!chunk.includes('name="stepID"') || !chunk.includes('value="3"')) continue;
    const wvVin = chunk.match(/name="wvVinID"\s+value="(\d+)"/)?.[1] || vehicleId;
    const regCell = chunk.match(/<strong>Reg No\.\s*([^<]+)<\/strong>/)?.[1]?.trim() || regNo;

    const trRe = /<tr[^>]*class="(?:Even|Odd)"[^>]*>([\s\S]*?)<\/tr>/gi;
    let tr;
    while ((tr = trRe.exec(chunk))) {
      const body = tr[1];
      const inputMatch = body.match(/<input[^>]*name="(w_note_\d+)"[^>]*value="([^"]*)"/i);
      if (!inputMatch) continue;
      const waiverNote = inputMatch[2];
      if (isNoiseWaiverText(waiverNote)) continue;

      const rightTds = [...body.matchAll(/<td[^>]*style="text-align:right;"[^>]*>([\s\S]*?)<\/td>/gi)];
      const emp = rightTds[0]?.[1].replace(/<[^>]+>/g, "").trim() || "";
      const timeAdded = rightTds[1]?.[1].replace(/<[^>]+>/g, "").trim() || "";

      const statusSelect = body.match(/<select[^>]*name="wv_status_\d+"[^>]*>([\s\S]*?)<\/select>/i);
      let status = "";
      if (statusSelect) {
        const inner = statusSelect[1];
        const opt =
          inner.match(/<option[^>]*\bSelected\b[^>]*>([^<]+)/i) ||
          inner.match(/<option[^>]*\bselected\b[^>]*>([^<]+)/i);
        status = (opt?.[1] || "").trim();
      }

      const waiverId = inputMatch[1].replace("w_note_", "");

      rows.push({
        vehicle_id: wvVin,
        reg_no: regCell,
        waiver_row_id: waiverId,
        waiver_text: waiverNote,
        employee_no: emp,
        time_added: timeAdded,
        status,
      });
    }
  }
  return rows;
}

async function mapPool(items, concurrency, fn) {
  const results = new Array(items.length);
  let i = 0;
  async function worker() {
    for (;;) {
      const idx = i++;
      if (idx >= items.length) return;
      results[idx] = await fn(items[idx], idx);
    }
  }
  const workers = Array.from({ length: Math.min(concurrency, items.length) }, () => worker());
  await Promise.all(workers);
  return results;
}

async function main() {
  const outPath = path.resolve(process.argv[2] || path.join(process.cwd(), "pats_waivers.csv"));
  const user = requireEnv("PATS_USER");
  const pass = requireEnv("PATS_PASS");
  const concurrency = Math.max(1, Math.min(20, Number(process.env.PATS_CONCURRENCY || "4")));
  const limitRaw = process.env.PATS_LIMIT;
  const limit = limitRaw ? Math.max(1, Number(limitRaw)) : 0;

  const browser = await chromium.launch({
    headless: true,
    args: ["--disable-blink-features=AutomationControlled"],
  });
  const context = await browser.newContext({
    userAgent:
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    viewport: { width: 1400, height: 900 },
  });
  const page = await context.newPage();

  // The public site's login markup nests a <table> inside <form> incorrectly, so the
  // username/password fields are not actually inside the <form>. Submit via HTTP POST.
  await page.goto(PATS_ENTRY, { waitUntil: "load", timeout: 120_000 });
  const loginRes = await context.request.post(BENCH_LOGIN, {
    form: { aID: "0001", login_name: user, login_passwd: pass },
    maxRedirects: 0,
    timeout: 120_000,
  });
  if (loginRes.status() !== 302) {
    console.error(`Unexpected bench login response: HTTP ${loginRes.status()}`);
    await browser.close();
    process.exit(1);
  }

  const waiverRes = await context.request.get(WAIVER_PAGE, { timeout: 120_000 });
  const waiverHtml = await waiverRes.text();

  let vehicles = parseRegOptions(waiverHtml);
  if (limit && vehicles.length > limit) vehicles = vehicles.slice(0, limit);
  if (!vehicles.length) {
    console.error("No registration numbers found on waiver page. Is PATS_USER / PATS_PASS correct?");
    await browser.close();
    process.exit(1);
  }

  console.log(`Found ${vehicles.length} registration numbers. Fetching waiver tables (${concurrency} concurrent)…`);

  const storage = await context.storageState();

  const api = await chromium.launch();
  const apiCtx = await api.newContext({ storageState: storage });
  const request = apiCtx.request;

  const allRows = [];

  await mapPool(vehicles, concurrency, async (v, idx) => {
    const url = `https://www.nei.net/pats/sm/index.cfm?aID=S116&stePID=4&wvvinID=${encodeURIComponent(v.id)}`;
    const res = await request.get(url, { timeout: 120_000 });
    const html = await res.text();
    const rows = parseWaiverRows(html, v.id, v.regNo);
    if (rows.length) {
      for (const r of rows) allRows.push(r);
    }
    if ((idx + 1) % 50 === 0 || idx === vehicles.length - 1) {
      process.stdout.write(`\rProgress: ${idx + 1}/${vehicles.length} (${allRows.length} waiver rows)`);
    }
  });

  console.log("");

  const header = [
    "vehicle_id",
    "reg_no",
    "waiver_row_id",
    "waiver_text",
    "employee_no",
    "time_added",
    "status",
  ];
  const lines = [header.join(",")];
  for (const r of allRows) {
    lines.push(header.map((h) => csvEscape(r[h])).join(","));
  }
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, lines.join("\n") + "\n", "utf8");

  await apiCtx.close();
  await api.close();
  await browser.close();

  console.log(`Wrote ${allRows.length} waiver rows to ${outPath}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

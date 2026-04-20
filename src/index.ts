import PostalMime from "postal-mime";
import type { Attachment } from "postal-mime";
import ExcelJS from "exceljs";
import {
  extractRawWorkOrdersFromBinary,
  extractYardCheckSource,
  generateYardCheckWorkbookBuffer,
} from "./yardCheck";
import { getChangelog, getWatchRowById, getWatchRowsLatest, ingestWorkOrderSnapshot } from "./workOrderWatch";

const DEFAULT_MAX_ATTACHMENT_BYTES = 30 * 1024 * 1024;
const SITE_TITLE = "Minot Vehicle ETIC";

type SheetSummary = {
  name: string;
  state: "visible" | "hidden" | "veryHidden";
  rowCount: number;
  columnCountEstimate: number;
  sampleHeaders: string[];
  nonEmptyCellCountEstimate: number;
  melMentions: number;
};

/** KPIs from worksheet "Asset Manager" row 2: F2–J2 (MC rate, fleet total, FMC, NMC, surplus). */
type AssetManagerKpis = {
  sheetFound: boolean;
  mcRatePercent: number | null;
  fleetTotal: number | null;
  fmc: number | null;
  nmc: number | null;
  surplus: number | null;
};

type AnalysisResult = {
  workbookFileName: string;
  receivedAtIso: string;
  dateKey: string;
  from: string;
  to: string;
  subject: string;
  workbookBytes: number;
  assetManager: AssetManagerKpis;
  sheetSummaries: SheetSummary[];
  totalVisibleSheets: number;
  totalHiddenSheets: number;
  totalRowsAcrossSheets: number;
  melMentionsBySheet: Record<string, number>;
  melMentionsTotal: number;
};

type HistoryEntryDiff = {
  previousDateKey: string | null;
  deltaTotalRows: number | null;
  deltaMelMentionsTotal: number | null;
  deltaSheetsVisible: number | null;
};

type HistoryEntry = {
  dateKey: string;
  receivedAtIso: string;
  workbookFileName: string;
  workbookKey: string;
  analysisKey: string;
  totalVisibleSheets: number;
  totalHiddenSheets: number;
  totalRowsAcrossSheets: number;
  melMentionsTotal: number;
  diff: HistoryEntryDiff;
};

type HistoryIndex = {
  updatedAtIso: string;
  entries: HistoryEntry[];
};

export default {
  async email(message: ForwardableEmailMessage, env: Env, _ctx: ExecutionContext): Promise<void> {
    if (!isAuthorizedSender(message.from, env)) {
      message.setReject("Unauthorized sender");
      return;
    }

    const parsed = await PostalMime.parse(message.raw, { attachmentEncoding: "arraybuffer" });
    const workbookAttachment = pickWorkbookAttachment(parsed.attachments, env.EXPECTED_ATTACHMENT_NAME);
    if (!workbookAttachment) {
      console.warn(
        JSON.stringify({
          level: "warn",
          message: "No .xlsx attachment found on inbound email",
          from: message.from,
          to: message.to,
          subject: parsed.subject ?? "",
        }),
      );
      return;
    }

    const workbookBytes = normalizeAttachmentBinary(workbookAttachment.content);
    if (workbookBytes.byteLength > parseMaxAttachmentBytes(env.MAX_ATTACHMENT_BYTES)) {
      message.setReject("Attachment too large");
      return;
    }

    const now = new Date();
    const subject = parsed.subject ?? "";
    const dateKey = resolveAnalysisDateKey(subject, now);
    const safeName = sanitizeFileName(workbookAttachment.filename ?? "vehicle-etic.xlsx");
    const workbookKey = `workbooks/${dateKey}/${safeName}`;
    const analysisKey = `analyses/${dateKey}.json`;

    await env.ETIC_BUCKET.put(workbookKey, workbookBytes, {
      httpMetadata: {
        contentType:
          workbookAttachment.mimeType ||
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      customMetadata: {
        from: message.from,
        to: message.to,
        subject,
      },
    });

    const analysis = await analyzeWorkbook({
      binary: workbookBytes,
      fileName: safeName,
      receivedAtIso: now.toISOString(),
      dateKey,
      from: message.from,
      to: message.to,
      subject,
    });

    await env.ETIC_BUCKET.put(analysisKey, JSON.stringify(analysis, null, 2), {
      httpMetadata: { contentType: "application/json; charset=utf-8" },
    });

    await upsertSnapshotRow(env, analysis, workbookKey);

    const rawWos = await extractRawWorkOrdersFromBinary(workbookBytes);
    await ingestWorkOrderSnapshot(env, dateKey, rawWos, now.toISOString());

    const history = await loadHistory(env);
    const upsertedHistory = upsertHistoryEntry(history, analysis, workbookKey, analysisKey);
    await env.ETIC_BUCKET.put(
      "history/index.json",
      JSON.stringify(upsertedHistory, null, 2),
      {
        httpMetadata: { contentType: "application/json; charset=utf-8" },
      },
    );
    await env.ETIC_BUCKET.put(
      "analyses/latest.json",
      JSON.stringify(analysis, null, 2),
      {
        httpMetadata: { contentType: "application/json; charset=utf-8" },
      },
    );
  },

  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    const url = new URL(request.url);

    if (url.pathname === "/healthz") {
      return Response.json({ ok: true, service: "etic-email-automation" });
    }

    if (url.pathname === "/api/history") {
      const history = await loadHistory(env);
      return Response.json(history, { headers: cacheHeaders() });
    }

    if (url.pathname === "/api/snapshots") {
      const sync = url.searchParams.get("sync");
      if (request.method === "POST" && sync === "1") {
        ctx.waitUntil(backfillSnapshotsFromHistory(env));
        return Response.json(
          { ok: true, message: "Rebuilding snapshot index from R2 (runs in background)." },
          { headers: cacheHeaders() },
        );
      }
      return handleSnapshotsList(env);
    }

    if (url.pathname === "/api/latest") {
      let analysis = await readJson<AnalysisResult>(env, "analyses/latest.json");
      if (!analysis) return new Response("Not Found", { status: 404 });
      analysis = await enrichAnalysisAssetManagerFromR2(env, analysis);
      return Response.json(analysis, { headers: cacheHeaders() });
    }

    if (url.pathname.startsWith("/api/analysis/")) {
      const dateKey = url.pathname.slice("/api/analysis/".length);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) {
        return new Response("Invalid date key", { status: 400 });
      }
      let analysis = await readJson<AnalysisResult>(env, `analyses/${dateKey}.json`);
      if (!analysis) return new Response("Not Found", { status: 404 });
      analysis = await enrichAnalysisAssetManagerFromR2(env, analysis);
      return Response.json(analysis, { headers: cacheHeaders() });
    }

    if (url.pathname === "/api/workbook.xlsx") {
      return handleWorkbookDownload(env, request);
    }

    if (url.pathname === "/api/watch") {
      return handleWatchApi(env, request, ctx);
    }

    if (url.pathname === "/api/work-order-changelog") {
      return handleWorkOrderChangelogApi(env, request);
    }

    if (url.pathname === "/api/yard-check.xlsx") {
      return handleYardCheckDownload(env, request);
    }

    if (url.pathname === "/api/yard-check-meta") {
      return handleYardCheckMeta(env, request);
    }

    if (url.pathname === "/" || url.pathname === "/index.html") {
      return new Response(renderDashboardHtml(), {
        headers: {
          "Content-Type": "text/html; charset=utf-8",
          "Cache-Control": "private, no-store, no-cache, must-revalidate, max-age=0",
          Pragma: "no-cache",
          Expires: "0",
          "CDN-Cache-Control": "no-store",
          "Vary": "*",
        },
      });
    }

    return new Response("Not Found", { status: 404 });
  },
} satisfies ExportedHandler<Env>;

function parseYardCheckDateParam(request: Request): string | null {
  const url = new URL(request.url);
  const d = url.searchParams.get("date");
  return d && /^\d{4}-\d{2}-\d{2}$/.test(d) ? d : null;
}

async function resolveYardCheckWorkbook(
  env: Env,
  dateKey: string | null,
): Promise<
  | { ok: true; workbookKey: string; sourceFileName: string; sourceDateKey: string }
  | { ok: false; error: string }
> {
  const history = await loadHistory(env);
  const latest = await readJson<AnalysisResult>(env, "analyses/latest.json");

  if (dateKey) {
    const entry = history.entries.find((e) => e.dateKey === dateKey);
    if (entry) {
      return {
        ok: true,
        workbookKey: entry.workbookKey,
        sourceFileName: entry.workbookFileName,
        sourceDateKey: entry.dateKey,
      };
    }
    return {
      ok: false,
      error: `No ETIC snapshot for ${dateKey}. Pick another date or ingest that workbook.`,
    };
  }

  const latestEntry = history.entries.length
    ? [...history.entries].sort((a, b) => a.dateKey.localeCompare(b.dateKey)).pop() ?? null
    : null;
  const workbookKey = latestEntry?.workbookKey ?? null;
  const sourceFileName = latestEntry?.workbookFileName ?? latest?.workbookFileName ?? "vehicle-etic.xlsx";
  const sourceDateKey = latestEntry?.dateKey ?? latest?.dateKey ?? "latest";

  if (!workbookKey) {
    return { ok: false, error: "No source workbook found yet. Send the ETIC email to ingest first." };
  }

  return { ok: true, workbookKey, sourceFileName, sourceDateKey };
}

async function handleYardCheckDownload(env: Env, request: Request): Promise<Response> {
  const paramDate = parseYardCheckDateParam(request);
  const resolved = await resolveYardCheckWorkbook(env, paramDate);
  if (!resolved.ok) {
    return Response.json({ ok: false, error: resolved.error }, { status: 404 });
  }
  const { workbookKey, sourceFileName, sourceDateKey } = resolved;

  const object = await env.ETIC_BUCKET.get(workbookKey);
  if (!object) {
    return Response.json(
      { ok: false, error: `Source workbook not found in R2: ${workbookKey}` },
      { status: 404 },
    );
  }

  const workbookBytes = await object.arrayBuffer();
  const source = await extractYardCheckSource(workbookBytes);
  if (!source) {
    return Response.json(
      {
        ok: false,
        error: "Could not locate a Work Orders sheet in the latest workbook.",
      },
      { status: 422 },
    );
  }

  const generatedAtIso = new Date().toISOString();
  const buffer = await generateYardCheckWorkbookBuffer(source, {
    sourceWorkbookFileName: sourceFileName,
    sourceDateKey,
    generatedAtIso,
  });

  const downloadName = `yard-check_${sourceDateKey}_${generatedAtIso.replace(/[:.]/g, "-")}.xlsx`;
  return new Response(buffer, {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${downloadName}"`,
      "Cache-Control": "no-store",
      "X-Source-Workbook": workbookKey,
      "X-Source-Date": sourceDateKey,
      "X-Work-Order-Sheet": source.workOrderSheet ?? "",
      "X-Fleet-Sheet": source.fleetSheet ?? "",
      "X-Total-Assets": String(source.totalAssets),
      "X-Total-Work-Orders": String(source.totalWorkOrders),
    },
  });
}

type YardCheckMetaOk = {
  ok: true;
  sourceDateKey: string;
  sourceFileName: string;
  workbookKey: string;
};

type YardCheckMetaErr = { ok: false; error: string };

async function handleYardCheckMeta(env: Env, request: Request): Promise<Response> {
  const meta = await buildYardCheckMeta(env, request);
  const status = meta.ok ? 200 : 404;
  return Response.json(meta, { status, headers: cacheHeaders() });
}

async function buildYardCheckMeta(env: Env, request: Request): Promise<YardCheckMetaOk | YardCheckMetaErr> {
  const paramDate = parseYardCheckDateParam(request);
  const resolved = await resolveYardCheckWorkbook(env, paramDate);
  if (!resolved.ok) {
    return { ok: false, error: resolved.error };
  }
  const { workbookKey, sourceFileName, sourceDateKey } = resolved;

  const object = await env.ETIC_BUCKET.get(workbookKey);
  if (!object) {
    return { ok: false, error: `Source workbook not found in R2: ${workbookKey}` };
  }

  return {
    ok: true,
    sourceDateKey,
    sourceFileName,
    workbookKey,
  };
}

function cacheHeaders(): HeadersInit {
  return {
    "Cache-Control": "no-store",
  };
}

function assetManagerFromAnalysis(analysis: AnalysisResult): AssetManagerKpis {
  const am = analysis.assetManager;
  if (am && typeof am === "object") return am;
  return {
    sheetFound: false,
    mcRatePercent: null,
    fleetTotal: null,
    fmc: null,
    nmc: null,
    surplus: null,
  };
}

async function upsertSnapshotRow(env: Env, analysis: AnalysisResult, workbookKey: string): Promise<void> {
  const am = assetManagerFromAnalysis(analysis);
  const now = new Date().toISOString();
  await env.ETIC_SNAPSHOTS.prepare(
    `INSERT INTO etic_snapshots (
      date_key, workbook_key, workbook_file_name, received_at_iso,
      mc_rate, fleet_total, fmc, nmc, surplus, asset_manager_ok,
      total_rows, mel_total, visible_sheets, hidden_sheets, updated_at_iso
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(date_key) DO UPDATE SET
      workbook_key = excluded.workbook_key,
      workbook_file_name = excluded.workbook_file_name,
      received_at_iso = excluded.received_at_iso,
      mc_rate = excluded.mc_rate,
      fleet_total = excluded.fleet_total,
      fmc = excluded.fmc,
      nmc = excluded.nmc,
      surplus = excluded.surplus,
      asset_manager_ok = excluded.asset_manager_ok,
      total_rows = excluded.total_rows,
      mel_total = excluded.mel_total,
      visible_sheets = excluded.visible_sheets,
      hidden_sheets = excluded.hidden_sheets,
      updated_at_iso = excluded.updated_at_iso`,
  )
    .bind(
      analysis.dateKey,
      workbookKey,
      analysis.workbookFileName,
      analysis.receivedAtIso,
      am.mcRatePercent,
      am.fleetTotal,
      am.fmc,
      am.nmc,
      am.surplus,
      am.sheetFound ? 1 : 0,
      analysis.totalRowsAcrossSheets,
      analysis.melMentionsTotal,
      analysis.totalVisibleSheets,
      analysis.totalHiddenSheets,
      now,
    )
    .run();
}

type SnapshotListRow = {
  dateKey: string;
  workbookFileName: string;
  workbookKey: string;
  receivedAtIso: string;
  mcRatePercent: number | null;
  fleetTotal: number | null;
  fmc: number | null;
  nmc: number | null;
  surplus: number | null;
  assetManagerOk: boolean;
  totalRows: number | null;
  melTotal: number | null;
  visibleSheets: number | null;
  hiddenSheets: number | null;
  updatedAtIso: string;
};

async function handleSnapshotsList(env: Env): Promise<Response> {
  const result = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT date_key, workbook_key, workbook_file_name, received_at_iso,
            mc_rate, fleet_total, fmc, nmc, surplus, asset_manager_ok,
            total_rows, mel_total, visible_sheets, hidden_sheets, updated_at_iso
     FROM etic_snapshots ORDER BY date_key DESC`,
  ).all<{
    date_key: string;
    workbook_key: string;
    workbook_file_name: string;
    received_at_iso: string;
    mc_rate: number | null;
    fleet_total: number | null;
    fmc: number | null;
    nmc: number | null;
    surplus: number | null;
    asset_manager_ok: number;
    total_rows: number | null;
    mel_total: number | null;
    visible_sheets: number | null;
    hidden_sheets: number | null;
    updated_at_iso: string;
  }>();

  const rows: SnapshotListRow[] = (result.results ?? []).map((r) => ({
    dateKey: r.date_key,
    workbookKey: r.workbook_key,
    workbookFileName: r.workbook_file_name,
    receivedAtIso: r.received_at_iso,
    mcRatePercent: r.mc_rate,
    fleetTotal: r.fleet_total,
    fmc: r.fmc,
    nmc: r.nmc,
    surplus: r.surplus,
    assetManagerOk: Boolean(r.asset_manager_ok),
    totalRows: r.total_rows,
    melTotal: r.mel_total,
    visibleSheets: r.visible_sheets,
    hiddenSheets: r.hidden_sheets,
    updatedAtIso: r.updated_at_iso,
  }));

  return Response.json({ updatedAtIso: new Date().toISOString(), snapshots: rows }, { headers: cacheHeaders() });
}

async function persistAnalysisIfChanged(env: Env, analysis: AnalysisResult): Promise<void> {
  await env.ETIC_BUCKET.put(
    `analyses/${analysis.dateKey}.json`,
    JSON.stringify(analysis, null, 2),
    { httpMetadata: { contentType: "application/json; charset=utf-8" } },
  );
  const history = await loadHistory(env);
  const entry = history.entries.find((e) => e.dateKey === analysis.dateKey);
  if (entry) {
    await upsertSnapshotRow(env, analysis, entry.workbookKey);
  }
  const latest = await readJson<AnalysisResult>(env, "analyses/latest.json");
  if (latest?.dateKey === analysis.dateKey) {
    await env.ETIC_BUCKET.put("analyses/latest.json", JSON.stringify(analysis, null, 2), {
      httpMetadata: { contentType: "application/json; charset=utf-8" },
    });
  }
}

async function enrichAnalysisAssetManagerFromR2(env: Env, analysis: AnalysisResult): Promise<AnalysisResult> {
  let am = assetManagerFromAnalysis(analysis);
  if (am.sheetFound) return { ...analysis, assetManager: am };

  const history = await loadHistory(env);
  const entry = history.entries.find((e) => e.dateKey === analysis.dateKey);
  if (entry) {
    const obj = await env.ETIC_BUCKET.get(entry.workbookKey);
    if (obj) {
      const bytes = await obj.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(bytes);
      am = extractAssetManagerKpis(workbook);
    }
  }

  if (!am.sheetFound) {
    const row = await env.ETIC_SNAPSHOTS.prepare(
      `SELECT mc_rate, fleet_total, fmc, nmc, surplus, asset_manager_ok FROM etic_snapshots WHERE date_key = ?`,
    )
      .bind(analysis.dateKey)
      .first<{
        mc_rate: number | null;
        fleet_total: number | null;
        fmc: number | null;
        nmc: number | null;
        surplus: number | null;
        asset_manager_ok: number;
      }>();
    if (row && row.asset_manager_ok) {
      am = {
        sheetFound: true,
        mcRatePercent: row.mc_rate,
        fleetTotal: row.fleet_total,
        fmc: row.fmc,
        nmc: row.nmc,
        surplus: row.surplus,
      };
    }
  }

  const merged = { ...analysis, assetManager: am };
  if (am.sheetFound && JSON.stringify(analysis.assetManager) !== JSON.stringify(am)) {
    await persistAnalysisIfChanged(env, merged);
  }
  return merged;
}

async function backfillSnapshotsFromHistory(env: Env): Promise<void> {
  const history = await loadHistory(env);
  for (const entry of history.entries) {
    let analysis = await readJson<AnalysisResult>(env, entry.analysisKey);
    if (!analysis) continue;
    analysis = await enrichAnalysisAssetManagerFromR2(env, analysis);
    await upsertSnapshotRow(env, analysis, entry.workbookKey);
  }
}

/** Replay all workbooks in date order so changelog + ETIC push counts are correct. */
async function rebuildWorkOrderWatchFromHistory(env: Env): Promise<void> {
  await env.ETIC_SNAPSHOTS.prepare(`DELETE FROM work_order_changelog`).run();
  await env.ETIC_SNAPSHOTS.prepare(`DELETE FROM work_order_state`).run();
  const history = await loadHistory(env);
  const sorted = [...history.entries].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  const nowIso = new Date().toISOString();
  for (const entry of sorted) {
    const obj = await env.ETIC_BUCKET.get(entry.workbookKey);
    if (!obj) continue;
    const bytes = await obj.arrayBuffer();
    const raw = await extractRawWorkOrdersFromBinary(bytes);
    await ingestWorkOrderSnapshot(env, entry.dateKey, raw, nowIso);
  }
}

async function handleWatchApi(env: Env, request: Request, ctx: ExecutionContext): Promise<Response> {
  if (request.method === "POST") {
    const url = new URL(request.url);
    if (url.searchParams.get("rebuild") === "1") {
      ctx.waitUntil(rebuildWorkOrderWatchFromHistory(env));
      return Response.json(
        {
          ok: true,
          message:
            "Work order watch is rebuilding from R2 in chronological order (clears prior changelog). Takes a few minutes for large history.",
        },
        { headers: cacheHeaders() },
      );
    }
    return new Response("Use POST /api/watch?rebuild=1", { status: 400 });
  }

  const url = new URL(request.url);
  const woLookup = url.searchParams.get("workOrderId")?.trim();
  let asOf = url.searchParams.get("date");
  if (asOf && !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) {
    return new Response("Invalid date", { status: 400 });
  }
  if (!asOf) {
    const history = await loadHistory(env);
    const latest = history.entries.length
      ? [...history.entries].sort((a, b) => b.dateKey.localeCompare(a.dateKey))[0]
      : null;
    asOf = latest?.dateKey ?? null;
  }
  if (!asOf) {
    return Response.json(
      { asOfDateKey: null, rows: [], staleCount: 0, row: null, found: false },
      { headers: cacheHeaders() },
    );
  }

  if (woLookup) {
    const row = await getWatchRowById(env, woLookup, asOf);
    return Response.json(
      {
        asOfDateKey: asOf,
        workOrderId: woLookup,
        found: row !== null,
        row: row ?? null,
      },
      { headers: cacheHeaders() },
    );
  }

  const rows = await getWatchRowsLatest(env, asOf);
  const staleOnly = url.searchParams.get("stale") === "1";
  const filtered = staleOnly ? rows.filter((r) => r.remarkStale) : rows;
  const staleCount = rows.filter((r) => r.remarkStale).length;
  return Response.json(
    { asOfDateKey: asOf, staleCount, total: rows.length, rows: filtered },
    { headers: cacheHeaders() },
  );
}

async function handleWorkOrderChangelogApi(env: Env, request: Request): Promise<Response> {
  const url = new URL(request.url);
  const wo = url.searchParams.get("workOrderId")?.trim();
  if (!wo) {
    return Response.json({ ok: false, error: "Missing workOrderId query parameter." }, { status: 400 });
  }
  const entries = await getChangelog(env, wo, 500);
  return Response.json({ workOrderId: wo, entries }, { headers: cacheHeaders() });
}

async function handleWorkbookDownload(env: Env, request: Request): Promise<Response> {
  const paramDate = parseYardCheckDateParam(request);
  const resolved = await resolveYardCheckWorkbook(env, paramDate);
  if (!resolved.ok) {
    return Response.json({ ok: false, error: resolved.error }, { status: 404 });
  }
  const { workbookKey, sourceFileName, sourceDateKey } = resolved;
  const object = await env.ETIC_BUCKET.get(workbookKey);
  if (!object) {
    return Response.json({ ok: false, error: `Workbook not found: ${workbookKey}` }, { status: 404 });
  }
  const body = await object.arrayBuffer();
  const safe = sourceFileName.replace(/[^A-Za-z0-9._-]/g, "_") || "Vehicle_ETIC.xlsx";
  const downloadName = `etic_${sourceDateKey}_${safe}`;
  return new Response(body, {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${downloadName}"`,
      "Cache-Control": "no-store",
      "X-Source-Date": sourceDateKey,
      "X-Source-Key": workbookKey,
    },
  });
}

async function readJson<T>(env: Env, key: string): Promise<T | null> {
  const object = await env.ETIC_BUCKET.get(key);
  if (!object) return null;
  try {
    const text = await object.text();
    return JSON.parse(text) as T;
  } catch (error) {
    console.error(
      JSON.stringify({
        level: "error",
        message: "Failed to parse R2 JSON object",
        key,
        error: error instanceof Error ? error.message : String(error),
      }),
    );
    return null;
  }
}

async function loadHistory(env: Env): Promise<HistoryIndex> {
  const existing = await readJson<HistoryIndex>(env, "history/index.json");
  if (existing && Array.isArray(existing.entries)) {
    return existing;
  }
  return { updatedAtIso: new Date().toISOString(), entries: [] };
}

function upsertHistoryEntry(
  history: HistoryIndex,
  analysis: AnalysisResult,
  workbookKey: string,
  analysisKey: string,
): HistoryIndex {
  const filtered = history.entries.filter((entry) => entry.dateKey !== analysis.dateKey);
  const sortedPrevious = [...filtered].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  const previous = sortedPrevious.length > 0 ? sortedPrevious[sortedPrevious.length - 1] : null;

  const diff: HistoryEntryDiff = {
    previousDateKey: previous ? previous.dateKey : null,
    deltaTotalRows: previous ? analysis.totalRowsAcrossSheets - previous.totalRowsAcrossSheets : null,
    deltaMelMentionsTotal: previous
      ? analysis.melMentionsTotal - previous.melMentionsTotal
      : null,
    deltaSheetsVisible: previous
      ? analysis.totalVisibleSheets - previous.totalVisibleSheets
      : null,
  };

  const entry: HistoryEntry = {
    dateKey: analysis.dateKey,
    receivedAtIso: analysis.receivedAtIso,
    workbookFileName: analysis.workbookFileName,
    workbookKey,
    analysisKey,
    totalVisibleSheets: analysis.totalVisibleSheets,
    totalHiddenSheets: analysis.totalHiddenSheets,
    totalRowsAcrossSheets: analysis.totalRowsAcrossSheets,
    melMentionsTotal: analysis.melMentionsTotal,
    diff,
  };

  const merged = [...filtered, entry].sort((a, b) => a.dateKey.localeCompare(b.dateKey));
  return {
    updatedAtIso: new Date().toISOString(),
    entries: merged,
  };
}

function isAuthorizedSender(sender: string, env: Env): boolean {
  const configured = env.ALLOWED_SENDERS?.trim();
  if (!configured) return true;
  if (configured === "*") return true;
  const normalizedSender = normalizeEmailAddress(sender);
  const allowed = configured
    .split(",")
    .map((part) => normalizeEmailAddress(part))
    .filter(Boolean);
  return allowed.includes(normalizedSender);
}

function pickWorkbookAttachment(
  attachments: Attachment[],
  expectedAttachmentName?: string,
): Attachment | undefined {
  const expected = expectedAttachmentName?.trim().toLowerCase();
  if (expected) {
    const exact = attachments.find(
      (attachment) => (attachment.filename?.trim().toLowerCase() ?? "") === expected,
    );
    if (exact) return exact;
  }

  return attachments.find((attachment) => {
    const filename = attachment.filename?.toLowerCase() ?? "";
    return filename.endsWith(".xlsx");
  });
}

function normalizeAttachmentBinary(content: Attachment["content"]): ArrayBuffer {
  if (content instanceof ArrayBuffer) return content;
  if (content instanceof Uint8Array) return cloneToArrayBuffer(content);
  const encoded = new TextEncoder().encode(content);
  return cloneToArrayBuffer(encoded);
}

function cloneToArrayBuffer(view: ArrayBufferView): ArrayBuffer {
  const clone = new Uint8Array(view.byteLength);
  clone.set(new Uint8Array(view.buffer, view.byteOffset, view.byteLength));
  return clone.buffer;
}

function sanitizeFileName(fileName: string): string {
  return fileName.replace(/[^A-Za-z0-9._-]/g, "_");
}

function normalizeEmailAddress(email: string): string {
  return email.replace(/[<>]/g, "").trim().toLowerCase();
}

function parseMaxAttachmentBytes(raw: string | undefined): number {
  if (!raw) return DEFAULT_MAX_ATTACHMENT_BYTES;
  const value = Number.parseInt(raw, 10);
  return Number.isFinite(value) && value > 0 ? value : DEFAULT_MAX_ATTACHMENT_BYTES;
}

function isoDateKey(date: Date): string {
  const p = (n: number) => String(n).padStart(2, "0");
  return `${date.getUTCFullYear()}-${p(date.getUTCMonth() + 1)}-${p(date.getUTCDate())}`;
}

const MONTH_TOKEN: Record<string, number> = {
  jan: 1,
  feb: 2,
  mar: 3,
  apr: 4,
  may: 5,
  jun: 6,
  jul: 7,
  aug: 8,
  sep: 9,
  oct: 10,
  nov: 11,
  dec: 12,
};

/** Last DD-MMM-YY (or DD-MMM-YYYY) token in the subject → YYYY-MM-DD for R2 paths. */
function parseReportDateKeyFromSubject(subject: string): string | null {
  const s = subject.trim();
  let reportDateKey: string | null = null;
  const dateRe = /\b(\d{1,2})-([A-Za-z]{3})\.?-(\d{2,4})\b/g;
  let dm: RegExpExecArray | null;
  while ((dm = dateRe.exec(s)) !== null) {
    const day = Number.parseInt(dm[1] ?? "", 10);
    const monToken = (dm[2] ?? "").toLowerCase().replace(/\.$/, "").slice(0, 3);
    const month = MONTH_TOKEN[monToken];
    if (!month || !Number.isFinite(day) || day < 1 || day > 31) continue;
    const yRaw = dm[3] ?? "";
    const yNum = Number.parseInt(yRaw, 10);
    if (!Number.isFinite(yNum)) continue;
    const year = yRaw.length <= 2 ? (yNum >= 70 ? 1900 + yNum : 2000 + yNum) : yNum;
    reportDateKey = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }
  return reportDateKey;
}

/** Prefer report date from subject (forwarded ETICs); else UTC day of receipt. */
function resolveAnalysisDateKey(subject: string, receivedAt: Date): string {
  return parseReportDateKeyFromSubject(subject) ?? isoDateKey(receivedAt);
}

async function analyzeWorkbook(input: {
  binary: ArrayBuffer;
  fileName: string;
  receivedAtIso: string;
  dateKey: string;
  from: string;
  to: string;
  subject: string;
}): Promise<AnalysisResult> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(input.binary);

  const sheetSummaries = workbook.worksheets.map((sheet): SheetSummary => {
    const state = normalizeSheetState(sheet.state);
    const rowCount = estimateRowCount(sheet);
    const columnCountEstimate = estimateColumnCount(sheet);
    const sampleHeaders = collectHeaders(sheet);
    const nonEmptyCellCountEstimate = countNonEmptyCells(sheet);
    const melMentions = countMelMentions(sheet);

    return {
      name: sheet.name,
      state,
      rowCount,
      columnCountEstimate,
      sampleHeaders,
      nonEmptyCellCountEstimate,
      melMentions,
    };
  });

  const totalVisibleSheets = sheetSummaries.filter((s) => s.state === "visible").length;
  const totalHiddenSheets = sheetSummaries.filter((s) => s.state !== "visible").length;
  const totalRowsAcrossSheets = sheetSummaries.reduce((sum, s) => sum + s.rowCount, 0);
  const melMentionsBySheet = Object.fromEntries(sheetSummaries.map((s) => [s.name, s.melMentions]));
  const melMentionsTotal = sheetSummaries.reduce((sum, s) => sum + s.melMentions, 0);
  const assetManager = extractAssetManagerKpis(workbook);

  return {
    workbookFileName: input.fileName,
    receivedAtIso: input.receivedAtIso,
    dateKey: input.dateKey,
    from: input.from,
    to: input.to,
    subject: input.subject,
    workbookBytes: input.binary.byteLength,
    assetManager,
    sheetSummaries,
    totalVisibleSheets,
    totalHiddenSheets,
    totalRowsAcrossSheets,
    melMentionsBySheet,
    melMentionsTotal,
  };
}

function normalizeWorkbookSheetName(name: string): string {
  return name.trim().toLowerCase().replace(/\s+/g, " ");
}

function findAssetManagerSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet | undefined {
  const target = "asset manager";
  return workbook.worksheets.find((ws) => normalizeWorkbookSheetName(ws.name) === target);
}

function parseCellNumber(cell: ExcelJS.Cell): number | null {
  const raw = readCellText(cell).trim();
  if (!raw) return null;
  const cleaned = raw.replace(/,/g, "").replace(/%/g, "").trim();
  const n = Number.parseFloat(cleaned);
  if (!Number.isFinite(n)) return null;
  return n;
}

function parseMcRatePercent(cell: ExcelJS.Cell): number | null {
  const raw = readCellText(cell).trim();
  if (!raw) return null;
  const hasPercent = raw.includes("%");
  const n = parseCellNumber(cell);
  if (n === null) return null;
  if (hasPercent) return n;
  if (n > 0 && n <= 1) return n * 100;
  return n;
}

function parseIntegerCell(cell: ExcelJS.Cell): number | null {
  const n = parseCellNumber(cell);
  if (n === null) return null;
  return Math.round(n);
}

function extractAssetManagerKpis(workbook: ExcelJS.Workbook): AssetManagerKpis {
  const empty: AssetManagerKpis = {
    sheetFound: false,
    mcRatePercent: null,
    fleetTotal: null,
    fmc: null,
    nmc: null,
    surplus: null,
  };
  const sheet = findAssetManagerSheet(workbook);
  if (!sheet) return empty;

  const row = sheet.getRow(2);
  const mcRatePercent = parseMcRatePercent(row.getCell(6));
  const fleetTotal = parseIntegerCell(row.getCell(7));
  const fmc = parseIntegerCell(row.getCell(8));
  const nmc = parseIntegerCell(row.getCell(9));
  const surplus = parseIntegerCell(row.getCell(10));

  const any =
    mcRatePercent !== null ||
    fleetTotal !== null ||
    fmc !== null ||
    nmc !== null ||
    surplus !== null;

  return {
    sheetFound: any,
    mcRatePercent,
    fleetTotal,
    fmc,
    nmc,
    surplus,
  };
}

function normalizeSheetState(state: ExcelJS.WorksheetState | undefined): SheetSummary["state"] {
  if (state === "hidden") return "hidden";
  if (state === "veryHidden") return "veryHidden";
  return "visible";
}

function estimateRowCount(sheet: ExcelJS.Worksheet): number {
  let rows = 0;
  sheet.eachRow({ includeEmpty: false }, () => {
    rows += 1;
  });
  return rows;
}

function estimateColumnCount(sheet: ExcelJS.Worksheet): number {
  let max = 0;
  sheet.eachRow({ includeEmpty: false }, (row) => {
    const actual = row.actualCellCount;
    if (actual > max) max = actual;
  });
  return max;
}

function collectHeaders(sheet: ExcelJS.Worksheet): string[] {
  const firstRow = sheet.getRow(1);
  const headers: string[] = [];
  firstRow.eachCell({ includeEmpty: false }, (cell) => {
    const value = readCellText(cell).trim();
    if (value) headers.push(value);
  });
  return headers.slice(0, 12);
}

function countNonEmptyCells(sheet: ExcelJS.Worksheet): number {
  let count = 0;
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, () => {
      count += 1;
    });
  });
  return count;
}

function countMelMentions(sheet: ExcelJS.Worksheet): number {
  let count = 0;
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const text = readCellText(cell).toLowerCase();
      if (text.includes("mel")) count += 1;
    });
  });
  return count;
}

function readCellText(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map((part) => String(part ?? "")).join("");

  if (typeof value === "object") {
    const anyValue = value as unknown as Record<string, unknown>;
    const richText = anyValue.richText;
    if (Array.isArray(richText)) {
      return richText
        .map((part) =>
          part && typeof part === "object" ? String((part as { text?: unknown }).text ?? "") : "",
        )
        .join("");
    }
    if (typeof anyValue.text === "string") return anyValue.text;
    if (anyValue.result !== null && anyValue.result !== undefined) return String(anyValue.result);
    if (typeof anyValue.formula === "string") return anyValue.formula;
    if (typeof anyValue.hyperlink === "string") return anyValue.hyperlink;
    if (typeof anyValue.error === "string") return anyValue.error;
  }

  try {
    return String(value);
  } catch {
    return "";
  }
}

function renderDashboardHtml(): string {
  const escapedTitle = escapeHtml(SITE_TITLE);
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta name="color-scheme" content="dark light" />
  <meta name="etic-ui" content="tabs-snapshot-wo-v2b" />
  <title>${escapedTitle}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&display=swap" rel="stylesheet" />
  <style>
    :root {
      color-scheme: dark;
      --bg0: #070b14;
      --bg1: #0d1526;
      --surface: rgba(17, 26, 46, 0.72);
      --surface-solid: #121c32;
      --border: rgba(106, 169, 255, 0.14);
      --border-strong: rgba(106, 169, 255, 0.28);
      --text: #eef3fb;
      --muted: #8b9ab5;
      --accent: #6aa9ff;
      --accent-dim: #4d87d9;
      --glow: rgba(106, 169, 255, 0.12);
      --radius: 16px;
      --font: "DM Sans", ui-sans-serif, system-ui, -apple-system, sans-serif;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: var(--font);
      color: var(--text);
      background:
        radial-gradient(ellipse 900px 500px at 15% -20%, var(--glow), transparent 55%),
        radial-gradient(ellipse 700px 400px at 95% 10%, rgba(139, 92, 246, 0.08), transparent 50%),
        linear-gradient(165deg, var(--bg1) 0%, var(--bg0) 45%, #050810 100%);
      background-attachment: fixed;
    }
    .app {
      max-width: 1040px;
      margin: 0 auto;
      padding: 28px 22px 56px;
    }
    .main-nav {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 22px;
    }
    .main-nav button {
      font-family: var(--font);
      font-size: 0.9rem;
      font-weight: 600;
      padding: 12px 20px;
      border-radius: 12px;
      border: 1px solid var(--border-strong);
      background: rgba(12, 18, 32, 0.75);
      color: var(--muted);
      cursor: pointer;
      transition: border-color 0.15s ease, color 0.15s ease, background 0.15s ease;
    }
    .main-nav button:hover {
      border-color: rgba(106, 169, 255, 0.35);
      color: var(--text);
    }
    .main-nav button.active {
      border-color: rgba(106, 169, 255, 0.55);
      background: rgba(18, 28, 48, 0.95);
      color: var(--text);
      box-shadow: 0 0 0 1px rgba(106, 169, 255, 0.12);
    }
    .wo-lookup-row {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: flex-end;
      margin-bottom: 16px;
    }
    .wo-lookup-row label { font-size: 0.72rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.08em; color: var(--muted); display: block; margin-bottom: 6px; }
    .wo-field-grow { flex: 1; min-width: 200px; }
    .wo-field-grow input {
      width: 100%;
      padding: 14px 16px;
      font-family: var(--font);
      font-size: 1rem;
      font-weight: 600;
      color: var(--text);
      background: var(--surface-solid);
      border: 1px solid var(--border-strong);
      border-radius: 12px;
    }
    .wo-field-grow input:focus {
      outline: none;
      border-color: var(--accent);
      box-shadow: 0 0 0 3px var(--glow);
    }
    .btn-primary {
      font-family: var(--font);
      font-weight: 600;
      font-size: 0.9rem;
      padding: 14px 22px;
      border-radius: 12px;
      border: 1px solid rgba(106, 169, 255, 0.45);
      background: rgba(106, 169, 255, 0.18);
      color: var(--text);
      cursor: pointer;
    }
    .btn-primary:hover { background: rgba(106, 169, 255, 0.28); }
    .wo-detail-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
      gap: 12px 18px;
      margin-bottom: 18px;
    }
    .wo-detail-grid dt { font-size: 0.68rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: var(--muted); margin: 0 0 4px; }
    .wo-detail-grid dd { margin: 0; font-size: 0.9rem; line-height: 1.4; word-break: break-word; }
    .wo-remarks {
      margin-top: 8px;
      padding: 14px 16px;
      border-radius: 12px;
      background: rgba(0,0,0,0.25);
      border: 1px solid var(--border);
      font-size: 0.88rem;
      line-height: 1.5;
      white-space: pre-wrap;
      word-break: break-word;
    }
    .wo-stale-banner {
      margin-bottom: 14px;
      padding: 12px 16px;
      border-radius: 10px;
      background: rgba(255, 100, 100, 0.12);
      border: 1px solid rgba(255, 138, 138, 0.35);
      font-size: 0.88rem;
    }
    .top {
      display: flex;
      flex-wrap: wrap;
      align-items: flex-end;
      justify-content: space-between;
      gap: 20px;
      margin-bottom: 28px;
    }
    .brand h1 {
      margin: 0 0 6px;
      font-size: clamp(1.55rem, 3vw, 1.85rem);
      font-weight: 700;
      letter-spacing: -0.03em;
      line-height: 1.15;
    }
    .brand p {
      margin: 0;
      color: var(--muted);
      font-size: 0.92rem;
      max-width: 42ch;
      line-height: 1.45;
    }
    .picker-wrap {
      display: flex;
      flex-direction: column;
      gap: 6px;
      min-width: min(100%, 280px);
    }
    .picker-wrap label {
      font-size: 0.72rem;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.11em;
      color: var(--muted);
    }
    select#etic-date {
      appearance: none;
      width: 100%;
      padding: 14px 44px 14px 16px;
      font-family: var(--font);
      font-size: 1rem;
      font-weight: 600;
      color: var(--text);
      background: var(--surface-solid) url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='20' height='20' viewBox='0 0 24 24' fill='none' stroke='%238b9ab5' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E") no-repeat right 14px center;
      border: 1px solid var(--border-strong);
      border-radius: 12px;
      cursor: pointer;
      box-shadow: 0 8px 32px rgba(0,0,0,0.25);
    }
    select#etic-date:focus {
      outline: none;
      border-color: var(--accent);
      box-shadow: 0 0 0 3px var(--glow), 0 8px 32px rgba(0,0,0,0.25);
    }
    .hero {
      position: relative;
      border-radius: var(--radius);
      padding: 28px 28px 26px;
      margin-bottom: 22px;
      background: var(--surface);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border: 1px solid var(--border);
      box-shadow: 0 24px 48px rgba(0,0,0,0.35);
      overflow: hidden;
    }
    .hero::before {
      content: "";
      position: absolute;
      inset: 0;
      background: linear-gradient(125deg, rgba(106,169,255,0.07) 0%, transparent 45%);
      pointer-events: none;
    }
    .hero-inner { position: relative; z-index: 1; }
    .hero-date {
      display: inline-flex;
      align-items: center;
      gap: 10px;
      font-size: 0.8rem;
      font-weight: 600;
      color: var(--accent);
      letter-spacing: 0.04em;
      text-transform: uppercase;
      margin-bottom: 10px;
    }
    .hero-date .dot {
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: var(--accent);
      box-shadow: 0 0 12px var(--accent);
    }
    .hero h2 {
      margin: 0 0 8px;
      font-size: clamp(1.35rem, 2.5vw, 1.65rem);
      font-weight: 700;
      letter-spacing: -0.02em;
    }
    .hero-file {
      margin: 0;
      color: var(--muted);
      font-size: 0.95rem;
    }
    .hero-file strong { color: #c5d4eb; font-weight: 500; }
    .kpi-strip {
      margin-bottom: 22px;
      padding: 22px 20px 20px;
      border-radius: var(--radius);
      background: #000;
      border: 1px solid rgba(255, 255, 255, 0.08);
      box-shadow: 0 16px 48px rgba(0, 0, 0, 0.45);
    }
    .kpi-strip .kpi-head {
      font-size: 0.65rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.14em;
      color: rgba(255, 255, 255, 0.45);
      margin-bottom: 14px;
    }
    .kpi-row {
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 8px 12px;
      align-items: end;
    }
    @media (max-width: 900px) {
      .kpi-row { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    }
    @media (max-width: 520px) {
      .kpi-row { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    }
    .kpi-cell { text-align: center; padding: 6px 4px; }
    .kpi-cell .lbl {
      font-size: 0.68rem;
      font-weight: 600;
      color: rgba(255, 255, 255, 0.88);
      line-height: 1.25;
      margin-bottom: 10px;
      letter-spacing: 0.02em;
    }
    .kpi-cell .lbl small { display: block; font-size: 0.62rem; font-weight: 600; opacity: 0.9; }
    .kpi-cell .val {
      font-size: clamp(1.35rem, 3.5vw, 1.85rem);
      font-weight: 800;
      letter-spacing: -0.03em;
      line-height: 1.1;
    }
    .kpi-cell .val.em { font-size: clamp(1.55rem, 4vw, 2.1rem); }
    .kpi-val-mc {
      background: linear-gradient(180deg, #ffffff 0%, #7eb8ff 45%, #4a9fff 100%);
      -webkit-background-clip: text;
      background-clip: text;
      color: transparent;
      filter: drop-shadow(0 0 20px rgba(106, 169, 255, 0.45));
    }
    .kpi-val-fleet { color: #7ec8ff; }
    .kpi-val-fmc { color: #5ee397; }
    .kpi-val-nmc { color: #ff8a8a; }
    .kpi-val-surplus { color: #f5d547; }
    .kpi-missing .val { color: rgba(255, 255, 255, 0.25); font-weight: 600; font-size: 1.1rem; }
    .card {
      border-radius: var(--radius);
      background: var(--surface);
      backdrop-filter: blur(12px);
      border: 1px solid var(--border);
      padding: 22px 22px 20px;
      box-shadow: 0 12px 40px rgba(0,0,0,0.22);
    }
    .card h3 {
      margin: 0 0 14px;
      font-size: 0.72rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      color: var(--muted);
    }
    .detail-rows { display: flex; flex-direction: column; gap: 12px; }
    .detail-row {
      display: grid;
      grid-template-columns: 120px 1fr;
      gap: 12px;
      font-size: 0.9rem;
      align-items: start;
    }
    @media (max-width: 520px) { .detail-row { grid-template-columns: 1fr; gap: 4px; } }
    .detail-row dt { color: var(--muted); font-weight: 500; margin: 0; }
    .detail-row dd { margin: 0; color: #d2ddf0; line-height: 1.45; word-break: break-word; }
    .ingest-card { margin-bottom: 22px; }
    .yard-card { margin-bottom: 8px; }
    .yard-card .yard-lead {
      margin: 0 0 18px;
      font-size: 0.92rem;
      color: var(--muted);
      line-height: 1.5;
      max-width: 52ch;
    }
    .yard-export-btn {
      display: flex;
      align-items: center;
      gap: 18px;
      width: 100%;
      margin: 0;
      padding: 20px 22px;
      border-radius: 14px;
      border: 1px solid var(--border-strong);
      background: linear-gradient(145deg, rgba(106, 169, 255, 0.12) 0%, rgba(7, 11, 20, 0.65) 55%);
      box-shadow: 0 8px 32px rgba(0, 0, 0, 0.28), inset 0 1px 0 rgba(255, 255, 255, 0.06);
      cursor: pointer;
      font-family: var(--font);
      text-align: left;
      transition: border-color 0.15s ease, box-shadow 0.15s ease, transform 0.12s ease;
    }
    .yard-export-btn:hover:not(:disabled) {
      border-color: rgba(106, 169, 255, 0.45);
      box-shadow: 0 12px 40px rgba(0, 0, 0, 0.32), 0 0 0 1px rgba(106, 169, 255, 0.12);
    }
    .yard-export-btn:active:not(:disabled) { transform: scale(0.992); }
    .yard-export-btn:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    .yard-export-btn .icon-wrap {
      flex-shrink: 0;
      width: 52px;
      height: 52px;
      border-radius: 14px;
      display: flex;
      align-items: center;
      justify-content: center;
      background: rgba(106, 169, 255, 0.18);
      border: 1px solid rgba(106, 169, 255, 0.28);
      color: var(--accent);
    }
    .yard-export-btn .icon-wrap svg { width: 26px; height: 26px; }
    .yard-export-btn .copy { flex: 1; min-width: 0; }
    .yard-export-btn .title {
      display: block;
      font-size: 1.08rem;
      font-weight: 700;
      letter-spacing: -0.02em;
      color: var(--text);
      margin-bottom: 4px;
    }
    .yard-export-btn .sub {
      display: block;
      font-size: 0.82rem;
      color: var(--muted);
      font-weight: 500;
    }
    .yard-export-btn .chev {
      flex-shrink: 0;
      color: var(--accent);
      opacity: 0.85;
    }
    .yard-export-btn .chev svg { width: 22px; height: 22px; display: block; }
    .action-row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
      margin-bottom: 10px;
    }
    @media (max-width: 560px) { .action-row { grid-template-columns: 1fr; } }
    .btn-etic {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 10px;
      width: 100%;
      padding: 16px 18px;
      border-radius: 14px;
      border: 1px solid rgba(139, 154, 181, 0.35);
      background: rgba(12, 18, 32, 0.85);
      color: var(--text);
      font-family: var(--font);
      font-size: 0.95rem;
      font-weight: 600;
      cursor: pointer;
      transition: border-color 0.15s ease, background 0.15s ease;
    }
    .btn-etic:hover:not(:disabled) {
      border-color: rgba(106, 169, 255, 0.45);
      background: rgba(18, 28, 48, 0.95);
    }
    .btn-etic:disabled { opacity: 0.5; cursor: not-allowed; }
    .compare-card { margin-bottom: 22px; }
    .compare-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      flex-wrap: wrap;
      margin-bottom: 12px;
    }
    .compare-head .hint { font-size: 0.8rem; color: var(--muted); margin: 0; }
    .table-wrap { overflow-x: auto; border-radius: 12px; border: 1px solid var(--border); }
    .compare-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.82rem;
      min-width: 720px;
    }
    .compare-table th, .compare-table td {
      padding: 10px 12px;
      text-align: right;
      border-bottom: 1px solid var(--border);
      white-space: nowrap;
    }
    .compare-table th:first-child, .compare-table td:first-child {
      text-align: left;
      position: sticky;
      left: 0;
      background: var(--surface-solid);
      z-index: 1;
    }
    .compare-table thead th {
      background: var(--surface-solid);
      color: var(--muted);
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      font-size: 0.68rem;
    }
    .compare-table tbody tr:hover td { background: rgba(106, 169, 255, 0.06); }
    .compare-table .mc { color: #9dc6ff; font-weight: 700; }
    .compare-table .fmc { color: #5ee397; }
    .compare-table .nmc { color: #ff8a8a; }
    .compare-table .surp { color: #f5d547; }
    .watch-card { margin-bottom: 22px; }
    .watch-toolbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; margin-bottom: 12px; }
    .watch-toolbar label { font-size: 0.8rem; color: var(--muted); display: flex; align-items: center; gap: 8px; cursor: pointer; }
    .watch-toolbar button.linkish {
      background: none; border: none; color: var(--accent); font-family: var(--font); font-size: 0.8rem;
      cursor: pointer; text-decoration: underline; padding: 0;
    }
    .watch-table { font-size: 0.78rem; min-width: 920px; }
    .watch-table th, .watch-table td { padding: 8px 10px; }
    .watch-table .stale { background: rgba(255, 100, 100, 0.12); }
    .watch-table .badge-tier { font-size: 0.62rem; font-weight: 700; text-transform: uppercase; padding: 2px 6px; border-radius: 4px; }
    .tier-below { background: rgba(255, 138, 138, 0.2); color: #ff9a9a; }
    .tier-at { background: rgba(255, 213, 71, 0.15); color: #f5d547; }
    .tier-above { background: rgba(94, 227, 151, 0.15); color: #5ee397; }
    .tier-unknown { background: rgba(139, 154, 181, 0.15); color: var(--muted); }
    .modal-overlay {
      position: fixed; inset: 0; background: rgba(0,0,0,0.65); z-index: 1000;
      display: flex; align-items: center; justify-content: center; padding: 20px;
    }
    .modal {
      background: var(--surface-solid); border: 1px solid var(--border-strong); border-radius: 14px;
      max-width: 640px; width: 100%; max-height: 85vh; overflow: hidden; display: flex; flex-direction: column;
      box-shadow: 0 24px 64px rgba(0,0,0,0.5);
    }
    .modal header { padding: 16px 18px; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; }
    .modal header h4 { margin: 0; font-size: 1rem; }
    .modal .body { padding: 12px 18px; overflow-y: auto; font-size: 0.82rem; }
    .modal .log-row { padding: 10px 0; border-bottom: 1px solid var(--border); }
    .modal .log-row:last-child { border-bottom: 0; }
    .modal .log-meta { color: var(--muted); font-size: 0.72rem; margin-bottom: 4px; }
    .modal-close { background: none; border: none; color: var(--muted); font-size: 1.25rem; cursor: pointer; line-height: 1; }
    .yard-status-wrap { margin-top: 14px; }
    .status {
      font-size: 0.88rem;
      color: var(--muted);
      min-height: 22px;
      margin: 0;
    }
    .status.err { color: #ff9e9e; }
    .status.ok { color: #8fd4a8; }
    .empty-state {
      text-align: center;
      padding: 48px 24px;
      color: var(--muted);
      font-size: 0.95rem;
      line-height: 1.55;
    }
    .empty-state strong { color: var(--text); }
    .hidden { display: none !important; }
  </style>
</head>
<body>
  <div class="app">
    <header class="top">
      <div class="brand">
        <h1>${escapedTitle}</h1>
        <p id="brand-sub">Snapshots, downloads, and work order tracking.</p>
      </div>
      <div class="picker-wrap hidden" id="snapshot-picker-wrap">
        <label for="etic-date">ETIC report date</label>
        <select id="etic-date" aria-label="Select ETIC report date"></select>
      </div>
    </header>

    <nav class="main-nav" id="main-nav" aria-label="Main sections">
      <button type="button" id="tab-snapshot" class="active">Snapshot</button>
      <button type="button" id="tab-work-orders">Work orders</button>
    </nav>

    <div id="view-empty" class="empty-state hidden">
      <p><strong>No ETIC files yet.</strong><br />Email the Vehicle ETIC workbook to your ingest address to get dates here.</p>
    </div>

    <div id="view-main" class="hidden">
      <div id="panel-snapshot">
      <section class="hero">
        <div class="hero-inner">
          <div class="hero-date"><span class="dot" aria-hidden="true"></span> Selected snapshot</div>
          <h2 id="hero-title">—</h2>
          <p class="hero-file" id="hero-file"></p>
        </div>
      </section>

      <div class="kpi-strip" id="kpi-strip" style="display:none" aria-label="Asset Manager summary">
        <div class="kpi-head">Asset Manager</div>
        <div class="kpi-row" id="kpi-row"></div>
      </div>

      <div class="card ingest-card">
        <h3>This snapshot</h3>
        <dl class="detail-rows" id="ingest-details"></dl>
      </div>

      <div class="card compare-card">
        <div class="compare-head">
          <h3 style="margin:0">Compare snapshots</h3>
          <p class="hint" id="compare-hint">Asset Manager KPIs by report date (D1 index).</p>
        </div>
        <div class="table-wrap">
          <table class="compare-table" id="compare-table" aria-label="Snapshot comparison">
            <thead><tr><th>Report date</th><th>MC %</th><th>Fleet</th><th>FMC</th><th>NMC</th><th>Surplus</th></tr></thead>
            <tbody id="compare-body"><tr><td colspan="6" style="text-align:center;color:var(--muted)">Loading…</td></tr></tbody>
          </table>
        </div>
      </div>

      <div class="card yard-card">
        <h3>Downloads</h3>
        <p class="yard-lead">Get the raw Vehicle ETIC file for this date, or the yard check export for walkarounds.</p>
        <div class="action-row">
          <button type="button" class="btn-etic" id="btn-download-etic">Download Vehicle ETIC (.xlsx)</button>
        </div>
        <p class="status" id="etic-dl-status" role="status"></p>
        <button type="button" class="yard-export-btn" id="btn-yard-check" aria-describedby="yard-status">
          <span class="icon-wrap" aria-hidden="true">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="1.75"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
          </span>
          <span class="copy">
            <span class="title">Export for compound walkaround</span>
            <span class="sub">Excel workbook · landscape · ready to print</span>
          </span>
          <span class="chev" aria-hidden="true">
            <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" /></svg>
          </span>
        </button>
        <div class="yard-status-wrap">
          <p class="status" id="yard-status" role="status"></p>
        </div>
      </div>

      </div>

      <div id="panel-work-orders" class="hidden">
        <div class="card">
          <h3>Look up a work order</h3>
          <p class="yard-lead" style="margin-bottom:14px">
            Enter a work order ID to see current remarks, parts status, ETIC, MEL tier, remark staleness (by MEL: below 3d · at 5d · above 10d), and ETIC push history. Use “As of” to score staleness against any report date in your history.
          </p>
          <div class="wo-lookup-row">
            <div class="wo-field-grow">
              <label for="wo-id-input">Work order ID</label>
              <input type="text" id="wo-id-input" placeholder="e.g. WO12345" autocomplete="off" />
            </div>
            <div class="picker-wrap" style="min-width:200px">
              <label for="wo-asof-date">As of (report date)</label>
              <select id="wo-asof-date" aria-label="Staleness as of report date"></select>
            </div>
            <button type="button" class="btn-primary" id="wo-lookup-btn">Look up</button>
          </div>
          <p class="status" id="wo-lookup-status"></p>
          <div id="wo-detail-empty" class="yard-lead" style="margin-top:8px;color:var(--muted)">Search to load details.</div>
          <div id="wo-detail" class="hidden">
            <div id="wo-stale-banner" class="wo-stale-banner hidden"></div>
            <dl class="wo-detail-grid" id="wo-detail-grid"></dl>
            <div>
              <dt style="font-size:0.68rem;font-weight:700;text-transform:uppercase;letter-spacing:0.06em;color:var(--muted);margin:0 0 8px">Remarks</dt>
              <div class="wo-remarks" id="wo-remarks-text"></div>
            </div>
            <div style="margin-top:18px;display:flex;flex-wrap:wrap;gap:12px;align-items:center">
              <button type="button" class="btn-primary" id="wo-open-changelog" style="background:rgba(12,18,32,0.85);border-color:var(--border-strong)">View full changelog</button>
              <button type="button" class="linkish" id="watch-rebuild" style="font-size:0.85rem">Rebuild history from all ETICs</button>
            </div>
            <p class="status" id="watch-rebuild-status" style="margin-top:10px"></p>
          </div>
        </div>
      </div>

      <div id="changelog-modal" class="modal-overlay hidden" role="dialog" aria-modal="true" aria-labelledby="changelog-title">
        <div class="modal">
          <header>
            <h4 id="changelog-title">Changelog</h4>
            <button type="button" class="modal-close" id="changelog-close" aria-label="Close">×</button>
          </header>
          <div class="body" id="changelog-body"></div>
        </div>
      </div>
    </div>
  </div>
  <script>
    function esc(s) {
      if (s == null || s === undefined) return "";
      return String(s)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
    }

    function fmtKpi(n) {
      if (n === null || n === undefined || typeof n !== "number") return "—";
      return n.toLocaleString(undefined, { maximumFractionDigits: 2 });
    }

    function fmtMc(n) {
      if (n === null || n === undefined || typeof n !== "number") return "—";
      return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "%";
    }

    function sortDesc(entries) {
      return [...entries].sort((a, b) => b.dateKey.localeCompare(a.dateKey));
    }

    function readHashRoute() {
      const raw = (location.hash || "").replace(/^#/, "");
      if (!raw) return { tab: "snapshot", dateKey: null, workOrderId: null };
      if (raw.indexOf("wo=") === 0) {
        const id = decodeURIComponent(raw.slice(3).trim());
        return { tab: "wo", dateKey: null, workOrderId: id || null };
      }
      if (/^\\d{4}-\\d{2}-\\d{2}$/.test(raw)) {
        return { tab: "snapshot", dateKey: raw, workOrderId: null };
      }
      return { tab: "snapshot", dateKey: null, workOrderId: null };
    }

    let historyEntries = [];
    let selectedDate = null;
    let snapshotRows = [];
    function setHashSnapshot(dateKey) {
      if (dateKey) location.hash = "#" + dateKey;
      else location.hash = "";
    }

    function setHashWorkOrder(woId) {
      const t = (woId || "").trim();
      if (!t) {
        location.hash = "";
        return;
      }
      location.hash = "#wo=" + encodeURIComponent(t);
    }

    async function loadHistory() {
      const res = await fetch("/api/history");
      if (!res.ok) throw new Error("Could not load history");
      const data = await res.json();
      return Array.isArray(data.entries) ? data.entries : [];
    }

    async function syncSnapshotsOnce() {
      try {
        if (sessionStorage.getItem("etic_d1_sync_v1")) return;
        await fetch("/api/snapshots?sync=1", { method: "POST" });
        sessionStorage.setItem("etic_d1_sync_v1", "1");
      } catch (_) {}
    }

    async function loadSnapshots() {
      const res = await fetch("/api/snapshots");
      if (!res.ok) return [];
      const data = await res.json();
      return Array.isArray(data.snapshots) ? data.snapshots : [];
    }

    function renderCompareTable() {
      const body = document.getElementById("compare-body");
      const hint = document.getElementById("compare-hint");
      if (!snapshotRows.length) {
        body.innerHTML = "<tr><td colspan='6' style='text-align:center;color:var(--muted)'>No rows in index yet. Reload after a moment, or ingest a new ETIC.</td></tr>";
        return;
      }
      hint.textContent = snapshotRows.length + " snapshots · newest first";
      body.innerHTML = snapshotRows.map(function (r) {
        const sel = r.dateKey === selectedDate ? " style='outline:1px solid rgba(106,169,255,0.5);'" : "";
        const dash = "—";
        const mc = r.mcRatePercent != null ? fmtMc(r.mcRatePercent) : dash;
        const ft = r.fleetTotal != null ? fmtKpi(r.fleetTotal) : dash;
        const fmc = r.fmc != null ? fmtKpi(r.fmc) : dash;
        const nmc = r.nmc != null ? fmtKpi(r.nmc) : dash;
        const sur = r.surplus != null ? fmtKpi(r.surplus) : dash;
        return (
          "<tr" + sel + ">" +
          "<td><strong>" + esc(r.dateKey) + "</strong></td>" +
          "<td class='mc'>" + esc(mc) + "</td>" +
          "<td>" + esc(ft) + "</td>" +
          "<td class='fmc'>" + esc(fmc) + "</td>" +
          "<td class='nmc'>" + esc(nmc) + "</td>" +
          "<td class='surp'>" + esc(sur) + "</td>" +
          "</tr>"
        );
      }).join("");
    }

    function fillSelect(entries) {
      const sel = document.getElementById("etic-date");
      sel.innerHTML = "";
      const sorted = sortDesc(entries);
      for (const e of sorted) {
        const opt = document.createElement("option");
        opt.value = e.dateKey;
        opt.textContent = e.dateKey;
        sel.appendChild(opt);
      }
    }

    async function loadAnalysis(dateKey) {
      const res = await fetch("/api/analysis/" + encodeURIComponent(dateKey));
      if (!res.ok) throw new Error("No analysis for " + dateKey);
      return res.json();
    }

    function renderKpis(am) {
      const row = document.getElementById("kpi-row");
      const strip = document.getElementById("kpi-strip");
      if (!am || !am.sheetFound) {
        strip.style.display = "none";
        row.innerHTML = "";
        return;
      }
      strip.style.display = "block";
      const cells = [
        { lbl: "MC Rate", sub: "", val: fmtMc(am.mcRatePercent), cls: "kpi-val-mc em", miss: am.mcRatePercent == null },
        { lbl: "Fleet Total", sub: "", val: fmtKpi(am.fleetTotal), cls: "kpi-val-fleet", miss: am.fleetTotal == null },
        { lbl: "No. Vehs", sub: "FMC", val: fmtKpi(am.fmc), cls: "kpi-val-fmc", miss: am.fmc == null },
        { lbl: "No. Vehs", sub: "NMC", val: fmtKpi(am.nmc), cls: "kpi-val-nmc", miss: am.nmc == null },
        { lbl: "No. Vehs", sub: "Surplus", val: fmtKpi(am.surplus), cls: "kpi-val-surplus", miss: am.surplus == null },
      ];
      row.innerHTML = cells.map(function (c) {
        const miss = c.miss ? " kpi-missing" : "";
        const sub = c.sub ? "<small>" + esc(c.sub) + "</small>" : "";
        return (
          "<div class='kpi-cell" + miss + "'><div class='lbl'>" + esc(c.lbl) + sub + "</div>" +
          "<div class='val " + c.cls + "'>" + esc(c.val) + "</div></div>"
        );
      }).join("");
    }

    function tierClass(t) {
      if (t === "below") return "tier-below";
      if (t === "at") return "tier-at";
      if (t === "above") return "tier-above";
      return "tier-unknown";
    }

    function fillWoAsOfSelect() {
      const sel = document.getElementById("wo-asof-date");
      sel.innerHTML = "";
      const sorted = sortDesc(historyEntries);
      for (const e of sorted) {
        const opt = document.createElement("option");
        opt.value = e.dateKey;
        opt.textContent = e.dateKey;
        sel.appendChild(opt);
      }
    }

    function syncWoAsOfFromSnapshot() {
      const woSel = document.getElementById("wo-asof-date");
      if (!woSel.options.length || !selectedDate) return;
      const ok = Array.prototype.some.call(woSel.options, function (o) { return o.value === selectedDate; });
      if (ok) woSel.value = selectedDate;
    }

    function setMainTab(which) {
      const snapBtn = document.getElementById("tab-snapshot");
      const woBtn = document.getElementById("tab-work-orders");
      const panelSnap = document.getElementById("panel-snapshot");
      const panelWo = document.getElementById("panel-work-orders");
      const pickerWrap = document.getElementById("snapshot-picker-wrap");
      const brandSub = document.getElementById("brand-sub");
      const isWo = which === "wo";
      snapBtn.classList.toggle("active", !isWo);
      woBtn.classList.toggle("active", isWo);
      panelSnap.classList.toggle("hidden", isWo);
      panelWo.classList.toggle("hidden", !isWo);
      pickerWrap.classList.toggle("hidden", isWo);
      brandSub.textContent = isWo
        ? "Look up any work order ID — staleness uses the “as of” report date."
        : "Pick a report date, compare KPIs, download the raw ETIC or yard check.";
    }

    async function fetchWoById(woId, asOfDate) {
      const res = await fetch(
        "/api/watch?workOrderId=" + encodeURIComponent(woId) + "&date=" + encodeURIComponent(asOfDate),
      );
      if (!res.ok) throw new Error("Lookup failed");
      return res.json();
    }

    function renderWoDetail(data) {
      const empty = document.getElementById("wo-detail-empty");
      const wrap = document.getElementById("wo-detail");
      const st = document.getElementById("wo-lookup-status");
      const ban = document.getElementById("wo-stale-banner");
      const grid = document.getElementById("wo-detail-grid");
      const remarksEl = document.getElementById("wo-remarks-text");
      if (!data.found || !data.row) {
        empty.classList.remove("hidden");
        wrap.classList.add("hidden");
        st.className = "status err";
        st.textContent =
          'No work order "' + esc(data.workOrderId || "") + '" in the index yet. Check the ID or run Rebuild history.';
        return;
      }
      const r = data.row;
      empty.classList.add("hidden");
      wrap.classList.remove("hidden");
      st.className = "status ok";
      st.textContent = "As of " + esc(data.asOfDateKey || "") + " · last seen on snapshot " + esc(r.lastSnapshotDate || "—");
      if (r.remarkStale) {
        ban.classList.remove("hidden");
        const intv = r.requiredIntervalDays != null ? r.requiredIntervalDays + " days" : "—";
        ban.textContent =
          "Stale remarks: no update in " +
          esc(String(r.daysSinceRemarkChange)) +
          " days (MEL tier " +
          esc(r.melTier) +
          " expects at least every " +
          intv +
          ").";
      } else {
        ban.classList.add("hidden");
        ban.textContent = "";
      }
      const eticDisp = (r.eticRaw && String(r.eticRaw).trim()) ? r.eticRaw : (r.eticDate || "—");
      const firstLast =
        esc(r.firstEticDate || "—") + " → " + esc(r.lastEticDate || "—");
      grid.innerHTML = [
        { dt: "Work order", dd: esc(r.workOrderId) },
        { dt: "Asset", dd: esc(r.assetId || "—") },
        { dt: "MEL tier", dd: "<span class='badge-tier " + tierClass(r.melTier) + "'>" + esc(r.melTier) + "</span>" },
        { dt: "Parts status", dd: esc(r.partsStatus || "—") },
        { dt: "ETIC (cell)", dd: esc(String(eticDisp)) },
        { dt: "ETIC (parsed date)", dd: esc(r.eticDate || "—") },
        { dt: "Remark last changed", dd: esc(r.lastRemarkChangeDate || "—") },
        {
          dt: "Days since remark change",
          dd:
            esc(String(r.daysSinceRemarkChange)) +
            (r.requiredIntervalDays != null ? " (expect ≤ " + esc(String(r.requiredIntervalDays)) + "d)" : ""),
        },
        { dt: "ETIC pushes", dd: esc(String(r.eticPushCount ?? 0)) },
        { dt: "Cumulative slip (calendar days)", dd: esc(String(r.cumulativeEticSlipDays ?? 0)) },
        { dt: "First → last ETIC date", dd: firstLast },
      ]
        .map(function (x) {
          return "<div><dt>" + esc(x.dt) + "</dt><dd>" + x.dd + "</dd></div>";
        })
        .join("");
      remarksEl.textContent = (r.remarks && String(r.remarks).trim()) ? r.remarks : "—";
    }

    async function runWoLookup(pushHash) {
      const input = document.getElementById("wo-id-input");
      const woSel = document.getElementById("wo-asof-date");
      const st = document.getElementById("wo-lookup-status");
      const id = (input.value || "").trim();
      if (!id) {
        st.className = "status err";
        st.textContent = "Enter a work order ID.";
        return;
      }
      const asOf = woSel.value || selectedDate;
      if (!asOf) {
        st.className = "status err";
        st.textContent = "No report dates loaded.";
        return;
      }
      st.className = "status";
      st.textContent = "Loading…";
      document.getElementById("wo-detail-empty").classList.remove("hidden");
      document.getElementById("wo-detail").classList.add("hidden");
      try {
        const data = await fetchWoById(id, asOf);
        renderWoDetail(data);
        if (pushHash) setHashWorkOrder(id);
      } catch (e) {
        st.className = "status err";
        st.textContent = String(e && e.message ? e.message : e);
      }
    }

    async function openChangelog(woId) {
      const modal = document.getElementById("changelog-modal");
      const title = document.getElementById("changelog-title");
      const body = document.getElementById("changelog-body");
      title.textContent = "Changelog — " + woId;
      body.innerHTML = "Loading…";
      modal.classList.remove("hidden");
      try {
        const res = await fetch("/api/work-order-changelog?workOrderId=" + encodeURIComponent(woId));
        const data = await res.json();
        const entries = data.entries || [];
        if (!entries.length) {
          body.innerHTML = "<p style='color:var(--muted)'>No changes recorded yet.</p>";
          return;
        }
        body.innerHTML = entries.map(function (e) {
          return (
            "<div class='log-row'>" +
            "<div class='log-meta'>" + esc(e.snapshot_date_key) + " · " + esc(e.field) + " · " + esc(e.changed_at_iso) + "</div>" +
            "<div>" + esc(e.old_value || "—") + " → " + esc(e.new_value || "—") + "</div>" +
            "</div>"
          );
        }).join("");
      } catch (err) {
        body.innerHTML = "<p style='color:#ff9a9a'>" + esc(String(err.message || err)) + "</p>";
      }
    }

    function renderDetails(analysis) {
      document.getElementById("hero-title").textContent = analysis.dateKey;
      document.getElementById("hero-file").innerHTML =
        "File <strong>" + esc(analysis.workbookFileName) + "</strong>";

      renderKpis(analysis.assetManager);
      renderCompareTable();
      fillWoAsOfSelect();
      syncWoAsOfFromSnapshot();

      const ing = document.getElementById("ingest-details");
      const recv = analysis.receivedAtIso
        ? new Date(analysis.receivedAtIso).toLocaleString(undefined, {
            dateStyle: "medium",
            timeStyle: "short",
          })
        : "—";
      ing.innerHTML = [
        "<div class='detail-row'><dt>Received</dt><dd>" + esc(recv) + "</dd></div>",
        "<div class='detail-row'><dt>From</dt><dd>" + esc(analysis.from) + "</dd></div>",
        "<div class='detail-row'><dt>To</dt><dd>" + esc(analysis.to) + "</dd></div>",
        "<div class='detail-row'><dt>Workbook size</dt><dd>" +
          (typeof analysis.workbookBytes === "number"
            ? (analysis.workbookBytes / 1024 / 1024).toFixed(2) + " MB"
            : "—") +
          "</dd></div>",
      ].join("");
    }

    async function selectDate(dateKey, pushHash) {
      selectedDate = dateKey;
      if (pushHash && dateKey) setHashSnapshot(dateKey);
      document.getElementById("yard-status").textContent = "";
      document.getElementById("yard-status").className = "status";

      if (!dateKey) return;

      try {
        const analysis = await loadAnalysis(dateKey);
        renderDetails(analysis);
      } catch (e) {
        document.getElementById("ingest-details").innerHTML =
          "<div class='detail-row'><dt>Error</dt><dd>" + esc(e.message || String(e)) + "</dd></div>";
      }
    }

    async function applyHashRoute() {
      const r = readHashRoute();
      if (r.tab === "wo" && r.workOrderId) {
        setMainTab("wo");
        document.getElementById("wo-id-input").value = r.workOrderId;
        fillWoAsOfSelect();
        syncWoAsOfFromSnapshot();
        await runWoLookup(false);
        return;
      }
      setMainTab("snapshot");
      const sel = document.getElementById("etic-date");
      if (r.dateKey && historyEntries.some((e) => e.dateKey === r.dateKey)) {
        if (sel.value !== r.dateKey) sel.value = r.dateKey;
        await selectDate(r.dateKey, false);
      } else {
        await selectDate(sel.value, false);
      }
    }

    async function downloadEticWorkbook() {
      const btn = document.getElementById("btn-download-etic");
      const st = document.getElementById("etic-dl-status");
      if (!selectedDate) return;
      btn.disabled = true;
      st.className = "status";
      st.textContent = "Downloading…";
      try {
        const url = "/api/workbook.xlsx?date=" + encodeURIComponent(selectedDate);
        const res = await fetch(url);
        if (!res.ok) {
          let msg = "Could not download.";
          try {
            const j = await res.json();
            if (j && j.error) msg = j.error;
          } catch (_) {}
          st.className = "status err";
          st.textContent = msg;
          return;
        }
        const disp = res.headers.get("Content-Disposition") || "";
        const m = /filename="?([^";]+)"?/i.exec(disp);
        const name = m ? m[1] : "etic.xlsx";
        const blob = await res.blob();
        const u = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = u;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(u);
        st.className = "status ok";
        st.textContent = "Saved " + name;
      } catch (e) {
        st.className = "status err";
        st.textContent = String(e && e.message ? e.message : e);
      } finally {
        btn.disabled = false;
      }
    }

    async function downloadYardCheck() {
      const btn = document.getElementById("btn-yard-check");
      const st = document.getElementById("yard-status");
      if (!selectedDate) return;
      btn.disabled = true;
      st.className = "status";
      st.textContent = "Building your file…";
      try {
        const url = "/api/yard-check.xlsx?date=" + encodeURIComponent(selectedDate);
        const res = await fetch(url);
        if (!res.ok) {
          let msg = "Could not generate file.";
          try {
            const j = await res.json();
            if (j && j.error) msg = j.error;
          } catch (_) {}
          st.className = "status err";
          st.textContent = msg;
          return;
        }
        const disp = res.headers.get("Content-Disposition") || "";
        const m = /filename="?([^";]+)"?/i.exec(disp);
        const name = m ? m[1] : "yard-check.xlsx";
        const blob = await res.blob();
        const u = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = u;
        a.download = name;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(u);
        st.className = "status ok";
        st.textContent = "Saved as " + name;
      } catch (e) {
        st.className = "status err";
        st.textContent = String(e && e.message ? e.message : e);
      } finally {
        btn.disabled = false;
      }
    }

    async function init() {
      try {
        historyEntries = await loadHistory();
      } catch (e) {
        document.getElementById("view-empty").classList.remove("hidden");
        document.getElementById("view-empty").querySelector("p").innerHTML =
          "<strong>Could not load data.</strong><br />" + esc(e.message || String(e));
        return;
      }

      if (!historyEntries.length) {
        document.getElementById("view-empty").classList.remove("hidden");
        return;
      }

      document.getElementById("view-main").classList.remove("hidden");
      fillSelect(historyEntries);

      snapshotRows = await loadSnapshots();
      renderCompareTable();
      await syncSnapshotsOnce();
      setTimeout(async function () {
        snapshotRows = await loadSnapshots();
        renderCompareTable();
      }, 3000);

      const sorted = sortDesc(historyEntries);
      const route = readHashRoute();
      const start =
        route.tab === "snapshot" && route.dateKey && sorted.some((e) => e.dateKey === route.dateKey)
          ? route.dateKey
          : sorted[0].dateKey;
      const sel = document.getElementById("etic-date");
      sel.value = start;
      selectedDate = start;
      await applyHashRoute();

      sel.addEventListener("change", async () => {
        await selectDate(sel.value, true);
      });

      window.addEventListener("hashchange", async () => {
        await applyHashRoute();
      });

      document.getElementById("tab-snapshot").addEventListener("click", function () {
        setHashSnapshot(selectedDate || sortDesc(historyEntries)[0].dateKey);
      });
      document.getElementById("tab-work-orders").addEventListener("click", function () {
        const cur = (document.getElementById("wo-id-input").value || "").trim();
        if (cur) setHashWorkOrder(cur);
        else setMainTab("wo");
      });

      document.getElementById("btn-yard-check").addEventListener("click", downloadYardCheck);
      document.getElementById("btn-download-etic").addEventListener("click", downloadEticWorkbook);

      document.getElementById("wo-lookup-btn").addEventListener("click", function () {
        runWoLookup(true);
      });
      document.getElementById("wo-id-input").addEventListener("keydown", function (ev) {
        if (ev.key === "Enter") runWoLookup(true);
      });
      document.getElementById("wo-asof-date").addEventListener("change", function () {
        const id = (document.getElementById("wo-id-input").value || "").trim();
        if (id) runWoLookup(true);
      });
      document.getElementById("wo-open-changelog").addEventListener("click", function () {
        const id = (document.getElementById("wo-id-input").value || "").trim();
        if (id) openChangelog(id);
      });
      document.getElementById("watch-rebuild").addEventListener("click", async function () {
        const st = document.getElementById("watch-rebuild-status");
        if (!confirm("Rebuild clears work order changelog and replays every ETIC in date order. Continue?")) return;
        st.textContent = "Starting rebuild…";
        try {
          const res = await fetch("/api/watch?rebuild=1", { method: "POST" });
          const j = await res.json();
          st.textContent = (j && j.message) ? j.message : "Rebuild started.";
          const id = (document.getElementById("wo-id-input").value || "").trim();
          if (id) setTimeout(function () { runWoLookup(false); }, 5000);
        } catch (e) {
          st.textContent = String(e && e.message ? e.message : e);
        }
      });
      document.getElementById("changelog-close").addEventListener("click", function () {
        document.getElementById("changelog-modal").classList.add("hidden");
      });
      document.getElementById("changelog-modal").addEventListener("click", function (ev) {
        if (ev.target.id === "changelog-modal") document.getElementById("changelog-modal").classList.add("hidden");
      });
    }

    init();
  </script>
</body>
</html>`;
}

function escapeHtml(input: string): string {
  return input
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export {
  analyzeWorkbook,
  countMelMentions,
  extractAssetManagerKpis,
  isoDateKey,
  parseMaxAttachmentBytes,
  parseReportDateKeyFromSubject,
  pickWorkbookAttachment,
  readCellText,
  resolveAnalysisDateKey,
  sanitizeFileName,
  upsertHistoryEntry,
};

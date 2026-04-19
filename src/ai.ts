/**
 * In-app AI assistant. Answers natural-language questions about ETIC data
 * (snapshots, work orders, meetings, changelog) using OpenAI tool/function
 * calling. All tools are READ-ONLY queries against D1 + the cached analysis
 * JSON in R2.
 *
 * Wire-up:
 *   - Set the API key once:  wrangler secret put OPENAI_API_KEY
 *   - POST /api/ask  body { messages: [{role,content}, ...] }
 *
 * The orchestration loop runs at most MAX_ROUNDS iterations to bound cost.
 */
import { getChangelog, getWatchRowsLatest, getWorkOrderTimeline, getWatchRowById } from "./workOrderWatch";
import { listMeetings, getMeetingWithNotes } from "./meeting";

type AiEnv = {
  ETIC_SNAPSHOTS: D1Database;
  ETIC_BUCKET: R2Bucket;
  OPENAI_API_KEY?: string;
};

type ChatMessage = {
  role: "system" | "user" | "assistant" | "tool";
  content: string | null;
  tool_calls?: Array<{
    id: string;
    type: "function";
    function: { name: string; arguments: string };
  }>;
  tool_call_id?: string;
  name?: string;
};

type AskTrace = {
  tool: string;
  args: Record<string, unknown>;
  ms: number;
  ok: boolean;
  bytes: number;
  preview?: string;
};

const MAX_ROUNDS = 6;
const MAX_TOOL_RESULT_CHARS = 18000;
const DEFAULT_MODEL = "gpt-4o-mini";

const SYSTEM_PROMPT = `You are the in-app analyst for the Minot AFB Vehicle ETIC tracker.

You have READ-ONLY tools to query a daily snapshot history of fleet KPIs,
per-unit MC%, individual work orders (with full change history), and ETIC
meetings. Always call a tool to get real numbers — never guess.

Vocabulary the user will use:
  - "MC rate" / "MC%"  = mission-capable rate, FMC / (FMC + NMC).
  - "Below MEL"        = a unit/MEL key whose total assets in maintenance
                         exceeds the Minimum Essential Level. Always urgent.
  - "ETIC"             = Estimated Time In Commission, i.e. the date a work
                         order should close. "Slipped" = ETIC moved to a
                         later date. "Pushed" = same thing.
  - "NCE"              = Nuclear Certified Equipment.
  - "Unit" examples    = "5 LRS", "5 CES", "5 SFS", "5 MUNS", etc.
  - "Shop" examples    = "GP" (general purpose), "MOA" (material control),
                         "ALLIED" (allied trades), "REFUEL", "SPEC".
  - "MEL Tier"         = "below_mel" | "at_mel" | "above_mel" | "unknown".

Date handling:
  - Snapshot dates are ISO YYYY-MM-DD. The user will say things like
    "last quarter" or "April"; resolve those to ISO ranges yourself.
  - Today's date is provided in the user message header.

Metric polarity (CRITICAL — do not get this wrong):
  - Fleet MC% / mcPercent / mcPct       : higher is BETTER.
  - FMC count                           : higher is BETTER.
  - NMC count                           : lower is BETTER.
  - Assets below MEL / keys below MEL   : lower is BETTER.
  - Assets to reach AT MEL              : lower is BETTER.
  - Recall delta                        : closer to 0 is BETTER.
  - Acc/Abus                            : lower is BETTER.
  - Days since remarks change (staleness): lower is BETTER.
  - ETIC pushes / cumulative slip days  : lower is BETTER.

When asked whether something is "better" or "worse" over a window:
  1. Compute the direction each metric moved.
  2. Map each direction through the polarity table above.
  3. Count how many indicators improved vs declined.
  4. Say "improving", "declining", or "mixed (X improved, Y declined)".
  5. NEVER call a DROP in below-MEL count a worsening trend.
  6. NEVER call a RISE in MC% a worsening trend.
  7. If the user's pinned context already provides first/last or compareTotals, use those values directly — don't re-derive from the series.

Style:
  - Be concise. Answer the actual question first, then a short rationale.
  - Use plain numbers and percentages with 1 decimal (e.g. "78.4%").
  - When showing trends, give first-value, last-value, and delta (pp).
  - Format work-order ids as monospace.
  - Use short markdown bullets/tables when it helps; otherwise prose.`;

const TOOLS = [
  {
    type: "function",
    function: {
      name: "list_snapshots",
      description:
        "Returns recent snapshot dates with top-line KPIs (MC%, FMC, NMC, fleet total). Use this first to discover what dates are available.",
      parameters: {
        type: "object",
        properties: {
          limit: { type: "integer", description: "Max rows (default 30, max 365).", minimum: 1, maximum: 365 },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_snapshot",
      description: "Returns the headline KPIs for one snapshot date.",
      parameters: {
        type: "object",
        properties: { date: { type: "string", description: "ISO YYYY-MM-DD." } },
        required: ["date"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_unit_breakdown",
      description:
        "Returns the per-unit MC% breakdown (and the NCE row if present) for one snapshot date. Use to answer 'how is unit X doing today'.",
      parameters: {
        type: "object",
        properties: { date: { type: "string", description: "ISO YYYY-MM-DD." } },
        required: ["date"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_unit_history",
      description:
        "Returns one row per snapshot in [from, to] for a single unit (or 'NCE'), with MC%, FMC, NMC. Use for trend questions like 'how has 5 CES done this last quarter?'.",
      parameters: {
        type: "object",
        properties: {
          unit: { type: "string", description: "Unit label exactly as it appears (e.g. '5 CES'), or 'NCE' for the NCE row." },
          from: { type: "string", description: "Inclusive ISO start date. Defaults to 90 days before to_date." },
          to: { type: "string", description: "Inclusive ISO end date. Defaults to the latest snapshot." },
        },
        required: ["unit"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "compare_snapshots",
      description: "Headline KPI deltas between two snapshot dates.",
      parameters: {
        type: "object",
        properties: {
          from: { type: "string", description: "ISO YYYY-MM-DD." },
          to: { type: "string", description: "ISO YYYY-MM-DD." },
        },
        required: ["from", "to"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_work_orders",
      description:
        "Latest-known state for all open work orders, with optional filters. Use for 'show me below-MEL work orders for 5 CES' style questions.",
      parameters: {
        type: "object",
        properties: {
          owning_unit: { type: "string", description: "Exact unit, e.g. '5 CES'." },
          shop: { type: "string", description: "Exact shop code, e.g. 'GP'." },
          mel_tier: { type: "string", enum: ["below_mel", "at_mel", "above_mel", "unknown"] },
          mel_key: { type: "string" },
          etic_pushed_min: { type: "integer", description: "Only WOs whose ETIC has been pushed at least this many times." },
          slipped_days_min: { type: "integer", description: "Only WOs whose cumulative ETIC slip exceeds this many days." },
          stale_days_min: { type: "integer", description: "Only WOs whose remarks haven't changed in at least this many days (vs latest snapshot)." },
          search: { type: "string", description: "Substring match on asset id, work order id, or remarks." },
          limit: { type: "integer", description: "Max rows (default 25, max 200).", minimum: 1, maximum: 200 },
        },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_work_order",
      description: "Single work order: latest state, recent changelog, full timeline (one row per snapshot).",
      parameters: {
        type: "object",
        properties: {
          id: { type: "string", description: "Work order id." },
          changelog_limit: { type: "integer", default: 25, minimum: 1, maximum: 200 },
        },
        required: ["id"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "list_meetings",
      description: "Recent ETIC meetings (most recent first).",
      parameters: {
        type: "object",
        properties: { limit: { type: "integer", default: 10, minimum: 1, maximum: 100 } },
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_meeting",
      description: "One meeting + its per-WO notes and due-outs.",
      parameters: {
        type: "object",
        properties: { id: { type: "integer", description: "Meeting id." } },
        required: ["id"],
      },
    },
  },
];

/* ───────────────────────── Tool implementations ─────────────────────────── */

type SnapshotRow = {
  date_key: string;
  mc_rate: number | null;
  fleet_total: number | null;
  fmc: number | null;
  nmc: number | null;
  surplus: number | null;
  asset_manager_breakdown: string;
};

async function getLatestSnapshotDate(env: AiEnv): Promise<string | null> {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    "SELECT date_key FROM etic_snapshots ORDER BY date_key DESC LIMIT 1",
  ).first<{ date_key: string }>();
  return r?.date_key ?? null;
}

async function tool_list_snapshots(env: AiEnv, args: { limit?: number }) {
  const limit = Math.max(1, Math.min(365, args.limit ?? 30));
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT date_key, mc_rate, fleet_total, fmc, nmc, surplus
     FROM etic_snapshots
     ORDER BY date_key DESC
     LIMIT ?`,
  ).bind(limit).all<{ date_key: string; mc_rate: number | null; fleet_total: number | null; fmc: number | null; nmc: number | null; surplus: number | null }>();
  return { rows: r.results ?? [] };
}

async function tool_get_snapshot(env: AiEnv, args: { date: string }) {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT date_key, mc_rate, fleet_total, fmc, nmc, surplus, mel_total, asset_manager_ok
     FROM etic_snapshots WHERE date_key = ?`,
  ).bind(args.date).first();
  if (!r) return { error: `No snapshot found for ${args.date}.` };
  return r;
}

type Breakdown = {
  label: string;
  mcRatePercent: number | null;
  fleetTotal: number | null;
  fmc: number | null;
  nmc: number | null;
  surplus: number | null;
  isNce: boolean;
  isTotal: boolean;
};

function parseBreakdown(json: string): Breakdown[] {
  if (!json) return [];
  try {
    const v = JSON.parse(json);
    return Array.isArray(v) ? v : [];
  } catch {
    return [];
  }
}

async function tool_get_unit_breakdown(env: AiEnv, args: { date: string }) {
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT date_key, asset_manager_breakdown FROM etic_snapshots WHERE date_key = ?`,
  ).bind(args.date).first<{ date_key: string; asset_manager_breakdown: string }>();
  if (!r) return { error: `No snapshot for ${args.date}.` };
  const rows = parseBreakdown(r.asset_manager_breakdown);
  return { date: r.date_key, units: rows.filter((x) => !x.isNce && !x.isTotal), nce: rows.find((x) => x.isNce) ?? null };
}

async function tool_get_unit_history(env: AiEnv, args: { unit: string; from?: string; to?: string }) {
  const latest = (await getLatestSnapshotDate(env)) ?? new Date().toISOString().slice(0, 10);
  const to = args.to ?? latest;
  const from = args.from ?? defaultFrom(to, 90);
  const r = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT date_key, asset_manager_breakdown
     FROM etic_snapshots
     WHERE date_key BETWEEN ? AND ?
     ORDER BY date_key ASC`,
  ).bind(from, to).all<{ date_key: string; asset_manager_breakdown: string }>();
  const wantNce = /^nce\b/i.test(args.unit.trim());
  const wantUnit = args.unit.trim().toLowerCase();
  const out: Array<{ date: string; mcRatePercent: number | null; fmc: number | null; nmc: number | null; fleetTotal: number | null }> = [];
  for (const row of r.results ?? []) {
    const rows = parseBreakdown(row.asset_manager_breakdown);
    const hit = rows.find((x) => (wantNce ? x.isNce : x.label.toLowerCase() === wantUnit));
    if (hit) {
      out.push({
        date: row.date_key,
        mcRatePercent: hit.mcRatePercent,
        fmc: hit.fmc,
        nmc: hit.nmc,
        fleetTotal: hit.fleetTotal,
      });
    }
  }
  if (out.length === 0) {
    return {
      error: `No data for unit "${args.unit}" between ${from} and ${to}. Use get_unit_breakdown on a recent date to see exact unit labels.`,
      from,
      to,
    };
  }
  const first = out[0];
  const last = out[out.length - 1];
  const deltaPp =
    first.mcRatePercent != null && last.mcRatePercent != null
      ? last.mcRatePercent - first.mcRatePercent
      : null;
  return { unit: args.unit, from, to, points: out.length, first, last, deltaPp, history: out };
}

async function tool_compare_snapshots(env: AiEnv, args: { from: string; to: string }) {
  const both = await env.ETIC_SNAPSHOTS.prepare(
    `SELECT date_key, mc_rate, fleet_total, fmc, nmc FROM etic_snapshots WHERE date_key IN (?, ?)`,
  ).bind(args.from, args.to).all<{ date_key: string; mc_rate: number | null; fleet_total: number | null; fmc: number | null; nmc: number | null }>();
  const rows = both.results ?? [];
  const a = rows.find((x) => x.date_key === args.from);
  const b = rows.find((x) => x.date_key === args.to);
  if (!a || !b) return { error: "One or both snapshots not found." };
  const delta = (x: number | null | undefined, y: number | null | undefined) =>
    x == null || y == null ? null : y - x;
  return {
    from: a,
    to: b,
    delta: {
      mc_rate_pp: delta(a.mc_rate, b.mc_rate),
      fleet_total: delta(a.fleet_total, b.fleet_total),
      fmc: delta(a.fmc, b.fmc),
      nmc: delta(a.nmc, b.nmc),
    },
  };
}

async function tool_list_work_orders(
  env: AiEnv,
  args: {
    owning_unit?: string;
    shop?: string;
    mel_tier?: string;
    mel_key?: string;
    etic_pushed_min?: number;
    slipped_days_min?: number;
    stale_days_min?: number;
    search?: string;
    limit?: number;
  },
) {
  const limit = Math.max(1, Math.min(200, args.limit ?? 25));
  const latest = (await getLatestSnapshotDate(env)) ?? new Date().toISOString().slice(0, 10);
  const all = await getWatchRowsLatest(env, latest);
  const needle = (args.search ?? "").toLowerCase();
  const filtered = all.filter((w) => {
    if (args.owning_unit && (w.owningUnit || "").toLowerCase() !== args.owning_unit.toLowerCase()) return false;
    if (args.shop && (w.shop || "").toLowerCase() !== args.shop.toLowerCase()) return false;
    if (args.mel_tier && w.melTier !== args.mel_tier) return false;
    if (args.mel_key && (w.melKey || "").toLowerCase() !== args.mel_key.toLowerCase()) return false;
    if (args.etic_pushed_min != null && w.eticPushCount < args.etic_pushed_min) return false;
    if (args.slipped_days_min != null && w.cumulativeEticSlipDays < args.slipped_days_min) return false;
    if (args.stale_days_min != null && (w.daysSinceRemarkChange ?? 0) < args.stale_days_min) return false;
    if (needle) {
      const blob = `${w.workOrderId} ${w.assetId} ${w.remarks}`.toLowerCase();
      if (!blob.includes(needle)) return false;
    }
    return true;
  });
  return {
    asOfDate: latest,
    matched: filtered.length,
    returned: Math.min(filtered.length, limit),
    rows: filtered.slice(0, limit).map((w) => ({
      workOrderId: w.workOrderId,
      assetId: w.assetId,
      vehNomen: w.vehNomen,
      makeModel: w.makeModel,
      owningUnit: w.owningUnit,
      mgmtCd: w.mgmtCd,
      shop: w.shop,
      melTier: w.melTier,
      melKey: w.melKey,
      partsStatus: w.partsStatus,
      eticDate: w.eticDate,
      eticPushCount: w.eticPushCount,
      cumulativeSlipDays: w.cumulativeEticSlipDays,
      daysSinceRemarkChange: w.daysSinceRemarkChange,
      remarksAgeIsLowerBoundOnly: w.historyBounded,
      firstSeenInOurData: w.firstSeenDate,
      establishedDate: w.establishedDateIso || w.establishedDate || null,
      reason: w.woReason || null,
      isNCE: w.nce,
      nceStatus: w.nceStatus || null,
      remarksPreview: (w.remarks || "").slice(0, 240),
    })),
  };
}

async function tool_get_work_order(env: AiEnv, args: { id: string; changelog_limit?: number }) {
  const latest = (await getLatestSnapshotDate(env)) ?? new Date().toISOString().slice(0, 10);
  const state = await getWatchRowById(env, args.id, latest);
  if (!state) return { error: `No work order ${args.id}.` };
  const [changelog, timeline] = await Promise.all([
    getChangelog(env, args.id, args.changelog_limit ?? 25),
    getWorkOrderTimeline(env, args.id),
  ]);
  return {
    state: {
      workOrderId: state.workOrderId,
      assetId: state.assetId,
      vehNomen: state.vehNomen,
      makeModel: state.makeModel,
      owningUnit: state.owningUnit,
      shop: state.shop,
      melTier: state.melTier,
      melKey: state.melKey,
      partsStatus: state.partsStatus,
      eticDate: state.eticDate,
      eticPushCount: state.eticPushCount,
      cumulativeSlipDays: state.cumulativeEticSlipDays,
      remarks: state.remarks,
      lastSnapshotDate: state.lastSnapshotDate,
    },
    changelog,
    timeline: timeline.map((t) => ({
      date: t.lastSnapshotDate,
      melTier: t.melTier,
      eticDate: t.eticDate,
      eticPushCount: t.eticPushCount,
      cumulativeSlipDays: t.cumulativeEticSlipDays,
      remarksPreview: (t.remarks || "").slice(0, 200),
    })),
  };
}

async function tool_list_meetings(env: AiEnv, args: { limit?: number }) {
  const list = await listMeetings(env, args.limit ?? 10);
  return { meetings: list };
}

async function tool_get_meeting(env: AiEnv, args: { id: number }) {
  const m = await getMeetingWithNotes(env, args.id);
  if (!m) return { error: `No meeting ${args.id}.` };
  return m;
}

const TOOL_TABLE: Record<string, (env: AiEnv, args: any) => Promise<unknown>> = {
  list_snapshots: tool_list_snapshots,
  get_snapshot: tool_get_snapshot,
  get_unit_breakdown: tool_get_unit_breakdown,
  get_unit_history: tool_get_unit_history,
  compare_snapshots: tool_compare_snapshots,
  list_work_orders: tool_list_work_orders,
  get_work_order: tool_get_work_order,
  list_meetings: tool_list_meetings,
  get_meeting: tool_get_meeting,
};

function defaultFrom(toIso: string, daysBack: number): string {
  const t = new Date(toIso + "T00:00:00Z");
  if (Number.isNaN(t.getTime())) return toIso;
  t.setUTCDate(t.getUTCDate() - daysBack);
  return t.toISOString().slice(0, 10);
}

/* ──────────────────────────── Orchestration ─────────────────────────────── */

export async function handleAskApi(env: AiEnv, request: Request): Promise<Response> {
  if (request.method !== "POST") {
    return Response.json({ error: "POST only" }, { status: 405 });
  }
  if (!env.OPENAI_API_KEY) {
    return Response.json(
      { error: "OPENAI_API_KEY is not configured. Run: wrangler secret put OPENAI_API_KEY" },
      { status: 500 },
    );
  }
  let body: { messages?: ChatMessage[]; model?: string };
  try {
    body = await request.json();
  } catch {
    return Response.json({ error: "Invalid JSON body" }, { status: 400 });
  }
  const userMessages = Array.isArray(body.messages) ? body.messages : [];
  if (userMessages.length === 0) {
    return Response.json({ error: "messages[] is required" }, { status: 400 });
  }

  const today = new Date().toISOString().slice(0, 10);
  const latestDate = (await getLatestSnapshotDate(env)) ?? "(none)";
  const messages: ChatMessage[] = [
    { role: "system", content: SYSTEM_PROMPT },
    { role: "system", content: `Today is ${today}. The most recent snapshot in the database is ${latestDate}.` },
    ...userMessages,
  ];

  const trace: AskTrace[] = [];
  const model = body.model || DEFAULT_MODEL;

  for (let round = 0; round < MAX_ROUNDS; round++) {
    const reply = await callOpenAi(env.OPENAI_API_KEY, model, messages);
    if (!reply.ok) {
      return Response.json({ error: reply.error, trace }, { status: 502 });
    }
    const choice = reply.choice;
    messages.push(choice);

    const toolCalls = choice.tool_calls ?? [];
    if (toolCalls.length === 0) {
      return Response.json({ answer: choice.content ?? "", trace });
    }

    for (const call of toolCalls) {
      const name = call.function.name;
      let args: Record<string, unknown> = {};
      try {
        args = JSON.parse(call.function.arguments || "{}");
      } catch {
        args = {};
      }
      const fn = TOOL_TABLE[name];
      const t0 = Date.now();
      let resultText: string;
      let ok = true;
      if (!fn) {
        resultText = JSON.stringify({ error: `Unknown tool ${name}` });
        ok = false;
      } else {
        try {
          const result = await fn(env, args);
          resultText = JSON.stringify(result);
        } catch (e) {
          ok = false;
          resultText = JSON.stringify({ error: e instanceof Error ? e.message : String(e) });
        }
      }
      if (resultText.length > MAX_TOOL_RESULT_CHARS) {
        resultText = resultText.slice(0, MAX_TOOL_RESULT_CHARS) + "\n…[truncated]";
      }
      trace.push({
        tool: name,
        args,
        ms: Date.now() - t0,
        ok,
        bytes: resultText.length,
        preview: resultText.slice(0, 280),
      });
      messages.push({
        role: "tool",
        tool_call_id: call.id,
        name,
        content: resultText,
      });
    }
  }

  return Response.json(
    {
      answer:
        "I hit the tool-call round limit before reaching a final answer. Try asking a more specific question or narrowing the date range.",
      trace,
    },
    { status: 200 },
  );
}

async function callOpenAi(
  apiKey: string,
  model: string,
  messages: ChatMessage[],
): Promise<{ ok: true; choice: ChatMessage } | { ok: false; error: string }> {
  const r = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      temperature: 0.2,
      messages,
      tools: TOOLS,
      tool_choice: "auto",
    }),
  });
  if (!r.ok) {
    const text = await r.text().catch(() => "");
    return { ok: false, error: `OpenAI ${r.status}: ${text.slice(0, 500)}` };
  }
  const data = (await r.json()) as {
    choices?: Array<{ message?: ChatMessage }>;
    error?: { message?: string };
  };
  const choice = data.choices?.[0]?.message;
  if (!choice) return { ok: false, error: data.error?.message || "OpenAI returned no choice" };
  return { ok: true, choice };
}

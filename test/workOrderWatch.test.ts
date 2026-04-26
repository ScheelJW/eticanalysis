import { describe, expect, it } from "vitest";
import type { RawWorkOrder } from "../src/yardCheck";
import {
  analyzeElmsScheduleMxFromRaw,
  calendarDaysBetween,
  classifyMelTier,
  computeMelRecallHintForRow,
  computeScheduleMxAssetStats,
  computeScheduleMxCommanderSummary,
  enrichScheduleMxRowsWithLatestEtic,
  elmsPlanRowKeyFromRaw,
  getChangelogForDisplay,
  ingestWorkOrderSnapshot,
  melMgmtCodesMatch,
  parseCsvTextToRowArrays,
  parseEticDate,
  parseScheduleMxCsvToPlanRows,
  parseScheduleMxCsvToRawByAsset,
} from "../src/workOrderWatch";
import type { ScheduleMxFleetRow, WatchRow } from "../src/workOrderWatch";

describe("classifyMelTier", () => {
  it("detects below / at / above from phrases", () => {
    expect(classifyMelTier("Below MEL")).toBe("below");
    expect(classifyMelTier("at MEL")).toBe("at");
    expect(classifyMelTier("Above MEL")).toBe("above");
  });

  it("maps numeric MEL levels", () => {
    expect(classifyMelTier("1")).toBe("below");
    expect(classifyMelTier("3")).toBe("at");
    expect(classifyMelTier("4")).toBe("above");
  });

  it("returns unknown when empty", () => {
    expect(classifyMelTier("")).toBe("unknown");
  });
});

describe("parseEticDate", () => {
  it("parses ISO prefix and US dates", () => {
    expect(parseEticDate("2026-04-20")).toBe("2026-04-20");
    expect(parseEticDate("4/20/2026")).toBe("2026-04-20");
  });
});

describe("parseCsvTextToRowArrays", () => {
  it("treats newline inside quotes as part of cell, not row break", () => {
    const csv =
      'Asset Id,Note\n' +
      'AF01,"line1\nline2"\n' +
      'AF02,x\n';
    const rows = parseCsvTextToRowArrays(csv);
    expect(rows.length).toBe(3);
    expect(rows[1]![1]).toContain("line1");
    expect(rows[1]![1]).toContain("line2");
  });
});

describe("enrichScheduleMxRowsWithLatestEtic", () => {
  it("waives overdue when latest ETIC shows an open WO", () => {
    const base = analyzeElmsScheduleMxFromRaw({}, "2026-04-25");
    const row = {
      assetId: "AF01",
      planRowKey: "p1",
      planId: "",
      planName: "Oil",
      planDesc: "",
      maintenanceScheduleId: "",
      itemDesc: "",
      location: "",
      makeModel: "",
      mgmtCd: "",
      workOrderCount: 0,
      nce: false,
      nceStatus: "",
      scheduleMxNceCritical: true,
      owningUnit: "",
      vehNomen: "",
      eticSnapshotDateKey: null,
      eticOpenWorkOrderIds: "",
      eticOpenInMaintenance: false,
      scheduleMxPlanEffectiveBucket: "overdue",
      scheduleMxPlanEffectiveNceCritical: true,
      scheduleMxSuppressedByOpenWo: false,
      ...base,
      scheduleMxBucket: "overdue",
    } as ScheduleMxFleetRow;
    const byAsset = new Map([
      [
        "AF01",
        {
          workOrderIds: ["WO-1"],
          owningUnit: "1 AMXS",
          makeModel: "F-150",
          vehNomen: "TRUCK",
          mgmtCd: "M16",
          nce: true,
          nceStatus: "Active",
          inMaintenance: false,
        },
      ],
    ]);
    const out = enrichScheduleMxRowsWithLatestEtic([row], "2026-04-26", byAsset)[0]!;
    expect(out.scheduleMxPlanEffectiveBucket).toBe("ok");
    expect(out.scheduleMxSuppressedByOpenWo).toBe(true);
    expect(out.workOrderCount).toBe(1);
    expect(out.owningUnit).toBe("1 AMXS");
    expect(out.scheduleMxPlanEffectiveNceCritical).toBe(false);
  });
});

describe("computeScheduleMxAssetStats", () => {
  it("counts each asset once using worst plan (not plan-row totals)", () => {
    const base = analyzeElmsScheduleMxFromRaw({}, "2026-04-25");
    const mk = (
      assetId: string,
      key: string,
      bucket: "overdue" | "ok" | "due_soon" | "missing" | "no_due",
      nceCrit?: boolean,
    ) => {
      const nc = !!nceCrit;
      return {
        assetId,
        planRowKey: key,
        planId: "",
        planName: "",
        planDesc: "",
        maintenanceScheduleId: "",
        itemDesc: "",
        location: "",
        makeModel: "",
        mgmtCd: "",
        workOrderCount: 0,
        nce: false,
        nceStatus: "",
        scheduleMxNceCritical: nc,
        owningUnit: "",
        vehNomen: "",
        eticSnapshotDateKey: null,
        eticOpenWorkOrderIds: "",
        eticOpenInMaintenance: false,
        scheduleMxPlanEffectiveBucket: bucket,
        scheduleMxPlanEffectiveNceCritical: nc,
        scheduleMxSuppressedByOpenWo: false,
        ...base,
        scheduleMxBucket: bucket,
      } as ScheduleMxFleetRow;
    };

    const rows = [
      mk("AF01", "p1", "ok"),
      mk("AF01", "p2", "overdue"),
      mk("AF02", "p3", "ok"),
      mk("AF02", "p4", "ok"),
    ];
    const s = computeScheduleMxAssetStats(rows);
    expect(s.distinctAssets).toBe(2);
    expect(s.planRows).toBe(4);
    expect(s.overdue).toBe(1);
    expect(s.ok).toBe(1);
    expect(s.dueSoon).toBe(0);
    expect(s.missing).toBe(0);
  });
});

describe("computeScheduleMxCommanderSummary", () => {
  it("rolls up by owning unit and counts NCE overdue separately", () => {
    const base = analyzeElmsScheduleMxFromRaw({}, "2026-04-25");
    const row = (
      assetId: string,
      unit: string,
      bucket: "overdue" | "ok",
      nceCrit: boolean,
      key: string,
    ): ScheduleMxFleetRow =>
      ({
        assetId,
        planRowKey: key,
        planId: "",
        planName: "",
        planDesc: "",
        maintenanceScheduleId: "",
        itemDesc: "",
        location: "",
        makeModel: "",
        mgmtCd: "",
        workOrderCount: 0,
        nce: false,
        nceStatus: "",
        scheduleMxNceCritical: nceCrit,
        owningUnit: unit,
        vehNomen: "",
        eticSnapshotDateKey: null,
        eticOpenWorkOrderIds: "",
        eticOpenInMaintenance: false,
        scheduleMxPlanEffectiveBucket: bucket,
        scheduleMxPlanEffectiveNceCritical: nceCrit,
        scheduleMxSuppressedByOpenWo: false,
        ...base,
        scheduleMxBucket: bucket,
      }) as ScheduleMxFleetRow;

    const rows = [
      row("A1", "91 MW", "ok", false, "p1"),
      row("A1", "91 MW", "ok", false, "p2"),
      row("B1", "5 CES", "overdue", false, "p3"),
      row("C1", "5 CES", "ok", false, "p4"),
      row("D1", "791 MXS", "overdue", true, "p5"),
    ];
    const c = computeScheduleMxCommanderSummary(rows);
    expect(c.wing.totalVehicles).toBe(4);
    expect(c.wing.overdue).toBe(2);
    expect(c.wing.nceOverdue).toBe(1);
    expect(c.wing.notOverdue).toBe(2);
    const u91 = c.units.find((x) => x.unit === "91 MW");
    expect(u91?.totalVehicles).toBe(1);
    expect(u91?.overdue).toBe(0);
    const u5 = c.units.find((x) => x.unit === "5 CES");
    expect(u5?.totalVehicles).toBe(2);
    expect(u5?.overdue).toBe(1);
    const u791 = c.units.find((x) => x.unit === "791 MXS");
    expect(u791?.nceOverdue).toBe(1);
    expect(u791?.overdue).toBe(1);
  });
});

describe("parseScheduleMxCsvToPlanRows", () => {
  it("maps headers to fleet.* keys and yields one row per CSV line", () => {
    const csv =
      "Asset Id,Maintenance Schedule Id,Plan Name,Next Maint Date\n" +
      "AF01B00001,SCH1,Oil,2026-05-01\n" +
      "AF01B00001,SCH2,Brakes,2026-06-01\n";
    const rows = parseScheduleMxCsvToPlanRows(csv);
    expect(rows.length).toBe(2);
    expect(rows[0]!.planRowKey).toBe("SCH1");
    expect(rows[1]!.planRowKey).toBe("SCH2");
    expect(rows[0]!.raw["fleet.next maint date"]).toBe("2026-05-01");
  });

  it("suffixes duplicate plan_row_key so D1 upserts do not collapse rows", () => {
    const csv =
      "Asset Id,Maintenance Schedule Id,Plan Name\n" +
      "AF01B00001,SAME,A\n" +
      "AF01B00001,SAME,B\n";
    const rows = parseScheduleMxCsvToPlanRows(csv);
    expect(rows.length).toBe(2);
    expect(rows[0]!.planRowKey).toBe("SAME");
    expect(rows[1]!.planRowKey).toBe("SAME#1");
  });

  it("collapses duplicate asset+plan id+plan name when one row has bogus schedule id", () => {
    const csv =
      "Asset Id,Maintenance Schedule Id,Plan Id,Plan Name\n" +
      "AF04L00057,AF04L00057,35AA,INSP/ WHEEL BEARINGS\n" +
      "AF04L00057,DF-VEHS5 LRS MINOT294497,35AA,INSP/ WHEEL BEARINGS\n";
    const rows = parseScheduleMxCsvToPlanRows(csv);
    expect(rows.length).toBe(1);
    expect(rows[0]!.planRowKey).toContain("MINOT");
  });

  it("strips UTF-8 BOM on first header", () => {
    const csv = "\uFEFFVehicle Id,Plan Id,Plan Name,Next Maint Date\nAF01B00001,P1,Oil,2026-05-01\n";
    const rows = parseScheduleMxCsvToPlanRows(csv);
    expect(rows.length).toBe(1);
    expect(rows[0]!.planRowKey).toContain("AF01B00001");
    expect(rows[0]!.raw["fleet.next maint date"]).toBe("2026-05-01");
  });
});

describe("elmsPlanRowKeyFromRaw", () => {
  it("does not use maintenance schedule cell when it equals asset id", () => {
    const raw = {
      "fleet.maintenance schedule id": "AF04L00057",
      "fleet.plan id": "35AA",
      "fleet.plan name": "Wheel",
    };
    const k = elmsPlanRowKeyFromRaw(raw, "AF04L00057", 0);
    expect(k).toContain("35AA");
    expect(k.toUpperCase()).not.toBe("AF04L00057");
  });
});

describe("parseScheduleMxCsvToRawByAsset (compat)", () => {
  it("keeps last row per asset only", () => {
    const csv =
      "Asset Id,Plan Id,Plan Name\n" +
      "AF01B00001,P1,A\n" +
      "AF01B00001,P2,B\n";
    const m = parseScheduleMxCsvToRawByAsset(csv);
    expect(m.get("AF01B00001")!["fleet.plan name"]).toBe("B");
  });
});

describe("analyzeElmsScheduleMxFromRaw", () => {
  it("marks overdue when next maint date is before as-of (no util interval to contradict)", () => {
    const raw = {
      "fleet.next maint date": "2026-04-01",
    };
    const a = analyzeElmsScheduleMxFromRaw(raw, "2026-04-25");
    expect(a.scheduleMxBucket).toBe("overdue");
    expect(a.scheduleMxOverdueByDays).toBe(24);
  });

  it("marks util overdue when current meter past next util", () => {
    const raw = {
      "fleet.next maint date": "2026-12-01",
      "fleet.current meter reading": "6000",
      "fleet.next util qty": "5000",
    };
    const a = analyzeElmsScheduleMxFromRaw(raw, "2026-04-25");
    expect(a.scheduleMxOverdueUtil).toBe(true);
    expect(a.scheduleMxBucket).toBe("overdue");
  });

  it("does not mark overdue from legacy slicer when ELMS util interval is clearly not due", () => {
    const raw = {
      "fleet.schedule mx slicer": "Overdue",
      "fleet.current meter reading": "1000",
      "fleet.next util qty": "5000",
    };
    const a = analyzeElmsScheduleMxFromRaw(raw, "2026-04-25");
    expect(a.scheduleMxBucket).toBe("ok");
  });

  it("trusts meter over a stale parsed next-maint date when still before next util qty", () => {
    const raw = {
      "fleet.next maint date": "2026-04-01",
      "fleet.current meter reading": "1000",
      "fleet.next util qty": "5000",
    };
    const a = analyzeElmsScheduleMxFromRaw(raw, "2026-04-25");
    expect(a.scheduleMxBucket).toBe("ok");
    expect(a.scheduleMxOverdueUtil).toBe(false);
  });
});

describe("calendarDaysBetween", () => {
  it("counts whole days", () => {
    expect(calendarDaysBetween("2026-04-01", "2026-04-04")).toBe(3);
  });
});

describe("melMgmtCodesMatch", () => {
  it("matches B216 style codes inside longer auth strings", () => {
    expect(melMgmtCodesMatch("Auth B216 / foo", "B216")).toBe(true);
    expect(melMgmtCodesMatch("b216", "B216")).toBe(true);
  });
  it("returns false when either side empty", () => {
    expect(melMgmtCodesMatch("", "B216")).toBe(false);
    expect(melMgmtCodesMatch("B216", "")).toBe(false);
  });
});

function minimalWatchRow(overrides: Partial<WatchRow>): WatchRow {
  return {
    workOrderId: "WO-1",
    assetId: "AF1",
    melTier: "above",
    partsStatus: "",
    eticRaw: "",
    eticDate: null,
    remarks: "",
    lastRemarkChangeDate: "2026-04-01",
    daysSinceRemarkChange: 0,
    requiredIntervalDays: 5,
    remarkStale: false,
    eticPushCount: 0,
    cumulativeEticSlipDays: 0,
    firstEticDate: null,
    lastEticDate: null,
    lastSnapshotDate: "2026-04-01",
    owningUnit: "",
    melKey: "MEL-A",
    shop: "",
    mgmtCd: "B216",
    makeModel: "",
    vehNomen: "",
    firstSeenDate: "",
    historyBounded: false,
    establishedDate: "",
    establishedDateIso: null,
    woReason: "",
    nce: false,
    nceStatus: "",
    scheduleMxStatus: "",
    scheduleMxDueIso: null,
    scheduleMxDaysUntil: null,
    scheduleMxOverdueByDays: null,
    scheduleMxBucket: "ok",
    scheduleMxNeedsEntry: false,
    daysDown: null,
    ...overrides,
  };
}

describe("computeMelRecallHintForRow", () => {
  it("returns a hint when another above-MEL line matches asset mgmt and donor has FMC cushion", () => {
    const melRows = [
      {
        mel_key: "MEL-A",
        unit: "U1",
        mgmt_code_name: "X100",
        fmc_count: 10,
        nmc_count: 2,
        mel_required: 6,
        mel_status: "above",
      },
      {
        mel_key: "MEL-B",
        unit: "U2",
        mgmt_code_name: "Auth B216",
        fmc_count: 8,
        nmc_count: 1,
        mel_required: 5,
        mel_status: "above",
      },
    ];
    const row = minimalWatchRow({ melKey: "MEL-A", mgmtCd: "B216", melTier: "above" });
    const hint = computeMelRecallHintForRow(melRows, row);
    expect(hint).not.toBeNull();
    expect(hint?.donorMelKey).toBe("MEL-A");
    expect(hint?.supporterMelKey).toBe("MEL-B");
    expect(hint?.otherNmcOnDonorMel).toBe(1);
  });

  it("returns null when WO is not above MEL tier", () => {
    const melRows = [
      {
        mel_key: "MEL-A",
        unit: "U1",
        mgmt_code_name: "X",
        fmc_count: 10,
        nmc_count: 2,
        mel_required: 6,
        mel_status: "above",
      },
      {
        mel_key: "MEL-B",
        unit: "U2",
        mgmt_code_name: "B216",
        fmc_count: 8,
        nmc_count: 1,
        mel_required: 5,
        mel_status: "above",
      },
    ];
    const row = minimalWatchRow({ melTier: "at" });
    expect(computeMelRecallHintForRow(melRows, row)).toBeNull();
  });

  it("returns null when donor FMC cushion is too thin", () => {
    const melRows = [
      {
        mel_key: "MEL-A",
        unit: "U1",
        mgmt_code_name: "X",
        fmc_count: 7,
        nmc_count: 1,
        mel_required: 6,
        mel_status: "above",
      },
      {
        mel_key: "MEL-B",
        unit: "U2",
        mgmt_code_name: "B216",
        fmc_count: 8,
        nmc_count: 1,
        mel_required: 5,
        mel_status: "above",
      },
    ];
    const row = minimalWatchRow({ melKey: "MEL-A" });
    expect(computeMelRecallHintForRow(melRows, row)).toBeNull();
  });
});

type MockRunResult = { meta: { changes: number } };

class MockD1PreparedStatement {
  constructor(
    private readonly db: MockD1Database,
    private readonly sql: string,
    private readonly params: unknown[] = [],
  ) {}

  bind(...params: unknown[]): MockD1PreparedStatement {
    return new MockD1PreparedStatement(this.db, this.sql, params);
  }

  async all<T>(): Promise<{ results: T[] }> {
    return { results: this.db.all<T>(this.sql, this.params) };
  }

  async run(): Promise<MockRunResult> {
    return { meta: { changes: this.db.run(this.sql, this.params) } };
  }
}

class MockD1Database {
  readonly snapshots: Map<string, Map<string, Record<string, unknown>>> = new Map();
  readonly changelog: Array<Record<string, unknown>> = [];
  readonly state: Map<string, Record<string, unknown>> = new Map();

  prepare(sql: string): MockD1PreparedStatement {
    return new MockD1PreparedStatement(this, sql);
  }

  async batch(statements: MockD1PreparedStatement[]): Promise<MockRunResult[]> {
    return Promise.all(statements.map((s) => s.run()));
  }

  all<T>(sql: string, params: unknown[]): T[] {
    if (sql.includes("FROM work_order_snapshot") && sql.includes("WHERE work_order_id = ?")) {
      const workOrderId = String(params[0]);
      const rows: Record<string, unknown>[] = [];
      for (const [snapshotDate, byWo] of this.snapshots) {
        const row = byWo.get(workOrderId);
        if (row) rows.push({ ...row, snapshot_date_key: snapshotDate });
      }
      rows.sort((a, b) => String(a.snapshot_date_key).localeCompare(String(b.snapshot_date_key)));
      return rows as T[];
    }
    if (sql.includes("FROM work_order_snapshot") && sql.includes("WHERE work_order_id IN")) {
      const dateKey = String(params[params.length - 1]);
      const workOrderIds = new Set(params.slice(0, -1).map(String));
      const rows: Record<string, unknown>[] = [];
      for (const [snapshotDate, byWo] of this.snapshots) {
        if (snapshotDate > dateKey) continue;
        for (const [wid, row] of byWo) {
          if (!workOrderIds.has(wid)) continue;
          rows.push({ ...row, last_snapshot_date: snapshotDate });
        }
      }
      rows.sort((a, b) => {
        const wid = String(a.work_order_id).localeCompare(String(b.work_order_id));
        if (wid !== 0) return wid;
        return String(b.last_snapshot_date).localeCompare(String(a.last_snapshot_date));
      });
      return rows as T[];
    }
    return [];
  }

  run(sql: string, params: unknown[]): number {
    if (sql.includes("INSERT OR REPLACE INTO work_order_changelog")) {
      const [work_order_id, snapshot_date_key, changed_at_iso, field, old_value, new_value] = params;
      const existing = this.changelog.findIndex(
        (row) =>
          row.work_order_id === work_order_id &&
          row.snapshot_date_key === snapshot_date_key &&
          row.field === field,
      );
      const row = { work_order_id, snapshot_date_key, changed_at_iso, field, old_value, new_value };
      if (existing >= 0) this.changelog[existing] = row;
      else this.changelog.push(row);
      return 1;
    }
    if (sql.includes("INSERT INTO work_order_snapshot")) {
      const row = this.workOrderRowFromParams(params);
      const dateKey = String(params[0]);
      const wid = String(params[1]);
      if (!this.snapshots.has(dateKey)) this.snapshots.set(dateKey, new Map());
      this.snapshots.get(dateKey)?.set(wid, row);
      return 1;
    }
    if (sql.includes("INSERT INTO work_order_state")) {
      const row = this.workOrderRowFromParams(params.slice(1));
      const wid = String(params[0]);
      const prev = this.state.get(wid);
      if (!prev || String(row.last_snapshot_date) >= String(prev.last_snapshot_date)) this.state.set(wid, row);
      return 1;
    }
    return 0;
  }

  private workOrderRowFromParams(params: unknown[]): Record<string, unknown> {
    return {
      snapshot_date_key: params[0],
      work_order_id: params[1],
      asset_id: params[2],
      remarks: params[3],
      parts_status: params[4],
      etic_raw: params[5],
      etic_date: params[6],
      mel_tier: params[7],
      last_remark_change_date: params[8],
      etic_push_count: params[9],
      first_etic_date: params[10],
      last_etic_date: params[11],
      cumulative_etic_slip_days: params[12],
      owning_unit: params[13],
      mel_key: params[14],
      shop: params[15],
      mgmt_cd: params[16],
      make_model: params[17],
      veh_nomen: params[18],
      raw_row_json: params[19],
      last_snapshot_date: params[0],
    };
  }
}

function wo(overrides: Partial<RawWorkOrder>): RawWorkOrder {
  return {
    assetId: "",
    workOrderId: "",
    remarks: "",
    shop: "",
    shop2: "",
    etiCLocation: "",
    makeModel: "",
    partsStatus: "",
    eticDue: "",
    currentMel: "",
    owningUnit: "",
    melKey: "",
    mgmtCd: "",
    vehNomen: "",
    rawColumns: {},
    ...overrides,
  };
}

describe("ingestWorkOrderSnapshot", () => {
  it("builds date-by-date changelog rows from each sequential ingest", async () => {
    const db = new MockD1Database();
    const env = { ETIC_SNAPSHOTS: db as unknown as D1Database };

    await ingestWorkOrderSnapshot(env, "2026-04-01", [
      wo({ workOrderId: "WO-1", assetId: "AF1", remarks: "Initial", eticDue: "4/5/2026", currentMel: "Below MEL" }),
    ], "2026-04-01T18:00:00.000Z");
    await ingestWorkOrderSnapshot(env, "2026-04-02", [
      wo({ workOrderId: "WO-1", assetId: "AF1", remarks: "Parts ordered", eticDue: "4/5/2026", currentMel: "Below MEL" }),
    ], "2026-04-02T18:00:00.000Z");
    await ingestWorkOrderSnapshot(env, "2026-04-03", [
      wo({ workOrderId: "WO-1", assetId: "AF1", remarks: "Parts ordered", eticDue: "4/7/2026", currentMel: "Below MEL" }),
    ], "2026-04-03T18:00:00.000Z");

    expect([...db.snapshots.keys()]).toEqual(["2026-04-01", "2026-04-02", "2026-04-03"]);
    expect(db.changelog.map((row) => [row.snapshot_date_key, row.field, row.old_value, row.new_value])).toEqual([
      ["2026-04-01", "initial", "", "first_seen"],
      ["2026-04-02", "remarks", "Initial", "Parts ordered"],
      ["2026-04-03", "etic", "4/5/2026", "4/7/2026"],
      ["2026-04-03", "etic_date_slip", "2026-04-05", "2026-04-07"],
    ]);
  });

  it("derives display changelog from snapshots when persisted changelog has stale duplicate initial rows", async () => {
    const db = new MockD1Database();
    const env = { ETIC_SNAPSHOTS: db as unknown as D1Database };

    await ingestWorkOrderSnapshot(env, "2026-02-26", [
      wo({ workOrderId: "WO-1", assetId: "AF1", remarks: "Initial", currentMel: "Below MEL" }),
    ], "2026-02-26T18:00:00.000Z");
    await ingestWorkOrderSnapshot(env, "2026-04-16", [
      wo({ workOrderId: "WO-1", assetId: "AF1", remarks: "Initial", currentMel: "Below MEL" }),
    ], "2026-04-16T18:00:00.000Z");

    db.changelog.push({
      work_order_id: "WO-1",
      snapshot_date_key: "2026-04-16",
      changed_at_iso: "2026-04-16T18:00:00.000Z",
      field: "initial",
      old_value: "",
      new_value: "first_seen",
    });

    const display = await getChangelogForDisplay(env, "WO-1");

    expect(display.filter((row) => row.field === "initial").map((row) => row.snapshot_date_key)).toEqual(["2026-02-26"]);
  });
});

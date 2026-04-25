import { describe, expect, it } from "vitest";
import { getRollingRoster, getYardRosterForDate, listOpenFindings, recordCheck } from "../src/yardSession";

type Row = Record<string, unknown>;

class MemoryStmt {
  private binds: unknown[] = [];

  constructor(private db: MemoryD1, private sql: string) {}

  bind(...values: unknown[]) {
    this.binds = values;
    return this;
  }

  async first<T>(): Promise<T | null> {
    const rows = this.db.query(this.sql, this.binds);
    return (rows[0] as T | undefined) ?? null;
  }

  async all<T>(): Promise<{ results: T[] }> {
    return { results: this.db.query(this.sql, this.binds) as T[] };
  }

  async run(): Promise<{ meta: { changes: number } }> {
    this.db.query(this.sql, this.binds);
    return { meta: { changes: 1 } };
  }
}

class MemoryD1 {
  checks: Row[] = [];
  fleetRows: Row[] = [
    {
      asset_id: "AF123",
      owning_unit: "OPS",
      shop: "Fleet Shop",
      mgmt_cd: "MGT",
      make_model: "Ford",
      veh_nomen: "Truck",
      mel_key: "MEL-A",
      raw_row_json: JSON.stringify({ "fleet.location": "Fleet lot", "fleet.vin": "VIN123" }),
    },
    {
      asset_id: "AF999",
      owning_unit: "LRS",
      shop: "Line",
      mgmt_cd: "M2",
      make_model: "Chevy",
      veh_nomen: "Van",
      mel_key: "MEL-B",
      raw_row_json: JSON.stringify({ "fleet.location": "Hangar 2", "fleet.vin": "VIN999" }),
    },
  ];
  workOrderRows: Row[] = [
    {
      asset_id: "AF123",
      owning_unit: "OPS",
      shop: "Bay 1",
      mgmt_cd: "MGT",
      make_model: "Ford",
      veh_nomen: "Truck",
      mel_key: "MEL-A",
      mel_tier: "below",
      raw_row_json: JSON.stringify({ Location: "Old lot", VIN: "VIN123" }),
    },
  ];
  nextId = 1;

  prepare(sql: string) {
    return new MemoryStmt(this, sql);
  }

  query(sql: string, binds: unknown[]): Row[] {
    if (sql.includes("FROM app_config")) return [];
    if (sql.includes("FROM etic_snapshots") && sql.includes("ORDER BY e.date_key DESC")) {
      return [{ date_key: "2026-04-02" }];
    }
    if (sql.includes("FROM etic_snapshots") && sql.includes("ORDER BY date_key DESC LIMIT 1")) {
      return [{ date_key: "2026-04-02" }];
    }
    if (sql.includes("SELECT 1 AS x FROM work_order_snapshot")) return [];
    if (sql.includes("FROM fleet_p_a_snapshot") && sql.includes("snapshot_date_key = ?")) {
      return this.fleetRows;
    }
    if (sql.includes("FROM work_order_snapshot") && sql.includes("snapshot_date_key = ?")) {
      if (sql.includes("COUNT(*)")) {
        const aid = String(binds[1] ?? "").trim().toUpperCase();
        const n = this.workOrderRows.filter((row) => String(row.asset_id ?? "").trim().toUpperCase() === aid).length;
        return [{ c: n }];
      }
      if (sql.includes("mel_tier = 'below'")) return [{ asset_id: "AF123" }];
      if (sql.includes("SELECT DISTINCT asset_id")) return this.workOrderRows.map((row) => ({ asset_id: row.asset_id }));
      return this.workOrderRows;
    }
    if (sql.includes("FROM yard_check") && sql.includes("ROW_NUMBER() OVER")) {
      return this.checks;
    }
    if (sql.includes("SELECT yc.* FROM yard_check yc")) {
      return this.checks;
    }
    if (sql.includes("FROM yard_finding_action")) return [];
    if (sql.includes("FROM yard_check yc") && sql.includes("COALESCE(location")) {
      return this.checks.filter((row) => String(row.location ?? "").trim()).map((row) => ({
        asset_id: row.asset_id,
        location: row.location,
      }));
    }
    if (sql.includes("FROM yard_photo")) return [];
    if (sql.includes("FROM etic_snapshots") && sql.includes("date_key < ?")) {
      return [];
    }
    if (sql.includes("INSERT INTO yard_check")) {
      const row: Row = {
        id: this.nextId++,
        asset_id: binds[0],
        location: binds[1],
        discrepancies: binds[2],
        status: binds[3],
        checked_by: binds[4],
        checked_at_iso: binds[5],
        source_date_key: binds[6],
        asset_snapshot_json: binds[7] ?? "",
      };
      this.checks.push(row);
      return [row];
    }
    throw new Error(`unexpected SQL: ${sql}`);
  }
}

describe("recordCheck", () => {
  it("stores the source snapshot key and frozen asset context", async () => {
    const db = new MemoryD1();
    const env = { ETIC_SNAPSHOTS: db, ETIC_BUCKET: {} };

    const check = await recordCheck(env as never, {
      assetId: "AF123",
      location: "Lot C",
      checkedBy: "Walker",
    });

    expect(check.sourceDateKey).toBe("2026-04-02");
    expect(check.assetSnapshotJson).toContain('"sourceDateKey":"2026-04-02"');
    expect(check.assetSnapshotJson).toContain('"assetId":"AF123"');
    expect(check.assetSnapshotJson).toContain('"previousLocation":"Old lot"');
  });
});

describe("Yard roster", () => {
  it("uses Fleet P&A as the searchable base and overlays open WO context", async () => {
    const db = new MemoryD1();
    const env = { ETIC_SNAPSHOTS: db, ETIC_BUCKET: {} };

    const roster = await getYardRosterForDate(env as never, "2026-04-02");

    expect(roster.assets.map((asset) => asset.assetId)).toEqual(["AF123", "AF999"]);
    const openWoAsset = roster.assets.find((asset) => asset.assetId === "AF123");
    expect(openWoAsset?.openWoCount).toBe(1);
    expect(openWoAsset?.melTier).toBe("below");
    expect(openWoAsset?.previousLocation).toBe("Old lot");
    const fleetOnlyAsset = roster.assets.find((asset) => asset.assetId === "AF999");
    expect(fleetOnlyAsset?.openWoCount).toBe(0);
    expect(fleetOnlyAsset?.previousLocation).toBe("Hangar 2");
  });

  it("keeps checked non-fleet assets out of the Fleet list", async () => {
    const db = new MemoryD1();
    const env = { ETIC_SNAPSHOTS: db, ETIC_BUCKET: {} };
    await recordCheck(env as never, { assetId: "AF-GHOST", location: "Lot Z", checkedBy: "Walker" });

    const roster = await getRollingRoster(env as never);

    expect(roster.assets.map((asset) => asset.assetId)).toEqual(["AF123", "AF999"]);
    expect(roster.assets.some((asset) => asset.assetId === "AF-GHOST")).toBe(false);
  });

  it("counts due checks only for latest open-WO assets older than the interval", async () => {
    const db = new MemoryD1();
    const env = { ETIC_SNAPSHOTS: db, ETIC_BUCKET: {} };
    await recordCheck(env as never, { assetId: "AF999", location: "Hangar 2", checkedBy: "Walker" });

    const roster = await getRollingRoster(env as never);
    const openWo = roster.assets.find((asset) => asset.assetId === "AF123");
    const fleetOnly = roster.assets.find((asset) => asset.assetId === "AF999");

    expect(openWo?.openWoCount).toBe(1);
    expect(openWo?.rollingState).toBe("never");
    expect(fleetOnly?.openWoCount).toBe(0);
    expect(fleetOnly?.rollingState).toBe("fresh");
    expect(roster.totals.never + roster.totals.due + roster.totals.overdue).toBe(1);
    expect(roster.totals.fresh).toBe(1);
  });

  it("does not use ETIC/Fleet locations as current yard locations", async () => {
    const db = new MemoryD1();
    const env = { ETIC_SNAPSHOTS: db, ETIC_BUCKET: {} };

    const roster = await getRollingRoster(env as never);
    const openWo = roster.assets.find((asset) => asset.assetId === "AF123");
    const fleetOnly = roster.assets.find((asset) => asset.assetId === "AF999");

    expect(openWo?.previousLocation).toBe("Old lot");
    expect(openWo?.lastLocation).toBe("");
    expect(fleetOnly?.previousLocation).toBe("Hangar 2");
    expect(fleetOnly?.lastLocation).toBe("");
  });

  it("flags checked assets without latest open WOs as unlisted findings", async () => {
    const db = new MemoryD1();
    const env = { ETIC_SNAPSHOTS: db, ETIC_BUCKET: {} };
    await recordCheck(env as never, { assetId: "AF999", location: "Hangar 2", checkedBy: "Walker" });
    await recordCheck(env as never, { assetId: "AF-GHOST", location: "Lot Z", checkedBy: "Walker" });

    const result = await listOpenFindings(env as never);
    const unlistedIds = result.findings
      .filter((finding) => finding.kind === "unlisted")
      .map((finding) => finding.assetId)
      .sort();

    expect(unlistedIds).toEqual(["AF-GHOST", "AF999"]);
    expect(result.totals.unlisted).toBe(2);
  });

  it("does not flag unlisted when the walker sighting was under an open WO per check snapshot", async () => {
    const db = new MemoryD1();
    db.workOrderRows = [];
    db.checks = [
      {
        id: 1,
        asset_id: "AF123",
        location: "Lot C",
        discrepancies: "",
        status: "present",
        checked_by: "Walker",
        checked_at_iso: "2026-04-10T12:00:00.000Z",
        source_date_key: "2026-04-02",
        asset_snapshot_json: JSON.stringify({
          sourceDateKey: "2026-04-02",
          asset: { assetId: "AF123", openWoCount: 1, melTier: "below" },
        }),
      },
    ];
    const env = { ETIC_SNAPSHOTS: db, ETIC_BUCKET: {} };
    const result = await listOpenFindings(env as never);
    expect(result.findings.some((f) => f.kind === "unlisted" && f.assetId === "AF123")).toBe(false);
  });

  it("does not flag unlisted when snapshot JSON is empty but source_date_key WO count > 0", async () => {
    const db = new MemoryD1();
    db.checks = [
      {
        id: 1,
        asset_id: "AF123",
        location: "Lot C",
        discrepancies: "",
        status: "present",
        checked_by: "Walker",
        checked_at_iso: "2026-04-10T12:00:00.000Z",
        source_date_key: "2026-04-02",
        asset_snapshot_json: "",
      },
    ];
    const env = { ETIC_SNAPSHOTS: db, ETIC_BUCKET: {} };
    const result = await listOpenFindings(env as never);
    expect(result.findings.some((f) => f.kind === "unlisted" && f.assetId === "AF123")).toBe(false);
  });
});

import { describe, expect, it } from "vitest";
import { recordCheck } from "../src/yardSession";

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
  nextId = 1;

  prepare(sql: string) {
    return new MemoryStmt(this, sql);
  }

  query(sql: string, binds: unknown[]): Row[] {
    if (sql.includes("FROM app_config")) return [];
    if (sql.includes("FROM etic_snapshots") && sql.includes("ORDER BY date_key DESC LIMIT 1")) {
      return [{ date_key: "2026-04-02" }];
    }
    if (sql.includes("SELECT 1 AS x FROM work_order_snapshot")) return [];
    if (sql.includes("FROM work_order_snapshot") && sql.includes("snapshot_date_key = ?")) {
      return [
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

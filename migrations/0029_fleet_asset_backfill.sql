-- Seed fleet_asset_current from latest work_order_state row per asset (one-time backfill).

INSERT OR REPLACE INTO fleet_asset_current (
  asset_id, owning_unit, shop, make_model, veh_nomen, mgmt_cd, mel_key, last_seen_snapshot_date, updated_at_iso
)
SELECT asset_id, owning_unit, shop, make_model, veh_nomen, mgmt_cd, mel_key, last_snapshot_date, updated_at_iso
FROM (
  SELECT
    asset_id,
    owning_unit,
    shop,
    make_model,
    veh_nomen,
    mgmt_cd,
    mel_key,
    last_snapshot_date,
    updated_at_iso,
    ROW_NUMBER() OVER (
      PARTITION BY asset_id
      ORDER BY datetime(updated_at_iso) DESC, work_order_id DESC
    ) AS rn
  FROM work_order_state
  WHERE trim(COALESCE(asset_id, '')) != ''
) t
WHERE rn = 1;

-- Backfill work_order_id on open abuse cases from latest work_order_state per asset.

UPDATE abuse_tracker_case
SET work_order_id = (
  SELECT w.work_order_id
  FROM work_order_state w
  WHERE w.asset_id = abuse_tracker_case.asset_id
  ORDER BY datetime(w.updated_at_iso) DESC
  LIMIT 1
)
WHERE abuse_tracker_case.closed_at_iso IS NULL
  AND trim(abuse_tracker_case.work_order_id) = ''
  AND EXISTS (
    SELECT 1 FROM work_order_state w2 WHERE w2.asset_id = abuse_tracker_case.asset_id
  );

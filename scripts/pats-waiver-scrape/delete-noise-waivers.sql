-- Remove placeholder PATS rows ("No waiver information", etc.).
-- Same predicate as WAIVER_ROW_IS_NOISE_SQL in src/waivers.ts.
-- waiver_verification rows cascade when import exists on FK (see migration).

DELETE FROM waiver
WHERE (
  lower(trim(ifnull(title, ''))) LIKE '%no waiver information%'
  OR lower(trim(ifnull(description, ''))) LIKE '%no waiver information%'
  OR lower(trim(ifnull(title, ''))) LIKE '%no waivered items%'
  OR lower(trim(ifnull(description, ''))) LIKE '%no waivered items%'
  OR lower(trim(ifnull(title, ''))) LIKE '%no waiverable items%'
  OR lower(trim(ifnull(description, ''))) LIKE '%no waiverable items%'
  OR (trim(ifnull(title, '')) = '' AND trim(ifnull(description, '')) = '')
);

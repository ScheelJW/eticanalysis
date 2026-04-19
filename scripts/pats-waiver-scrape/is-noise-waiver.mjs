/**
 * PATS placeholder / empty-card rows — do not import or scrape into CSV.
 */
export function isNoiseWaiverText(s) {
  const t = String(s ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
  if (!t) return true;
  if (t.includes("no waiver information")) return true;
  if (t.includes("no waivered items")) return true;
  if (t.includes("no waiverable items")) return true;
  return false;
}

/** Owning org for ELMS/ETIC assets whose id starts with MSX (case-insensitive). */
export const MSX_FORCED_OWNING_UNIT = "791 MXS";

/** All asset ids starting with MSX are displayed and rolled up as 791 MXS. */
export function owningUnitForAssetId(assetId: string, extractUnit: string): string {
  const aid = (assetId ?? "").trim();
  if (aid.toUpperCase().startsWith("MSX")) return MSX_FORCED_OWNING_UNIT;
  return (extractUnit ?? "").trim();
}

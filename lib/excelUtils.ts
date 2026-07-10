export function cellText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && "text" in (value as Record<string, unknown>)) {
    return String((value as { text: unknown }).text ?? "");
  }
  if (typeof value === "object" && "result" in (value as Record<string, unknown>)) {
    return String((value as { result: unknown }).result ?? "");
  }
  return String(value);
}

export function toNumber(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return value;
  if (typeof value === "object" && value !== null && "result" in (value as Record<string, unknown>)) {
    // ExcelJS formula cell result
    return toNumber((value as { result: unknown }).result);
  }
  const n = Number(String(value).trim());
  return Number.isFinite(n) ? n : null;
}

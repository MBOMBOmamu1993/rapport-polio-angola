export function fmtInt(n: number | null | undefined): string {
  if (n === null || n === undefined || !Number.isFinite(n)) return "0";
  return Math.round(n).toLocaleString("fr-FR");
}

export function fmtPct(n: number | null | undefined, digits = 1): string {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  return `${n.toFixed(digits).replace(".", ",")} %`;
}

export function pct(part: number, whole: number): number | null {
  if (!whole) return null;
  return (part / whole) * 100;
}

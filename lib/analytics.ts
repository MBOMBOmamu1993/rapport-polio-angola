/**
 * Agrégation des données du masque vers les indicateurs du rapport (polio uniquement).
 * Applique les filtres en cascade (Province → Antenne → ZS → AS) avec la même
 * logique que le dashboard : un niveau sélectionné restreint le périmètre, et
 * l'agrégation des graphiques se fait au niveau immédiatement inférieur.
 */

import type { ASRecord, MasqueData } from "./parse-masque";
import type { Filters } from "./store";
import { pct } from "./format";

export interface CascadeOptions {
  provinces: string[];
  antennes: string[];
  zones: string[];
  aires: string[];
}

export function cascadeOptions(data: MasqueData, f: Filters): CascadeOptions {
  const provinces = uniq(data.records.map((r) => r.province));
  const antennes = uniq(
    data.records.filter((r) => !f.province || r.province === f.province).map((r) => r.antenne)
  );
  const zones = uniq(
    data.records
      .filter((r) => (!f.province || r.province === f.province) && (!f.antenne || r.antenne === f.antenne))
      .map((r) => r.zs)
  );
  const aires = uniq(
    data.records
      .filter(
        (r) =>
          (!f.province || r.province === f.province) &&
          (!f.antenne || r.antenne === f.antenne) &&
          (!f.zs || r.zs === f.zs)
      )
      .map((r) => r.as)
  );
  return { provinces, antennes, zones, aires };
}

export function applyFilters(data: MasqueData, f: Filters): ASRecord[] {
  return data.records.filter(
    (r) =>
      (!f.province || r.province === f.province) &&
      (!f.antenne || r.antenne === f.antenne) &&
      (!f.zs || r.zs === f.zs) &&
      (!f.as || r.as === f.as)
  );
}

export type DrillLevel = "province" | "antenne" | "zs" | "as";

export function resolveDrillLevel(f: Filters): { level: DrillLevel; label: string } {
  if (f.zs) return { level: "as", label: "Aire de Santé" };
  if (f.antenne) return { level: "zs", label: "Zone de Santé" };
  if (f.province) return { level: "zs", label: "Zone de Santé" };
  // Niveau national (aucune province sélectionnée) → agrégation par province.
  return { level: "province", label: "Province" };
}

export function scopeLabel(f: Filters): string {
  if (f.as) return `Aire de Santé : ${f.as}`;
  if (f.zs) return `Zone de Santé : ${f.zs}`;
  if (f.antenne) return `Antenne : ${f.antenne}`;
  if (f.province) return `Province : ${f.province}`;
  return "Niveau national — toutes les provinces";
}

export function scopeName(f: Filters, data: MasqueData): string {
  return f.as ?? f.zs ?? f.antenne ?? f.province ?? data.meta.province ?? "Province";
}

/** Clé d'agrégation pour les graphiques par unité (selon le niveau de drill). */
function keyOf(r: ASRecord, level: DrillLevel): string {
  switch (level) {
    case "province": return r.province;
    case "antenne": return r.antenne || "—";
    case "zs": return r.zs || "—";
    case "as": return r.as || "—";
  }
}

export interface UnitAgg {
  unit: string;
  nvpo2Vacc: number;
  nvpo2Cible: number;
  vpobVacc: number;
  vpobCible: number;
  nvpo2FlaconsUtil: number;
  nvpo2Perdus: number;
  vpobFlaconsUtil: number;
  vpobPerdus: number;
  vaccAttendus: number;
  vaccRecus: number;
  recup: number;
  nvpo2ZeroDose: number;
  vpobZeroDose: number;
}

export function aggregateByUnit(records: ASRecord[], level: DrillLevel): UnitAgg[] {
  const map = new Map<string, UnitAgg>();
  for (const r of records) {
    const k = keyOf(r, level);
    let a = map.get(k);
    if (!a) {
      a = {
        unit: k,
        nvpo2Vacc: 0, nvpo2Cible: 0, vpobVacc: 0, vpobCible: 0,
        nvpo2FlaconsUtil: 0, nvpo2Perdus: 0, vpobFlaconsUtil: 0, vpobPerdus: 0,
        vaccAttendus: 0, vaccRecus: 0, recup: 0, nvpo2ZeroDose: 0, vpobZeroDose: 0,
      };
      map.set(k, a);
    }
    a.nvpo2Vacc += r.nvpo2Vacc;
    a.nvpo2Cible += r.nvpo2CibleExtrap;
    a.vpobVacc += r.vpobVacc;
    a.vpobCible += r.vpobCibleExtrap;
    a.nvpo2FlaconsUtil += r.nvpo2FlaconsUtil;
    a.nvpo2Perdus += r.nvpo2Perdus;
    a.vpobFlaconsUtil += r.vpobFlaconsUtil;
    a.vpobPerdus += r.vpobPerdus;
    a.vaccAttendus += r.vaccAttendus;
    a.vaccRecus += r.vaccRecus;
    a.recup += r.recup;
    a.nvpo2ZeroDose += r.nvpo2ZeroDose;
    a.vpobZeroDose += r.vpobZeroDose;
  }
  return Array.from(map.values()).sort((a, b) => a.unit.localeCompare(b.unit, "fr"));
}

export interface Totals {
  nvpo2Vacc: number;
  nvpo2Cible: number;
  nvpo2CV: number | null;
  vpobVacc: number;
  vpobCible: number;
  vpobCV: number | null;
  nvpo2FlaconsUtil: number;
  nvpo2TauxPerte: number | null;
  vpobFlaconsUtil: number;
  vpobTauxPerte: number | null;
  vaccAttendus: number;
  vaccRecus: number;
  completude: number | null;
  recup: number;
  nvpo2ZeroDose: number;
  vpobZeroDose: number;
  refusSignales: number;
  refusGeres: number;
  refusGeresPct: number | null;
  menagesPrevus: number;
  menagesVisites: number;
  menagesVisitesPct: number | null;
  pers15: number;
  mapiMineures: number;
  mapiGraves: number;
}

export function totals(records: ASRecord[]): Totals {
  const s = (sel: (r: ASRecord) => number) => records.reduce((acc, r) => acc + sel(r), 0);
  const nvpo2Vacc = s((r) => r.nvpo2Vacc);
  const nvpo2Cible = s((r) => r.nvpo2CibleExtrap);
  const vpobVacc = s((r) => r.vpobVacc);
  const vpobCible = s((r) => r.vpobCibleExtrap);
  const nvpo2FlaconsUtil = s((r) => r.nvpo2FlaconsUtil);
  const nvpo2Perdus = s((r) => r.nvpo2Perdus);
  const vpobFlaconsUtil = s((r) => r.vpobFlaconsUtil);
  const vpobPerdus = s((r) => r.vpobPerdus);
  const vaccAttendus = s((r) => r.vaccAttendus);
  const vaccRecus = s((r) => r.vaccRecus);
  const refusSignales = s((r) => r.refusSignales);
  const refusGeres = s((r) => r.refusGeres);
  const menagesPrevus = s((r) => r.menagesPrevus);
  const menagesVisites = s((r) => r.menagesVisites);

  void nvpo2Perdus;
  void vpobPerdus;
  return {
    nvpo2Vacc, nvpo2Cible, nvpo2CV: pct(nvpo2Vacc, nvpo2Cible),
    vpobVacc, vpobCible, vpobCV: pct(vpobVacc, vpobCible),
    nvpo2FlaconsUtil, nvpo2TauxPerte: tauxPerte(nvpo2Vacc, nvpo2FlaconsUtil, NVPO2_DOSES_PAR_FLACON),
    vpobFlaconsUtil, vpobTauxPerte: tauxPerte(vpobVacc, vpobFlaconsUtil, VPOB_DOSES_PAR_FLACON),
    vaccAttendus, vaccRecus, completude: pct(vaccRecus, vaccAttendus),
    recup: s((r) => r.recup),
    nvpo2ZeroDose: s((r) => r.nvpo2ZeroDose),
    vpobZeroDose: s((r) => r.vpobZeroDose),
    refusSignales, refusGeres, refusGeresPct: pct(refusGeres, refusSignales),
    menagesPrevus, menagesVisites, menagesVisitesPct: pct(menagesVisites, menagesPrevus),
    pers15: s((r) => r.pers15),
    mapiMineures: s((r) => r.mapiMineures),
    mapiGraves: s((r) => r.mapiGraves),
  };
}

// Doses par flacon (co-administration polio). Le modèle officiel calcule le taux
// de perte sur la base des doses : taux = 1 − vaccinés / (flacons utilisés × doses).
export const NVPO2_DOSES_PAR_FLACON = 50;
export const VPOB_DOSES_PAR_FLACON = 20;

export function tauxPerte(vacc: number, flaconsUtil: number, dosesParFlacon: number): number | null {
  if (!flaconsUtil) return null;
  return (1 - vacc / (flaconsUtil * dosesParFlacon)) * 100;
}

export interface CoverageRow {
  unit: string;
  cible: number;
  vacc: number;
  cv: number | null;
}

export interface GestionRow {
  unit: string;
  flaconsUtil: number;
  perdus: number;
  vacc: number;
  taux: number | null;
}

export function nvpo2Coverage(byUnit: UnitAgg[]): CoverageRow[] {
  return byUnit.map((a) => ({ unit: a.unit, cible: a.nvpo2Cible, vacc: a.nvpo2Vacc, cv: pct(a.nvpo2Vacc, a.nvpo2Cible) }));
}
export function vpobCoverage(byUnit: UnitAgg[]): CoverageRow[] {
  return byUnit.map((a) => ({ unit: a.unit, cible: a.vpobCible, vacc: a.vpobVacc, cv: pct(a.vpobVacc, a.vpobCible) }));
}
export function nvpo2Gestion(byUnit: UnitAgg[]): GestionRow[] {
  return byUnit.map((a) => ({
    unit: a.unit, flaconsUtil: a.nvpo2FlaconsUtil, perdus: a.nvpo2Perdus, vacc: a.nvpo2Vacc,
    taux: tauxPerte(a.nvpo2Vacc, a.nvpo2FlaconsUtil, NVPO2_DOSES_PAR_FLACON),
  }));
}
export function vpobGestion(byUnit: UnitAgg[]): GestionRow[] {
  return byUnit.map((a) => ({
    unit: a.unit, flaconsUtil: a.vpobFlaconsUtil, perdus: a.vpobPerdus, vacc: a.vpobVacc,
    taux: tauxPerte(a.vpobVacc, a.vpobFlaconsUtil, VPOB_DOSES_PAR_FLACON),
  }));
}

function uniq(arr: string[]): string[] {
  return Array.from(new Set(arr.filter(Boolean))).sort((a, b) => a.localeCompare(b, "fr"));
}

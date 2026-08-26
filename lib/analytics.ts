/**
 * Agrégation des données du masque vers les indicateurs du rapport (campagne
 * intégrée RR‑POLIO). Filtres en cascade Antenne → Zone de Santé → Aire de Santé ;
 * les tableaux du rapport sont produits au niveau immédiatement inférieur au
 * périmètre sélectionné, avec en plus une vue par Antenne quand plusieurs
 * antennes sont présentes.
 */

import {
  ANTIGENES,
  DOSES_PAR_FLACON,
  POLIO_AGE_LABELS,
  RR_AGE_LABELS,
  isRecapRow,
  type ASRecord,
  type MasqueData,
  type VaccineKey,
  type VaccineStats,
} from "./parse-masque";
import type { Filters } from "./store";
import { pct } from "./format";

export interface CascadeOptions {
  provinces: string[];
  antennes: string[];
  zones: string[];
  aires: string[];
}

function inProvinces(r: ASRecord, f: Filters): boolean {
  return f.provinces.length === 0 || f.provinces.includes(r.province);
}
function isRealUnit(r: ASRecord): boolean {
  return !isRecapRow(r.province, r.antenne, r.zs, r.as);
}

export function cascadeOptions(data: MasqueData, f: Filters): CascadeOptions {
  const records = data.records.filter(isRealUnit);
  const provinces = uniq(records.map((r) => r.province));
  const antennes = uniq(records.filter((r) => inProvinces(r, f)).map((r) => r.antenne));
  const zones = uniq(
    records.filter((r) => inProvinces(r, f) && (!f.antenne || r.antenne === f.antenne)).map((r) => r.zs)
  );
  const aires = uniq(
    records
      .filter((r) => inProvinces(r, f) && (!f.antenne || r.antenne === f.antenne) && (!f.zs || r.zs === f.zs))
      .map((r) => r.as)
  );
  return { provinces, antennes, zones, aires };
}

export function applyFilters(data: MasqueData, f: Filters): ASRecord[] {
  return data.records.filter(
    (r) =>
      isRealUnit(r) &&
      inProvinces(r, f) &&
      (!f.antenne || r.antenne === f.antenne) &&
      (!f.zs || r.zs === f.zs) &&
      (!f.as || r.as === f.as)
  );
}

export type DrillLevel = "province" | "antenne" | "zs" | "as" | "all";

/**
 * Niveau de désagrégation principal du rapport :
 *  - ZS filtrée → Aires de Santé ;
 *  - périmètre couvrant plusieurs provinces (sans antenne / ZS filtrée) →
 *    Provinces (les vues par Antenne et par ZS restent en complément) ;
 *  - sinon → Zones de Santé (la vue par Antenne est ajoutée en complément).
 */
export function resolveDrillLevel(f: Filters, multiProvince = false): { level: DrillLevel; label: string } {
  if (f.zs) return { level: "as", label: "Aire de Santé" };
  if (multiProvince && !f.antenne) return { level: "province", label: "Province" };
  return { level: "zs", label: "Zone de Santé" };
}

export function scopeLabel(f: Filters, province = "Kasaï Central"): string {
  if (f.as) return `Aire de Santé : ${f.as}`;
  if (f.zs) return `Zone de Santé : ${f.zs}`;
  if (f.antenne) return `Antenne : ${f.antenne}`;
  // Au-delà de 3 provinces, on reprend le style du rapport national Bloc 3
  // (le nom de fichier et les titres n'énumèrent plus toutes les provinces).
  if (f.provinces.length > 3) return `Campagne integree Bloc3 Aout 2026 J1-J5 (${f.provinces.length} provinces)`;
  if (f.provinces.length > 1) return `Provinces : ${f.provinces.map(prettyName).join(", ")}`;
  return `Province : ${province}`;
}

function prettyName(p: string): string {
  const s = (p || "").trim();
  if (/kasa[iï]/i.test(s) && /central/i.test(s)) return "Kasaï Central";
  return s.toLowerCase().replace(/[-_]+/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());
}

export function keyOf(r: ASRecord, level: DrillLevel): string {
  switch (level) {
    case "province": return r.province;
    case "antenne": return r.antenne || "—";
    case "zs": return r.zs || "—";
    case "as": return r.as || "—";
    case "all": return "Total";
  }
}

/* ─── Agrégat par unité ────────────────────────────────────────────────── */

export interface VaccineAgg {
  vacc: number;
  cible: number;
  cv: number | null;
  zeroDose: number;
  garcons: number;
  filles: number;
  ages: number[];
  flaconsRecus: number;
  flaconsUtil: number;
  flaconsRendus: number;
  flaconsPerdus: number;
  tauxPerte: number | null;
  daily: number[];
}

export interface UnitAgg {
  unit: string;
  /** Antenne de rattachement (pour les libellés « Antenne | ZS »). */
  antenne: string;
  nbAires: number;
  popTotale: number;
  vaccAttendus: number;
  vaccRecus: number;
  completude: number | null;
  attendusDaily: number[];
  recusDaily: number[];
  ciblePolio: number;
  cibleRR: number;
  rr: VaccineAgg;
  nvpo2: VaccineAgg;
  vpob: VaccineAgg;
  survPFA: number;
  survRougeole: number;
  survFJ: number;
  survTNN: number;
  mapiNonGraves: number;
  mapiGraves: number;
  /** MAPI (toutes) pour 100 000 doses administrées (RR + nVPO2 + VPOb). */
  mapiPour100k: number | null;
  aidantsIdent: number[];
  aidantsRecup: number[];
  pevIdent: number[];
  pevRecup: number[];
  refusSignales: number;
  refusGeres: number;
  menagesPrevus: number;
  menagesVisites: number;
}

export function tauxPerte(vacc: number, flaconsUtil: number, doses: number): number | null {
  if (!flaconsUtil) return null;
  return (1 - vacc / (flaconsUtil * doses)) * 100;
}

function emptyVacc(nbJours: number, nbAges: number): VaccineAgg {
  return {
    vacc: 0, cible: 0, cv: null, zeroDose: 0, garcons: 0, filles: 0,
    ages: new Array(nbAges).fill(0),
    flaconsRecus: 0, flaconsUtil: 0, flaconsRendus: 0, flaconsPerdus: 0,
    tauxPerte: null, daily: new Array(nbJours).fill(0),
  };
}
function addVacc(a: VaccineAgg, v: VaccineStats, cible: number): void {
  a.vacc += v.vacc;
  a.cible += cible;
  a.zeroDose += v.zeroDose;
  a.garcons += v.garcons;
  a.filles += v.filles;
  for (let i = 0; i < a.ages.length; i++) a.ages[i] += v.ages?.[i] ?? 0;
  a.flaconsRecus += v.flaconsRecus;
  a.flaconsUtil += v.flaconsUtil;
  a.flaconsRendus += v.flaconsRendus;
  a.flaconsPerdus += v.flaconsPerdus;
  for (let i = 0; i < a.daily.length; i++) a.daily[i] += v.daily?.[i] ?? 0;
}
function finalizeVacc(a: VaccineAgg, key: VaccineKey): void {
  a.cv = pct(a.vacc, a.cible);
  a.tauxPerte = tauxPerte(a.vacc, a.flaconsUtil, DOSES_PAR_FLACON[key]);
}

export function aggregateByUnit(records: ASRecord[], level: DrillLevel): UnitAgg[] {
  const map = new Map<string, UnitAgg>();
  const nbJours = Math.max(0, ...records.map((r) => r.dailyReports?.length ?? 0));

  for (const r of records) {
    const k = keyOf(r, level);
    let a = map.get(k);
    if (!a) {
      a = {
        unit: k,
        antenne: r.antenne,
        nbAires: 0, popTotale: 0,
        vaccAttendus: 0, vaccRecus: 0, completude: null,
        attendusDaily: new Array(nbJours).fill(0),
        recusDaily: new Array(nbJours).fill(0),
        ciblePolio: 0, cibleRR: 0,
        rr: emptyVacc(nbJours, RR_AGE_LABELS.length),
        nvpo2: emptyVacc(nbJours, POLIO_AGE_LABELS.length),
        vpob: emptyVacc(nbJours, POLIO_AGE_LABELS.length),
        survPFA: 0, survRougeole: 0, survFJ: 0, survTNN: 0,
        mapiNonGraves: 0, mapiGraves: 0, mapiPour100k: null,
        aidantsIdent: [0, 0, 0], aidantsRecup: [0, 0, 0],
        pevIdent: new Array(ANTIGENES.length).fill(0),
        pevRecup: new Array(ANTIGENES.length).fill(0),
        refusSignales: 0, refusGeres: 0, menagesPrevus: 0, menagesVisites: 0,
      };
      map.set(k, a);
    }
    a.nbAires += 1;
    a.popTotale += r.popTotale;
    a.vaccAttendus += r.vaccAttendus;
    a.vaccRecus += r.vaccRecus;
    for (let i = 0; i < nbJours; i++) {
      a.attendusDaily[i] += r.dailyReports?.[i]?.attendus ?? 0;
      a.recusDaily[i] += r.dailyReports?.[i]?.recus ?? 0;
    }
    a.ciblePolio += r.ciblePolio;
    a.cibleRR += r.cibleRR;
    addVacc(a.rr, r.rr, r.cibleRR);
    addVacc(a.nvpo2, r.nvpo2, r.ciblePolio);
    addVacc(a.vpob, r.vpob, r.ciblePolio);
    a.survPFA += r.survPFA;
    a.survRougeole += r.survRougeole;
    a.survFJ += r.survFJ;
    a.survTNN += r.survTNN;
    a.mapiNonGraves += r.mapiNonGraves;
    a.mapiGraves += r.mapiGraves;
    for (let j = 0; j < 3; j++) { a.aidantsIdent[j] += r.aidantsIdent?.[j] ?? 0; a.aidantsRecup[j] += r.aidantsRecup?.[j] ?? 0; }
    for (let j = 0; j < ANTIGENES.length; j++) { a.pevIdent[j] += r.pevIdent?.[j] ?? 0; a.pevRecup[j] += r.pevRecup?.[j] ?? 0; }
    a.refusSignales += r.refusSignales;
    a.refusGeres += r.refusGeres;
    a.menagesPrevus += r.menagesPrevus;
    a.menagesVisites += r.menagesVisites;
  }
  const out = Array.from(map.values());
  for (const a of out) {
    a.completude = pct(a.vaccRecus, a.vaccAttendus);
    finalizeVacc(a.rr, "rr");
    finalizeVacc(a.nvpo2, "nvpo2");
    finalizeVacc(a.vpob, "vpob");
    a.mapiPour100k = mapiPour100k(a.mapiNonGraves + a.mapiGraves, a.rr.vacc + a.nvpo2.vacc + a.vpob.vacc);
  }
  return out.sort((a, b) => a.unit.localeCompare(b.unit, "fr"));
}

export function mapiPour100k(mapi: number, doses: number): number | null {
  if (!doses) return null;
  return (mapi / doses) * 100000;
}

/** Totaux du périmètre = agrégat à une seule unité (tous niveaux confondus). */
export function totals(records: ASRecord[]): UnitAgg {
  const [t] = aggregateByUnit(records, "all");
  if (t) return { ...t, unit: "Total" };
  return emptyUnit();
}

function emptyUnit(): UnitAgg {
  return {
    unit: "Total", antenne: "", nbAires: 0, popTotale: 0, vaccAttendus: 0, vaccRecus: 0, completude: null,
    attendusDaily: [], recusDaily: [], ciblePolio: 0, cibleRR: 0,
    rr: emptyVacc(0, RR_AGE_LABELS.length), nvpo2: emptyVacc(0, POLIO_AGE_LABELS.length), vpob: emptyVacc(0, POLIO_AGE_LABELS.length),
    survPFA: 0, survRougeole: 0, survFJ: 0, survTNN: 0, mapiNonGraves: 0, mapiGraves: 0, mapiPour100k: null,
    aidantsIdent: [0, 0, 0], aidantsRecup: [0, 0, 0],
    pevIdent: new Array(ANTIGENES.length).fill(0), pevRecup: new Array(ANTIGENES.length).fill(0),
    refusSignales: 0, refusGeres: 0, menagesPrevus: 0, menagesVisites: 0,
  };
}

/* ─── Helpers de tri ───────────────────────────────────────────────────── */

export function sortByDesc<T>(rows: T[], sel: (r: T) => number | null): T[] {
  return [...rows].sort((a, b) => (sel(b) ?? -Infinity) - (sel(a) ?? -Infinity));
}
export function sortByAsc<T>(rows: T[], sel: (r: T) => number | null): T[] {
  return [...rows].sort((a, b) => (sel(a) ?? Infinity) - (sel(b) ?? Infinity));
}

function uniq(arr: string[]): string[] {
  return Array.from(new Set(arr.filter(Boolean))).sort((a, b) => a.localeCompare(b, "fr"));
}

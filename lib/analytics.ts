/**
 * Agrégation des données du masque vers les indicateurs du rapport (polio uniquement).
 * Applique les filtres en cascade (Province → Antenne → ZS → AS) avec la même
 * logique que le dashboard : un niveau sélectionné restreint le périmètre, et
 * l'agrégation des graphiques se fait au niveau immédiatement inférieur.
 */

import { ANTIGENES, isRecapRow, type ASRecord, type DailyValue, type MasqueData } from "./parse-masque";
import type { Filters } from "./store";
import { pct } from "./format";

export interface CascadeOptions {
  provinces: string[];
  antennes: string[];
  zones: string[];
  aires: string[];
}

// Une province vide (aucune sélection) signifie « toutes les provinces ».
function inProvinces(r: ASRecord, f: Filters): boolean {
  return f.provinces.length === 0 || f.provinces.includes(r.province);
}

/**
 * Écarte les lignes titres/récapitulatives où un parent est reporté dans une
 * colonne enfant (« Ant.LUIZA » dans la colonne Zone de Santé, nom de la Province
 * dans la colonne Antenne au niveau national…). Garde-fou pour les compilations
 * déjà stockées avant le correctif du parseur : sans cela, le parent réapparaît
 * comme enfant dans les listes déroulantes et dans le rapport généré.
 */
function isRealUnit(r: ASRecord): boolean {
  return !isRecapRow(r.province, r.antenne, r.zs, r.as);
}

export function cascadeOptions(data: MasqueData, f: Filters): CascadeOptions {
  const records = data.records.filter(isRealUnit);
  const provinces = uniq(records.map((r) => r.province));
  const antennes = uniq(
    records.filter((r) => inProvinces(r, f)).map((r) => r.antenne)
  );
  const zones = uniq(
    records
      .filter((r) => inProvinces(r, f) && (!f.antenne || r.antenne === f.antenne))
      .map((r) => r.zs)
  );
  const aires = uniq(
    records
      .filter(
        (r) =>
          inProvinces(r, f) &&
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
      isRealUnit(r) &&
      inProvinces(r, f) &&
      (!f.antenne || r.antenne === f.antenne) &&
      (!f.zs || r.zs === f.zs) &&
      (!f.as || r.as === f.as)
  );
}

export type DrillLevel = "province" | "antenne" | "zs" | "as";

/**
 * Niveau d'agrégation des tableaux du rapport.
 * - ZS filtrée → détail par Aire de Santé.
 * - Antenne filtrée, ou une seule province sélectionnée → détail par Zone de Santé.
 * - Aucune sélection (niveau pays) ou plusieurs provinces → agrégation par Province
 *   (par défaut le rapport présente les provinces et non les ZS). Un jeu de données
 *   mono-province (import local) tombe sur les ZS, l'agrégation par province n'ayant
 *   alors qu'une seule ligne.
 */
export function resolveDrillLevel(
  f: Filters,
  provinceCount: number
): { level: DrillLevel; label: string } {
  if (f.zs) return { level: "as", label: "Aire de Santé" };
  if (f.antenne) return { level: "zs", label: "Zone de Santé" };
  if (f.provinces.length === 1) return { level: "zs", label: "Zone de Santé" };
  if (f.provinces.length === 0 && provinceCount <= 1) return { level: "zs", label: "Zone de Santé" };
  return { level: "province", label: "Province" };
}

export function scopeLabel(f: Filters): string {
  if (f.as) return `Aire de Santé : ${f.as}`;
  if (f.zs) return `Zone de Santé : ${f.zs}`;
  if (f.antenne) return `Antenne : ${f.antenne}`;
  if (f.provinces.length === 1) return `Province : ${f.provinces[0]}`;
  if (f.provinces.length > 1) return `Provinces : ${f.provinces.join(", ")}`;
  return "Niveau national — toutes les provinces";
}

export function scopeName(f: Filters, data: MasqueData): string {
  if (f.as) return f.as;
  if (f.zs) return f.zs;
  if (f.antenne) return f.antenne;
  if (f.provinces.length === 1) return f.provinces[0];
  if (f.provinces.length > 1) return "Provinces sélectionnées";
  return data.meta.province ?? "Province";
}

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
  nvpo2FlaconsRecus: number;
  nvpo2FlaconsUtil: number;
  nvpo2FlaconsRendus: number;
  nvpo2Perdus: number;
  vpobFlaconsRecus: number;
  vpobFlaconsUtil: number;
  vpobFlaconsRendus: number;
  vpobPerdus: number;
  vaccAttendus: number;
  vaccRecus: number;
  recup: number;
  nvpo2ZeroDose: number;
  vpobZeroDose: number;
  /** Vaccinés nVPO2 par jour de campagne (cumul agrégé par unité). */
  nvpo2Daily: number[];
  /** Vaccinés VPOb par jour de campagne. */
  vpobDaily: number[];
  /** Rapports reçus par jour de campagne. */
  rapportsRecusDaily: number[];
  /** Rapports attendus par jour de campagne. */
  rapportsAttendusDaily: number[];
  /** Enfants vaccinés par antigène (somme par unité) — ordre = ANTIGENES. */
  antigenesEV: number[];
  /** Enfants identifiés pour récupération par antigène (somme par unité) — ordre = ANTIGENES. */
  antigenesIdentifies: number[];
  survPFA: number;
  survRougeole: number;
  survFJ: number;
  survTNN: number;
}

export function aggregateByUnit(records: ASRecord[], level: DrillLevel): UnitAgg[] {
  const map = new Map<string, UnitAgg>();
  const nbJours = Math.max(0, ...records.map((r) => r.nvpo2Daily.length));

  for (const r of records) {
    const k = keyOf(r, level);
    let a = map.get(k);
    if (!a) {
      a = {
        unit: k,
        nvpo2Vacc: 0, nvpo2Cible: 0, vpobVacc: 0, vpobCible: 0,
        nvpo2FlaconsRecus: 0, nvpo2FlaconsUtil: 0, nvpo2FlaconsRendus: 0, nvpo2Perdus: 0,
        vpobFlaconsRecus: 0, vpobFlaconsUtil: 0, vpobFlaconsRendus: 0, vpobPerdus: 0,
        vaccAttendus: 0, vaccRecus: 0, recup: 0, nvpo2ZeroDose: 0, vpobZeroDose: 0,
        nvpo2Daily: new Array(nbJours).fill(0),
        vpobDaily: new Array(nbJours).fill(0),
        rapportsRecusDaily: new Array(nbJours).fill(0),
        rapportsAttendusDaily: new Array(nbJours).fill(0),
        antigenesEV: new Array(ANTIGENES.length).fill(0),
        antigenesIdentifies: new Array(ANTIGENES.length).fill(0),
        survPFA: 0, survRougeole: 0, survFJ: 0, survTNN: 0,
      };
      map.set(k, a);
    }
    a.nvpo2Vacc += r.nvpo2Vacc;
    a.nvpo2Cible += r.nvpo2CibleExtrap;
    a.vpobVacc += r.vpobVacc;
    a.vpobCible += r.vpobCibleExtrap;
    a.nvpo2FlaconsRecus += r.nvpo2FlaconsRecus;
    a.nvpo2FlaconsUtil += r.nvpo2FlaconsUtil;
    a.nvpo2FlaconsRendus += r.nvpo2FlaconsRendus;
    a.nvpo2Perdus += r.nvpo2Perdus;
    a.vpobFlaconsRecus += r.vpobFlaconsRecus;
    a.vpobFlaconsUtil += r.vpobFlaconsUtil;
    a.vpobFlaconsRendus += r.vpobFlaconsRendus;
    a.vpobPerdus += r.vpobPerdus;
    a.vaccAttendus += r.vaccAttendus;
    a.vaccRecus += r.vaccRecus;
    a.recup += r.recup;
    a.nvpo2ZeroDose += r.nvpo2ZeroDose;
    a.vpobZeroDose += r.vpobZeroDose;
    for (let i = 0; i < nbJours; i++) {
      a.nvpo2Daily[i] += r.nvpo2Daily[i]?.vaccines ?? 0;
      a.vpobDaily[i] += r.vpobDaily[i]?.vaccines ?? 0;
      a.rapportsRecusDaily[i] += r.nvpo2Daily[i]?.rapportsRecus ?? 0;
      a.rapportsAttendusDaily[i] += r.nvpo2Daily[i]?.rapportsAttendus ?? 0;
    }
    for (let j = 0; j < a.antigenesEV.length; j++) a.antigenesEV[j] += r.antigenesEV?.[j] ?? 0;
    for (let j = 0; j < a.antigenesIdentifies.length; j++) a.antigenesIdentifies[j] += r.antigenesIdentifies?.[j] ?? 0;
    a.survPFA += r.survPFA ?? 0;
    a.survRougeole += r.survRougeole ?? 0;
    a.survFJ += r.survFJ ?? 0;
    a.survTNN += r.survTNN ?? 0;
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
  nvpo2FlaconsRecus: number;
  nvpo2FlaconsUtil: number;
  nvpo2FlaconsRendus: number;
  nvpo2Perdus: number;
  nvpo2TauxPerte: number | null;
  vpobFlaconsRecus: number;
  vpobFlaconsUtil: number;
  vpobFlaconsRendus: number;
  vpobPerdus: number;
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
  /** Enfants vaccinés par antigène (total) — ordre = ANTIGENES. */
  antigenesEV: number[];
  /** Enfants identifiés pour récupération par antigène (total) — ordre = ANTIGENES. */
  antigenesIdentifies: number[];
  survPFA: number;
  survRougeole: number;
  survFJ: number;
  survTNN: number;
  /** Cumul Vaccinés nVPO2 par jour. */
  nvpo2VaccDaily: number[];
  /** Cumul Vaccinés VPOb par jour. */
  vpobVaccDaily: number[];
  /** Cumul Rapports Reçus par jour (utilisé pour la complétude journalière). */
  recusDaily: number[];
  /** Cumul Rapports Attendus par jour de campagne (= mosoAttendus × jour). */
  attendusDaily: number[];
}

export function totals(records: ASRecord[]): Totals {
  const s = (sel: (r: ASRecord) => number) => records.reduce((acc, r) => acc + sel(r), 0);
  const nvpo2Vacc = s((r) => r.nvpo2Vacc);
  const nvpo2Cible = s((r) => r.nvpo2CibleExtrap);
  const vpobVacc = s((r) => r.vpobVacc);
  const vpobCible = s((r) => r.vpobCibleExtrap);
  const nvpo2FlaconsRecus = s((r) => r.nvpo2FlaconsRecus);
  const nvpo2FlaconsUtil = s((r) => r.nvpo2FlaconsUtil);
  const nvpo2FlaconsRendus = s((r) => r.nvpo2FlaconsRendus);
  const nvpo2Perdus = s((r) => r.nvpo2Perdus);
  const vpobFlaconsRecus = s((r) => r.vpobFlaconsRecus);
  const vpobFlaconsUtil = s((r) => r.vpobFlaconsUtil);
  const vpobFlaconsRendus = s((r) => r.vpobFlaconsRendus);
  const vpobPerdus = s((r) => r.vpobPerdus);
  const vaccAttendus = s((r) => r.vaccAttendus);
  const vaccRecus = s((r) => r.vaccRecus);
  const refusSignales = s((r) => r.refusSignales);
  const refusGeres = s((r) => r.refusGeres);
  const menagesPrevus = s((r) => r.menagesPrevus);
  const menagesVisites = s((r) => r.menagesVisites);

  const nbJours = Math.max(0, ...records.map((r) => r.nvpo2Daily.length));
  const nvpo2VaccDaily = new Array(nbJours).fill(0);
  const vpobVaccDaily = new Array(nbJours).fill(0);
  const recusDaily = new Array(nbJours).fill(0);
  const attendusDaily = new Array(nbJours).fill(0);
  for (const r of records) {
    for (let i = 0; i < nbJours; i++) {
      nvpo2VaccDaily[i] += r.nvpo2Daily[i]?.vaccines ?? 0;
      vpobVaccDaily[i] += r.vpobDaily[i]?.vaccines ?? 0;
      recusDaily[i] += r.nvpo2Daily[i]?.rapportsRecus ?? 0;
      attendusDaily[i] += r.mosoAttendus;
    }
  }

  return {
    nvpo2Vacc, nvpo2Cible, nvpo2CV: pct(nvpo2Vacc, nvpo2Cible),
    vpobVacc, vpobCible, vpobCV: pct(vpobVacc, vpobCible),
    nvpo2FlaconsRecus, nvpo2FlaconsUtil, nvpo2FlaconsRendus, nvpo2Perdus,
    nvpo2TauxPerte: tauxPerte(nvpo2Vacc, nvpo2FlaconsUtil, NVPO2_DOSES_PAR_FLACON),
    vpobFlaconsRecus, vpobFlaconsUtil, vpobFlaconsRendus, vpobPerdus,
    vpobTauxPerte: tauxPerte(vpobVacc, vpobFlaconsUtil, VPOB_DOSES_PAR_FLACON),
    vaccAttendus, vaccRecus, completude: pct(vaccRecus, vaccAttendus),
    recup: s((r) => r.recup),
    nvpo2ZeroDose: s((r) => r.nvpo2ZeroDose),
    vpobZeroDose: s((r) => r.vpobZeroDose),
    refusSignales, refusGeres, refusGeresPct: pct(refusGeres, refusSignales),
    menagesPrevus, menagesVisites, menagesVisitesPct: pct(menagesVisites, menagesPrevus),
    pers15: s((r) => r.pers15),
    mapiMineures: s((r) => r.mapiMineures),
    mapiGraves: s((r) => r.mapiGraves),
    antigenesEV: ANTIGENES.map((_, j) => records.reduce((acc, r) => acc + (r.antigenesEV?.[j] ?? 0), 0)),
    antigenesIdentifies: ANTIGENES.map((_, j) => records.reduce((acc, r) => acc + (r.antigenesIdentifies?.[j] ?? 0), 0)),
    survPFA: s((r) => r.survPFA ?? 0),
    survRougeole: s((r) => r.survRougeole ?? 0),
    survFJ: s((r) => r.survFJ ?? 0),
    survTNN: s((r) => r.survTNN ?? 0),
    nvpo2VaccDaily, vpobVaccDaily, recusDaily, attendusDaily,
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

export interface CoverageDailyRow {
  unit: string;
  cibleCampagne: number;
  /** Cible journalière calculée = Cible Campagne ÷ Nb jours de campagne. */
  cibleJournaliere: number;
  /** Pour chaque jour : { vaccines, couverture % cumulée ou ponctuelle } */
  daily: { vaccines: number; couvJour: number | null }[];
  totalVacc: number;
  couvGlobale: number | null;
}

/**
 * Construit le tableau « par ZS × par jour » à l'identique du modèle Power BI :
 *   ZS | Cible Campagne | Cible Journalière | Vaccinés J1 | Couvert. J1 | … | Couvert. Globale.
 * La colonne « Couvert. JourN » correspond au taux journalier (vaccinés_jour ÷ cible).
 * Le total global cumule tous les vaccinés des jours saisis sur la cible campagne.
 */
export function coverageByDay(
  byUnit: UnitAgg[],
  vaccine: "nvpo2" | "vpob",
  nbJours: number
): CoverageDailyRow[] {
  return byUnit
    .map((u) => {
      const cible = vaccine === "nvpo2" ? u.nvpo2Cible : u.vpobCible;
      const daily = vaccine === "nvpo2" ? u.nvpo2Daily : u.vpobDaily;
      const cibleJournaliere = nbJours > 0 ? cible / nbJours : 0;
      const dailyRows = Array.from({ length: nbJours }, (_, i) => {
        const vaccines = daily[i] ?? 0;
        return { vaccines, couvJour: cible > 0 ? (vaccines / cible) * 100 : null };
      });
      const totalVacc = dailyRows.reduce((a, b) => a + b.vaccines, 0);
      const couvGlobale = cible > 0 ? (totalVacc / cible) * 100 : null;
      return { unit: u.unit, cibleCampagne: cible, cibleJournaliere, daily: dailyRows, totalVacc, couvGlobale };
    })
    // Tri décroissant par couverture globale (comme le modèle Power BI).
    .sort((a, b) => (b.couvGlobale ?? -1) - (a.couvGlobale ?? -1));
}

export interface GestionRow {
  unit: string;
  flaconsRecus: number;
  flaconsUtil: number;
  flaconsRendus: number;
  perdus: number;
  vacc: number;
  taux: number | null;
}

export function nvpo2Gestion(byUnit: UnitAgg[]): GestionRow[] {
  return byUnit
    .map((a) => ({
      unit: a.unit,
      flaconsRecus: a.nvpo2FlaconsRecus,
      flaconsUtil: a.nvpo2FlaconsUtil,
      flaconsRendus: a.nvpo2FlaconsRendus,
      perdus: a.nvpo2Perdus,
      vacc: a.nvpo2Vacc,
      taux: tauxPerte(a.nvpo2Vacc, a.nvpo2FlaconsUtil, NVPO2_DOSES_PAR_FLACON),
    }))
    .sort((a, b) => (b.taux ?? -999) - (a.taux ?? -999));
}

export function vpobGestion(byUnit: UnitAgg[]): GestionRow[] {
  return byUnit
    .map((a) => ({
      unit: a.unit,
      flaconsRecus: a.vpobFlaconsRecus,
      flaconsUtil: a.vpobFlaconsUtil,
      flaconsRendus: a.vpobFlaconsRendus,
      perdus: a.vpobPerdus,
      vacc: a.vpobVacc,
      taux: tauxPerte(a.vpobVacc, a.vpobFlaconsUtil, VPOB_DOSES_PAR_FLACON),
    }))
    .sort((a, b) => (b.taux ?? -999) - (a.taux ?? -999));
}

// Helpers de compatibilité (utilisés par d'autres modules) ────────────────────
export interface CoverageRow { unit: string; cible: number; vacc: number; cv: number | null }
export function nvpo2Coverage(byUnit: UnitAgg[]): CoverageRow[] {
  return byUnit.map((a) => ({ unit: a.unit, cible: a.nvpo2Cible, vacc: a.nvpo2Vacc, cv: pct(a.nvpo2Vacc, a.nvpo2Cible) }));
}
export function vpobCoverage(byUnit: UnitAgg[]): CoverageRow[] {
  return byUnit.map((a) => ({ unit: a.unit, cible: a.vpobCible, vacc: a.vpobVacc, cv: pct(a.vpobVacc, a.vpobCible) }));
}

function uniq(arr: string[]): string[] {
  return Array.from(new Set(arr.filter(Boolean))).sort((a, b) => a.localeCompare(b, "fr"));
}

// Types réexportés pour les consommateurs.
export type { DailyValue };

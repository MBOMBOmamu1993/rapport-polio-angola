/**
 * Parseur du masque de saisie de la campagne intégrée RR‑POLIO (Kasaï Central).
 *
 * Le masque (Google Sheets exporté en .xlsx) contient, par Aire de Santé :
 *  - « Synthèse » : tous les indicateurs cumulés de la campagne (5 lignes d'en‑tête,
 *    données à partir de la ligne 6) ;
 *  - « Jour1 » … « Jour6 » : la même structure de colonnes, jour par jour.
 *    Le Jour6 est le **ratissage** (campagne = 5 jours + ratissage) ;
 *  - « J(-1) » : préparatifs (non exploités) ; « Synth_ZS » : totaux par ZS ;
 *  - « Donnees de base » : période, province, cibles.
 *
 * Trois vaccins sont suivis : nVPO2 et VPOb (cible polio 0‑9 ans ≈ 33,3 % de la
 * population) et RR (cible 6 mois‑14 ans ≈ 46 % de la population). Les cibles sont
 * lues telles quelles dans le masque (colonnes « Cible Polio » et « Cible RR »).
 */

import * as XLSX from "xlsx";

/* ─── Colonnes (1‑based) de la feuille « Synthèse » / « JourN » ─────────── */

export const C = {
  province: 1,
  antenne: 2,
  zs: 3,
  as: 4,
  popTotale: 5,
  menagesPrevus: 6,
  menagesVisites: 7,
  mosoAttendus: 8,
  mosoRecus: 9,
  pers15Total: 10,
  refusSignales: 18,
  refusGeres: 19,
  // Complétude des rapports de vaccination
  vaccAttendus: 45,
  vaccRecus: 46,
  // Cibles de la campagne
  ciblePolio: 48,
  cibleRR: 49,
  // Gestion des vaccins (Quantité reçue / utilisée / rendue / perdue) — flacons
  nvpo2Gestion: 98,
  vpobGestion: 102,
  rrGestion: 106,
  diluantGestion: 110,
  sabGestion: 114,
  sdGestion: 118,
  receptaclesGestion: 122,
  // Surveillance des MPV
  survPFA: 126,
  survRougeole: 127,
  survFJ: 128,
  survTNN: 129,
  // Récupération par les aidants communautaires (identifiés / récupérés : nVPO2, VPOb, RR)
  aidantsIdent: 130,
  aidantsRecup: 133,
  // MAPI
  mapiNonGraves: 136,
  mapiGraves: 137,
  // PEV systématique : identifiés (22 antigènes) puis récupérés (22 antigènes)
  pevIdent: 138,
  pevRecup: 160,
  // Synthèse de la vaccination
  nvpo2Synth: 182,
  vpobSynth: 206,
  rrSynth: 230,
  // Taux de perte calculés par le masque
  nvpo2TauxPerte: 255,
  vpobTauxPerte: 256,
  rrTauxPerte: 257,
} as const;

/** Antigènes du bloc « PEV systématique » (identifiés / récupérés), dans l'ordre du masque. */
export const ANTIGENES: { key: string; label: string }[] = [
  { key: "BCG", label: "BCG" },
  { key: "DTC1", label: "DTC1" },
  { key: "PCV1", label: "PCV1" },
  { key: "ROTA1", label: "Rota1" },
  { key: "DTC2", label: "DTC2" },
  { key: "PCV2", label: "PCV2" },
  { key: "ROTA2", label: "Rota2" },
  { key: "DTC3", label: "DTC3" },
  { key: "PCV3", label: "PCV3" },
  { key: "ROTA3", label: "Rota3" },
  { key: "VPI1", label: "VPI1" },
  { key: "VAP1", label: "VAP1" },
  { key: "VAP2", label: "VAP2" },
  { key: "VAP3", label: "VAP3" },
  { key: "VPI2", label: "VPI2" },
  { key: "VAA", label: "VAA" },
  { key: "VAP4", label: "VAP4" },
  { key: "TD1", label: "Td1" },
  { key: "TD2", label: "Td2" },
  { key: "TD3", label: "Td3" },
  { key: "TD4", label: "Td4" },
  { key: "TD5", label: "Td5" },
];

/** Tranches d'âge de la vaccination polio (nVPO2 / VPOb) — ordre du tableau `ages`. */
export const POLIO_AGE_LABELS = ["0-5 mois", "6-11 mois", "12-59 mois", "5-9 ans"];
/** Tranches d'âge de la vaccination RR — ordre du tableau `ages`. */
export const RR_AGE_LABELS = ["6-11 mois", "12-59 mois", "5-14 ans"];

export type VaccineKey = "nvpo2" | "vpob" | "rr";
export const VACCINE_KEYS: VaccineKey[] = ["rr", "nvpo2", "vpob"];
export const VACCINE_LABELS: Record<VaccineKey, string> = { nvpo2: "nVPO2", vpob: "VPOb", rr: "RR" };
/** Doses par flacon (base du taux de perte : 1 − vaccinés ÷ (flacons utilisés × doses)). */
export const DOSES_PAR_FLACON: Record<VaccineKey, number> = { nvpo2: 50, vpob: 20, rr: 10 };
/** Seuil de perte acceptable (%), par vaccin. */
export const SEUIL_PERTE: Record<VaccineKey, number> = { nvpo2: 10, vpob: 10, rr: 10 };

export interface VaccineStats {
  /** Enfants vaccinés (toute la cible). */
  vacc: number;
  zeroDose: number;
  garcons: number;
  filles: number;
  /** Vaccinés par tranche d'âge (POLIO_AGE_LABELS ou RR_AGE_LABELS). */
  ages: number[];
  flaconsRecus: number;
  flaconsUtil: number;
  flaconsRendus: number;
  flaconsPerdus: number;
  /** Vaccinés par jour de campagne (index = jour − 1 ; le dernier peut être le ratissage). */
  daily: number[];
}

export interface DailyReport {
  /** Étiquette du jour (« J1 » … « J5 », « Ratissage »). */
  label: string;
  attendus: number;
  recus: number;
}

export interface ASRecord {
  province: string;
  antenne: string;
  zs: string;
  as: string;
  popTotale: number;
  menagesPrevus: number;
  menagesVisites: number;
  mosoAttendus: number;
  mosoRecus: number;
  pers15: number;
  refusSignales: number;
  refusGeres: number;
  vaccAttendus: number;
  vaccRecus: number;
  /** Complétude journalière des rapports de vaccination. */
  dailyReports: DailyReport[];
  ciblePolio: number;
  cibleRR: number;
  nvpo2: VaccineStats;
  vpob: VaccineStats;
  rr: VaccineStats;
  survPFA: number;
  survRougeole: number;
  survFJ: number;
  survTNN: number;
  mapiNonGraves: number;
  mapiGraves: number;
  /** Récupération par les aidants communautaires — [nVPO2, VPOb, RR]. */
  aidantsIdent: number[];
  aidantsRecup: number[];
  /** PEV systématique : identifiés / récupérés par antigène (ordre = ANTIGENES). */
  pevIdent: number[];
  pevRecup: number[];
}

export interface MasqueData {
  meta: {
    pays: string;
    periode: string;
    dateDebut: string;
    dateFin: string;
    province: string;
    antennes: string[];
    zones: string[];
    importedAt: string;
    fileName: string;
    nbAires: number;
    /** Nombre de jours de campagne effectivement renseignés (données présentes). */
    nbJours: number;
    /** Étiquettes des jours (« J1 » … « J5 », « Ratissage »). */
    jourLabels: string[];
    /** Version du format d'enregistrement (permet d'invalider les anciens caches). */
    schema: number;
  };
  records: ASRecord[];
}

export const MASQUE_SCHEMA = 4;

/* ─── Helpers ──────────────────────────────────────────────────────────── */

function num(v: unknown): number {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", ".").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}
function str(v: unknown): string {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}
function cell(row: unknown[], col1: number): unknown {
  return row[col1 - 1];
}
function sum(row: unknown[], from: number, count: number): number {
  let s = 0;
  for (let i = 0; i < count; i++) s += num(cell(row, from + i));
  return s;
}
/** Valeur de synthèse si renseignée, sinon repli sur la valeur recalculée. */
function pick(synth: number, raw: number): number {
  return synth > 0 ? synth : raw;
}

const TOTAL_RE = /^\s*total\b/i;

/**
 * Lignes de récapitulation (titres, sous‑totaux) insérées dans les feuilles :
 * « Total BENA LEKA », « Prvce. KASAI-CENTRAL », « Ant.LUIZA »… Ce ne sont pas
 * des unités réelles et les inclure doublerait les totaux.
 */
const RECAP_RE = /^\s*(ant|antenne|zs|zone|prov|prvce|province|total)\b[\s.:\-]*\S/i;
export function isRecapLabel(s: string): boolean {
  return RECAP_RE.test(s);
}
function normLabel(s: string): string {
  return s.normalize("NFD").replace(/[̀-ͯ]/g, "").trim().toUpperCase();
}
export function isRecapRow(province: string, antenne: string, zs: string, as: string): boolean {
  if (isRecapLabel(antenne) || isRecapLabel(zs) || isRecapLabel(as)) return true;
  if (TOTAL_RE.test(zs) || TOTAL_RE.test(as)) return true;
  if (antenne && province && normLabel(antenne) === normLabel(province)) return true;
  if (zs && province && normLabel(zs) === normLabel(province)) return true;
  return false;
}

/**
 * Écarte les sous‑totaux résiduels par concordance numérique : une ligne dont les
 * effectifs égalent la somme d'au moins deux voisines du même parent est un
 * sous‑total. On exige la concordance sur plusieurs indicateurs indépendants.
 */
export function sanitizeRecords(records: ASRecord[]): ASRecord[] {
  const kept = records.filter((r) => !isRecapRow(r.province, r.antenne, r.zs, r.as));
  const metricsOf = (r: ASRecord): number[] => [r.popTotale, r.vaccAttendus, r.ciblePolio, r.cibleRR];
  const approxEq = (a: number, b: number): boolean => Math.abs(a - b) <= Math.max(1, Math.abs(b) * 0.001);
  const removed = new Set<ASRecord>();
  const detect = (groupKey: (r: ASRecord) => string): void => {
    const groups = new Map<string, ASRecord[]>();
    for (const r of kept) {
      if (removed.has(r)) continue;
      const k = groupKey(r);
      const g = groups.get(k);
      if (g) g.push(r);
      else groups.set(k, [r]);
    }
    for (const group of groups.values()) {
      if (group.length < 3) continue;
      for (const cand of group) {
        if (removed.has(cand)) continue;
        const others = group.filter((x) => x !== cand && !removed.has(x));
        if (others.length < 2) continue;
        const m = metricsOf(cand);
        const sums = m.map((_, i) => others.reduce((acc, x) => acc + metricsOf(x)[i], 0));
        const nonTrivial = m.some((v, i) => v > 0 && sums[i] > 0);
        if (nonTrivial && m.every((v, i) => approxEq(v, sums[i]))) removed.add(cand);
      }
    }
  };
  detect((r) => `${normLabel(r.province)}||${normLabel(r.antenne)}`);
  detect((r) => normLabel(r.province));
  detect((r) => `${normLabel(r.province)}||${normLabel(r.antenne)}||${normLabel(r.zs)}`);
  return kept.filter((r) => !removed.has(r));
}

/** Clé d'identification d'une Aire de Santé (insensible casse / accents / ponctuation). */
export function asKey(province: string, zs: string, as: string): string {
  const norm = (s: string) =>
    s.normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  return `${norm(province)}|${norm(zs)}|${norm(as)}`;
}

/* ─── Lecture d'une ligne de vaccination ───────────────────────────────── */

type VaccCore = Omit<VaccineStats, "daily">;
interface RowVacc { nvpo2: VaccCore; vpob: VaccCore; rr: VaccCore }

/**
 * Bloc « Synthèse de la vaccination » polio (24 colonnes à partir de `base`) :
 *  +0..2  0‑5 mois (ZD, 1+, Total)   +3..5  6‑11 mois   +6..8  0‑11 mois
 *  +9..11 12‑59 mois                 +12..14 0‑59 mois  +15..17 0‑59 (G, F, T)
 *  +18..20 0‑9 ans (ZD, 1+, Total)   +21..23 0‑9 ans (G, F, T)
 * Repli sur le bloc brut « enfants vaccinés par tranche d'âge et sexe » (col. 50‑97 :
 * 0 dose F/M puis 1 dose+ F/M par tranche) si la synthèse n'est pas calculée.
 */
function readPolio(
  row: unknown[],
  base: number,
  rawCols: { m0_5: number; m6_11: number; m12_59: number; a5_9: number },
  gest: number
): VaccCore {
  const raw4 = (c: number) => sum(row, c, 4);
  const rawZD = (c: number) => sum(row, c, 2);
  const rawF = (c: number) => num(cell(row, c)) + num(cell(row, c + 2));
  const rawM = (c: number) => num(cell(row, c + 1)) + num(cell(row, c + 3));
  const tr = [rawCols.m0_5, rawCols.m6_11, rawCols.m12_59, rawCols.a5_9];

  const s0_5 = num(cell(row, base + 2));
  const s6_11 = num(cell(row, base + 5));
  const s12_59 = num(cell(row, base + 11));
  const s0_59 = num(cell(row, base + 14));
  const s0_9 = num(cell(row, base + 20));
  const s5_9 = Math.max(0, s0_9 - s0_59);

  const r0_5 = raw4(tr[0]);
  const r6_11 = raw4(tr[1]);
  const r12_59 = raw4(tr[2]);
  const r5_9 = raw4(tr[3]);
  const rawTotal = r0_5 + r6_11 + r12_59 + r5_9;

  const vacc = pick(s0_9, rawTotal);
  const ages = s0_9 > 0 ? [s0_5, s6_11, s12_59, s5_9] : [r0_5, r6_11, r12_59, r5_9];
  const zeroDose = pick(num(cell(row, base + 18)), tr.reduce((a, c) => a + rawZD(c), 0));
  const garcons = pick(num(cell(row, base + 21)), tr.reduce((a, c) => a + rawM(c), 0));
  const filles = pick(num(cell(row, base + 22)), tr.reduce((a, c) => a + rawF(c), 0));

  return {
    vacc, zeroDose, garcons, filles, ages,
    flaconsRecus: num(cell(row, gest)),
    flaconsUtil: num(cell(row, gest + 1)),
    flaconsRendus: num(cell(row, gest + 2)),
    flaconsPerdus: num(cell(row, gest + 3)),
  };
}

/**
 * Bloc « Synthèse de la vaccination » RR (15 colonnes à partir de `base`) :
 *  +0..2 6‑11 mois (ZD, 1+, T)  +3..5 12‑59 mois  +6..8 6‑59 mois (G, F, T)
 *  +9..11 6 mois‑14 ans (ZD, 1+, T)  +12..14 6 mois‑14 ans (G, F, T)
 * Repli sur le bloc brut : 6‑11 mois col. 66, 12‑59 mois col. 78, 5‑9 ans col. 90,
 * 10‑14 ans col. 94.
 */
function readRR(row: unknown[], base: number, gest: number): VaccCore {
  const raw4 = (c: number) => sum(row, c, 4);
  const rawZD = (c: number) => sum(row, c, 2);
  const rawF = (c: number) => num(cell(row, c)) + num(cell(row, c + 2));
  const rawM = (c: number) => num(cell(row, c + 1)) + num(cell(row, c + 3));
  const tr = [66, 78, 90, 94];

  const s6_11 = num(cell(row, base + 2));
  const s12_59 = num(cell(row, base + 5));
  const sTotal = num(cell(row, base + 11));
  const s5_14 = Math.max(0, sTotal - s6_11 - s12_59);

  const r6_11 = raw4(66);
  const r12_59 = raw4(78);
  const r5_14 = raw4(90) + raw4(94);
  const rawTotal = r6_11 + r12_59 + r5_14;

  const vacc = pick(sTotal, rawTotal);
  const ages = sTotal > 0 ? [s6_11, s12_59, s5_14] : [r6_11, r12_59, r5_14];
  const zeroDose = pick(num(cell(row, base + 9)), tr.reduce((a, c) => a + rawZD(c), 0));
  const garcons = pick(num(cell(row, base + 12)), tr.reduce((a, c) => a + rawM(c), 0));
  const filles = pick(num(cell(row, base + 13)), tr.reduce((a, c) => a + rawF(c), 0));

  return {
    vacc, zeroDose, garcons, filles, ages,
    flaconsRecus: num(cell(row, gest)),
    flaconsUtil: num(cell(row, gest + 1)),
    flaconsRendus: num(cell(row, gest + 2)),
    flaconsPerdus: num(cell(row, gest + 3)),
  };
}

function readRowVacc(row: unknown[]): RowVacc {
  return {
    nvpo2: readPolio(row, C.nvpo2Synth, { m0_5: 50, m6_11: 58, m12_59: 70, a5_9: 82 }, C.nvpo2Gestion),
    vpob: readPolio(row, C.vpobSynth, { m0_5: 54, m6_11: 62, m12_59: 74, a5_9: 86 }, C.vpobGestion),
    rr: readRR(row, C.rrSynth, C.rrGestion),
  };
}

/** Première ligne de données : on cherche l'en‑tête « Province » en colonne A. */
function firstDataIndex(rows: unknown[][]): number {
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    if (/^province$/i.test(str(cell(rows[i] ?? [], 1)))) return i + 1;
  }
  return 5;
}

/** Étiquette d'un jour de campagne : Jour1..Jour5 → J1..J5, Jour6 → Ratissage. */
export function jourLabel(day: number): string {
  return day >= 6 ? "Ratissage" : `J${day}`;
}

/* ─── Parseur principal ────────────────────────────────────────────────── */

export function parseMasque(buffer: ArrayBuffer, fileName: string): MasqueData {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });

  const synthName = wb.SheetNames.find((n) => /^synth[eè]se$/i.test(n.trim()));
  const synth = synthName ? wb.Sheets[synthName] : undefined;
  if (!synth) throw new Error("Feuille « Synthèse » introuvable dans le masque de saisie.");
  const rows = XLSX.utils.sheet_to_json<unknown[]>(synth, { header: 1, blankrows: false });
  const start = firstDataIndex(rows);

  // ── Feuilles « JourN » (J1..J5 + Jour6 = ratissage) ─────────────────────
  const jourSheets: { day: number; sheet: XLSX.WorkSheet }[] = [];
  for (const name of wb.SheetNames) {
    const m = /^jour\s*(\d+)$/i.exec(name.trim());
    if (m) jourSheets.push({ day: Number(m[1]), sheet: wb.Sheets[name] });
  }
  jourSheets.sort((a, b) => a.day - b.day);

  type DailyCell = { nvpo2: number; vpob: number; rr: number; attendus: number; recus: number };
  const dailyMaps: Map<string, DailyCell>[] = jourSheets.map(({ sheet }) => {
    const m = new Map<string, DailyCell>();
    const jrows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, blankrows: false });
    const js = firstDataIndex(jrows);
    for (let i = js; i < jrows.length; i++) {
      const r = jrows[i];
      if (!r) continue;
      const province = str(cell(r, C.province));
      const antenne = str(cell(r, C.antenne));
      const zs = str(cell(r, C.zs));
      const as = str(cell(r, C.as));
      if (!province || !as) continue;
      if (isRecapRow(province, antenne, zs, as)) continue;
      const v = readRowVacc(r);
      m.set(asKey(province, zs, as), {
        nvpo2: v.nvpo2.vacc,
        vpob: v.vpob.vacc,
        rr: v.rr.vacc,
        attendus: num(cell(r, C.vaccAttendus)),
        recus: num(cell(r, C.vaccRecus)),
      });
    }
    return m;
  });
  // Jours effectivement renseignés : on garde jusqu'au dernier jour portant des
  // données (rapports reçus ou vaccinés), au minimum le J1.
  let lastDay = 0;
  dailyMaps.forEach((m, idx) => {
    for (const v of m.values()) {
      if (v.recus > 0 || v.nvpo2 > 0 || v.vpob > 0 || v.rr > 0) { lastDay = idx + 1; break; }
    }
  });
  const nbJours = Math.max(1, Math.min(lastDay || 1, jourSheets.length || 1));
  const jourLabels = jourSheets.slice(0, nbJours).map((j) => jourLabel(j.day));

  // ── Feuille « Synthèse » ────────────────────────────────────────────────
  const records: ASRecord[] = [];
  for (let i = start; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const province = str(cell(row, C.province));
    const antenne = str(cell(row, C.antenne));
    const zs = str(cell(row, C.zs));
    const as = str(cell(row, C.as));
    if (!province || !as) continue;
    if (isRecapRow(province, antenne, zs, as)) continue;

    const key = asKey(province, zs, as);
    const v = readRowVacc(row);
    const dailyOf = (sel: (d: DailyCell) => number): number[] =>
      dailyMaps.slice(0, nbJours).map((m) => { const d = m.get(key); return d ? sel(d) : 0; });

    records.push({
      province, antenne, zs, as,
      popTotale: num(cell(row, C.popTotale)),
      menagesPrevus: num(cell(row, C.menagesPrevus)),
      menagesVisites: num(cell(row, C.menagesVisites)),
      mosoAttendus: num(cell(row, C.mosoAttendus)),
      mosoRecus: num(cell(row, C.mosoRecus)),
      pers15: num(cell(row, C.pers15Total)),
      refusSignales: num(cell(row, C.refusSignales)),
      refusGeres: num(cell(row, C.refusGeres)),
      vaccAttendus: num(cell(row, C.vaccAttendus)),
      vaccRecus: num(cell(row, C.vaccRecus)),
      dailyReports: jourLabels.map((label, idx) => {
        const d = dailyMaps[idx]?.get(key);
        return { label, attendus: d?.attendus ?? 0, recus: d?.recus ?? 0 };
      }),
      ciblePolio: num(cell(row, C.ciblePolio)),
      cibleRR: num(cell(row, C.cibleRR)),
      nvpo2: { ...v.nvpo2, daily: dailyOf((d) => d.nvpo2) },
      vpob: { ...v.vpob, daily: dailyOf((d) => d.vpob) },
      rr: { ...v.rr, daily: dailyOf((d) => d.rr) },
      survPFA: num(cell(row, C.survPFA)),
      survRougeole: num(cell(row, C.survRougeole)),
      survFJ: num(cell(row, C.survFJ)),
      survTNN: num(cell(row, C.survTNN)),
      mapiNonGraves: num(cell(row, C.mapiNonGraves)),
      mapiGraves: num(cell(row, C.mapiGraves)),
      aidantsIdent: [0, 1, 2].map((k) => num(cell(row, C.aidantsIdent + k))),
      aidantsRecup: [0, 1, 2].map((k) => num(cell(row, C.aidantsRecup + k))),
      pevIdent: ANTIGENES.map((_, k) => num(cell(row, C.pevIdent + k))),
      pevRecup: ANTIGENES.map((_, k) => num(cell(row, C.pevRecup + k))),
    });
  }

  if (records.length === 0) {
    throw new Error("Aucune Aire de Santé trouvée dans la feuille « Synthèse ». Vérifiez que le bon masque de saisie est importé.");
  }
  const cleanRecords = sanitizeRecords(records);

  // ── Méta‑données (feuille « Donnees de base ») ─────────────────────────
  let periode = "";
  let pays = "RD CONGO";
  let dateDebut = "";
  let dateFin = "";
  const baseName = wb.SheetNames.find((n) => /donn[eé]es de base/i.test(n));
  if (baseName) {
    const r0 = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[baseName], { header: 1, blankrows: false });
    const l1 = r0[0] ?? [];
    const l2 = r0[1] ?? [];
    periode = str(cell(l1, 2));
    const paysVal = str(cell(l1, 5));
    if (paysVal) pays = paysVal;
    dateDebut = toISODate(cell(l2, 2));
    dateFin = toISODate(cell(l2, 4));
  }

  const province = mostCommon(cleanRecords.map((r) => r.province));
  return {
    meta: {
      pays, periode, dateDebut, dateFin, province,
      antennes: unique(cleanRecords.map((r) => r.antenne).filter(Boolean)),
      zones: unique(cleanRecords.map((r) => r.zs).filter(Boolean)),
      importedAt: new Date().toISOString(),
      fileName,
      nbAires: cleanRecords.length,
      nbJours,
      jourLabels,
      schema: MASQUE_SCHEMA,
    },
    records: cleanRecords,
  };
}

function toISODate(v: unknown): string {
  if (v instanceof Date && !isNaN(v.getTime())) {
    // xlsx produit des dates à minuit (fuseau ambigu) : on recentre à midi pour
    // éviter le glissement d'un jour selon le fuseau horaire.
    const c = new Date(v.getTime() + 12 * 3600 * 1000);
    const y = c.getUTCFullYear();
    const m = String(c.getUTCMonth() + 1).padStart(2, "0");
    const d = String(c.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof v === "number" && Number.isFinite(v)) {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  const s = str(v);
  const m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/.exec(s);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  return s;
}

function unique(arr: string[]): string[] {
  return Array.from(new Set(arr)).sort((a, b) => a.localeCompare(b, "fr"));
}
function mostCommon(arr: string[]): string {
  const counts = new Map<string, number>();
  for (const v of arr) counts.set(v, (counts.get(v) ?? 0) + 1);
  let best = "";
  let max = -1;
  for (const [k, c] of counts) if (c > max) { max = c; best = k; }
  return best;
}

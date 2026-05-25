/**
 * Parseur du masque de saisie de la campagne polio synchronisée avec l'Angola.
 *
 * Le masque est un classeur Excel identique pour toutes les provinces / antennes /
 * zones de santé. La feuille « Synthèse » contient, par Aire de Santé, l'ensemble
 * des indicateurs déjà agrégés sur la campagne. Les feuilles « Jour1 », « Jour2 »
 * etc. permettent de reconstituer le détail journalier (vaccinés et complétude
 * par jour). La feuille « Donnees de base » fournit l'en-tête (période, pays).
 *
 * Seule la composante POLIO (nVPO2 et VPOb) est conservée — la Rougeole-Rubéole
 * (RR) est entièrement ignorée puisqu'elle ne fait pas partie de cette campagne
 * synchronisée avec l'Angola (co-administration polio uniquement).
 */

import * as XLSX from "xlsx";

/** Indices de colonnes (1-based) de la feuille « Synthèse » — voir masque modèle. */
const C = {
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
  refusSignales: 30,
  refusGeres: 31,
  // Complétude rapports vaccination (bloc vaccination)
  vaccAttendus: 90,
  vaccRecus: 91,
  // Cible 0-59 mois (commune nVPO2 / VPOb — co-administration, même tranche d'âge)
  cible059: 93,
  // nVPO2 — vaccination 0-59 mois
  nvpo2ZeroDose011: 122,
  nvpo2ZeroDose1259: 125,
  nvpo2Vacc059Total: 130,
  // nVPO2 — performances équipes
  nvpo2Rural: 137,
  nvpo2Urbain: 139,
  // nVPO2 — cibles (repli sur cible 0-59 si bloc dédié vide)
  nvpo2CibleDenombre: 144,
  // nVPO2 — gestion vaccin (flacons reçues / utilisées / rendues / perdus)
  nvpo2FlaconsRecus: 150,
  nvpo2FlaconsRendus: 151,
  nvpo2Perdus: 153,
  nvpo2TauxPerte: 154,
  nvpo2FlaconsUtil: 155,
  // VPOb — vaccination 0-59 mois
  vpobZeroDose011: 184,
  vpobZeroDose1259: 187,
  vpobVacc059Total: 192,
  // VPOb — performances équipes
  vpobRural: 199,
  vpobUrbain: 201,
  // VPOb — cibles
  vpobCibleDenombre: 206,
  // VPOb — gestion vaccin
  vpobFlaconsRecus: 212,
  vpobFlaconsRendus: 213,
  vpobPerdus: 215,
  vpobTauxPerte: 216,
  vpobFlaconsUtil: 217,
  // MAPI
  mapiMineures: 222,
  mapiGraves: 223,
  // Surveillance des MPV — Recherche active des cas de MEV (bloc vaccination).
  survPFA: 218,
  survRougeole: 219,
  survFJ: 220,
  survTNN: 221,
  // Récupérations PEV de routine (co-administration)
  recup011: 258,
  recup1223: 292,
  recup2459: 326,
} as const;

/**
 * Antigènes du « Renforcement PEV systématique » récupérés pendant la campagne.
 * `col` = colonne EV (enfants vaccinés) de la tranche 0-11 mois. Le masque répète
 * chaque antigène pour 0-11, 12-23 et 24-59 mois avec un pas de 34 colonnes ; le
 * total « enfants vaccinés » somme les trois tranches d'âge.
 */
export const ANTIGENES: { key: string; label: string; col: number }[] = [
  { key: "BCG", label: "BCG", col: 225 },
  { key: "VPI1", label: "VPI 1", col: 227 },
  { key: "DTC1", label: "DTC 1", col: 229 },
  { key: "DTC2", label: "DTC 2", col: 231 },
  { key: "DTC3", label: "DTC 3", col: 233 },
  { key: "PCV1", label: "PCV 1", col: 236 },
  { key: "PCV2", label: "PCV 2", col: 238 },
  { key: "PCV3", label: "PCV 3", col: 240 },
  { key: "ROTA1", label: "Rota 1", col: 243 },
  { key: "ROTA2", label: "Rota 2", col: 245 },
  { key: "ROTA3", label: "Rota 3", col: 247 },
  { key: "VAR1", label: "VAR 1", col: 250 },
  { key: "VAR2", label: "VAR 2", col: 252 },
  { key: "VAA", label: "VAA", col: 254 },
  { key: "VPI2", label: "VPI 2", col: 256 },
];
const ANTIGENE_OFFSETS = [0, 34, 68]; // tranches 0-11, 12-23, 24-59 mois

export interface DailyValue {
  /** Index du jour (1-based : Jour1 = 1, Jour2 = 2, …) */
  day: number;
  /** Étiquette « Jour1 », « Jour2 », … */
  label: string;
  vaccines: number;
  rapportsRecus: number;
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
  // nVPO2
  nvpo2Vacc: number;
  nvpo2CibleExtrap: number;
  nvpo2CibleDenombre: number;
  nvpo2ZeroDose: number;
  nvpo2FlaconsRecus: number;
  nvpo2FlaconsRendus: number;
  nvpo2Perdus: number;
  nvpo2FlaconsUtil: number;
  nvpo2TauxPerte: number | null;
  nvpo2Rural: number;
  nvpo2Urbain: number;
  /** Vaccinés nVPO2 par jour de campagne (1-indexé). */
  nvpo2Daily: DailyValue[];
  // VPOb
  vpobVacc: number;
  vpobCibleExtrap: number;
  vpobCibleDenombre: number;
  vpobZeroDose: number;
  vpobFlaconsRecus: number;
  vpobFlaconsRendus: number;
  vpobPerdus: number;
  vpobFlaconsUtil: number;
  vpobTauxPerte: number | null;
  vpobRural: number;
  vpobUrbain: number;
  vpobDaily: DailyValue[];
  // récup + MAPI
  recup: number;
  mapiMineures: number;
  mapiGraves: number;
  /** Enfants vaccinés par antigène (EV), somme des tranches d'âge — ordre = ANTIGENES. */
  antigenesEV: number[];
  // Surveillance des MPV (cas notifiés)
  survPFA: number;
  survRougeole: number;
  survFJ: number;
  survTNN: number;
}

export interface MasqueData {
  meta: {
    pays: string;
    periode: string;
    province: string;
    antennes: string[];
    zones: string[];
    importedAt: string;
    fileName: string;
    nbAires: number;
    /** Nombre de jours de campagne effectivement saisis dans le masque. */
    nbJours: number;
    /** Étiquettes des jours (« Jour1 », « Jour2 »…). */
    jourLabels: string[];
  };
  records: ASRecord[];
}

function num(v: unknown): number {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", ".").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function str(v: unknown): string {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

/** Lit une cellule (1-based col) d'une ligne array-of-arrays. */
function cell(row: unknown[], col1: number): unknown {
  return row[col1 - 1];
}

const TOTAL_RE = /total/i;

/** Clé d'identification d'une Aire de Santé (insensible casse / accents / ponctuation). */
function asKey(province: string, zs: string, as: string): string {
  const norm = (s: string) =>
    s
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .replace(/[^A-Z0-9]/g, "");
  return `${norm(province)}|${norm(zs)}|${norm(as)}`;
}

export function parseMasque(buffer: ArrayBuffer, fileName: string): MasqueData {
  const wb = XLSX.read(buffer, { type: "array" });

  const synthName = wb.SheetNames.find((n) => /^synth/i.test(n.trim()) && !/ps/i.test(n)) ?? "Synthèse";
  const synth = wb.Sheets[synthName];
  if (!synth) throw new Error("Feuille « Synthèse » introuvable dans le masque de saisie.");

  const rows = XLSX.utils.sheet_to_json<unknown[]>(synth, { header: 1, blankrows: false });

  // ── 1. Extraction des feuilles « JourN » (vaccination quotidienne) ─────────
  const jourSheets: { day: number; label: string; sheet: XLSX.WorkSheet }[] = [];
  for (const name of wb.SheetNames) {
    const m = /^jour\s*(\d+)$/i.exec(name.trim());
    if (m && !/^ps/i.test(name)) {
      jourSheets.push({ day: Number(m[1]), label: name.trim(), sheet: wb.Sheets[name] });
    }
  }
  jourSheets.sort((a, b) => a.day - b.day);

  // Pré-calcul : pour chaque jour, une map AS → {vacc nvpo2, vacc vpob, rapports reçus}.
  const dailyMaps: Map<string, { nvpo2: number; vpob: number; recus: number }>[] = jourSheets.map(
    ({ sheet }) => {
      const m = new Map<string, { nvpo2: number; vpob: number; recus: number }>();
      const jrows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, blankrows: false });
      for (let i = 3; i < jrows.length; i++) {
        const r = jrows[i];
        if (!r) continue;
        const province = str(cell(r, C.province));
        const zs = str(cell(r, C.zs));
        const as = str(cell(r, C.as));
        if (!province || !as) continue;
        if (TOTAL_RE.test(zs) || TOTAL_RE.test(as)) continue;
        m.set(asKey(province, zs, as), {
          nvpo2: num(cell(r, C.nvpo2Vacc059Total)),
          vpob: num(cell(r, C.vpobVacc059Total)),
          recus: num(cell(r, C.vaccRecus)),
        });
      }
      return m;
    }
  );

  // ── 2. Parcours de la feuille « Synthèse » ────────────────────────────────
  const records: ASRecord[] = [];
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const province = str(cell(row, C.province));
    const zs = str(cell(row, C.zs));
    const as = str(cell(row, C.as));
    if (!as) continue;
    if (TOTAL_RE.test(zs) || TOTAL_RE.test(as)) continue;
    if (!province) continue;

    const nvpo2TauxPerteRaw = cell(row, C.nvpo2TauxPerte);
    const vpobTauxPerteRaw = cell(row, C.vpobTauxPerte);
    const key = asKey(province, zs, as);

    const nvpo2Daily: DailyValue[] = jourSheets.map((j, idx) => ({
      day: j.day,
      label: j.label,
      vaccines: dailyMaps[idx].get(key)?.nvpo2 ?? 0,
      rapportsRecus: dailyMaps[idx].get(key)?.recus ?? 0,
    }));
    const vpobDaily: DailyValue[] = jourSheets.map((j, idx) => ({
      day: j.day,
      label: j.label,
      vaccines: dailyMaps[idx].get(key)?.vpob ?? 0,
      rapportsRecus: dailyMaps[idx].get(key)?.recus ?? 0,
    }));

    records.push({
      province,
      antenne: str(cell(row, C.antenne)),
      zs,
      as,
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
      nvpo2Vacc: num(cell(row, C.nvpo2Vacc059Total)),
      nvpo2CibleExtrap: num(cell(row, C.cible059)),
      nvpo2CibleDenombre: num(cell(row, C.nvpo2CibleDenombre)),
      nvpo2ZeroDose: num(cell(row, C.nvpo2ZeroDose011)) + num(cell(row, C.nvpo2ZeroDose1259)),
      nvpo2FlaconsRecus: num(cell(row, C.nvpo2FlaconsRecus)),
      nvpo2FlaconsRendus: num(cell(row, C.nvpo2FlaconsRendus)),
      nvpo2Perdus: num(cell(row, C.nvpo2Perdus)),
      nvpo2FlaconsUtil: num(cell(row, C.nvpo2FlaconsUtil)),
      nvpo2TauxPerte:
        nvpo2TauxPerteRaw === "" || nvpo2TauxPerteRaw == null ? null : num(nvpo2TauxPerteRaw),
      nvpo2Rural: num(cell(row, C.nvpo2Rural)),
      nvpo2Urbain: num(cell(row, C.nvpo2Urbain)),
      nvpo2Daily,
      vpobVacc: num(cell(row, C.vpobVacc059Total)),
      vpobCibleExtrap: num(cell(row, C.cible059)),
      vpobCibleDenombre: num(cell(row, C.vpobCibleDenombre)),
      vpobZeroDose: num(cell(row, C.vpobZeroDose011)) + num(cell(row, C.vpobZeroDose1259)),
      vpobFlaconsRecus: num(cell(row, C.vpobFlaconsRecus)),
      vpobFlaconsRendus: num(cell(row, C.vpobFlaconsRendus)),
      vpobPerdus: num(cell(row, C.vpobPerdus)),
      vpobFlaconsUtil: num(cell(row, C.vpobFlaconsUtil)),
      vpobTauxPerte:
        vpobTauxPerteRaw === "" || vpobTauxPerteRaw == null ? null : num(vpobTauxPerteRaw),
      vpobRural: num(cell(row, C.vpobRural)),
      vpobUrbain: num(cell(row, C.vpobUrbain)),
      vpobDaily,
      recup:
        num(cell(row, C.recup011)) +
        num(cell(row, C.recup1223)) +
        num(cell(row, C.recup2459)),
      mapiMineures: num(cell(row, C.mapiMineures)),
      mapiGraves: num(cell(row, C.mapiGraves)),
      antigenesEV: ANTIGENES.map((a) =>
        ANTIGENE_OFFSETS.reduce((sum, off) => sum + num(cell(row, a.col + off)), 0)
      ),
      survPFA: num(cell(row, C.survPFA)),
      survRougeole: num(cell(row, C.survRougeole)),
      survFJ: num(cell(row, C.survFJ)),
      survTNN: num(cell(row, C.survTNN)),
    });
  }

  if (records.length === 0) {
    throw new Error(
      "Aucune Aire de Santé trouvée dans la feuille « Synthèse ». Vérifiez que le bon masque de saisie est importé."
    );
  }

  // ── 3. Méta-données du classeur ──────────────────────────────────────────
  let periode = "";
  let pays = "RD CONGO";
  const baseName = wb.SheetNames.find((n) => /donn[eé]es de base/i.test(n));
  if (baseName) {
    const base = wb.Sheets[baseName];
    const baseRows = XLSX.utils.sheet_to_json<unknown[]>(base, { header: 1, blankrows: false });
    const r0 = baseRows[0] ?? [];
    periode = str(cell(r0, 2));
    const paysVal = str(cell(r0, 5));
    if (paysVal) pays = paysVal;
  }

  const province = mostCommon(records.map((r) => r.province));
  const antennes = unique(records.map((r) => r.antenne).filter(Boolean));
  const zones = unique(records.map((r) => r.zs).filter(Boolean));
  const jourLabels = jourSheets.map((j) => j.label);

  return {
    meta: {
      pays,
      periode,
      province,
      antennes,
      zones,
      importedAt: new Date().toISOString(),
      fileName,
      nbAires: records.length,
      nbJours: jourSheets.length,
      jourLabels,
    },
    records,
  };
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

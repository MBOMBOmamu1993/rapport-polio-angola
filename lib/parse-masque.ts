/**
 * Parseur du masque de saisie de la campagne polio synchronisée avec l'Angola.
 *
 * Le masque est un classeur Excel identique pour toutes les provinces / antennes /
 * zones de santé. La feuille « Synthèse » contient, par Aire de Santé, l'ensemble
 * des indicateurs déjà agrégés sur la campagne. La feuille « Donnees de base »
 * fournit l'en-tête (période, pays, province).
 *
 * On ne conserve que la partie POLIO (nVPO2 et VPOb) — toute la composante
 * Rougeole-Rubéole (RR) est ignorée.
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
  // nVPO2 — gestion vaccin
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
  vpobPerdus: 215,
  vpobTauxPerte: 216,
  vpobFlaconsUtil: 217,
  // MAPI
  mapiMineures: 222,
  mapiGraves: 223,
  // Récupérations PEV de routine (co-administration)
  recup011: 258,
  recup1223: 292,
  recup2459: 326,
} as const;

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
  nvpo2Perdus: number;
  nvpo2FlaconsUtil: number;
  nvpo2TauxPerte: number | null;
  nvpo2Rural: number;
  nvpo2Urbain: number;
  // VPOb
  vpobVacc: number;
  vpobCibleExtrap: number;
  vpobCibleDenombre: number;
  vpobZeroDose: number;
  vpobPerdus: number;
  vpobFlaconsUtil: number;
  vpobTauxPerte: number | null;
  vpobRural: number;
  vpobUrbain: number;
  // récup + MAPI
  recup: number;
  mapiMineures: number;
  mapiGraves: number;
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

export function parseMasque(buffer: ArrayBuffer, fileName: string): MasqueData {
  const wb = XLSX.read(buffer, { type: "array" });

  const synthName = wb.SheetNames.find((n) => /synth/i.test(n) && !/ps/i.test(n)) ?? "Synthèse";
  const synth = wb.Sheets[synthName];
  if (!synth) throw new Error("Feuille « Synthèse » introuvable dans le masque de saisie.");

  const rows = XLSX.utils.sheet_to_json<unknown[]>(synth, { header: 1, blankrows: false });

  const records: ASRecord[] = [];
  // Les 3 premières lignes sont des en-têtes ; les données commencent en ligne 4 (index 3).
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const province = str(cell(row, C.province));
    const zs = str(cell(row, C.zs));
    const as = str(cell(row, C.as));
    if (!as) continue;
    if (TOTAL_RE.test(zs) || TOTAL_RE.test(as)) continue; // lignes de sous-total
    if (!province) continue;

    const nvpo2TauxPerte = cell(row, C.nvpo2TauxPerte);
    const vpobTauxPerte = cell(row, C.vpobTauxPerte);

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
      nvpo2Perdus: num(cell(row, C.nvpo2Perdus)),
      nvpo2FlaconsUtil: num(cell(row, C.nvpo2FlaconsUtil)),
      nvpo2TauxPerte: nvpo2TauxPerte === "" || nvpo2TauxPerte == null ? null : num(nvpo2TauxPerte),
      nvpo2Rural: num(cell(row, C.nvpo2Rural)),
      nvpo2Urbain: num(cell(row, C.nvpo2Urbain)),
      vpobVacc: num(cell(row, C.vpobVacc059Total)),
      vpobCibleExtrap: num(cell(row, C.cible059)),
      vpobCibleDenombre: num(cell(row, C.vpobCibleDenombre)),
      vpobZeroDose: num(cell(row, C.vpobZeroDose011)) + num(cell(row, C.vpobZeroDose1259)),
      vpobPerdus: num(cell(row, C.vpobPerdus)),
      vpobFlaconsUtil: num(cell(row, C.vpobFlaconsUtil)),
      vpobTauxPerte: vpobTauxPerte === "" || vpobTauxPerte == null ? null : num(vpobTauxPerte),
      vpobRural: num(cell(row, C.vpobRural)),
      vpobUrbain: num(cell(row, C.vpobUrbain)),
      recup: num(cell(row, C.recup011)) + num(cell(row, C.recup1223)) + num(cell(row, C.recup2459)),
      mapiMineures: num(cell(row, C.mapiMineures)),
      mapiGraves: num(cell(row, C.mapiGraves)),
    });
  }

  if (records.length === 0) {
    throw new Error(
      "Aucune Aire de Santé trouvée dans la feuille « Synthèse ». Vérifiez que le bon masque de saisie est importé."
    );
  }

  // Méta : période depuis « Donnees de base », province majoritaire depuis les données.
  let periode = "";
  let pays = "RD CONGO";
  const baseName = wb.SheetNames.find((n) => /donnees de base/i.test(n) || /données de base/i.test(n));
  if (baseName) {
    const base = wb.Sheets[baseName];
    const baseRows = XLSX.utils.sheet_to_json<unknown[]>(base, { header: 1, blankrows: false });
    const r0 = baseRows[0] ?? [];
    // « PERIODE: » en col1, valeur en col2 ; « PAYS : » col4, valeur col5.
    periode = str(cell(r0, 2));
    const paysVal = str(cell(r0, 5));
    if (paysVal) pays = paysVal;
  }

  const province = mostCommon(records.map((r) => r.province));
  const antennes = unique(records.map((r) => r.antenne).filter(Boolean));
  const zones = unique(records.map((r) => r.zs).filter(Boolean));

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

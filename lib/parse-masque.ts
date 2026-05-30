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
  // nVPO2 — gestion vaccin. Colonnes réelles du masque :
  //  149 = TOTAL flacons reçus journalièrement par les équipes (mouvement de flacons).
  //        C'est le total reçu faisant foi ET la base du taux de perte du masque :
  //        taux = 1 − vaccinés / (149 × 50).
  //  150 = Flacons complémentaires reçus — colonne AUTO « Ne pas remplir » : valeur
  //        d'assistance calculée par le masque, à NE PAS additionner au total reçu
  //        (col. 149 le contient déjà). L'ajouter gonflait le nombre de flacons reçus.
  //  152 = Flacons inutilisables (entamés, cassés, virés…)
  //  153 = Flacons perdus
  //  154 = Taux de perte % (calculé par le masque)
  //  155 = Total de flacons utilisables RESTANTS dans la CDF (stock retourné)
  nvpo2FlaconsRecusJour: 149,
  nvpo2Inutilisables: 152,
  nvpo2Perdus: 153,
  nvpo2TauxPerte: 154,
  nvpo2StockCDF: 155,
  // VPOb — vaccination 0-59 mois
  vpobZeroDose011: 184,
  vpobZeroDose1259: 187,
  vpobVacc059Total: 192,
  // VPOb — performances équipes
  vpobRural: 199,
  vpobUrbain: 201,
  // VPOb — cibles
  vpobCibleDenombre: 206,
  // VPOb — gestion vaccin (mêmes colonnes que nVPO2, doses/flacon = 20).
  //  211 = TOTAL flacons reçus journalièrement (fait foi). 212 = complémentaires
  //  « Ne pas remplir » (auto) — à NE PAS additionner au total reçu.
  vpobFlaconsRecusJour: 211,
  vpobInutilisables: 214,
  vpobPerdus: 215,
  vpobTauxPerte: 216,
  vpobStockCDF: 217,
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
  /** Rapports de vaccination reçus ce jour (feuille JourN, col. « Recus »). */
  rapportsRecus: number;
  /** Rapports de vaccination attendus ce jour (feuille JourN, col. « Attendus »). */
  rapportsAttendus: number;
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

/**
 * Normalise un taux de perte lu dans le masque. Le masque le stocke en fraction
 * (ex. 0,0392 affiché « 3,9 % ») ; on le ramène en pourcentage. Renvoie null si vide.
 */
function normTaux(v: unknown): number | null {
  if (v === "" || v == null) return null;
  const n = num(v);
  if (!Number.isFinite(n)) return null;
  return Math.abs(n) <= 1.5 ? n * 100 : n;
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

/**
 * Sous-titres de récapitulation insérés dans la Synthèse, surtout dans les
 * masques importés au niveau provincial : avant de lister les Zones de Santé
 * d'une antenne, le masque place une ligne titre reprenant le nom du parent
 * préfixé (« Ant.LUIZA », « ant Luiza », « ZS.KALOMBA »…). Ce ne sont pas des
 * unités réelles : laissées en place, le parent réapparaît comme enfant dans les
 * listes déroulantes et le rapport, et ses effectifs sont comptés deux fois.
 * On exige un séparateur (point, espace, tiret, deux-points) après le mot-clé
 * pour ne pas écarter de vrais noms commençant par ces lettres (ex. « ANTONIO »).
 */
const RECAP_RE = /^\s*(ant|antenne|zs|zone|prov|province)\b[\s.:\-]+\S/i;
export function isRecapLabel(s: string): boolean {
  return RECAP_RE.test(s);
}

/** Normalise un libellé pour comparaison (sans accents/casse/espaces superflus). */
function normLabel(s: string): string {
  return s
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}

/**
 * Vrai si la ligne est un titre/sous-titre de récapitulation et non une Aire de
 * Santé réelle. Deux formes rencontrées :
 *  - un libellé préfixé du niveau parent dans une colonne enfant
 *    (« Ant.LUIZA », « ant Luiza », « Prov. KASAÏ »…) ;
 *  - un titre du niveau national où une colonne enfant reprend exactement le nom
 *    de son parent (la colonne Antenne contient le nom de la Province avant de
 *    lister ses antennes). On compare uniquement Antenne = Province, car au
 *    niveau inférieur un vrai chef-lieu porte légitimement le nom de sa zone
 *    (Aire de Santé LUIZA dans la Zone de Santé LUIZA).
 */
export function isRecapRow(province: string, antenne: string, zs: string, as: string): boolean {
  if (isRecapLabel(antenne) || isRecapLabel(zs) || isRecapLabel(as)) return true;
  if (antenne && province && normLabel(antenne) === normLabel(province)) return true;
  return false;
}

/**
 * Nettoie une liste d'Aires de Santé des lignes d'agrégat/sous-total qui s'y sont
 * glissées et qui faussaient les calculs (double comptage). Deux détections
 * complémentaires et indépendantes du libellé exact :
 *
 *  1. Récapitulatif par libellé (`isRecapRow`) — préfixe parent ou Antenne =
 *     Province.
 *  2. Détection numérique : une ligne dont les effectifs égalent la somme d'au
 *     moins deux voisines partageant le même parent est, par construction, un
 *     sous-total ; l'inclure double les totaux. On exige une concordance sur
 *     plusieurs indicateurs indépendants (rapports attendus, cible, vaccinés
 *     nVPO2 et VPOb) afin qu'aucune Aire de Santé réelle — même un chef-lieu
 *     homonyme de sa zone — ne soit écartée par coïncidence.
 *
 * La détection numérique fonctionne même si le sous-total n'a aucun préfixe
 * reconnaissable : c'est le garde-fou contre les anomalies non anticipées.
 */
export function sanitizeRecords(records: ASRecord[]): ASRecord[] {
  const kept = records.filter((r) => !isRecapRow(r.province, r.antenne, r.zs, r.as));

  const metricsOf = (r: ASRecord): number[] => [
    r.vaccAttendus,
    r.nvpo2CibleExtrap,
    r.nvpo2Vacc,
    r.vpobVacc,
  ];
  const approxEq = (a: number, b: number): boolean =>
    Math.abs(a - b) <= Math.max(1, Math.abs(b) * 0.001);

  const removed = new Set<ASRecord>();
  const detectSubtotals = (groupKey: (r: ASRecord) => string): void => {
    const groups = new Map<string, ASRecord[]>();
    for (const r of kept) {
      if (removed.has(r)) continue;
      const k = groupKey(r);
      const g = groups.get(k);
      if (g) g.push(r);
      else groups.set(k, [r]);
    }
    for (const group of groups.values()) {
      // Besoin d'au moins un candidat sous-total et deux enfants réels.
      if (group.length < 3) continue;
      for (const cand of group) {
        if (removed.has(cand)) continue;
        const others = group.filter((x) => x !== cand && !removed.has(x));
        if (others.length < 2) continue;
        const m = metricsOf(cand);
        const sums = m.map((_, i) => others.reduce((acc, x) => acc + metricsOf(x)[i], 0));
        const nonTrivial = m.some((v, i) => v > 0 && sums[i] > 0);
        const allMatch = m.every((v, i) => approxEq(v, sums[i]));
        if (nonTrivial && allMatch) removed.add(cand);
      }
    }
  };

  // Du parent le plus haut au plus bas : antenne (résume ses ZS), province
  // (résume ses antennes), zone (résume ses AS).
  detectSubtotals((r) => `${normLabel(r.province)}||${normLabel(r.antenne)}`);
  detectSubtotals((r) => normLabel(r.province));
  detectSubtotals((r) => `${normLabel(r.province)}||${normLabel(r.antenne)}||${normLabel(r.zs)}`);

  return kept.filter((r) => !removed.has(r));
}

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

  // Pré-calcul : pour chaque jour, une map AS → {vacc nvpo2, vacc vpob, rapports
  // attendus et reçus DU JOUR}. Les feuilles JourN ont la même structure de
  // colonnes que la Synthèse : col. 90 = Attendus du jour, col. 91 = Reçus du jour.
  type DailyCell = { nvpo2: number; vpob: number; recus: number; attendus: number };
  const dailyMaps: Map<string, DailyCell>[] = jourSheets.map(
    ({ sheet }) => {
      const m = new Map<string, DailyCell>();
      const jrows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, blankrows: false });
      for (let i = 3; i < jrows.length; i++) {
        const r = jrows[i];
        if (!r) continue;
        const province = str(cell(r, C.province));
        const antenne = str(cell(r, C.antenne));
        const zs = str(cell(r, C.zs));
        const as = str(cell(r, C.as));
        if (!province || !as) continue;
        if (TOTAL_RE.test(zs) || TOTAL_RE.test(as)) continue;
        if (isRecapRow(province, antenne, zs, as)) continue;
        m.set(asKey(province, zs, as), {
          nvpo2: num(cell(r, C.nvpo2Vacc059Total)),
          vpob: num(cell(r, C.vpobVacc059Total)),
          recus: num(cell(r, C.vaccRecus)),
          attendus: num(cell(r, C.vaccAttendus)),
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
    const antenne = str(cell(row, C.antenne));
    const zs = str(cell(row, C.zs));
    const as = str(cell(row, C.as));
    if (!as) continue;
    if (TOTAL_RE.test(zs) || TOTAL_RE.test(as)) continue;
    if (!province) continue;
    if (isRecapRow(province, antenne, zs, as)) continue;

    const nvpo2TauxPerteRaw = cell(row, C.nvpo2TauxPerte);
    const vpobTauxPerteRaw = cell(row, C.vpobTauxPerte);
    const key = asKey(province, zs, as);

    const nvpo2Daily: DailyValue[] = jourSheets.map((j, idx) => ({
      day: j.day,
      label: j.label,
      vaccines: dailyMaps[idx].get(key)?.nvpo2 ?? 0,
      rapportsRecus: dailyMaps[idx].get(key)?.recus ?? 0,
      rapportsAttendus: dailyMaps[idx].get(key)?.attendus ?? 0,
    }));
    const vpobDaily: DailyValue[] = jourSheets.map((j, idx) => ({
      day: j.day,
      label: j.label,
      vaccines: dailyMaps[idx].get(key)?.vpob ?? 0,
      rapportsRecus: dailyMaps[idx].get(key)?.recus ?? 0,
      rapportsAttendus: dailyMaps[idx].get(key)?.attendus ?? 0,
    }));

    records.push({
      province,
      antenne,
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
      // Total reçu = col. 149 seule (« Total Flacons reçus journalièrement »). On
      // n'ajoute plus la col. 150 « complémentaires (Ne pas remplir) » : c'est une
      // colonne auto déjà comprise dans le total, dont l'ajout doublait les reçus.
      nvpo2FlaconsRecus: num(cell(row, C.nvpo2FlaconsRecusJour)),
      nvpo2FlaconsRendus: num(cell(row, C.nvpo2StockCDF)),
      nvpo2Perdus: num(cell(row, C.nvpo2Perdus)),
      // Base du taux de perte du masque : flacons reçus journalièrement (× 50 doses).
      nvpo2FlaconsUtil: num(cell(row, C.nvpo2FlaconsRecusJour)),
      nvpo2TauxPerte: normTaux(nvpo2TauxPerteRaw),
      nvpo2Rural: num(cell(row, C.nvpo2Rural)),
      nvpo2Urbain: num(cell(row, C.nvpo2Urbain)),
      nvpo2Daily,
      vpobVacc: num(cell(row, C.vpobVacc059Total)),
      vpobCibleExtrap: num(cell(row, C.cible059)),
      vpobCibleDenombre: num(cell(row, C.vpobCibleDenombre)),
      vpobZeroDose: num(cell(row, C.vpobZeroDose011)) + num(cell(row, C.vpobZeroDose1259)),
      // Total reçu = col. 211 seule ; on n'ajoute plus la col. 212 « complémentaires
      // (Ne pas remplir) », colonne auto qui doublait le nombre de flacons reçus.
      vpobFlaconsRecus: num(cell(row, C.vpobFlaconsRecusJour)),
      vpobFlaconsRendus: num(cell(row, C.vpobStockCDF)),
      vpobPerdus: num(cell(row, C.vpobPerdus)),
      // Base du taux de perte du masque : flacons reçus journalièrement (× 20 doses).
      vpobFlaconsUtil: num(cell(row, C.vpobFlaconsRecusJour)),
      vpobTauxPerte: normTaux(vpobTauxPerteRaw),
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

  // Écarte les sous-totaux résiduels (y compris ceux sans préfixe reconnaissable)
  // détectés par concordance numérique, pour ne jamais doubler les totaux.
  const cleanRecords = sanitizeRecords(records);

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

  const province = mostCommon(cleanRecords.map((r) => r.province));
  const antennes = unique(cleanRecords.map((r) => r.antenne).filter(Boolean));
  const zones = unique(cleanRecords.map((r) => r.zs).filter(Boolean));
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
      nbAires: cleanRecords.length,
      nbJours: jourSheets.length,
      jourLabels,
    },
    records: cleanRecords,
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

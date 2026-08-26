/**
 * Types et fusion des extractions DHIS2 de la campagne intégrée RR-POLIO
 * (rdccampagne.hispwca.org). Module isomorphe (client + serveur) : la partie
 * réseau vit dans `dhis2-campagne.ts` (serveur uniquement).
 */

import { MASQUE_SCHEMA, type ASRecord, type MasqueData } from "./parse-masque";

/** Version du format des blocs DHIS2 mis en cache (KV) — à incrémenter si la structure change. */
export const DHIS2_BLOCK_SCHEMA = 3;

export const JOUR_LABELS = ["J1", "J2", "J3", "J4", "J5", "Ratissage"];
export const NB_JOURS = JOUR_LABELS.length;

/** Province de la campagne (ligne du sélecteur). */
export interface ProvinceInfo {
  id: string;
  /** Nom propre (« Sud Kivu », sans le préfixe DHIS2 « sk … Province »). */
  name: string;
  cibleRR: number;
  rrVacc: number;
  polioVacc: number;
  /** Vrai si la province a intégré le volet polio (RR-POLIO) ; faux = RR seule. */
  polio: boolean;
}

export interface ProvinceListPayload {
  ok: boolean;
  reason?: string;
  fetchedAt: string;
  schema: number;
  provinces: ProvinceInfo[];
}

/** Extraction complète d'une province (toutes ses Aires de Santé). */
export interface ProvinceBlock {
  ok: boolean;
  reason?: string;
  schema: number;
  fetchedAt: string;
  provinceId: string;
  province: string;
  /** Date J1 de la province (ISO) — détectée sur le volume journalier. */
  j1: string;
  /** Vrai si la province a intégré le volet polio. */
  polio: boolean;
  /**
   * Origine des résultats : "dhis2" (extraction analytics) ou "masque"
   * (dernier masque de saisie importé — prioritaire quand la province saisit
   * ses résultats dans le masque et non dans DHIS2, ex. Kasaï Central).
   */
  source?: "dhis2" | "masque";
  records: ASRecord[];
  /** Marqué par l'API quand la valeur servie provient d'un cache antérieur. */
  stale?: boolean;
}

function fmtJourFR(iso: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
  return m ? `${m[3]}/${m[2]}` : iso;
}

/** Fusionne les blocs provinciaux sélectionnés en un `MasqueData` exploitable par le pipeline. */
export function mergeBlocks(blocks: ProvinceBlock[]): MasqueData | null {
  const valid = blocks.filter((b) => b && b.ok && Array.isArray(b.records));
  if (valid.length === 0) return null;
  const records = valid.flatMap((b) => b.records);
  const provinces = uniq(valid.map((b) => b.province));
  const j1s = valid.map((b) => b.j1).filter(Boolean).sort();
  const dateDebut = j1s[0] ?? "2026-08-11";
  const dateFinJ5 = j1s.length ? addDays(j1s[j1s.length - 1], 4) : "2026-08-15";
  const periode =
    j1s.length && j1s[0] === j1s[j1s.length - 1]
      ? `du ${fmtJourFR(dateDebut)} au ${fmtJourFR(addDays(dateDebut, 4))}/2026 (+ ratissage)`
      : `J1 → J5 + ratissage (lancements du ${fmtJourFR(dateDebut)} au ${fmtJourFR(j1s[j1s.length - 1] ?? dateDebut)}/2026 selon les provinces)`;
  const latest = valid.reduce((m, b) => (b.fetchedAt > m ? b.fetchedAt : m), "");
  return {
    meta: {
      pays: "RD CONGO",
      periode,
      dateDebut,
      dateFin: dateFinJ5,
      province: provinces.length === 1 ? provinces[0] : "RD Congo",
      antennes: uniq(records.map((r) => r.antenne)),
      zones: uniq(records.map((r) => r.zs)),
      importedAt: latest || new Date().toISOString(),
      fileName: "DHIS2 rdccampagne.hispwca.org — dataset PEV_Campagne RR et Polio",
      nbAires: records.length,
      nbJours: NB_JOURS,
      jourLabels: JOUR_LABELS,
      schema: MASQUE_SCHEMA,
    },
    records,
  };
}

function addDays(iso: string, days: number): string {
  const d = new Date(`${iso}T00:00:00Z`);
  if (Number.isNaN(d.getTime())) return iso;
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function uniq(arr: string[]): string[] {
  return Array.from(new Set(arr.filter(Boolean))).sort((a, b) => a.localeCompare(b, "fr"));
}

/**
 * Accès serveur au formulaire ODK de supervision (api.whonghub.org — API Ona).
 *
 * Variables d'environnement (toutes optionnelles — valeurs par défaut = configuration
 * de la campagne intégrée RR‑Polio, identique au dépôt
 * `rr-polio-independent-monitoring-dashboard`) :
 *  - ODK_BASE_URL   (https://api.whonghub.org)
 *  - ODK_USERNAME / ODK_PASSWORD
 *  - ODK_SUPERVISION_FORM_ID (17559)
 *  - ODK_PROVINCE   (Kasai_Central — valeur du champ « Province » du formulaire)
 *  - ODK_DATE_MIN   (2026-08-17 — première date de supervision retenue)
 */

import { SUPERVISION_INDICATORS, type SupervisionPayload, type SupervisionRecord } from "./odk-supervision";
import { kvGetJSON, kvSetJSON } from "./kv-store";

const DEFAULTS = {
  baseUrl: "https://api.whonghub.org",
  username: "drcmdd",
  password: "C0ng0@mdd",
  formId: 17559,
  province: "Kasai_Central",
  dateMin: "2026-08-17",
  formTitle: "POLIO-RR SUPERVISION EQUIPES INTEGREES BLOC 3",
};

export function odkConfig() {
  return {
    baseUrl: process.env.ODK_BASE_URL || DEFAULTS.baseUrl,
    username: process.env.ODK_USERNAME || DEFAULTS.username,
    password: process.env.ODK_PASSWORD || DEFAULTS.password,
    formId: Number(process.env.ODK_SUPERVISION_FORM_ID || DEFAULTS.formId),
    province: process.env.ODK_PROVINCE || DEFAULTS.province,
    dateMin: process.env.ODK_DATE_MIN || DEFAULTS.dateMin,
  };
}

const FIELDS = [
  "_id",
  "_geolocation",
  "_submission_time",
  "group_identification/date_supervision",
  "group_identification/Province",
  "group_identification/Antenne",
  "group_identification/ZS",
  "group_identification/aire_sante",
  "group_identification/nom_site",
  "group_identification/nom_superviseur",
  "group_identification/profil_superviseur",
  "group_offre/type_equipe",
  "group_conclusion/actions_correctrices",
  "group_conclusion/recommandations",
  ...SUPERVISION_INDICATORS.map((i) => i.field),
];

type Raw = Record<string, unknown>;

function s(v: unknown): string {
  return v == null ? "" : String(v).trim();
}
function yn(v: unknown): "oui" | "non" | null {
  const t = s(v).toLowerCase();
  if (t === "oui" || t === "yes" || t === "1") return "oui";
  if (t === "non" || t === "no" || t === "0") return "non";
  return null;
}
function geo(r: Raw): { lat: number | null; lon: number | null } {
  const g = r._geolocation;
  if (Array.isArray(g) && g.length >= 2 && typeof g[0] === "number" && typeof g[1] === "number") {
    return { lat: g[0], lon: g[1] };
  }
  const gps = s(r["group_identification/gps"] ?? r["group_conclusion/gps2"]);
  const m = /^(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)/.exec(gps);
  if (m) return { lat: parseFloat(m[1]), lon: parseFloat(m[2]) };
  return { lat: null, lon: null };
}

function toRecord(r: Raw): SupervisionRecord {
  const answers: Record<string, "oui" | "non" | null> = {};
  for (const ind of SUPERVISION_INDICATORS) answers[ind.key] = yn(r[ind.field]);
  const { lat, lon } = geo(r);
  return {
    id: Number(r._id),
    date: s(r["group_identification/date_supervision"]) || s(r._submission_time).slice(0, 10),
    province: s(r["group_identification/Province"]),
    antenne: s(r["group_identification/Antenne"]),
    zs: s(r["group_identification/ZS"]),
    as: s(r["group_identification/aire_sante"]),
    site: s(r["group_identification/nom_site"]),
    superviseur: s(r["group_identification/nom_superviseur"]),
    profil: s(r["group_identification/profil_superviseur"]),
    typeEquipe: s(r["group_offre/type_equipe"]),
    lat, lon,
    answers,
    actions: s(r["group_conclusion/actions_correctrices"]),
    recommandations: s(r["group_conclusion/recommandations"]),
  };
}

let cache: { key: string; at: number; payload: SupervisionPayload } | null = null;
const TTL_MS = 5 * 60 * 1000;

/**
 * Récupère les supervisions de la province depuis la date minimale. Résultat mis
 * en cache 5 minutes côté serveur (les données ODK évoluent en continu pendant la
 * campagne, mais un rafraîchissement par génération de rapport suffit).
 */
export async function fetchSupervision(opts: { dateMin?: string; force?: boolean } = {}): Promise<SupervisionPayload> {
  const cfg = odkConfig();
  const dateMin = /^\d{4}-\d{2}-\d{2}$/.test(opts.dateMin ?? "") ? (opts.dateMin as string) : cfg.dateMin;
  const key = `${cfg.formId}|${cfg.province}|${dateMin}`;
  if (!opts.force && cache && cache.key === key && Date.now() - cache.at < TTL_MS) return cache.payload;

  const query = JSON.stringify({
    "group_identification/Province": cfg.province,
    "group_identification/date_supervision": { $gte: dateMin },
  });
  const url = `${cfg.baseUrl}/api/v1/data/${cfg.formId}.json?query=${encodeURIComponent(query)}&fields=${encodeURIComponent(JSON.stringify(FIELDS))}`;
  const auth = Buffer.from(`${cfg.username}:${cfg.password}`).toString("base64");
  const kvKey = `rrpolio:odk:${key}`;
  let raw: Raw[];
  try {
    const res = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: "application/json" }, cache: "no-store", signal: AbortSignal.timeout(55_000) });
    if (!res.ok) throw new Error(`ODK ${res.status} ${res.statusText}`);
    raw = (await res.json()) as Raw[];
  } catch (e) {
    // Serveur ODK lent / indisponible : on ressert la dernière extraction réussie
    // (mémoire du processus, puis stockage partagé KV s'il est configuré).
    if (cache && cache.key === key) return { ...cache.payload, stale: true };
    const saved = await kvGetJSON<SupervisionPayload>(kvKey);
    if (saved?.ok) return { ...saved, stale: true };
    throw e;
  }
  const records = (Array.isArray(raw) ? raw : [])
    .map(toRecord)
    .filter((r) => r.date >= dateMin)
    .sort((a, b) => a.date.localeCompare(b.date) || a.id - b.id);

  const payload: SupervisionPayload = {
    ok: true,
    fetchedAt: new Date().toISOString(),
    dateMin,
    province: cfg.province,
    formId: cfg.formId,
    formTitle: DEFAULTS.formTitle,
    total: records.length,
    records,
  };
  cache = { key, at: Date.now(), payload };
  void kvSetJSON(kvKey, payload);
  return payload;
}

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
const inflight = new Map<string, Promise<SupervisionPayload>>();
const TTL_MS = 5 * 60 * 1000;
/** Délai maximal accordé au serveur ODK (il peut mettre plusieurs minutes). */
const ODK_TIMEOUT_MS = 280_000;

function cacheKey(cfg: ReturnType<typeof odkConfig>, dateMin: string): string {
  return `${cfg.formId}|${cfg.province}|${dateMin}`;
}
function resolveDateMin(cfg: ReturnType<typeof odkConfig>, dateMin?: string): string {
  return /^\d{4}-\d{2}-\d{2}$/.test(dateMin ?? "") ? (dateMin as string) : cfg.dateMin;
}

/** Interroge réellement le serveur ODK (long) et met à jour les caches (mémoire + KV). */
async function refreshFromOdk(dateMin: string): Promise<SupervisionPayload> {
  const cfg = odkConfig();
  const key = cacheKey(cfg, dateMin);
  const running = inflight.get(key);
  if (running) return running;
  const job = (async () => {
    const query = JSON.stringify({
      "group_identification/Province": cfg.province,
      "group_identification/date_supervision": { $gte: dateMin },
    });
    const url = `${cfg.baseUrl}/api/v1/data/${cfg.formId}.json?query=${encodeURIComponent(query)}&fields=${encodeURIComponent(JSON.stringify(FIELDS))}`;
    const auth = Buffer.from(`${cfg.username}:${cfg.password}`).toString("base64");
    const res = await fetch(url, { headers: { Authorization: `Basic ${auth}`, Accept: "application/json" }, cache: "no-store", signal: AbortSignal.timeout(ODK_TIMEOUT_MS) });
    if (!res.ok) throw new Error(`ODK ${res.status} ${res.statusText}`);
    const raw = (await res.json()) as Raw[];
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
    await kvSetJSON(`rrpolio:odk:${key}`, payload);
    return payload;
  })();
  inflight.set(key, job);
  job.finally(() => inflight.delete(key)).catch(() => undefined);
  return job;
}

export interface SupervisionLookup {
  /** Données disponibles (fraîches ou antérieures), ou null si aucune extraction n'a encore abouti. */
  payload: SupervisionPayload | null;
  /** Vrai si un rafraîchissement doit être lancé en arrière‑plan. */
  needsRefresh: boolean;
  /** Tâche de rafraîchissement à confier à `waitUntil` (ou à attendre). */
  refresh: () => Promise<SupervisionPayload>;
}

/**
 * Lecture rapide (sans appel long) : cache mémoire frais → KV (dernière extraction
 * réussie, marquée `stale`) → rien. Le serveur ODK est très lent (plusieurs minutes
 * pour ce formulaire) : l'appel réseau est fait en arrière‑plan et le client
 * interroge à nouveau la route jusqu'à obtenir les données.
 */
export async function lookupSupervision(opts: { dateMin?: string; force?: boolean } = {}): Promise<SupervisionLookup> {
  const cfg = odkConfig();
  const dateMin = resolveDateMin(cfg, opts.dateMin);
  const key = cacheKey(cfg, dateMin);
  const refresh = () => refreshFromOdk(dateMin);
  if (!opts.force && cache && cache.key === key && Date.now() - cache.at < TTL_MS) {
    return { payload: cache.payload, needsRefresh: false, refresh };
  }
  const saved = cache && cache.key === key ? cache.payload : await kvGetJSON<SupervisionPayload>(`rrpolio:odk:${key}`);
  if (saved?.ok) {
    if (!cache || cache.key !== key) cache = { key, at: Date.parse(saved.fetchedAt) || 0, payload: saved };
    const ageMs = Date.now() - (Date.parse(saved.fetchedAt) || 0);
    return { payload: { ...saved, stale: ageMs > TTL_MS }, needsRefresh: opts.force || ageMs > TTL_MS, refresh };
  }
  return { payload: null, needsRefresh: true, refresh };
}

/** Récupère les supervisions (attend l'appel ODK si nécessaire) — scripts / tests. */
export async function fetchSupervision(opts: { dateMin?: string; force?: boolean } = {}): Promise<SupervisionPayload> {
  const l = await lookupSupervision(opts);
  if (l.payload && !l.needsRefresh) return l.payload;
  try {
    return await l.refresh();
  } catch (e) {
    if (l.payload) return l.payload;
    throw e;
  }
}

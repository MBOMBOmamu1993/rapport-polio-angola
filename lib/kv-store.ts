/**
 * Stockage partagé côté serveur (Vercel KV) pour la compilation provinciale
 * (toutes les antennes / ZS du Kasaï Central).
 *
 * Granularité : une clé par Zone de Santé — `polio:zs:<province>__<antenne>__<zs>`.
 * Réimporter une entité écrase uniquement ses ZS ; les imports simultanés de
 * provinces/antennes/ZS différentes écrivent des clés distinctes (pas de collision).
 */

import { gzipSync, gunzipSync } from "node:zlib";
import { createClient, type VercelKV } from "@vercel/kv";
import { MASQUE_SCHEMA, sanitizeRecords, type ASRecord, type MasqueData } from "./parse-masque";

const INDEX_KEY = "rrpolio:zs:index";

/**
 * Résout les identifiants du store, quel que soit le nommage de l'intégration :
 * - Vercel KV « classique » : KV_REST_API_URL / KV_REST_API_TOKEN
 * - Upstash for Redis (Marketplace) : UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN
 */
function kvCreds(): { url?: string; token?: string } {
  const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
  return { url, token };
}

let _kv: VercelKV | null = null;
function kvClient(): VercelKV {
  if (_kv) return _kv;
  const { url, token } = kvCreds();
  if (!url || !token) throw new Error("Vercel KV non configuré");
  _kv = createClient({ url, token });
  return _kv;
}

export interface ZSBlock {
  province: string;
  antenne: string;
  zs: string;
  periode: string;
  fileName: string;
  importedAt: string;
  /** Version du format des enregistrements (les blocs d'un ancien format sont ignorés). */
  schema?: number;
  dateDebut?: string;
  dateFin?: string;
  jourLabels?: string[];
  records: ASRecord[];
}

/**
 * Lecture / écriture génériques (cache de secours des données ODK, etc.).
 *
 * Les valeurs sont compressées (gzip → base64) et découpées en morceaux :
 * le plan Upstash Free refuse toute requête de plus de 10 Mo, et le paquet
 * ODK complet a fini par dépasser ce plafond (échec d'écriture silencieux,
 * cache figé). La clé principale ne porte qu'un manifeste `{ __gz, n }` ;
 * les morceaux vivent dans `<clé>:gz:<i>`. Une valeur non compressée écrite
 * par une ancienne version reste lisible telle quelle.
 */
const GZ_PREFIX = "gz64:";
/** ~3 Mo par morceau : marge large sous la limite de 10 Mo/requête. */
const GZ_CHUNK_CHARS = 3_000_000;
/** Nombre de morceaux résiduels balayés après une écriture plus courte. */
const GZ_SWEEP_EXTRA = 20;

interface GzManifest { __gz: 1; n: number }
function isGzManifest(v: unknown): v is GzManifest {
  return typeof v === "object" && v !== null && (v as GzManifest).__gz === 1 && typeof (v as GzManifest).n === "number";
}
function gzChunkKey(key: string, i: number): string {
  return `${key}:gz:${i}`;
}

export async function kvGetJSON<T>(key: string): Promise<T | null> {
  if (!kvAvailable()) return null;
  try {
    const kv = kvClient();
    const head = await kv.get<unknown>(key);
    if (head == null) return null;
    if (!isGzManifest(head)) return head as T; // ancienne valeur non compressée
    if (head.n < 1) return null;
    const parts = await kv.mget<(string | null)[]>(...Array.from({ length: head.n }, (_, i) => gzChunkKey(key, i)));
    const b64 = parts.map((p) => (typeof p === "string" && p.startsWith(GZ_PREFIX) ? p.slice(GZ_PREFIX.length) : null));
    if (b64.some((p) => p == null)) return null; // morceau manquant → cache invalide
    return JSON.parse(gunzipSync(Buffer.from(b64.join(""), "base64")).toString("utf8")) as T;
  } catch (err) {
    console.error(`KV : échec de lecture de « ${key} »`, err);
    return null;
  }
}

export async function kvSetJSON(key: string, value: unknown): Promise<void> {
  if (!kvAvailable()) return;
  try {
    const kv = kvClient();
    const b64 = gzipSync(Buffer.from(JSON.stringify(value), "utf8")).toString("base64");
    const chunks: string[] = [];
    for (let i = 0; i < b64.length; i += GZ_CHUNK_CHARS) chunks.push(b64.slice(i, i + GZ_CHUNK_CHARS));
    // Morceaux d'abord, manifeste ensuite : un lecteur concurrent ne voit
    // jamais un manifeste pointant vers des morceaux absents. Écriture
    // séquentielle : l'auto-pipelining du client regrouperait des SET
    // simultanés en une seule requête HTTP, qui redépasserait les 10 Mo.
    for (let i = 0; i < chunks.length; i++) {
      await kv.set(gzChunkKey(key, i), GZ_PREFIX + chunks[i]);
    }
    await kv.set(key, { __gz: 1, n: chunks.length } satisfies GzManifest);
    const stale = Array.from({ length: GZ_SWEEP_EXTRA }, (_, i) => gzChunkKey(key, chunks.length + i));
    await kv.del(...(stale as [string, ...string[]]));
  } catch (err) {
    console.error(`KV : échec d'écriture de « ${key} »`, err);
  }
}

export function kvAvailable(): boolean {
  const { url, token } = kvCreds();
  return Boolean(url && token);
}

/**
 * Verrou best-effort inter-instances (SET NX + expiration) : évite que
 * plusieurs lambdas lancent en parallèle la même extraction lourde (ODK).
 * Sans KV, le verrou est réputé acquis (une seule instance en dev local).
 */
export async function kvTryLock(name: string, ttlSeconds: number): Promise<boolean> {
  if (!kvAvailable()) return true;
  try {
    const res = await kvClient().set(`lock:${name}`, "1", { nx: true, ex: ttlSeconds });
    return res === "OK";
  } catch {
    return true; // KV en panne → ne pas bloquer l'extraction
  }
}

export async function kvUnlock(name: string): Promise<void> {
  if (!kvAvailable()) return;
  try {
    await kvClient().del(`lock:${name}`);
  } catch {
    /* le TTL fera le ménage */
  }
}

function slug(s: string): string {
  return (s || "—").trim().toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_|_$/g, "") || "NA";
}

function keyFor(province: string, antenne: string, zs: string): string {
  return `rrpolio:zs:${slug(province)}__${slug(antenne)}__${slug(zs)}`;
}

/** Enregistre / met à jour les ZS contenues dans un import (remplacement par ZS). */
export async function upsertImport(data: MasqueData): Promise<{ updatedZones: string[] }> {
  const groups = new Map<string, ZSBlock>();
  for (const r of data.records) {
    const k = keyFor(r.province, r.antenne, r.zs);
    let g = groups.get(k);
    if (!g) {
      g = {
        province: r.province,
        antenne: r.antenne,
        zs: r.zs,
        periode: data.meta.periode,
        fileName: data.meta.fileName,
        importedAt: data.meta.importedAt,
        schema: MASQUE_SCHEMA,
        dateDebut: data.meta.dateDebut,
        dateFin: data.meta.dateFin,
        jourLabels: data.meta.jourLabels,
        records: [],
      };
      groups.set(k, g);
    }
    g.records.push(r);
  }

  const kv = kvClient();
  const keys = Array.from(groups.keys());
  await Promise.all([
    ...Array.from(groups.entries()).map(([k, block]) => kv.set(k, block)),
    keys.length ? kv.sadd(INDEX_KEY, ...(keys as [string, ...string[]])) : Promise.resolve(),
  ]);

  return { updatedZones: Array.from(groups.values()).map((g) => `${g.province} · ${g.antenne} · ${g.zs}`) };
}

/** Lit tous les blocs ZS du pays. */
export async function readNationalBlocks(): Promise<ZSBlock[]> {
  const kv = kvClient();
  const keys = await kv.smembers(INDEX_KEY);
  if (!keys || keys.length === 0) return [];
  const values = await kv.mget<ZSBlock[]>(...keys);
  return values.filter((v): v is ZSBlock => Boolean(v && v.records && v.schema === MASQUE_SCHEMA));
}

/**
 * Réinitialise toute la compilation nationale : supprime chaque bloc ZS et
 * l'index. Action d'administration destructive (tous les imports de toutes les
 * provinces sont effacés et chacun devra réimporter à zéro).
 */
export async function resetNational(): Promise<{ deleted: number }> {
  const kv = kvClient();
  const keys = await kv.smembers(INDEX_KEY);
  if (keys && keys.length > 0) {
    await kv.del(...(keys as [string, ...string[]]));
  }
  await kv.del(INDEX_KEY);
  return { deleted: keys?.length ?? 0 };
}

/** Construit un MasqueData consolidé (toutes provinces) + la liste des entités importées. */
export async function readNational(): Promise<{
  data: MasqueData;
  entities: { province: string; antenne: string; zs: string; importedAt: string; nbAires: number; periode: string }[];
}> {
  const blocks = await readNationalBlocks();
  // Assainit la compilation consolidée : écarte les sous-totaux/titres résiduels
  // (sinon le parent est recompté au niveau province).
  const records: ASRecord[] = sanitizeRecords(blocks.flatMap((b) => b.records));

  const provinces = uniq(records.map((r) => r.province));
  const antennes = uniq(records.map((r) => r.antenne));
  const zones = uniq(records.map((r) => r.zs));
  const latest = blocks.reduce((m, b) => (b.importedAt > m ? b.importedAt : m), "");
  const periode = blocks.length ? mostCommon(blocks.map((b) => b.periode).filter(Boolean)) : "";

  const data: MasqueData = {
    meta: {
      pays: "RD CONGO",
      periode,
      dateDebut: blocks.length ? mostCommon(blocks.map((b) => b.dateDebut ?? "").filter(Boolean)) : "",
      dateFin: blocks.length ? mostCommon(blocks.map((b) => b.dateFin ?? "").filter(Boolean)) : "",
      province: provinces.length === 1 ? provinces[0] : "Niveau national",
      antennes,
      zones,
      importedAt: latest || new Date().toISOString(),
      fileName: "Compilation nationale",
      nbAires: records.length,
      nbJours: Math.max(0, ...records.map((r) => r.dailyReports?.length ?? 0)),
      jourLabels: maxJourLabels(blocks),
      schema: MASQUE_SCHEMA,
    },
    records,
  };

  const entities = blocks
    .map((b) => ({
      province: b.province,
      antenne: b.antenne,
      zs: b.zs,
      importedAt: b.importedAt,
      nbAires: b.records.length,
      periode: b.periode,
    }))
    .sort((a, b) => b.importedAt.localeCompare(a.importedAt));

  return { data, entities };
}

function uniq(arr: string[]): string[] {
  return Array.from(new Set(arr.filter(Boolean))).sort((a, b) => a.localeCompare(b, "fr"));
}
function mostCommon(arr: string[]): string {
  const c = new Map<string, number>();
  for (const v of arr) c.set(v, (c.get(v) ?? 0) + 1);
  let best = "";
  let max = -1;
  for (const [k, n] of c) if (n > max) { max = n; best = k; }
  return best;
}

/** Récupère les étiquettes de jour les plus longues parmi les blocs (pour la compilation nationale). */
function maxJourLabels(blocks: ZSBlock[]): string[] {
  let best: string[] = [];
  for (const b of blocks) {
    const first = b.records[0];
    const labels = b.jourLabels ?? first?.dailyReports?.map((d) => d.label) ?? [];
    if (labels.length > best.length) best = labels;
  }
  return best;
}

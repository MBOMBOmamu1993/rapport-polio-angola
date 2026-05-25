/**
 * Stockage partagé côté serveur (Vercel KV) pour la compilation nationale.
 *
 * Granularité : une clé par Zone de Santé — `polio:zs:<province>__<antenne>__<zs>`.
 * Réimporter une entité écrase uniquement ses ZS ; les imports simultanés de
 * provinces/antennes/ZS différentes écrivent des clés distinctes (pas de collision).
 */

import { kv } from "@vercel/kv";
import type { ASRecord, MasqueData } from "./parse-masque";

const INDEX_KEY = "polio:zs:index";

export interface ZSBlock {
  province: string;
  antenne: string;
  zs: string;
  periode: string;
  fileName: string;
  importedAt: string;
  records: ASRecord[];
}

export function kvAvailable(): boolean {
  return Boolean(process.env.KV_REST_API_URL && process.env.KV_REST_API_TOKEN);
}

function slug(s: string): string {
  return (s || "—").trim().toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_|_$/g, "") || "NA";
}

function keyFor(province: string, antenne: string, zs: string): string {
  return `polio:zs:${slug(province)}__${slug(antenne)}__${slug(zs)}`;
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
        records: [],
      };
      groups.set(k, g);
    }
    g.records.push(r);
  }

  const keys = Array.from(groups.keys());
  await Promise.all([
    ...Array.from(groups.entries()).map(([k, block]) => kv.set(k, block)),
    keys.length ? kv.sadd(INDEX_KEY, ...(keys as [string, ...string[]])) : Promise.resolve(),
  ]);

  return { updatedZones: Array.from(groups.values()).map((g) => `${g.province} · ${g.antenne} · ${g.zs}`) };
}

/** Lit tous les blocs ZS du pays. */
export async function readNationalBlocks(): Promise<ZSBlock[]> {
  const keys = await kv.smembers(INDEX_KEY);
  if (!keys || keys.length === 0) return [];
  const values = await kv.mget<ZSBlock[]>(...keys);
  return values.filter((v): v is ZSBlock => Boolean(v && v.records));
}

/** Construit un MasqueData consolidé (toutes provinces) + la liste des entités importées. */
export async function readNational(): Promise<{
  data: MasqueData;
  entities: { province: string; antenne: string; zs: string; importedAt: string; nbAires: number; periode: string }[];
}> {
  const blocks = await readNationalBlocks();
  const records: ASRecord[] = blocks.flatMap((b) => b.records);

  const provinces = uniq(records.map((r) => r.province));
  const antennes = uniq(records.map((r) => r.antenne));
  const zones = uniq(records.map((r) => r.zs));
  const latest = blocks.reduce((m, b) => (b.importedAt > m ? b.importedAt : m), "");
  const periode = blocks.length ? mostCommon(blocks.map((b) => b.periode).filter(Boolean)) : "";

  const data: MasqueData = {
    meta: {
      pays: "RD CONGO",
      periode,
      province: provinces.length === 1 ? provinces[0] : "Niveau national",
      antennes,
      zones,
      importedAt: latest || new Date().toISOString(),
      fileName: "Compilation nationale",
      nbAires: records.length,
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

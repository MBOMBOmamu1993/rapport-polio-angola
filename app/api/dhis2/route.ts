import { NextRequest, NextResponse } from "next/server";
import { waitUntil } from "@vercel/functions";
import { buildProvinceBlock, listCampaignProvinces } from "@/lib/dhis2-campagne";
import { DHIS2_BLOCK_SCHEMA, type ProvinceBlock, type ProvinceListPayload } from "@/lib/dhis2-shared";
import { kvGetJSON, kvSetJSON } from "@/lib/kv-store";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

/**
 * Données de la campagne intégrée RR-POLIO depuis DHIS2 (rdccampagne).
 *  - `?list=1` : provinces ayant des données pour la campagne (cibles / vaccinés) ;
 *  - `?province=<uid>` : extraction complète de la province (toutes ses AS) ;
 *  - `?force=1` : forcer le rafraîchissement.
 *
 * Réponses : 200 + payload (`stale: true` si extraction antérieure, un
 * rafraîchissement tourne alors en arrière-plan via waitUntil) ; 202 +
 * `{ pending: true }` pendant la toute première extraction.
 */

const TTL_MS = 10 * 60 * 1000;

interface CacheEntry<T> { at: number; value: T }
const mem = new Map<string, CacheEntry<unknown>>();
const inflight = new Map<string, Promise<unknown>>();

function refresh<T>(key: string, job: () => Promise<T>): Promise<T> {
  const running = inflight.get(key);
  if (running) return running as Promise<T>;
  const p = (async () => {
    const value = await job();
    mem.set(key, { at: Date.now(), value });
    await kvSetJSON(key, value);
    return value;
  })();
  inflight.set(key, p);
  p.finally(() => inflight.delete(key)).catch(() => undefined);
  return p;
}

async function lookup<T extends { fetchedAt: string; schema: number }>(
  key: string,
  force: boolean,
  job: () => Promise<T>
): Promise<{ value: (T & { stale?: boolean }) | null; needsRefresh: boolean; refresh: () => Promise<T> }> {
  const run = () => refresh(key, job);
  const cached = mem.get(key) as CacheEntry<T> | undefined;
  if (!force && cached && Date.now() - cached.at < TTL_MS && cached.value.schema === DHIS2_BLOCK_SCHEMA) {
    return { value: cached.value, needsRefresh: false, refresh: run };
  }
  const saved = cached?.value ?? (await kvGetJSON<T>(key));
  if (saved && saved.schema === DHIS2_BLOCK_SCHEMA) {
    const age = Date.now() - (Date.parse(saved.fetchedAt) || 0);
    if (!mem.has(key)) mem.set(key, { at: Date.parse(saved.fetchedAt) || 0, value: saved });
    return { value: { ...saved, stale: force || age > TTL_MS }, needsRefresh: force || age > TTL_MS, refresh: run };
  }
  return { value: null, needsRefresh: true, refresh: run };
}

export async function GET(req: NextRequest) {
  const provinceId = req.nextUrl.searchParams.get("province");
  const force = req.nextUrl.searchParams.get("force") === "1";
  const noStore = { "cache-control": "no-store" };

  try {
    if (provinceId) {
      if (!/^[A-Za-z][A-Za-z0-9]{10}$/.test(provinceId)) {
        return NextResponse.json({ ok: false, reason: "identifiant de province invalide" }, { status: 400 });
      }
      const key = `rrpolio:dhis2:v${DHIS2_BLOCK_SCHEMA}:prov:${provinceId}`;
      const l = await lookup<ProvinceBlock>(key, force, () => buildProvinceBlock(provinceId));
      if (l.needsRefresh) waitUntil(l.refresh().catch(() => undefined));
      if (l.value) return NextResponse.json(l.value, { headers: noStore });
      return NextResponse.json({ ok: false, pending: true }, { status: 202, headers: noStore });
    }

    const key = `rrpolio:dhis2:v${DHIS2_BLOCK_SCHEMA}:list`;
    const l = await lookup<ProvinceListPayload>(key, force, async () => ({
      ok: true,
      fetchedAt: new Date().toISOString(),
      schema: DHIS2_BLOCK_SCHEMA,
      provinces: await listCampaignProvinces(),
    }));
    if (l.needsRefresh) waitUntil(l.refresh().catch(() => undefined));
    if (l.value) return NextResponse.json(l.value, { headers: noStore });
    return NextResponse.json({ ok: false, pending: true }, { status: 202, headers: noStore });
  } catch (e) {
    return NextResponse.json(
      { ok: false, reason: e instanceof Error ? e.message : "erreur DHIS2" },
      { status: 502, headers: noStore }
    );
  }
}

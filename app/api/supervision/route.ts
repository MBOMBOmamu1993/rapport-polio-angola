import { gunzipSync } from "node:zlib";
import { NextRequest, NextResponse } from "next/server";
import { waitUntil } from "@vercel/functions";
import { lookupSupervision, seedSupervision } from "@/lib/odk-server";
import {
  aggregateSupervisionByZS,
  indicatorConformity,
  type SupervisionCompact,
  type SupervisionPayload,
  type ZSRef,
} from "@/lib/odk-supervision";
import { getProvinceBlockCached } from "@/lib/dhis2-campagne";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
// Le serveur ODK (whonghub) est très lent sur ce formulaire (plusieurs minutes) :
// l'appel réseau se fait en arrière‑plan (waitUntil) et la réponse est immédiate.
export const maxDuration = 300;

/**
 * Supervisions ODK de la province (Kasaï Central) depuis la date minimale.
 *  - `?dateMin=YYYY-MM-DD` : première date retenue ; `?force=1` : forcer le rafraîchissement.
 *  - 200 + payload : données disponibles (`stale: true` si extraction antérieure, un
 *    rafraîchissement tourne alors en arrière‑plan) ;
 *  - 202 + `{ pending: true }` : première extraction en cours — réinterroger dans
 *    quelques secondes.
 */
export async function GET(req: NextRequest) {
  const dateMin = req.nextUrl.searchParams.get("dateMin") ?? undefined;
  const force = req.nextUrl.searchParams.get("force") === "1";
  // Province du formulaire ODK (onglet DHIS2 : « Maniema », « Sud Kivu »…) ;
  // sans paramètre, la province par défaut (Kasai_Central) est conservée.
  const province = req.nextUrl.searchParams.get("province")?.slice(0, 40) ?? undefined;
  // Mode compact (rapport national) : agrégation par ZS côté serveur, jointe aux
  // Aires de Santé du bloc provincial (masque ou DHIS2) — quelques centaines de Ko
  // au lieu des dizaines de Mo d'enregistrements bruts.
  const compact = req.nextUrl.searchParams.get("compact") === "1";
  const provinceId = req.nextUrl.searchParams.get("provinceId") ?? "";
  try {
    const l = await lookupSupervision({ dateMin, force, province });
    if (l.needsRefresh) {
      const job = l.refresh().catch(() => undefined);
      waitUntil(job);
    }
    if (l.payload && compact) {
      if (!/^[A-Za-z][A-Za-z0-9]{10}$/.test(provinceId)) {
        return NextResponse.json({ ok: false, reason: "provinceId requis en mode compact" }, { status: 400 });
      }
      const block = await getProvinceBlockCached(provinceId);
      if (!block) {
        return NextResponse.json({ ok: false, reason: "bloc provincial indisponible" }, { status: 502 });
      }
      const zsMap = new Map<string, ZSRef>();
      for (const r of block.records) {
        let z = zsMap.get(r.zs);
        if (!z) { z = { zs: r.zs, antenne: r.antenne, aires: [] }; zsMap.set(r.zs, z); }
        z.aires.push(r.as);
      }
      const recs = l.payload.records;
      const out: SupervisionCompact = {
        ok: true,
        compact: true,
        provinceId,
        province: block.province,
        fetchedAt: l.payload.fetchedAt,
        dateMin: l.payload.dateMin,
        formTitle: l.payload.formTitle,
        total: recs.length,
        byZS: aggregateSupervisionByZS(recs, Array.from(zsMap.values())),
        conformity: indicatorConformity(recs),
        points: recs.filter((r) => r.lat != null && r.lon != null).map((r) => ({ lat: r.lat as number, lon: r.lon as number })),
      };
      return NextResponse.json(out, { headers: { "cache-control": "no-store" } });
    }
    if (l.payload) {
      return NextResponse.json(l.payload, { headers: { "cache-control": "no-store" } });
    }
    return NextResponse.json({ ok: false, pending: true, reason: "pending", records: [], total: 0 }, { status: 202, headers: { "cache-control": "no-store" } });
  } catch (e) {
    return NextResponse.json(
      { ok: false, reason: e instanceof Error ? e.message : "error", records: [], total: 0 },
      { status: 502 }
    );
  }
}

/**
 * Injection d'une extraction ODK figée (campagne terminée) — administration.
 * POST /api/supervision avec l'en-tête `x-seed-code` (ODK_SEED_CODE) et pour
 * corps le JSON SupervisionPayload, éventuellement gzippé.
 */
export async function POST(req: NextRequest) {
  const code = req.headers.get("x-seed-code") ?? "";
  const expected = process.env.ODK_SEED_CODE || "";
  if (!expected || code !== expected) {
    return NextResponse.json({ ok: false, reason: "non autorisé" }, { status: 401 });
  }
  try {
    const buf = Buffer.from(await req.arrayBuffer());
    const text = buf.length > 2 && buf[0] === 0x1f && buf[1] === 0x8b ? gunzipSync(buf).toString("utf8") : buf.toString("utf8");
    const payload = JSON.parse(text) as SupervisionPayload;
    const res = await seedSupervision(payload);
    return NextResponse.json({ ok: true, ...res });
  } catch (e) {
    return NextResponse.json({ ok: false, reason: e instanceof Error ? e.message : "erreur" }, { status: 400 });
  }
}

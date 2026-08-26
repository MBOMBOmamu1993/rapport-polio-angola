import { gunzipSync } from "node:zlib";
import { NextRequest, NextResponse } from "next/server";
import { waitUntil } from "@vercel/functions";
import { lookupSupervision, seedSupervision } from "@/lib/odk-server";
import type { SupervisionPayload } from "@/lib/odk-supervision";

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
  try {
    const l = await lookupSupervision({ dateMin, force, province });
    if (l.needsRefresh) {
      const job = l.refresh().catch(() => undefined);
      waitUntil(job);
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

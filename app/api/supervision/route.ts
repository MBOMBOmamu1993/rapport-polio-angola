import { NextRequest, NextResponse } from "next/server";
import { waitUntil } from "@vercel/functions";
import { lookupSupervision } from "@/lib/odk-server";

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
  try {
    const l = await lookupSupervision({ dateMin, force });
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

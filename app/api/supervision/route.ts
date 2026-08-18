import { NextRequest, NextResponse } from "next/server";
import { fetchSupervision } from "@/lib/odk-server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
// L'API ODK (whonghub) peut être lente : on laisse jusqu'à 60 s à la fonction.
export const maxDuration = 60;

/**
 * Supervisions ODK de la province (Kasaï Central) depuis la date minimale.
 * `?dateMin=YYYY-MM-DD` pour changer la date, `?force=1` pour ignorer le cache.
 */
export async function GET(req: NextRequest) {
  const dateMin = req.nextUrl.searchParams.get("dateMin") ?? undefined;
  const force = req.nextUrl.searchParams.get("force") === "1";
  try {
    const payload = await fetchSupervision({ dateMin, force });
    return NextResponse.json(payload, { headers: { "cache-control": "no-store" } });
  } catch (e) {
    return NextResponse.json(
      { ok: false, reason: e instanceof Error ? e.message : "error", records: [], total: 0 },
      { status: 502 }
    );
  }
}

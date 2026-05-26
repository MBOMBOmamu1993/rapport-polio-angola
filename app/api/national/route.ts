import { NextRequest, NextResponse } from "next/server";
import { kvAvailable, readNational, resetNational } from "@/lib/kv-store";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET() {
  if (!kvAvailable()) {
    return NextResponse.json({ ok: false, reason: "kv_unavailable" }, { status: 503 });
  }
  try {
    const { data, entities } = await readNational();
    return NextResponse.json({ ok: true, data, entities });
  } catch (e) {
    return NextResponse.json({ ok: false, reason: e instanceof Error ? e.message : "error" }, { status: 500 });
  }
}

/**
 * Réinitialisation administrateur de la compilation nationale. Protégée par un
 * code secret défini dans la variable d'environnement `ADMIN_RESET_CODE` (à
 * configurer côté hébergeur). Sans cette variable, l'endpoint refuse l'action :
 * la réinitialisation n'est jamais ouverte par défaut.
 */
export async function DELETE(req: NextRequest) {
  if (!kvAvailable()) {
    return NextResponse.json({ ok: false, reason: "kv_unavailable" }, { status: 503 });
  }
  const expected = process.env.ADMIN_RESET_CODE;
  if (!expected) {
    return NextResponse.json({ ok: false, reason: "admin_not_configured" }, { status: 403 });
  }
  const provided =
    req.headers.get("x-admin-code") ??
    ((await req.json().catch(() => ({}))) as { code?: string }).code ??
    "";
  if (provided !== expected) {
    return NextResponse.json({ ok: false, reason: "bad_code" }, { status: 401 });
  }
  try {
    const { deleted } = await resetNational();
    return NextResponse.json({ ok: true, deleted });
  } catch (e) {
    return NextResponse.json({ ok: false, reason: e instanceof Error ? e.message : "error" }, { status: 500 });
  }
}

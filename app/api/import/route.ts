import { NextRequest, NextResponse } from "next/server";
import { kvAvailable, upsertImport } from "@/lib/kv-store";
import type { MasqueData } from "@/lib/parse-masque";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(req: NextRequest) {
  if (!kvAvailable()) {
    return NextResponse.json({ ok: false, reason: "kv_unavailable" }, { status: 503 });
  }
  try {
    const body = (await req.json()) as MasqueData;
    if (!body?.records?.length) {
      return NextResponse.json({ ok: false, reason: "empty" }, { status: 400 });
    }
    const { updatedZones } = await upsertImport(body);
    return NextResponse.json({ ok: true, updatedZones });
  } catch (e) {
    return NextResponse.json({ ok: false, reason: e instanceof Error ? e.message : "error" }, { status: 500 });
  }
}

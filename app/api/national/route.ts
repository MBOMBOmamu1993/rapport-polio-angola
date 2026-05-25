import { NextResponse } from "next/server";
import { kvAvailable, readNational } from "@/lib/kv-store";

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

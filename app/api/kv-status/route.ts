import { NextResponse } from "next/server";
import { kvAvailable } from "@/lib/kv-store";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

/**
 * Diagnostic non sensible de la configuration KV : indique si le stockage est
 * détecté et quelles variables d'environnement (NOMS uniquement, jamais les
 * valeurs) ressemblent à une intégration KV/Upstash/Redis. Permet d'identifier
 * un problème de nommage, de portée (Production) ou de projet Vercel.
 */
export async function GET() {
  const detectedVarNames = Object.keys(process.env)
    .filter((k) => /KV_|UPSTASH|REDIS/i.test(k))
    .sort();

  return NextResponse.json({
    available: kvAvailable(),
    expected: {
      KV_REST_API_URL: Boolean(process.env.KV_REST_API_URL),
      KV_REST_API_TOKEN: Boolean(process.env.KV_REST_API_TOKEN),
      UPSTASH_REDIS_REST_URL: Boolean(process.env.UPSTASH_REDIS_REST_URL),
      UPSTASH_REDIS_REST_TOKEN: Boolean(process.env.UPSTASH_REDIS_REST_TOKEN),
    },
    detectedVarNames,
    vercelEnv: process.env.VERCEL_ENV ?? null,
  });
}

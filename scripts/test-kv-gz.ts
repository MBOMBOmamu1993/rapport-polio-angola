/**
 * Test du cache KV compressé (kvSetJSON / kvGetJSON) contre un faux serveur
 * Upstash (interception de fetch) qui rejoue le protocole REST et REFUSE
 * toute requête de plus de 10 Mo — exactement la limite du plan Free qui a
 * déclenché l'alerte. Vérifie : aller-retour > 10 Mo, découpage en morceaux,
 * lecture d'une ancienne valeur non compressée, balayage des morceaux
 * résiduels après une réécriture plus petite.
 */

const MAX_REQUEST = 10_000_000;
const store = new Map<string, string>();
let maxSeen = 0;
let requests = 0;

function runCommand(cmd: unknown[]): unknown {
  const op = String(cmd[0]).toUpperCase();
  if (op === "SET") { store.set(String(cmd[1]), String(cmd[2])); return "OK"; }
  if (op === "GET") { return store.get(String(cmd[1])) ?? null; }
  if (op === "MGET") { return cmd.slice(1).map((k) => store.get(String(k)) ?? null); }
  if (op === "DEL") { let n = 0; for (const k of cmd.slice(1)) if (store.delete(String(k))) n++; return n; }
  if (op === "SADD") { return 0; }
  throw new Error(`commande non gérée : ${op}`);
}

const realFetch = globalThis.fetch;
globalThis.fetch = (async (input: any, init?: any) => {
  const url = String(typeof input === "string" ? input : input?.url ?? input);
  if (!url.startsWith("https://fake-upstash.test")) return realFetch(input, init);
  requests++;
  const body = String(init?.body ?? "");
  maxSeen = Math.max(maxSeen, Buffer.byteLength(body));
  if (Buffer.byteLength(body) > MAX_REQUEST) {
    return new Response(JSON.stringify({ error: "ERR max request size exceeded" }), { status: 400 });
  }
  const headers = new Headers(init?.headers ?? {});
  const b64 = (headers.get("upstash-encoding") ?? "").toLowerCase() === "base64";
  const enc = (r: unknown): unknown => {
    if (r === null || typeof r === "number") return r;
    if (Array.isArray(r)) return r.map(enc);
    return b64 ? Buffer.from(String(r), "utf8").toString("base64") : String(r);
  };
  const parsed = JSON.parse(body);
  const payload = url.includes("/pipeline") || url.includes("/multi-exec")
    ? (parsed as unknown[][]).map((c) => ({ result: enc(runCommand(c)) }))
    : { result: enc(runCommand(parsed as unknown[])) };
  return new Response(JSON.stringify(payload), { status: 200, headers: { "content-type": "application/json" } });
}) as typeof fetch;

process.env.KV_REST_API_URL = "https://fake-upstash.test";
process.env.KV_REST_API_TOKEN = "test-token";

const KEY = "rrpolio:test:gz-roundtrip";

function randomHex(n: number): string {
  let s = "";
  while (s.length < n) s += Math.random().toString(16).slice(2);
  return s.slice(0, n);
}

async function main() {
  const { kvGetJSON, kvSetJSON, kvAvailable } = await import("../lib/kv-store");
  if (!kvAvailable()) throw new Error("KV non configuré (env manquant)");

  // 1) Charge > 10 Mo de JSON peu compressible : sans compression/découpage,
  //    le faux serveur renverrait 400 comme le vrai plan Free.
  const records = Array.from({ length: 3000 }, (_, i) => ({ id: i, zs: `ZS_${i % 26}`, blob: randomHex(4000) }));
  const payload = { ok: true, fetchedAt: new Date().toISOString(), total: records.length, records };
  const jsonSize = Buffer.byteLength(JSON.stringify(payload));
  console.log(`Taille JSON brut : ${(jsonSize / 1e6).toFixed(1)} Mo`);
  if (jsonSize < 10_000_000) throw new Error("charge de test trop petite");

  await kvSetJSON(KEY, payload);
  const chunkKeys = Array.from(store.keys()).filter((k) => k.startsWith(`${KEY}:gz:`));
  if (!store.has(KEY)) throw new Error("ÉCHEC : manifeste absent (écriture refusée ?)");
  console.log(`OK : écrit en ${chunkKeys.length} morceau(x) ; plus grosse requête vue : ${(maxSeen / 1e6).toFixed(1)} Mo (< 10 Mo).`);
  if (chunkKeys.length < 2) throw new Error("ÉCHEC : le découpage en morceaux n'a pas été exercé");

  const back = await kvGetJSON<typeof payload>(KEY);
  if (!back) throw new Error("ÉCHEC : relecture nulle");
  if (JSON.stringify(back) !== JSON.stringify(payload)) throw new Error("ÉCHEC : contenu différent après aller-retour");
  console.log(`OK : aller-retour ${back.total} enregistrements, ${(jsonSize / 1e6).toFixed(1)} Mo, intact.`);

  // 2) Compatibilité : une ancienne valeur non compressée doit rester lisible
  store.set(`${KEY}:legacy`, JSON.stringify({ ok: true, total: 2 }));
  const legacy = await kvGetJSON<{ total: number }>(`${KEY}:legacy`);
  if (legacy?.total !== 2) throw new Error("ÉCHEC : valeur héritée illisible");
  console.log("OK : ancienne valeur non compressée lue telle quelle.");

  // 3) Réécriture plus petite → l'ancien surplus de morceaux doit disparaître
  await kvSetJSON(KEY, { ok: true, total: 1 });
  const small = await kvGetJSON<{ total: number }>(KEY);
  if (small?.total !== 1) throw new Error("ÉCHEC : réécriture plus petite illisible");
  const leftovers = Array.from(store.keys()).filter((k) => k.startsWith(`${KEY}:gz:`) && k !== `${KEY}:gz:0`);
  if (leftovers.length > 0) throw new Error(`ÉCHEC : morceaux résiduels non balayés : ${leftovers.join(", ")}`);
  console.log(`OK : réécriture plus petite, morceaux résiduels balayés. (${requests} requêtes au total)`);
  console.log("TOUT EST OK.");
}

main().catch((e) => { console.error(e); process.exit(1); });

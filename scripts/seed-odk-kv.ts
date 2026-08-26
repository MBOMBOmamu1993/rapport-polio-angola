/**
 * Injecte l'extraction ODK Maniema du 25/08 (cache disque du même formulaire
 * 17559) dans le KV de production, sous la clé canonique de l'onglet DHIS2.
 * La reprise incrémentale (curseur lastSubmissionTime) complétera le delta.
 *   npx tsx --env-file=<env> scripts/seed-odk-kv.ts
 */
import { readFileSync } from "node:fs";
import { kvSetJSON, kvGetJSON, kvAvailable } from "../lib/kv-store";
import type { SupervisionPayload } from "../lib/odk-supervision";

async function main() {
  if (!kvAvailable()) throw new Error("KV non configuré (variables d'environnement absentes)");
  const sup = JSON.parse(readFileSync("scripts/.sup-maniema-cache.json", "utf8")) as SupervisionPayload;
  if (!sup.ok || sup.formId !== 17559 || sup.province !== "Maniema") throw new Error("cache inattendu");
  const key = "rrpolio:odk:17559|Maniema|2026-08-11";
  const existing = await kvGetJSON<SupervisionPayload>(key);
  if (existing?.ok && existing.total >= sup.total) {
    console.log(`KV déjà à jour (${existing.total} ≥ ${sup.total}) — rien à faire.`);
    return;
  }
  await kvSetJSON(key, sup);
  const check = await kvGetJSON<SupervisionPayload>(key);
  console.log(`Semé ${key}: ${check?.total} supervisions (extraction du ${sup.fetchedAt}, curseur ${sup.lastSubmissionTime ?? "—"}).`);
}
main().catch((e) => { console.error(e); process.exit(1); });

"use client";

import type { MasqueData } from "./parse-masque";

export interface EntityInfo {
  province: string;
  antenne: string;
  zs: string;
  importedAt: string;
  nbAires: number;
  periode: string;
}

export interface NationalResult {
  data: MasqueData;
  entities: EntityInfo[];
}

/** Envoie l'import au stockage national. Retourne true si synchronisé. */
export async function pushImport(data: MasqueData): Promise<{ ok: boolean; reason?: string }> {
  try {
    const res = await fetch("/api/import", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(data),
    });
    const json = await res.json().catch(() => ({}));
    return { ok: res.ok && json.ok === true, reason: json.reason };
  } catch (e) {
    return { ok: false, reason: e instanceof Error ? e.message : "network" };
  }
}

/**
 * Réinitialise la compilation nationale (action admin). Le code secret est
 * vérifié côté serveur contre la variable d'environnement `ADMIN_RESET_CODE`.
 */
export async function resetNational(
  code: string
): Promise<{ ok: boolean; deleted?: number; reason?: string }> {
  try {
    const res = await fetch("/api/national", {
      method: "DELETE",
      headers: { "content-type": "application/json", "x-admin-code": code },
    });
    const json = await res.json().catch(() => ({}));
    return { ok: res.ok && json.ok === true, deleted: json.deleted, reason: json.reason };
  } catch (e) {
    return { ok: false, reason: e instanceof Error ? e.message : "network" };
  }
}

/** Récupère la compilation nationale. Retourne null si le stockage n'est pas configuré. */
export async function fetchNational(): Promise<NationalResult | null> {
  try {
    const res = await fetch("/api/national", { cache: "no-store" });
    if (!res.ok) return null;
    const json = await res.json();
    if (!json.ok) return null;
    return { data: json.data as MasqueData, entities: json.entities as EntityInfo[] };
  } catch {
    return null;
  }
}

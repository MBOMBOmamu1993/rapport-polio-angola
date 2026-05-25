"use client";

import { create } from "zustand";
import { persist } from "zustand/middleware";
import type { MasqueData } from "./parse-masque";

export interface Filters {
  /** Provinces sélectionnées (multi-sélection). Vide = toutes les provinces. */
  provinces: string[];
  antenne: string | null;
  zs: string | null;
  as: string | null;
}

interface AppState {
  data: MasqueData | null;
  filters: Filters;
  setData: (d: MasqueData) => void;
  clearData: () => void;
  setProvinces: (provinces: string[]) => void;
  toggleProvince: (name: string) => void;
  setFilter: (key: "antenne" | "zs" | "as", value: string | null) => void;
  resetFilter: (key: keyof Filters) => void;
  resetFilters: () => void;
}

const emptyFilters: Filters = { provinces: [], antenne: null, zs: null, as: null };

export const useApp = create<AppState>()(
  persist(
    (set) => ({
      data: null,
      filters: emptyFilters,
      // Réimporter remplace intégralement l'ancienne version et réinitialise les filtres.
      setData: (d) => set({ data: d, filters: emptyFilters }),
      clearData: () => set({ data: null, filters: emptyFilters }),
      // Changer la sélection de provinces réinitialise les niveaux inférieurs.
      setProvinces: (provinces) =>
        set((s) => ({ filters: { ...s.filters, provinces, antenne: null, zs: null, as: null } })),
      toggleProvince: (name) =>
        set((s) => {
          const provinces = s.filters.provinces.includes(name)
            ? s.filters.provinces.filter((p) => p !== name)
            : [...s.filters.provinces, name];
          return { filters: { ...s.filters, provinces, antenne: null, zs: null, as: null } };
        }),
      setFilter: (key, value) =>
        set((s) => {
          const next = { ...s.filters, [key]: value };
          // Réinitialise les niveaux inférieurs quand un niveau supérieur change.
          if (key === "antenne") { next.zs = null; next.as = null; }
          if (key === "zs") { next.as = null; }
          return { filters: next };
        }),
      // Réinitialise un filtre précis et, par cohérence de cascade, ses niveaux inférieurs.
      resetFilter: (key) =>
        set((s) => {
          const next = { ...s.filters };
          if (key === "provinces") { next.provinces = []; next.antenne = null; next.zs = null; next.as = null; }
          if (key === "antenne") { next.antenne = null; next.zs = null; next.as = null; }
          if (key === "zs") { next.zs = null; next.as = null; }
          if (key === "as") { next.as = null; }
          return { filters: next };
        }),
      resetFilters: () => set({ filters: emptyFilters }),
    }),
    {
      name: "polio-angola-masque",
      version: 2,
      // Normalise les états persistés antérieurs (province unique → tableau de provinces).
      migrate: (persisted) => {
        const state = (persisted ?? {}) as { data?: MasqueData | null; filters?: Record<string, unknown> };
        const f = state.filters ?? {};
        const legacyProvince = typeof f.province === "string" ? f.province : null;
        state.filters = {
          provinces: Array.isArray(f.provinces)
            ? (f.provinces as string[])
            : legacyProvince
            ? [legacyProvince]
            : [],
          antenne: (f.antenne as string | null) ?? null,
          zs: (f.zs as string | null) ?? null,
          as: (f.as as string | null) ?? null,
        };
        return state as unknown as AppState;
      },
    }
  )
);

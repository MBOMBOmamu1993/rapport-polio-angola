"use client";

import { create } from "zustand";
import { persist } from "zustand/middleware";
import type { MasqueData } from "./parse-masque";

export interface Filters {
  province: string | null;
  antenne: string | null;
  zs: string | null;
  as: string | null;
}

interface AppState {
  data: MasqueData | null;
  filters: Filters;
  setData: (d: MasqueData) => void;
  clearData: () => void;
  setFilter: (key: keyof Filters, value: string | null) => void;
  resetFilters: () => void;
}

const emptyFilters: Filters = { province: null, antenne: null, zs: null, as: null };

export const useApp = create<AppState>()(
  persist(
    (set) => ({
      data: null,
      filters: emptyFilters,
      // Réimporter remplace intégralement l'ancienne version et réinitialise les filtres.
      setData: (d) => set({ data: d, filters: emptyFilters }),
      clearData: () => set({ data: null, filters: emptyFilters }),
      setFilter: (key, value) =>
        set((s) => {
          const next = { ...s.filters, [key]: value };
          // Réinitialise les niveaux inférieurs quand un niveau supérieur change.
          if (key === "province") { next.antenne = null; next.zs = null; next.as = null; }
          if (key === "antenne") { next.zs = null; next.as = null; }
          if (key === "zs") { next.as = null; }
          return { filters: next };
        }),
      resetFilters: () => set({ filters: emptyFilters }),
    }),
    { name: "polio-angola-masque" }
  )
);

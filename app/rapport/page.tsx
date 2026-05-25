"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { useApp } from "@/lib/store";
import {
  applyFilters,
  aggregateByUnit,
  cascadeOptions,
  coverageByDay,
  nvpo2Gestion,
  vpobGestion,
  resolveDrillLevel,
  scopeLabel,
  totals,
} from "@/lib/analytics";
import { fmtInt, fmtPct } from "@/lib/format";
import { fetchNational } from "@/lib/national";
import type { MasqueData } from "@/lib/parse-masque";
import type { CompletudeRow, ProblemeRow, ReportData } from "@/lib/export-report-pptx";

const DEFAULT_PROBLEMES: ProblemeRow[] = [
  {
    probleme: "Écart entre données de vaccination nVPO2 et VPOb",
    causes: "Différents flacons reçus le matin dans les sites de stockage",
    zs: "—",
    solutions: "Concorder les doses lors de la transmission des vaccins aux équipes",
  },
  {
    probleme: "Faible remontée des données dans les ZS sans réseau Internet",
    causes: "Perturbation de la connectivité, faible rendement des collecteurs",
    zs: "—",
    solutions: "Traquer les collecteurs pour la transmission des données aux BCZ",
  },
  {
    probleme: "Perturbation de l'horaire de travail",
    causes: "Pluies",
    zs: "—",
    solutions: "Reprendre les activités après la pluie",
  },
  {
    probleme: "Taux de perte hors seuil (nVPO2 / VPOb)",
    causes: "Saisie irrégulière des flacons utilisés",
    zs: "—",
    solutions: "Revoir la saisie des flacons utilisés dans l'outil de collecte",
  },
];

export default function RapportPage() {
  const { data: localData, filters, setFilter, resetFilters } = useApp();
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(false);
  const [problemes, setProblemes] = useState<ProblemeRow[]>(DEFAULT_PROBLEMES);
  const [nationalData, setNationalData] = useState<MasqueData | null>(null);
  const [nationalLoaded, setNationalLoaded] = useState(false);
  const [source, setSource] = useState<"local" | "national">("local");

  useEffect(() => {
    let alive = true;
    fetchNational().then((nat) => {
      if (!alive) return;
      setNationalLoaded(true);
      if (nat && nat.data.records.length > 0) {
        setNationalData(nat.data);
        setSource("national");
      }
    });
    return () => { alive = false; };
  }, []);

  const data = source === "national" && nationalData ? nationalData : localData;

  const opts = useMemo(() => (data ? cascadeOptions(data, filters) : null), [data, filters]);
  const filtered = useMemo(() => (data ? applyFilters(data, filters) : []), [data, filters]);
  const drill = useMemo(() => resolveDrillLevel(filters), [filters]);
  const t = useMemo(() => totals(filtered), [filtered]);
  const byUnit = useMemo(() => aggregateByUnit(filtered, drill.level), [filtered, drill.level]);

  if (!data) {
    return (
      <div className="rounded-2xl border border-surface-200 bg-white p-10 text-center shadow-card">
        <div className="mb-3 text-5xl">📭</div>
        <h2 className="mb-2 text-lg font-bold text-navy-700">Aucun masque de saisie importé</h2>
        <p className="mb-4 text-sm text-surface-500">Importez d&apos;abord le masque de saisie pour générer le rapport.</p>
        <Link href="/import" className="inline-block rounded-xl bg-navy-700 px-6 py-3 text-sm font-semibold text-white hover:bg-navy-800">
          📥 Importer le masque de saisie
        </Link>
      </div>
    );
  }

  function buildReport(): ReportData {
    const nbJours = data!.meta.nbJours;
    const jourLabels = data!.meta.jourLabels;

    // Construction du tableau complétude par jour pour chaque unité.
    const completudeByUnit: CompletudeRow[] = byUnit.map((a) => {
      const daily = Array.from({ length: nbJours }, (_, i) => {
        const recus = a.rapportsRecusDaily[i] ?? 0;
        const att = a.vaccAttendus / Math.max(1, nbJours);
        return { recus, couv: att > 0 ? (recus / att) * 100 : null };
      });
      return {
        unit: a.unit,
        attendus: a.vaccAttendus,
        recus: a.vaccRecus,
        couv: a.vaccAttendus > 0 ? (a.vaccRecus / a.vaccAttendus) * 100 : null,
        daily,
      };
    });
    const recupByUnit = byUnit.map((a) => ({ unit: a.unit, value: a.recup }));

    return {
      province: filters.province ?? data!.meta.province,
      periode: data!.meta.periode,
      dateLabel: new Date().toLocaleDateString("fr-FR", { day: "2-digit", month: "long", year: "numeric" }),
      scopeLabel: scopeLabel(filters),
      byUnitLabel: drill.label,
      jourLabels,
      nbJours,
      saillants: {
        completude: t.completude,
        completudeAttendus: t.vaccAttendus,
        completudeRecus: t.vaccRecus,
        nvpo2Vacc: t.nvpo2Vacc,
        nvpo2Cible: t.nvpo2Cible,
        nvpo2CV: t.nvpo2CV,
        vpobVacc: t.vpobVacc,
        vpobCible: t.vpobCible,
        vpobCV: t.vpobCV,
        nvpo2FlaconsRecus: t.nvpo2FlaconsRecus,
        nvpo2FlaconsUtil: t.nvpo2FlaconsUtil,
        nvpo2FlaconsRendus: t.nvpo2FlaconsRendus,
        nvpo2Perdus: t.nvpo2Perdus,
        nvpo2Perte: t.nvpo2TauxPerte,
        vpobFlaconsRecus: t.vpobFlaconsRecus,
        vpobFlaconsUtil: t.vpobFlaconsUtil,
        vpobFlaconsRendus: t.vpobFlaconsRendus,
        vpobPerdus: t.vpobPerdus,
        vpobPerte: t.vpobTauxPerte,
        recup: t.recup,
        mapiMineures: t.mapiMineures,
        mapiGraves: t.mapiGraves,
      },
      completudeByUnit,
      nvpo2Daily: coverageByDay(byUnit, "nvpo2", nbJours),
      vpobDaily: coverageByDay(byUnit, "vpob", nbJours),
      nvpo2Gestion: nvpo2Gestion(byUnit),
      vpobGestion: vpobGestion(byUnit),
      recupByUnit,
      problemes,
    };
  }

  async function handleDownload() {
    setBusy(true);
    setDone(false);
    try {
      const report = buildReport();
      try {
        const { renderZSMap, normalizeZS } = await import("@/lib/zs-map");
        const byZS = new Map<string, { att: number; rec: number }>();
        for (const r of filtered) {
          const k = normalizeZS(r.zs);
          if (!k) continue;
          const acc = byZS.get(k) ?? { att: 0, rec: 0 };
          acc.att += r.vaccAttendus;
          acc.rec += r.vaccRecus;
          byZS.set(k, acc);
        }
        const completudeByZS = new Map<string, number>();
        for (const [k, v] of byZS) if (v.att > 0) completudeByZS.set(k, (v.rec / v.att) * 100);
        report.completudeMapPng = completudeByZS.size > 0 ? await renderZSMap(completudeByZS) : null;
      } catch {
        report.completudeMapPng = null;
      }
      const { exportReportPPT } = await import("@/lib/export-report-pptx");
      await exportReportPPT(report);
      setDone(true);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <h1 className="text-2xl font-bold text-navy-700 md:text-3xl">Télécharger le rapport</h1>
          <p className="text-sm text-surface-500">
            {scopeLabel(filters)} · {data.meta.periode || "Période de la campagne"} · agrégation par {drill.label}
          </p>
        </div>
        <div className="inline-flex rounded-xl border border-navy-100 bg-white p-1 shadow-card">
          <button
            onClick={() => { setSource("local"); resetFilters(); }}
            disabled={!localData}
            className={`rounded-lg px-3 py-2 text-xs font-semibold transition ${
              source === "local" ? "bg-navy-700 text-white shadow" : "text-navy-700 hover:bg-navy-50 disabled:opacity-40"
            }`}
          >
            💾 Mon import
          </button>
          <button
            onClick={() => { setSource("national"); resetFilters(); }}
            disabled={!nationalData}
            title={!nationalLoaded ? "Chargement…" : !nationalData ? "Compilation nationale non disponible" : ""}
            className={`rounded-lg px-3 py-2 text-xs font-semibold transition ${
              source === "national" ? "bg-navy-700 text-white shadow" : "text-navy-700 hover:bg-navy-50 disabled:opacity-40"
            }`}
          >
            🌍 Compilation nationale
          </button>
        </div>
      </div>

      {source === "national" && nationalData && (
        <div className="rounded-xl border border-navy-100 bg-navy-50 px-4 py-2.5 text-xs text-navy-700">
          🌍 Rapport <strong>niveau pays</strong> — {nationalData.meta.zones.length} ZS consolidées sur{" "}
          {Array.from(new Set(nationalData.records.map((r) => r.province))).length} province(s). Sélectionnez une
          province pour cibler, ou laissez « Tous » pour télécharger la situation de toutes les provinces à la fois.
        </div>
      )}

      {/* Filtres */}
      <section className="rounded-2xl border border-surface-200 bg-white p-4 shadow-card">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-bold text-navy-700">🔎 Filtres</h2>
          <button onClick={resetFilters} className="text-xs font-medium text-accent-600 hover:underline">
            Réinitialiser
          </button>
        </div>
        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
          <Select label="Province" value={filters.province} options={opts!.provinces} onChange={(v) => setFilter("province", v)} />
          <Select label="Antenne" value={filters.antenne} options={opts!.antennes} onChange={(v) => setFilter("antenne", v)} />
          <Select label="Zone de Santé" value={filters.zs} options={opts!.zones} onChange={(v) => setFilter("zs", v)} />
          <Select label="Aire de Santé" value={filters.as} options={opts!.aires} onChange={(v) => setFilter("as", v)} />
        </div>
      </section>

      {/* Résumé */}
      <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
        <Mini label="Aires couvertes" value={fmtInt(filtered.length)} />
        <Mini label="Complétude" value={fmtPct(t.completude)} tone={t.completude} />
        <Mini label="CV nVPO2" value={fmtPct(t.nvpo2CV)} tone={t.nvpo2CV} />
        <Mini label="CV VPOb" value={fmtPct(t.vpobCV)} tone={t.vpobCV} />
      </div>

      {/* Slides inclus */}
      <section className="rounded-2xl border border-surface-200 bg-white p-5 shadow-card">
        <h2 className="mb-4 text-sm font-bold text-navy-700">📊 Contenu du rapport PowerPoint</h2>
        <div className="grid grid-cols-1 gap-1.5 sm:grid-cols-2">
          {[
            "Page de garde (image polio + logos)",
            "Plan de présentation",
            "Points saillants",
            "Complétude des rapports par ZS (jauge + tableau J1..Jn)",
            "Couvertures vaccinales nVPO2 par jour (cible / vaccinés / couverture)",
            "Couvertures vaccinales VPOb par jour",
            "Récupération PEV de routine (co-administration)",
            "Gestion du vaccin nVPO2 (flacons reçus / utilisés / rendus / perdus)",
            "Gestion du vaccin VPOb",
            "Surveillance des MAPI",
            "Problèmes / Actions correctrices",
            "Merci pour votre attention",
          ].map((s, i) => (
            <div key={s} className="flex items-center gap-2 rounded-lg border border-surface-200 bg-surface-50 px-3 py-2 text-xs text-navy-700">
              <span className="flex h-5 w-5 items-center justify-center rounded-full bg-navy-700 text-[10px] font-bold text-white">{i + 1}</span>
              {s}
            </div>
          ))}
        </div>
        <div className="mt-5">
          <button
            onClick={handleDownload}
            disabled={busy}
            className="flex w-full items-center justify-center gap-2 rounded-xl bg-navy-700 py-3.5 text-sm font-semibold text-white shadow-card transition hover:bg-navy-800 disabled:opacity-50"
          >
            {busy ? (
              <>
                <span className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                Génération du rapport en cours…
              </>
            ) : (
              <>📊 Télécharger le rapport en PowerPoint (.pptx)</>
            )}
          </button>
          {done && <p className="mt-2 text-center text-xs font-medium text-good-600">✅ Rapport généré avec succès !</p>}
        </div>
      </section>

      {/* Problèmes éditables */}
      <section className="rounded-2xl border border-surface-200 bg-white p-5 shadow-card">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-bold text-navy-700">📝 Problèmes rencontrés / Actions correctrices</h2>
          <button
            onClick={() => setProblemes((p) => [...p, { probleme: "", causes: "", zs: "", solutions: "" }])}
            className="rounded-lg bg-accent-50 px-3 py-1.5 text-xs font-semibold text-accent-700 hover:bg-accent-100"
          >
            + Ajouter une ligne
          </button>
        </div>
        <div className="space-y-2">
          {problemes.map((p, i) => (
            <div key={i} className="grid grid-cols-1 gap-2 rounded-lg border border-surface-200 p-2 md:grid-cols-[1fr_1fr_120px_1fr_auto]">
              <Field placeholder="Problème identifié" value={p.probleme} onChange={(v) => editRow(setProblemes, i, "probleme", v)} />
              <Field placeholder="Causes" value={p.causes} onChange={(v) => editRow(setProblemes, i, "causes", v)} />
              <Field placeholder="ZS" value={p.zs} onChange={(v) => editRow(setProblemes, i, "zs", v)} />
              <Field placeholder="Solutions proposées" value={p.solutions} onChange={(v) => editRow(setProblemes, i, "solutions", v)} />
              <button
                onClick={() => setProblemes((arr) => arr.filter((_, j) => j !== i))}
                className="rounded-lg px-2 text-danger-500 hover:bg-danger-50"
                title="Supprimer"
              >
                ✕
              </button>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

function editRow(
  setter: React.Dispatch<React.SetStateAction<ProblemeRow[]>>,
  index: number,
  key: keyof ProblemeRow,
  value: string
) {
  setter((arr) => arr.map((r, j) => (j === index ? { ...r, [key]: value } : r)));
}

function Select({
  label,
  value,
  options,
  onChange,
}: {
  label: string;
  value: string | null;
  options: string[];
  onChange: (v: string | null) => void;
}) {
  return (
    <label className="block">
      <span className="mb-1 block text-[11px] font-semibold uppercase tracking-wide text-surface-400">{label}</span>
      <select
        value={value ?? ""}
        onChange={(e) => onChange(e.target.value || null)}
        className="w-full rounded-lg border border-surface-300 bg-white px-3 py-2.5 text-sm font-medium text-navy-700 focus:border-navy-500 focus:outline-none focus:ring-2 focus:ring-navy-200"
      >
        <option value="">Tous</option>
        {options.map((o) => (
          <option key={o} value={o}>{o}</option>
        ))}
      </select>
    </label>
  );
}

function Mini({ label, value, tone }: { label: string; value: string; tone?: number | null }) {
  const color =
    tone == null ? "text-navy-700" :
    tone > 100 ? "text-accent-600" :
    tone >= 95 ? "text-good-600" :
    tone >= 80 ? "text-warn-600" : "text-danger-500";
  return (
    <div className="rounded-xl border border-surface-200 bg-white p-3 shadow-card">
      <div className="text-[11px] uppercase tracking-wide text-surface-400">{label}</div>
      <div className={`text-2xl font-bold ${color}`}>{value}</div>
    </div>
  );
}

function Field({ placeholder, value, onChange }: { placeholder: string; value: string; onChange: (v: string) => void }) {
  return (
    <input
      value={value}
      placeholder={placeholder}
      onChange={(e) => onChange(e.target.value)}
      className="w-full rounded border border-surface-200 px-2 py-1.5 text-xs text-navy-700 focus:border-navy-500 focus:outline-none focus:ring-1 focus:ring-navy-200"
    />
  );
}

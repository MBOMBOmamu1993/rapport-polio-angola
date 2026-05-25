"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { useApp } from "@/lib/store";
import {
  applyFilters,
  aggregateByUnit,
  cascadeOptions,
  resolveDrillLevel,
  scopeLabel,
  totals,
} from "@/lib/analytics";
import { fmtInt, fmtPct } from "@/lib/format";
import { fetchNational } from "@/lib/national";
import type { MasqueData } from "@/lib/parse-masque";
import type { ProblemeRow, ReportData, UnitValue } from "@/lib/export-report-pptx";

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
        // Par défaut, privilégier la compilation nationale si disponible.
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
      <div className="rounded-2xl border border-surface-200 bg-white p-10 text-center">
        <div className="mb-3 text-5xl">📭</div>
        <h2 className="mb-2 text-lg font-bold text-oms-800">Aucun masque de saisie importé</h2>
        <p className="mb-4 text-sm text-surface-500">Importez d&apos;abord le masque de saisie pour générer le rapport.</p>
        <Link href="/import" className="inline-block rounded-xl bg-oms-500 px-6 py-3 text-sm font-semibold text-white hover:bg-oms-600">
          📥 Importer le masque de saisie
        </Link>
      </div>
    );
  }

  function buildReport(): ReportData {
    const toUV = (sel: (a: ReturnType<typeof aggregateByUnit>[number]) => number | null): UnitValue[] =>
      byUnit.map((a) => ({ unit: a.unit, value: sel(a) }));

    const pctOrNull = (p: number, w: number) => (w ? (p / w) * 100 : null);

    return {
      province: filters.province ?? data!.meta.province,
      periode: data!.meta.periode,
      dateLabel: new Date().toLocaleDateString("fr-FR", { day: "2-digit", month: "long", year: "numeric" }),
      scopeLabel: scopeLabel(filters),
      byUnitLabel: drill.label,
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
        nvpo2Flacons: t.nvpo2FlaconsUtil,
        nvpo2Perte: t.nvpo2TauxPerte,
        vpobFlacons: t.vpobFlaconsUtil,
        vpobPerte: t.vpobTauxPerte,
        recup: t.recup,
        mapiMineures: t.mapiMineures,
        mapiGraves: t.mapiGraves,
      },
      completudeByUnit: toUV((a) => pctOrNull(a.vaccRecus, a.vaccAttendus)),
      nvpo2CVByUnit: toUV((a) => pctOrNull(a.nvpo2Vacc, a.nvpo2Cible)),
      vpobCVByUnit: toUV((a) => pctOrNull(a.vpobVacc, a.vpobCible)),
      recupByUnit: toUV((a) => a.recup),
      nvpo2PerteByUnit: toUV((a) => pctOrNull(a.nvpo2Perdus, a.nvpo2FlaconsUtil)),
      vpobPerteByUnit: toUV((a) => pctOrNull(a.vpobPerdus, a.vpobFlaconsUtil)),
      problemes,
      commentNvpo2: `Couverture vaccinale nVPO2 de ${fmtPct(t.nvpo2CV)} sur le périmètre sélectionné${
        t.nvpo2CV !== null && t.nvpo2CV < 95 ? " — des zones restent sous le seuil de 95 %." : "."
      }`,
      commentVpob: `Couverture vaccinale VPOb de ${fmtPct(t.vpobCV)} sur le périmètre sélectionné${
        t.vpobCV !== null && t.vpobCV < 95 ? " — poursuivre les passages de rattrapage." : "."
      }`,
    };
  }

  async function handleDownload() {
    setBusy(true);
    setDone(false);
    try {
      const { exportReportPPT } = await import("@/lib/export-report-pptx");
      await exportReportPPT(buildReport());
      setDone(true);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <h1 className="text-2xl font-bold text-oms-800">Télécharger le rapport</h1>
          <p className="text-sm text-surface-500">
            {scopeLabel(filters)} · {data.meta.periode || "Période de la campagne"} · agrégation par {drill.label}
          </p>
        </div>
        {/* Source des données : import local vs compilation nationale */}
        <div className="inline-flex rounded-xl border border-surface-200 bg-white p-1 shadow-card">
          <button
            onClick={() => { setSource("local"); resetFilters(); }}
            disabled={!localData}
            className={`rounded-lg px-3 py-2 text-xs font-semibold transition ${
              source === "local" ? "bg-oms-500 text-white" : "text-oms-700 hover:bg-oms-50 disabled:opacity-40"
            }`}
          >
            💾 Mon import
          </button>
          <button
            onClick={() => { setSource("national"); resetFilters(); }}
            disabled={!nationalData}
            title={!nationalLoaded ? "Chargement…" : !nationalData ? "Compilation nationale non disponible" : ""}
            className={`rounded-lg px-3 py-2 text-xs font-semibold transition ${
              source === "national" ? "bg-oms-500 text-white" : "text-oms-700 hover:bg-oms-50 disabled:opacity-40"
            }`}
          >
            🌍 Compilation nationale
          </button>
        </div>
      </div>

      {source === "national" && nationalData && (
        <div className="rounded-xl border border-oms-100 bg-oms-50 px-4 py-2.5 text-xs text-oms-700">
          🌍 Rapport <strong>niveau pays</strong> — {nationalData.meta.zones.length} ZS consolidées sur {" "}
          {Array.from(new Set(nationalData.records.map((r) => r.province))).length} province(s). Sélectionnez une
          province pour cibler, ou laissez « Tous » pour télécharger la situation de toutes les provinces à la fois.
        </div>
      )}

      {/* Filtres */}
      <section className="rounded-2xl border border-surface-200 bg-white p-4 shadow-card">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-bold text-oms-800">🔎 Filtres</h2>
          <button onClick={resetFilters} className="text-xs font-medium text-oms-600 hover:underline">
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
        <Mini label="Complétude" value={fmtPct(t.completude)} />
        <Mini label="CV nVPO2" value={fmtPct(t.nvpo2CV)} />
        <Mini label="CV VPOb" value={fmtPct(t.vpobCV)} />
      </div>

      {/* Slides inclus */}
      <section className="rounded-2xl border border-surface-200 bg-white p-4 shadow-card">
        <h2 className="mb-3 text-sm font-bold text-oms-800">📊 Contenu du rapport PowerPoint</h2>
        <div className="grid grid-cols-1 gap-1.5 sm:grid-cols-2">
          {[
            "Page de garde (image polio + logos)",
            "Plan de présentation",
            "Points saillants",
            "Complétude des rapports par ZS",
            "Couvertures vaccinales nVPO2",
            "Couvertures vaccinales VPOb",
            "Récupération PEV de routine (co-administration)",
            "Gestion du vaccin nVPO2",
            "Gestion du vaccin VPOb",
            "Surveillance des MAPI",
            "Problèmes / Actions correctrices",
            "Merci pour votre attention",
          ].map((s, i) => (
            <div key={s} className="flex items-center gap-2 rounded border border-surface-200 bg-surface-50 px-3 py-1.5 text-xs text-oms-800">
              <span className="flex h-5 w-5 items-center justify-center rounded-full bg-oms-100 text-[10px] font-bold text-oms-700">{i + 1}</span>
              {s}
            </div>
          ))}
        </div>
        <div className="mt-4">
          <button
            onClick={handleDownload}
            disabled={busy}
            className="flex w-full items-center justify-center gap-2 rounded-xl bg-oms-500 py-3.5 text-sm font-semibold text-white transition hover:bg-oms-600 disabled:opacity-50"
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
      <section className="rounded-2xl border border-surface-200 bg-white p-4 shadow-card">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-bold text-oms-800">📝 Problèmes rencontrés / Actions correctrices</h2>
          <button
            onClick={() => setProblemes((p) => [...p, { probleme: "", causes: "", zs: "", solutions: "" }])}
            className="rounded-lg bg-oms-50 px-3 py-1.5 text-xs font-semibold text-oms-700 hover:bg-oms-100"
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
        className="w-full rounded-lg border border-surface-300 bg-white px-3 py-2 text-sm text-oms-800 focus:border-oms-500 focus:outline-none"
      >
        <option value="">Tous</option>
        {options.map((o) => (
          <option key={o} value={o}>{o}</option>
        ))}
      </select>
    </label>
  );
}

function Mini({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-xl border border-surface-200 bg-white p-3 shadow-card">
      <div className="text-[11px] uppercase tracking-wide text-surface-400">{label}</div>
      <div className="text-lg font-bold text-oms-800">{value}</div>
    </div>
  );
}

function Field({ placeholder, value, onChange }: { placeholder: string; value: string; onChange: (v: string) => void }) {
  return (
    <input
      value={value}
      placeholder={placeholder}
      onChange={(e) => onChange(e.target.value)}
      className="w-full rounded border border-surface-200 px-2 py-1.5 text-xs text-oms-800 focus:border-oms-500 focus:outline-none"
    />
  );
}

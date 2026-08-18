"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { useApp } from "@/lib/store";
import { applyFilters, cascadeOptions, resolveDrillLevel, scopeLabel, totals } from "@/lib/analytics";
import { fmtInt, fmtPct } from "@/lib/format";
import { fetchNational } from "@/lib/national";
import type { MasqueData } from "@/lib/parse-masque";
import { normUnit, type SupervisionPayload } from "@/lib/odk-supervision";
import {
  buildReportData,
  computeProblemes,
  prettyProvince,
  type ActionPCRow,
  type ProblemeRow,
} from "@/lib/report-data";

const SLIDES = [
  "Page de garde (logos MinSanté / PEV, photos, partenaires)",
  "Plan de présentation",
  "Suivi des points d'action des précédents PC (éditable ci‑dessous)",
  "Points saillants (complétude, CV RR / nVPO2 / VPOb, vaccinés, flacons, MAPI)",
  "Synthèse des principaux indicateurs par Antenne puis par Zone de Santé",
  "Complétude journalière et globale (tableaux + graphiques colorés)",
  "Couverture vaccinale RR et taux de perte (par Antenne, par ZS)",
  "Couverture vaccinale nVPO2 et taux de perte",
  "Couverture vaccinale VPOb et taux de perte",
  "Notification des MAPI (par Antenne / ZS / AS) + proportion pour 100 000 doses",
  "Surveillance MPV (PFA, rougeole, fièvre jaune, TNN)",
  "Récupération PEV de routine par antigène",
  "Proportion des vaccinés par tranche d'âge (RR, nVPO2, VPOb)",
  "Proportion des vaccinés par sexe",
  "Supervision des équipes (ODK) : % AS visitées, cartes, scores, conformité",
  "Problèmes rencontrés / Actions correctrices (générés + éditables)",
  "Merci",
];

const ACTIONS_KEY = "rrpolio-actions-pc";
const SETTINGS_KEY = "rrpolio-settings";

interface Settings {
  dateLancement: string;
  odkDateMin: string;
}

function loadJSON<T>(key: string, fallback: T): T {
  if (typeof window === "undefined") return fallback;
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
  }
}

export default function RapportPage() {
  const { data: localData, filters, setFilter, resetFilter, resetFilters } = useApp();
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [nationalData, setNationalData] = useState<MasqueData | null>(null);
  const [nationalLoaded, setNationalLoaded] = useState(false);
  const [source, setSource] = useState<"local" | "national">("local");
  const [actions, setActions] = useState<ActionPCRow[]>([]);
  const [settings, setSettings] = useState<Settings>({ dateLancement: "", odkDateMin: "2026-08-17" });
  const [problemes, setProblemes] = useState<ProblemeRow[]>([]);
  const [problemesTouched, setProblemesTouched] = useState(false);
  const [sup, setSup] = useState<SupervisionPayload | null>(null);
  const [supState, setSupState] = useState<"idle" | "loading" | "ok" | "error">("idle");
  const [supError, setSupError] = useState<string | undefined>(undefined);

  // Persistance légère des saisies utilisateur.
  useEffect(() => {
    setActions(loadJSON<ActionPCRow[]>(ACTIONS_KEY, []));
    setSettings((s) => ({ ...s, ...loadJSON<Partial<Settings>>(SETTINGS_KEY, {}) }));
  }, []);
  useEffect(() => { try { window.localStorage.setItem(ACTIONS_KEY, JSON.stringify(actions)); } catch { /* ignore */ } }, [actions]);
  useEffect(() => { try { window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)); } catch { /* ignore */ } }, [settings]);

  useEffect(() => {
    let alive = true;
    fetchNational().then((nat) => {
      if (!alive) return;
      setNationalLoaded(true);
      if (nat && nat.data.records.length > 0) {
        setNationalData(nat.data);
        if (!localData) setSource("national");
      }
    });
    return () => { alive = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const loadSupervision = useCallback(async (force = false) => {
    setSupState("loading");
    setSupError(undefined);
    try {
      const qs = new URLSearchParams();
      if (settings.odkDateMin) qs.set("dateMin", settings.odkDateMin);
      if (force) qs.set("force", "1");
      // Le serveur ODK peut être lent : deux tentatives.
      let res = await fetch(`/api/supervision?${qs.toString()}`, { cache: "no-store" });
      if (!res.ok) res = await fetch(`/api/supervision?${qs.toString()}`, { cache: "no-store" });
      const json = (await res.json()) as SupervisionPayload;
      if (!res.ok || !json.ok) {
        setSup(null);
        setSupError(json.reason ?? `HTTP ${res.status}`);
        setSupState("error");
        return;
      }
      setSup(json);
      setSupState("ok");
    } catch (e) {
      setSup(null);
      setSupError(e instanceof Error ? e.message : "réseau");
      setSupState("error");
    }
  }, [settings.odkDateMin]);

  useEffect(() => { void loadSupervision(false); }, [loadSupervision]);

  const data = source === "national" && nationalData ? nationalData : localData;
  const opts = useMemo(() => (data ? cascadeOptions(data, filters) : null), [data, filters]);
  const filtered = useMemo(() => (data ? applyFilters(data, filters) : []), [data, filters]);
  const drill = useMemo(() => resolveDrillLevel(filters), [filters]);
  const t = useMemo(() => totals(filtered), [filtered]);

  // Données de rapport (sans les cartes) — sert aussi à l'aperçu et aux problèmes auto.
  const report = useMemo(() => {
    if (!data) return null;
    return buildReportData({
      data,
      filters,
      supervision: sup,
      supervisionReason: supError,
      actionsPC: actions,
      dateLancement: settings.dateLancement || undefined,
    });
  }, [data, filters, sup, supError, actions, settings.dateLancement]);

  useEffect(() => {
    if (!report || problemesTouched) return;
    setProblemes(computeProblemes(report.units, report.total, report.byUnitLabel, report.supervision));
  }, [report, problemesTouched]);

  // Supervisions du périmètre (aperçu).
  const supPreview = useMemo(() => {
    if (!report) return null;
    const s = report.supervision;
    const visited = s.byZS.reduce((a, z) => a + z.nbASVisitees, 0);
    const totalAS = s.byZS.reduce((a, z) => a + z.nbASTotal, 0);
    return { total: s.total, visited, totalAS, pct: totalAS ? (visited / totalAS) * 100 : null, under80: s.byZS.filter((z) => (z.pctASVisitees ?? 0) < 80).length, nbZS: s.byZS.length };
  }, [report]);

  if (!data) {
    return (
      <div className="rounded-2xl border border-surface-200 bg-white p-10 text-center shadow-card">
        <div className="mb-3 text-5xl">📭</div>
        <h2 className="mb-2 text-lg font-bold text-navy-700">Aucun masque de saisie importé</h2>
        <p className="mb-4 text-sm text-surface-500">
          Importez d&apos;abord le masque de saisie intégré RR‑Polio du Kasaï Central pour générer le rapport.
          {nationalLoaded && !nationalData && " La compilation partagée est vide pour l'instant."}
        </p>
        <Link href="/import" className="inline-block rounded-xl bg-navy-700 px-6 py-3 text-sm font-semibold text-white hover:bg-navy-800">
          📥 Importer le masque de saisie
        </Link>
      </div>
    );
  }

  async function handleDownload() {
    if (!report) return;
    setBusy(true);
    setDone(null);
    setError(null);
    try {
      const rep = { ...report, problemes };
      // Cartes de supervision (fond ZS du Kasaï Central) — rendu navigateur.
      try {
        const { renderProvinceMap, colorFor3 } = await import("@/lib/zs-map");
        const fill = new Map<string, string>();
        for (const z of rep.supervision.byZS) fill.set(normUnit(z.zs), colorFor3(z.pctASVisitees));
        const [mapPng, pointsMapPng] = await Promise.all([
          renderProvinceMap({ province: rep.province, fillByZS: fill, labels: true }),
          renderProvinceMap({ province: rep.province, points: rep.supervision.points, dark: true, labels: true }),
        ]);
        rep.supervision = { ...rep.supervision, mapPng, pointsMapPng };
      } catch {
        rep.supervision = { ...rep.supervision, mapPng: null, pointsMapPng: null };
      }
      const { exportReportPPT } = await import("@/lib/export-report-pptx");
      await exportReportPPT(rep);
      setDone(`Rapport généré (${new Date().toLocaleTimeString("fr-FR")}).`);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Échec de la génération.");
    } finally {
      setBusy(false);
    }
  }

  const provinceLabel = prettyProvince(data.meta.province);

  return (
    <div className="space-y-6">
      {/* En‑tête */}
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <div className="text-[11px] font-semibold uppercase tracking-wider text-accent-600">Étape 2 / 2</div>
          <h1 className="text-2xl font-bold text-navy-700 md:text-3xl">Télécharger le rapport PowerPoint</h1>
          <p className="text-sm text-surface-500">
            {scopeLabel(filters, provinceLabel)} · {data.meta.jourLabels.length ? `${data.meta.jourLabels[0]} → ${data.meta.jourLabels[data.meta.jourLabels.length - 1]}` : "—"} · agrégation par {drill.label}
          </p>
        </div>
        <div className="inline-flex rounded-xl border border-navy-100 bg-white p-1 shadow-card">
          <button
            onClick={() => { setSource("local"); resetFilters(); }}
            disabled={!localData}
            className={`rounded-lg px-3 py-2 text-xs font-semibold transition ${source === "local" ? "bg-navy-700 text-white shadow" : "text-navy-700 hover:bg-navy-50 disabled:opacity-40"}`}
          >
            💾 Mon import
          </button>
          <button
            onClick={() => { setSource("national"); resetFilters(); }}
            disabled={!nationalData}
            title={!nationalLoaded ? "Chargement…" : !nationalData ? "Compilation partagée non disponible" : ""}
            className={`rounded-lg px-3 py-2 text-xs font-semibold transition ${source === "national" ? "bg-navy-700 text-white shadow" : "text-navy-700 hover:bg-navy-50 disabled:opacity-40"}`}
          >
            🌍 Compilation provinciale
          </button>
        </div>
      </div>

      {source === "national" && nationalData && (
        <div className="rounded-xl border border-navy-100 bg-navy-50 px-4 py-2.5 text-xs text-navy-700">
          🌍 Compilation partagée — {nationalData.meta.zones.length} ZS et {fmtInt(nationalData.meta.nbAires)} AS consolidées (tous les imports des antennes / ZS du Kasaï Central).
        </div>
      )}

      {/* Filtres */}
      <section className="rounded-2xl border border-surface-200 bg-white p-4 shadow-card">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-bold text-navy-700">🔎 Périmètre du rapport</h2>
          <button onClick={resetFilters} className="text-xs font-medium text-accent-600 hover:underline">Réinitialiser</button>
        </div>
        <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
          <Select label="Antenne" value={filters.antenne} options={opts!.antennes} onChange={(v) => setFilter("antenne", v)} onReset={() => resetFilter("antenne")} />
          <Select label="Zone de Santé" value={filters.zs} options={opts!.zones} onChange={(v) => setFilter("zs", v)} onReset={() => resetFilter("zs")} />
          <Select label="Aire de Santé" value={filters.as} options={opts!.aires} onChange={(v) => setFilter("as", v)} onReset={() => resetFilter("as")} />
        </div>
      </section>

      {/* KPI */}
      <div className="grid grid-cols-2 gap-3 md:grid-cols-4 lg:grid-cols-6">
        <Mini label="Aires de Santé" value={fmtInt(filtered.length)} />
        <Mini label="Complétude" value={fmtPct(t.completude)} tone={t.completude} />
        <Mini label="CV RR" value={fmtPct(t.rr.cv)} tone={t.rr.cv} />
        <Mini label="CV nVPO2" value={fmtPct(t.nvpo2.cv)} tone={t.nvpo2.cv} />
        <Mini label="CV VPOb" value={fmtPct(t.vpob.cv)} tone={t.vpob.cv} />
        <Mini label="MAPI graves" value={fmtInt(t.mapiGraves)} tone={t.mapiGraves > 0 ? 0 : null} />
      </div>

      {/* Supervision ODK */}
      <section className="rounded-2xl border border-surface-200 bg-white p-4 shadow-card">
        <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
          <h2 className="text-sm font-bold text-navy-700">🛰️ Supervision des équipes (ODK — formulaire Bloc 3)</h2>
          <div className="flex items-center gap-2">
            <SupBadge state={supState} />
            <button onClick={() => void loadSupervision(true)} className="rounded-lg border border-surface-300 px-3 py-1.5 text-xs font-medium text-navy-700 hover:bg-surface-100">↻ Actualiser</button>
          </div>
        </div>
        {supState === "error" && (
          <p className="mb-2 rounded-lg bg-danger-50 px-3 py-2 text-xs text-danger-700">
            Données ODK indisponibles ({supError}). Le rapport contiendra une diapositive « données de supervision indisponibles ».
          </p>
        )}
        {supPreview && supState === "ok" && (
          <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
            <Mini label="Supervisions (périmètre)" value={fmtInt(supPreview.total)} />
            <Mini label="AS visitées" value={`${fmtInt(supPreview.visited)} / ${fmtInt(supPreview.totalAS)}`} />
            <Mini label="% AS visitées" value={fmtPct(supPreview.pct)} tone={supPreview.pct == null ? null : supPreview.pct >= 80 ? 96 : supPreview.pct >= 50 ? 85 : 20} />
            <Mini label="ZS < 80 % d'AS visitées" value={`${supPreview.under80} / ${supPreview.nbZS}`} />
          </div>
        )}
        <div className="mt-3 grid grid-cols-1 gap-3 sm:grid-cols-2">
          <label className="block text-xs">
            <span className="mb-1 block font-semibold uppercase tracking-wide text-surface-400">Date de lancement (page de garde / synthèse)</span>
            <input type="date" value={settings.dateLancement} onChange={(e) => setSettings((s) => ({ ...s, dateLancement: e.target.value }))} className="w-full rounded-lg border border-surface-300 px-3 py-2 text-sm text-navy-700" />
            <span className="mt-1 block text-[11px] text-surface-400">Par défaut : {data.meta.dateDebut || "date du masque"}.</span>
          </label>
          <label className="block text-xs">
            <span className="mb-1 block font-semibold uppercase tracking-wide text-surface-400">Supervisions ODK à partir du</span>
            <input type="date" value={settings.odkDateMin} onChange={(e) => setSettings((s) => ({ ...s, odkDateMin: e.target.value }))} className="w-full rounded-lg border border-surface-300 px-3 py-2 text-sm text-navy-700" />
            <span className="mt-1 block text-[11px] text-surface-400">Province Kasaï Central uniquement (17/08/2026 par défaut).</span>
          </label>
        </div>
      </section>

      {/* Génération */}
      <section className="rounded-2xl border border-surface-200 bg-white p-5 shadow-card">
        <h2 className="mb-4 text-sm font-bold text-navy-700">📊 Contenu du rapport PowerPoint (design du modèle Bloc 3)</h2>
        <div className="grid grid-cols-1 gap-1.5 sm:grid-cols-2">
          {SLIDES.map((s, i) => (
            <div key={s} className="flex items-center gap-2 rounded-lg border border-surface-200 bg-surface-50 px-3 py-2 text-xs text-navy-700">
              <span className="flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-navy-700 text-[10px] font-bold text-white">{i + 1}</span>
              {s}
            </div>
          ))}
        </div>
        <div className="mt-5">
          <button
            onClick={handleDownload}
            disabled={busy || supState === "loading"}
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
          {done && <p className="mt-2 text-center text-xs font-medium text-good-600">✅ {done}</p>}
          {error && <p className="mt-2 text-center text-xs font-medium text-danger-600">⚠️ {error}</p>}
        </div>
      </section>

      {/* Points d'action des précédents PC */}
      <section className="rounded-2xl border border-surface-200 bg-white p-5 shadow-card">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="text-sm font-bold text-navy-700">📌 Suivi des points d&apos;action des précédents PC</h2>
          <button
            onClick={() => setActions((a) => [...a, { activite: "", responsable: "", statut: "En cours", echeance: "" }])}
            className="rounded-lg bg-accent-50 px-3 py-1.5 text-xs font-semibold text-accent-700 hover:bg-accent-100"
          >
            + Ajouter une ligne
          </button>
        </div>
        {actions.length === 0 && <p className="text-xs text-surface-400">Aucune ligne — la diapositive indiquera « à compléter ». Ajoutez les activités décidées lors du précédent poste de commandement.</p>}
        <div className="space-y-2">
          {actions.map((a, i) => (
            <div key={i} className="grid grid-cols-1 gap-2 rounded-lg border border-surface-200 p-2 md:grid-cols-[2fr_1fr_1fr_1fr_auto]">
              <Field placeholder="Activité prévue" value={a.activite} onChange={(v) => setActions((arr) => arr.map((r, j) => (j === i ? { ...r, activite: v } : r)))} />
              <Field placeholder="Responsable" value={a.responsable} onChange={(v) => setActions((arr) => arr.map((r, j) => (j === i ? { ...r, responsable: v } : r)))} />
              <select value={a.statut} onChange={(e) => setActions((arr) => arr.map((r, j) => (j === i ? { ...r, statut: e.target.value } : r)))} className="rounded border border-surface-200 px-2 py-1.5 text-xs text-navy-700">
                {["Réalisée", "En cours", "En retard", "Non réalisée"].map((s) => <option key={s}>{s}</option>)}
              </select>
              <Field placeholder="Échéance" value={a.echeance} onChange={(v) => setActions((arr) => arr.map((r, j) => (j === i ? { ...r, echeance: v } : r)))} />
              <button onClick={() => setActions((arr) => arr.filter((_, j) => j !== i))} className="rounded-lg px-2 text-danger-500 hover:bg-danger-50" title="Supprimer">✕</button>
            </div>
          ))}
        </div>
      </section>

      {/* Problèmes éditables */}
      <section className="rounded-2xl border border-surface-200 bg-white p-5 shadow-card">
        <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
          <h2 className="text-sm font-bold text-navy-700">📝 Problèmes rencontrés / Actions correctrices (générés automatiquement, modifiables)</h2>
          <div className="flex gap-2">
            <button onClick={() => { setProblemesTouched(false); if (report) setProblemes(computeProblemes(report.units, report.total, report.byUnitLabel, report.supervision)); }} className="rounded-lg border border-surface-300 px-3 py-1.5 text-xs font-medium text-navy-700 hover:bg-surface-100">↻ Régénérer</button>
            <button onClick={() => { setProblemesTouched(true); setProblemes((p) => [...p, { probleme: "", causes: "", zs: "", solutions: "" }]); }} className="rounded-lg bg-accent-50 px-3 py-1.5 text-xs font-semibold text-accent-700 hover:bg-accent-100">+ Ajouter une ligne</button>
          </div>
        </div>
        <div className="space-y-2">
          {problemes.map((p, i) => (
            <div key={i} className="grid grid-cols-1 gap-2 rounded-lg border border-surface-200 p-2 md:grid-cols-[1fr_1fr_140px_1fr_auto]">
              <Field placeholder="Problème identifié" value={p.probleme} onChange={(v) => { setProblemesTouched(true); editRow(setProblemes, i, "probleme", v); }} />
              <Field placeholder="Causes" value={p.causes} onChange={(v) => { setProblemesTouched(true); editRow(setProblemes, i, "causes", v); }} />
              <Field placeholder={`${drill.label}(s) concernée(s)`} value={p.zs} onChange={(v) => { setProblemesTouched(true); editRow(setProblemes, i, "zs", v); }} />
              <Field placeholder="Solutions proposées" value={p.solutions} onChange={(v) => { setProblemesTouched(true); editRow(setProblemes, i, "solutions", v); }} />
              <button onClick={() => { setProblemesTouched(true); setProblemes((arr) => arr.filter((_, j) => j !== i)); }} className="rounded-lg px-2 text-danger-500 hover:bg-danger-50" title="Supprimer">✕</button>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

function editRow(setter: React.Dispatch<React.SetStateAction<ProblemeRow[]>>, index: number, key: keyof ProblemeRow, value: string) {
  setter((arr) => arr.map((r, j) => (j === index ? { ...r, [key]: value } : r)));
}

function SupBadge({ state }: { state: "idle" | "loading" | "ok" | "error" }) {
  const map = {
    idle: ["bg-surface-100 text-surface-500", "En attente"],
    loading: ["bg-navy-50 text-navy-700", "Connexion ODK…"],
    ok: ["bg-good-50 text-good-600", "Connecté à ODK"],
    error: ["bg-danger-50 text-danger-600", "ODK indisponible"],
  } as const;
  const [cls, label] = map[state];
  return <span className={`rounded-full px-2.5 py-1 text-[11px] font-semibold ${cls}`}>{label}</span>;
}

function FilterHeader({ label, active, onReset }: { label: string; active: boolean; onReset: () => void }) {
  return (
    <div className="mb-1 flex items-center justify-between gap-2">
      <span className="text-[11px] font-semibold uppercase tracking-wide text-surface-400">{label}</span>
      {active && (
        <button type="button" onClick={onReset} title={`Réinitialiser : ${label}`} className="inline-flex h-5 w-5 items-center justify-center rounded-full text-sm leading-none text-surface-400 transition hover:bg-accent-50 hover:text-accent-600">↺</button>
      )}
    </div>
  );
}

function Select({ label, value, options, onChange, onReset }: { label: string; value: string | null; options: string[]; onChange: (v: string | null) => void; onReset: () => void }) {
  return (
    <div className="block">
      <FilterHeader label={label} active={value != null} onReset={onReset} />
      <select
        value={value ?? ""}
        onChange={(e) => onChange(e.target.value || null)}
        className="w-full rounded-lg border border-surface-300 bg-white px-3 py-2.5 text-sm font-medium text-navy-700 focus:border-navy-500 focus:outline-none focus:ring-2 focus:ring-navy-200"
      >
        <option value="">Tous</option>
        {options.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
    </div>
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
      <div className={`text-xl font-bold md:text-2xl ${color}`}>{value}</div>
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

"use client";

/**
 * Onglet « RR-polio DHIS2 Bloc 3 » : génération automatique du rapport PowerPoint
 * (modèle Bloc 3, identique à l'onglet Téléchargement) à partir du DHIS2 de
 * campagne (rdccampagne.hispwca.org), à tous les niveaux :
 *  - plusieurs provinces → situation par province + détail par antenne ;
 *  - une province → situation par antenne + détail par ZS ;
 *  - une antenne → détail par ZS ; une ZS → détail par AS.
 * Les provinces en campagne RR seule (sans volet polio) s'affichent « — » sur
 * les indicateurs polio.
 */

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { applyFilters, cascadeOptions, resolveDrillLevel, totals } from "@/lib/analytics";
import { fmtInt, fmtPct } from "@/lib/format";
import { mergeBlocks, type ProvinceBlock, type ProvinceInfo, type ProvinceListPayload } from "@/lib/dhis2-shared";
import { normUnit, type SupervisionPayload } from "@/lib/odk-supervision";
import type { Filters } from "@/lib/store";
import {
  buildReportData,
  computeProblemes,
  fmtDateFR,
  prettyProvince,
  type ActionPCRow,
  type ProblemeRow,
} from "@/lib/report-data";

const SLIDES = [
  "Page de garde (logos MinSanté / PEV, photos, partenaires)",
  "Plan de présentation",
  "Suivi des points d'action des précédents PC (éditable ci-dessous)",
  "Points saillants (complétude, CV RR / nVPO2 / VPOb, vaccinés, flacons, MAPI)",
  "Synthèse des principaux indicateurs (par Province / Antenne / ZS selon le périmètre)",
  "Complétude journalière et globale (tableaux + graphiques colorés)",
  "Couverture vaccinale RR et taux de perte",
  "Couverture vaccinale nVPO2 et VPOb (provinces RR-POLIO uniquement)",
  "Notification des MAPI + proportion pour 100 000 doses",
  "Surveillance MPV (PFA, rougeole, fièvre jaune, TNN)",
  "Récupération PEV de routine par antigène",
  "Proportion des vaccinés par tranche d'âge et par sexe",
  "Supervision des équipes (ODK) : % AS visitées, cartes, scores, conformité",
  "Problèmes rencontrés / Actions correctrices (générés + éditables)",
  "Merci",
];

const ACTIONS_KEY = "rrpolio-dhis2-actions-pc";
const SEL_KEY = "rrpolio-dhis2-provinces";

function loadJSON<T>(key: string, fallback: T): T {
  if (typeof window === "undefined") return fallback;
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
  }
}

type SubFilters = { antenne: string | null; zs: string | null; as: string | null };

export default function Dhis2Page() {
  const [list, setList] = useState<ProvinceInfo[] | null>(null);
  const [listError, setListError] = useState<string | null>(null);
  const [selected, setSelected] = useState<string[]>([]);
  const [blocks, setBlocks] = useState<Record<string, ProvinceBlock>>({});
  const [loading, setLoading] = useState<Record<string, boolean>>({});
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [sub, setSub] = useState<SubFilters>({ antenne: null, zs: null, as: null });
  const [actions, setActions] = useState<ActionPCRow[]>([]);
  const [problemes, setProblemes] = useState<ProblemeRow[]>([]);
  const [problemesTouched, setProblemesTouched] = useState(false);
  const [sup, setSup] = useState<SupervisionPayload | null>(null);
  const [supState, setSupState] = useState<"idle" | "loading" | "ok" | "error" | "off">("idle");
  const [supError, setSupError] = useState<string | undefined>(undefined);
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const alive = useRef(true);

  useEffect(() => {
    alive.current = true;
    setActions(loadJSON<ActionPCRow[]>(ACTIONS_KEY, []));
    setSelected(loadJSON<string[]>(SEL_KEY, []));
    return () => { alive.current = false; };
  }, []);
  useEffect(() => { try { window.localStorage.setItem(ACTIONS_KEY, JSON.stringify(actions)); } catch { /* ignore */ } }, [actions]);
  useEffect(() => { try { window.localStorage.setItem(SEL_KEY, JSON.stringify(selected)); } catch { /* ignore */ } }, [selected]);

  /* Liste des provinces de la campagne (l'API répond 202 pendant la 1re extraction). */
  const loadList = useCallback(async (force = false) => {
    setListError(null);
    const deadline = Date.now() + 4 * 60 * 1000;
    let qs = force ? "&force=1" : "";
    try {
      for (;;) {
        const res = await fetch(`/api/dhis2?list=1${qs}`, { cache: "no-store" });
        if (!alive.current) return;
        if (res.status === 202 && Date.now() < deadline) {
          qs = "";
          await new Promise((r) => setTimeout(r, 5000));
          continue;
        }
        const json = (await res.json().catch(() => ({}))) as ProvinceListPayload & { reason?: string };
        if (res.ok && json.ok) {
          setList(json.provinces);
          return;
        }
        setListError(json.reason ?? `HTTP ${res.status}`);
        return;
      }
    } catch (e) {
      if (alive.current) setListError(e instanceof Error ? e.message : "réseau");
    }
  }, []);
  useEffect(() => { void loadList(false); }, [loadList]);

  /* Extraction d'une province (avec relance tant que 202 / stale). */
  const loadBlock = useCallback(async (id: string, force = false) => {
    setLoading((l) => ({ ...l, [id]: true }));
    setErrors((e) => { const n = { ...e }; delete n[id]; return n; });
    const deadline = Date.now() + 5 * 60 * 1000;
    let qs = force ? "&force=1" : "";
    try {
      for (;;) {
        const res = await fetch(`/api/dhis2?province=${id}${qs}`, { cache: "no-store" });
        if (!alive.current) return;
        if (res.status === 202 && Date.now() < deadline) {
          qs = "";
          await new Promise((r) => setTimeout(r, 6000));
          continue;
        }
        const json = (await res.json().catch(() => ({}))) as ProvinceBlock & { reason?: string };
        if (res.ok && json.ok) {
          setBlocks((b) => ({ ...b, [id]: json }));
          // Valeur antérieure servie pendant qu'un rafraîchissement tourne : relire une fois.
          if (json.stale && Date.now() < deadline && !qs) {
            await new Promise((r) => setTimeout(r, 20000));
            const res2 = await fetch(`/api/dhis2?province=${id}`, { cache: "no-store" });
            if (!alive.current) return;
            const json2 = (await res2.json().catch(() => ({}))) as ProvinceBlock;
            if (res2.ok && json2.ok) setBlocks((b) => ({ ...b, [id]: json2 }));
          }
          return;
        }
        setErrors((e) => ({ ...e, [id]: json.reason ?? `HTTP ${res.status}` }));
        return;
      }
    } catch (e) {
      if (alive.current) setErrors((er) => ({ ...er, [id]: e instanceof Error ? e.message : "réseau" }));
    } finally {
      if (alive.current) setLoading((l) => ({ ...l, [id]: false }));
    }
  }, []);

  useEffect(() => {
    for (const id of selected) {
      if (!blocks[id] && !loading[id] && !errors[id]) void loadBlock(id);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selected, blocks, loadBlock]);

  const toggleProvince = (id: string) => {
    setSub({ antenne: null, zs: null, as: null });
    setSelected((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  };

  const selectedBlocks = useMemo(
    () => selected.map((id) => blocks[id]).filter((b): b is ProvinceBlock => Boolean(b?.ok)),
    [selected, blocks]
  );
  const pendingCount = selected.filter((id) => !blocks[id]).length;
  const data = useMemo(() => mergeBlocks(selectedBlocks), [selectedBlocks]);

  const filters: Filters = useMemo(
    () => ({ provinces: selectedBlocks.map((b) => b.province), antenne: sub.antenne, zs: sub.zs, as: sub.as }),
    [selectedBlocks, sub]
  );
  const opts = useMemo(() => (data ? cascadeOptions(data, filters) : null), [data, filters]);
  const filtered = useMemo(() => (data ? applyFilters(data, filters) : []), [data, filters]);
  const nbProv = useMemo(() => new Set(filtered.map((r) => r.province)).size, [filtered]);
  const drill = useMemo(() => resolveDrillLevel(filters, nbProv > 1), [filters, nbProv]);
  const t = useMemo(() => totals(filtered), [filtered]);

  /* Supervision ODK — chargée pour une seule province à la fois (serveur très lent).
     Supervisions retenues à partir du 11/08/2026, sauf Kasaï Central et Nord Kivu
     (lancements décalés) : à partir du 17/08/2026. */
  const supProvince = selectedBlocks.length === 1 ? selectedBlocks[0].province : null;
  const supDateMin = useMemo(() => {
    if (!supProvince) return null;
    const key = supProvince.normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[^A-Z]/g, "");
    return key === "KASAICENTRAL" || key === "NORDKIVU" ? "2026-08-17" : "2026-08-11";
  }, [supProvince]);
  const loadSupervision = useCallback(async (force = false) => {
    if (!supProvince || !supDateMin) {
      setSup(null);
      setSupState("off");
      return;
    }
    setSupState("loading");
    setSupError(undefined);
    const qs = new URLSearchParams({ province: supProvince, dateMin: supDateMin });
    if (force) qs.set("force", "1");
    // Première extraction d'une province : le serveur ODK peut prendre très longtemps.
    const deadline = Date.now() + 12 * 60 * 1000;
    try {
      for (;;) {
        const res = await fetch(`/api/supervision?${qs.toString()}`, { cache: "no-store" });
        if (!alive.current) return;
        const json = (await res.json().catch(() => ({}))) as SupervisionPayload & { pending?: boolean };
        if (res.ok && json.ok) {
          setSup(json);
          setSupState("ok");
          return;
        }
        if (res.status === 202 && Date.now() < deadline) {
          qs.delete("force");
          await new Promise((r) => setTimeout(r, 8000));
          continue;
        }
        setSup(null);
        setSupError(json.reason ?? `HTTP ${res.status}`);
        setSupState("error");
        return;
      }
    } catch (e) {
      if (!alive.current) return;
      setSup(null);
      setSupError(e instanceof Error ? e.message : "réseau");
      setSupState("error");
    }
  }, [supProvince, supDateMin]);
  useEffect(() => { void loadSupervision(false); }, [loadSupervision]);

  const dateLancement = useMemo(() => {
    const j1s = selectedBlocks.map((b) => b.j1).filter(Boolean).sort();
    return j1s[0] ?? "2026-08-11";
  }, [selectedBlocks]);

  const titre = useMemo(() => {
    if (selectedBlocks.length === 1) return `Campagne intégrée RR-POLIO ${prettyProvince(selectedBlocks[0].province)}`;
    return "Campagne intégrée RR-POLIO RD Congo";
  }, [selectedBlocks]);

  const report = useMemo(() => {
    if (!data) return null;
    return buildReportData({
      data,
      filters,
      supervision: supState === "ok" ? sup : null,
      supervisionReason:
        supState === "off"
          ? "supervision ODK affichée pour une province à la fois — sélectionnez une seule province"
          : supError,
      actionsPC: actions,
      dateLancement,
      titre,
    });
  }, [data, filters, sup, supState, supError, actions, dateLancement, titre]);

  useEffect(() => {
    if (!report || problemesTouched) return;
    setProblemes(computeProblemes(report.units, report.total, report.byUnitLabel, report.supervision));
  }, [report, problemesTouched]);

  async function handleDownload() {
    if (!report) return;
    setBusy(true);
    setDone(null);
    setError(null);
    try {
      const rep = { ...report, problemes };
      // Cartes de supervision : rendues seulement quand une seule province est sélectionnée.
      if (selectedBlocks.length === 1) {
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
      } else {
        rep.supervision = { ...rep.supervision, mapPng: null, pointsMapPng: null };
      }
      const { exportReportPPT } = await import("@/lib/export-report-pptx");
      // Sélection large (≥ 4 provinces) : nom de fichier calqué sur le rapport
      // national Bloc 3 (« Résultats_Partiels_Campagne intégrée_Bloc3_Aout 2026_J1-J5 »).
      const fileName =
        selectedBlocks.length >= 4 && !sub.antenne && !sub.zs && !sub.as
          ? `Resultats_Partiels_Campagne_integree_Bloc3_Aout_2026_J1-J5_${rep.dateMaj.replace(/\//g, "-")}.pptx`
          : undefined;
      const hasMasque = selectedBlocks.some((b) => b.source === "masque");
      const hasDhis2 = selectedBlocks.some((b) => b.source !== "masque");
      await exportReportPPT(rep, {
        fileName,
        sourceText: hasMasque && hasDhis2
          ? "DHIS2 de campagne (rdccampagne.hispwca.org) + masque de saisie importé (Kasaï Central)"
          : hasMasque
            ? "Masque de saisie de la campagne (dernier import)"
            : "DHIS2 de campagne (rdccampagne.hispwca.org) — dataset PEV_Campagne RR et Polio",
      });
      setDone(`Rapport généré (${new Date().toLocaleTimeString("fr-FR")}).`);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Échec de la génération.");
    } finally {
      setBusy(false);
    }
  }

  const supPreview = useMemo(() => {
    if (!report || supState !== "ok") return null;
    const s = report.supervision;
    const visited = s.byZS.reduce((a, z) => a + z.nbASVisitees, 0);
    const totalAS = s.byZS.reduce((a, z) => a + z.nbASTotal, 0);
    return { total: s.total, visited, totalAS, pct: totalAS ? (visited / totalAS) * 100 : null };
  }, [report, supState]);

  return (
    <div className="space-y-6">
      <div>
        <div className="text-[11px] font-semibold uppercase tracking-wider text-accent-600">Connexion directe DHIS2 — rdccampagne.hispwca.org</div>
        <h1 className="text-2xl font-bold text-navy-700 md:text-3xl">RR-polio DHIS2 Bloc 3</h1>
        <p className="text-sm text-surface-500">
          Rapport PowerPoint automatique (modèle Bloc 3) à tous les niveaux : Provinces → Antennes → Zones de Santé → Aires de Santé.
          Les provinces en campagne <b>RR seule</b> s&apos;affichent « — » sur le volet polio.
        </p>
      </div>

      {/* 1. Provinces (multi-sélection) */}
      <section className="rounded-2xl border border-surface-200 bg-white p-4 shadow-card">
        <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
          <h2 className="text-sm font-bold text-navy-700">🗺️ Provinces de la campagne (multi-sélection)</h2>
          <div className="flex items-center gap-2">
            {list && (
              <>
                <button
                  onClick={() => { setSub({ antenne: null, zs: null, as: null }); setSelected(list.map((p) => p.id)); }}
                  className="rounded-lg bg-accent-50 px-3 py-1.5 text-xs font-semibold text-accent-700 hover:bg-accent-100"
                >
                  Tout sélectionner
                </button>
                <button
                  onClick={() => { setSub({ antenne: null, zs: null, as: null }); setSelected([]); }}
                  className="rounded-lg border border-surface-300 px-3 py-1.5 text-xs font-medium text-navy-700 hover:bg-surface-100"
                >
                  Vider
                </button>
              </>
            )}
            <button
              onClick={() => { void loadList(true); for (const id of selected) void loadBlock(id, true); }}
              className="rounded-lg border border-surface-300 px-3 py-1.5 text-xs font-medium text-navy-700 hover:bg-surface-100"
            >
              ↻ Actualiser DHIS2
            </button>
          </div>
        </div>
        {!list && !listError && (
          <p className="flex items-center gap-2 text-xs text-surface-500">
            <span className="h-3.5 w-3.5 animate-spin rounded-full border-2 border-navy-500 border-t-transparent" />
            Connexion au DHIS2 de campagne… (première extraction : jusqu&apos;à une minute)
          </p>
        )}
        {listError && (
          <p className="rounded-lg bg-danger-50 px-3 py-2 text-xs text-danger-700">
            DHIS2 indisponible ({listError}).{" "}
            <button onClick={() => void loadList(true)} className="font-semibold underline">Réessayer</button>
          </p>
        )}
        {list && (
          <div className="grid grid-cols-1 gap-2 sm:grid-cols-2 lg:grid-cols-3">
            {list.map((p) => {
              const on = selected.includes(p.id);
              const b = blocks[p.id];
              const busy = Boolean(loading[p.id]);
              const err = errors[p.id];
              return (
                <button
                  key={p.id}
                  onClick={() => toggleProvince(p.id)}
                  className={`flex items-center justify-between gap-2 rounded-xl border px-3 py-2.5 text-left transition ${
                    on ? "border-navy-500 bg-navy-50 ring-1 ring-navy-300" : "border-surface-200 bg-white hover:bg-surface-50"
                  }`}
                >
                  <span>
                    <span className="block text-sm font-semibold text-navy-700">
                      {on ? "☑" : "☐"} {prettyProvince(p.name)}
                    </span>
                    <span className="block text-[11px] text-surface-400">
                      {on && b?.source === "masque"
                        ? `Résultats du dernier masque importé · J1 ${fmtDateFR(b.j1)} · ${fmtInt(b.records.length)} AS`
                        : <>Cible RR {fmtInt(p.cibleRR)} · {fmtInt(p.rrVacc)} vaccinés RR{on && b ? ` · J1 ${fmtDateFR(b.j1)} · ${fmtInt(b.records.length)} AS` : ""}</>}
                      {err ? ` · ⚠️ ${err}` : ""}
                    </span>
                  </span>
                  <span className="flex shrink-0 items-center gap-1.5">
                    {busy && <span className="h-3.5 w-3.5 animate-spin rounded-full border-2 border-navy-500 border-t-transparent" />}
                    <span
                      className={`rounded-full px-2 py-0.5 text-[10px] font-bold ${
                        p.polio ? "bg-good-50 text-good-600" : "bg-warn-50 text-warn-600"
                      }`}
                    >
                      {p.polio ? "RR-POLIO" : "RR seule"}
                    </span>
                  </span>
                </button>
              );
            })}
          </div>
        )}
        {pendingCount > 0 && (
          <p className="mt-3 flex items-center gap-2 text-xs text-navy-700">
            <span className="h-3.5 w-3.5 animate-spin rounded-full border-2 border-navy-500 border-t-transparent" />
            Extraction DHIS2 en cours pour {pendingCount} province{pendingCount > 1 ? "s" : ""}… (jusqu&apos;à une minute par province la première fois)
          </p>
        )}
      </section>

      {data && report && (
        <>
          {/* 2. Filtres en cascade */}
          <section className="rounded-2xl border border-surface-200 bg-white p-4 shadow-card">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-sm font-bold text-navy-700">🔎 Périmètre du rapport — {report.scopeLabel} · agrégation par {drill.label}</h2>
              <button onClick={() => setSub({ antenne: null, zs: null, as: null })} className="text-xs font-medium text-accent-600 hover:underline">Réinitialiser</button>
            </div>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
              <Select label="Antenne" value={sub.antenne} options={opts!.antennes} onChange={(v) => setSub({ antenne: v, zs: null, as: null })} />
              <Select label="Zone de Santé" value={sub.zs} options={opts!.zones} onChange={(v) => setSub((s) => ({ ...s, zs: v, as: null }))} />
              <Select label="Aire de Santé" value={sub.as} options={opts!.aires} onChange={(v) => setSub((s) => ({ ...s, as: v }))} />
            </div>
          </section>

          {/* 3. Aperçu (KPI) */}
          <div className="grid grid-cols-2 gap-3 md:grid-cols-4 lg:grid-cols-7">
            <Mini label="Provinces" value={fmtInt(nbProv)} />
            <Mini label="Aires de Santé" value={fmtInt(filtered.length)} />
            <Mini label="Complétude" value={fmtPct(t.completude)} tone={t.completude} />
            <Mini label="CV RR" value={fmtPct(t.rr.cv)} tone={t.rr.cv} />
            <Mini label="CV nVPO2" value={fmtPct(t.nvpo2.cv)} tone={t.nvpo2.cv} />
            <Mini label="CV VPOb" value={fmtPct(t.vpob.cv)} tone={t.vpob.cv} />
            <Mini label="MAPI graves" value={fmtInt(t.mapiGraves)} tone={t.mapiGraves > 0 ? 0 : null} />
          </div>

          {/* 4. Supervision ODK */}
          <section className="rounded-2xl border border-surface-200 bg-white p-4 shadow-card">
            <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
              <h2 className="text-sm font-bold text-navy-700">🛰️ Supervision des équipes (ODK — formulaire Bloc 3)</h2>
              {supProvince && (
                <button onClick={() => void loadSupervision(true)} className="rounded-lg border border-surface-300 px-3 py-1.5 text-xs font-medium text-navy-700 hover:bg-surface-100">↻ Actualiser</button>
              )}
            </div>
            {supState === "off" && (
              <p className="text-xs text-surface-500">
                La supervision ODK est chargée pour <b>une province à la fois</b> (serveur ODK lent) : sélectionnez une seule
                province pour l&apos;inclure. Sinon, le rapport contiendra une diapositive « supervision indisponible ».
              </p>
            )}
            {supState === "loading" && (
              <p className="flex items-center gap-2 text-xs text-navy-700">
                <span className="h-3.5 w-3.5 animate-spin rounded-full border-2 border-navy-500 border-t-transparent" />
                Connexion ODK ({supProvince}, supervisions dès le {fmtDateFR(supDateMin ?? "")})… le serveur peut mettre plusieurs minutes.
              </p>
            )}
            {supState === "error" && (
              <p className="rounded-lg bg-danger-50 px-3 py-2 text-xs text-danger-700">
                Données ODK indisponibles ({supError}). Le rapport contiendra une diapositive « données de supervision indisponibles ».
              </p>
            )}
            {supState === "ok" && supPreview && (
              <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
                <Mini label="Supervisions (périmètre)" value={fmtInt(supPreview.total)} />
                <Mini label="AS visitées" value={`${fmtInt(supPreview.visited)} / ${fmtInt(supPreview.totalAS)}`} />
                <Mini label="% AS visitées" value={fmtPct(supPreview.pct)} tone={supPreview.pct == null ? null : supPreview.pct >= 80 ? 96 : supPreview.pct >= 50 ? 85 : 20} />
                <Mini label="Extraction ODK" value={sup?.fetchedAt ? new Date(sup.fetchedAt).toLocaleString("fr-FR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }) : "—"} />
              </div>
            )}
          </section>

          {/* 5. Contenu + téléchargement */}
          <section className="rounded-2xl border border-surface-200 bg-white p-5 shadow-card">
            <h2 className="mb-1 text-sm font-bold text-navy-700">📊 Contenu du rapport PowerPoint (design du modèle Bloc 3)</h2>
            <p className="mb-4 text-xs text-surface-500">
              {report.coverEntity} · lancement {report.dateLancement} · source : DHIS2 rdccampagne (extraction {data.meta.importedAt ? new Date(data.meta.importedAt).toLocaleString("fr-FR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }) : "—"})
            </p>
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
                disabled={busy || pendingCount > 0 || filtered.length === 0}
                className="flex w-full items-center justify-center gap-2 rounded-xl bg-navy-700 py-3.5 text-sm font-semibold text-white shadow-card transition hover:bg-navy-800 disabled:opacity-50"
              >
                {busy ? (
                  <>
                    <span className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" />
                    Génération du rapport en cours…
                  </>
                ) : (
                  <>📊 Télécharger le rapport du niveau sélectionné (.pptx)</>
                )}
              </button>
              {done && <p className="mt-2 text-center text-xs font-medium text-good-600">✅ {done}</p>}
              {error && <p className="mt-2 text-center text-xs font-medium text-danger-600">⚠️ {error}</p>}
            </div>
          </section>

          {/* 6. Points d'action des précédents PC */}
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
            {actions.length === 0 && <p className="text-xs text-surface-400">Aucune ligne — la diapositive indiquera « à compléter ».</p>}
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

          {/* 7. Problèmes éditables */}
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
                  <Field placeholder="Problème identifié" value={p.probleme} onChange={(v) => { setProblemesTouched(true); setProblemes((arr) => arr.map((r, j) => (j === i ? { ...r, probleme: v } : r))); }} />
                  <Field placeholder="Causes" value={p.causes} onChange={(v) => { setProblemesTouched(true); setProblemes((arr) => arr.map((r, j) => (j === i ? { ...r, causes: v } : r))); }} />
                  <Field placeholder={`${drill.label}(s) concernée(s)`} value={p.zs} onChange={(v) => { setProblemesTouched(true); setProblemes((arr) => arr.map((r, j) => (j === i ? { ...r, zs: v } : r))); }} />
                  <Field placeholder="Solutions proposées" value={p.solutions} onChange={(v) => { setProblemesTouched(true); setProblemes((arr) => arr.map((r, j) => (j === i ? { ...r, solutions: v } : r))); }} />
                  <button onClick={() => { setProblemesTouched(true); setProblemes((arr) => arr.filter((_, j) => j !== i)); }} className="rounded-lg px-2 text-danger-500 hover:bg-danger-50" title="Supprimer">✕</button>
                </div>
              ))}
            </div>
          </section>
        </>
      )}

      {!data && selected.length === 0 && list && (
        <div className="rounded-2xl border border-surface-200 bg-white p-10 text-center shadow-card">
          <div className="mb-3 text-5xl">🗺️</div>
          <h2 className="mb-2 text-lg font-bold text-navy-700">Sélectionnez une ou plusieurs provinces</h2>
          <p className="text-sm text-surface-500">
            Toutes les provinces / antennes / ZS / AS ayant des données de campagne dans DHIS2 sont disponibles.
            Plusieurs provinces → situation par province puis détail par antenne ; une province → détail par antenne et ZS.
          </p>
        </div>
      )}
    </div>
  );
}

function Select({ label, value, options, onChange }: { label: string; value: string | null; options: string[]; onChange: (v: string | null) => void }) {
  return (
    <div className="block">
      <div className="mb-1 text-[11px] font-semibold uppercase tracking-wide text-surface-400">{label}</div>
      <select
        value={value ?? ""}
        onChange={(e) => onChange(e.target.value || null)}
        className="w-full rounded-lg border border-surface-300 bg-white px-3 py-2.5 text-sm font-medium text-navy-700 focus:border-navy-500 focus:outline-none focus:ring-2 focus:ring-navy-200"
      >
        <option value="">Toutes</option>
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

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
  tauxPerte,
  totals,
  NVPO2_DOSES_PAR_FLACON,
  VPOB_DOSES_PAR_FLACON,
  type UnitAgg,
  type Totals,
} from "@/lib/analytics";
import { fmtInt, fmtPct } from "@/lib/format";
import { fetchNational } from "@/lib/national";
import { ANTIGENES, type MasqueData } from "@/lib/parse-masque";
import type { CompletudeRow, ProblemeRow, ReportData } from "@/lib/export-report-pptx";

/** Concatène une liste de ZS en limitant la longueur affichée. */
function joinUnits(names: string[], max = 8): string {
  if (names.length === 0) return "—";
  if (names.length <= max) return names.join(", ");
  return `${names.slice(0, max).join(", ")} … (+${names.length - max})`;
}

/**
 * Déduit les problèmes rencontrés directement des analyses, par unité (ZS le plus
 * souvent). Chaque problème liste les unités réellement concernées et propose une
 * piste de solution cohérente avec l'anomalie détectée.
 */
function computeProblemes(byUnit: UnitAgg[], t: Totals): ProblemeRow[] {
  const out: ProblemeRow[] = [];
  const NV_SEUIL = 11; // seuil de perte nVPO2 (%)
  const VP_SEUIL = 10; // seuil de perte VPOb (%)

  // 1. Complétude des rapports insuffisante (< 95 %).
  const complBas = byUnit
    .filter((u) => u.vaccAttendus > 0 && (u.vaccRecus / u.vaccAttendus) * 100 < 95)
    .sort((a, b) => a.vaccRecus / a.vaccAttendus - b.vaccRecus / b.vaccAttendus)
    .map((u) => u.unit);
  if (complBas.length > 0) {
    out.push({
      probleme: "Complétude des rapports insuffisante (< 95 %)",
      causes: "Faible remontée des données : zones sans réseau Internet, retard de transmission des collecteurs",
      zs: joinUnits(complBas),
      solutions: "Tracer les collecteurs et BCZ retardataires, sécuriser la transmission (relais radio/moto), valider les rapports manquants au J+1",
    });
  }

  // 2. Couverture nVPO2 sous l'objectif (< 95 %).
  const nvBas = byUnit
    .filter((u) => u.nvpo2Cible > 0 && (u.nvpo2Vacc / u.nvpo2Cible) * 100 < 95)
    .sort((a, b) => a.nvpo2Vacc / a.nvpo2Cible - b.nvpo2Vacc / b.nvpo2Cible)
    .map((u) => u.unit);
  if (nvBas.length > 0) {
    out.push({
      probleme: "Couverture vaccinale nVPO2 sous l'objectif (< 95 %)",
      causes: "Enfants absents/non atteints, refus, sites difficiles d'accès, démarrage tardif des équipes",
      zs: joinUnits(nvBas),
      solutions: "Organiser des passages de ratissage ciblés, renforcer la mobilisation sociale et le porte-à-porte dans les aires faibles",
    });
  }

  // 3. Couverture VPOb sous l'objectif (< 95 %).
  const vpBas = byUnit
    .filter((u) => u.vpobCible > 0 && (u.vpobVacc / u.vpobCible) * 100 < 95)
    .sort((a, b) => a.vpobVacc / a.vpobCible - b.vpobVacc / b.vpobCible)
    .map((u) => u.unit);
  if (vpBas.length > 0) {
    out.push({
      probleme: "Couverture vaccinale VPOb sous l'objectif (< 95 %)",
      causes: "Co-administration incomplète, ruptures ponctuelles de VPOb, enfants déjà partis",
      zs: joinUnits(vpBas),
      solutions: "Assurer la co-administration systématique nVPO2 + VPOb, réapprovisionner les sites et planifier le rattrapage",
    });
  }

  // 4. Taux de perte nVPO2 hors seuil (> 11 % ou négatif).
  const nvPerte = byUnit
    .filter((u) => {
      const x = tauxPerte(u.nvpo2Vacc, u.nvpo2FlaconsUtil, NVPO2_DOSES_PAR_FLACON);
      return x != null && (x > NV_SEUIL || x < 0);
    })
    .map((u) => u.unit);
  if (nvPerte.length > 0) {
    out.push({
      probleme: `Taux de perte nVPO2 hors seuil (> ${NV_SEUIL} % ou négatif)`,
      causes: "Saisie irrégulière des flacons utilisés/rendus, rupture de chaîne du froid, flacons entamés non terminés",
      zs: joinUnits(nvPerte),
      solutions: "Reprendre la saisie flacons reçus/utilisés/rendus, renforcer la gestion de la chaîne du froid et la politique des flacons entamés",
    });
  }

  // 5. Taux de perte VPOb hors seuil (> 10 % ou négatif).
  const vpPerte = byUnit
    .filter((u) => {
      const x = tauxPerte(u.vpobVacc, u.vpobFlaconsUtil, VPOB_DOSES_PAR_FLACON);
      return x != null && (x > VP_SEUIL || x < 0);
    })
    .map((u) => u.unit);
  if (vpPerte.length > 0) {
    out.push({
      probleme: `Taux de perte VPOb hors seuil (> ${VP_SEUIL} % ou négatif)`,
      causes: "Saisie irrégulière des flacons, flacons multidoses partiellement utilisés, conservation inadéquate",
      zs: joinUnits(vpPerte),
      solutions: "Fiabiliser la saisie des flacons VPOb, respecter la durée d'utilisation après ouverture et la chaîne du froid",
    });
  }

  // 6. MAPI graves notifiées.
  if (t.mapiGraves > 0) {
    out.push({
      probleme: `MAPI graves notifiées (${t.mapiGraves})`,
      causes: "Manifestations indésirables post-immunisation requérant investigation",
      zs: "Voir slide Surveillance des MAPI",
      solutions: "Investiguer chaque cas sous 48 h, notifier au niveau supérieur, assurer la prise en charge médicale et la communication de crise",
    });
  }

  const vaccTotal = t.nvpo2Vacc + t.vpobVacc;

  // 7. Non-rapportage des enfants récupérés en PEV de routine.
  const recupTotal = t.antigenesEV.reduce((a, b) => a + b, 0);
  if (vaccTotal > 0 && t.recup === 0 && recupTotal === 0) {
    out.push({
      probleme: "Non-rapportage des enfants récupérés en PEV de routine",
      causes: "Volet récupération PEV non rempli dans le masque malgré une campagne réalisée (co-administration non tracée)",
      zs: "Toutes les ZS du périmètre",
      solutions: "Sensibiliser les équipes à enregistrer systématiquement les antigènes de routine administrés et compléter le volet récupération PEV",
    });
  }

  // 8. Non-notification de la surveillance des MEV (MPV).
  const survTotal = t.survPFA + t.survRougeole + t.survFJ + t.survTNN;
  if (vaccTotal > 0 && survTotal === 0) {
    out.push({
      probleme: "Non-notification de la surveillance des MEV (MPV)",
      causes: "Aucun cas PFA / Rougeole / Fièvre Jaune / TNN notifié — recherche active probablement non documentée",
      zs: "Toutes les ZS du périmètre",
      solutions: "Renforcer la recherche active des cas de MEV pendant la campagne et documenter même les notifications « zéro cas »",
    });
  }

  // 9. Non-notification des MAPI.
  if (vaccTotal > 0 && t.mapiMineures === 0 && t.mapiGraves === 0) {
    out.push({
      probleme: "Non-notification des MAPI",
      causes: "Aucune MAPI (mineure ou grave) notifiée malgré le volume de doses administrées — sous-notification probable",
      zs: "Toutes les ZS du périmètre",
      solutions: "Rappeler aux équipes la notification systématique des MAPI, y compris les manifestations mineures, et documenter le « zéro cas »",
    });
  }

  // 10. Incohérences des données (taux de perte négatif / complétude > 100 %).
  const incoh = byUnit
    .filter((u) => {
      const nv = tauxPerte(u.nvpo2Vacc, u.nvpo2FlaconsUtil, NVPO2_DOSES_PAR_FLACON);
      const vp = tauxPerte(u.vpobVacc, u.vpobFlaconsUtil, VPOB_DOSES_PAR_FLACON);
      const perteNeg = (nv != null && nv < 0) || (vp != null && vp < 0);
      const complAberr = u.vaccAttendus > 0 && u.vaccRecus > u.vaccAttendus;
      return perteNeg || complAberr;
    })
    .map((u) => u.unit);
  if (incoh.length > 0) {
    out.push({
      probleme: "Incohérences des données saisies",
      causes: "Enfants vaccinés supérieurs aux doses disponibles (taux de perte négatif) ou rapports reçus > attendus : erreurs de saisie flacons / dénombrement",
      zs: joinUnits(incoh),
      solutions: "Vérifier et corriger la saisie des flacons utilisés et des cibles dans le masque, recroiser avec les fiches de pointage",
    });
  }

  return out;
}

export default function RapportPage() {
  const { data: localData, filters, setFilter, resetFilters } = useApp();
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(false);
  const [problemes, setProblemes] = useState<ProblemeRow[]>([]);
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

  // Problèmes déduits automatiquement des analyses (recalculés à chaque changement
  // de périmètre). L'utilisateur peut ensuite les ajuster manuellement.
  const autoProblemes = useMemo(() => computeProblemes(byUnit, t), [byUnit, t]);
  useEffect(() => {
    setProblemes(autoProblemes);
  }, [autoProblemes]);

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

    // Entité de plus bas niveau filtrée (page de garde) — dynamique.
    const coverEntity =
      filters.as ? `Aire de Santé : ${filters.as}` :
      filters.zs ? `Zone de Santé : ${filters.zs}` :
      filters.antenne ? `Antenne : ${filters.antenne}` :
      filters.province ? `Province : ${filters.province}` :
      `Province : ${data!.meta.province}`;

    // Période dynamique : du premier au dernier jour de campagne saisi.
    const shortJ = (l: string) => l.replace(/jour\s*/i, "J").trim();
    const periodeJours =
      jourLabels.length > 0
        ? `Du ${shortJ(jourLabels[0])} au ${shortJ(jourLabels[jourLabels.length - 1])}`
        : "";
    const periode =
      periodeJours && data!.meta.periode
        ? `${periodeJours} — ${data!.meta.periode}`
        : periodeJours || data!.meta.periode;

    // Construction du tableau complétude par jour pour chaque unité.
    const completudeByUnit: CompletudeRow[] = byUnit.map((a) => {
      // Complétude journalière = rapports reçus ce jour ÷ rapports attendus
      // (total campagne). Le cumul des jours redonne la complétude globale et
      // ne peut donc pas dépasser 100 % quand reçus ≤ attendus.
      const daily = Array.from({ length: nbJours }, (_, i) => {
        const recus = a.rapportsRecusDaily[i] ?? 0;
        return { recus, couv: a.vaccAttendus > 0 ? (recus / a.vaccAttendus) * 100 : null };
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
    const recupAntigenByUnit = byUnit.map((a) => ({ unit: a.unit, ev: a.antigenesEV }));
    const survByUnit = byUnit.map((a) => ({
      unit: a.unit, pfa: a.survPFA, rougeole: a.survRougeole, fj: a.survFJ, tnn: a.survTNN,
    }));

    return {
      province: filters.province ?? data!.meta.province,
      coverEntity,
      periode,
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
      antigenLabels: ANTIGENES.map((a) => a.label),
      recupAntigenByUnit,
      recupAntigenTotals: ANTIGENES.map((_, j) => t.antigenesEV[j] ?? 0),
      survByUnit,
      survTotals: { pfa: t.survPFA, rougeole: t.survRougeole, fj: t.survFJ, tnn: t.survTNN },
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
        // Carte : ZS du périmètre colorées selon leur complétude obtenue.
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
        for (const [k, v] of byZS) completudeByZS.set(k, v.att > 0 ? (v.rec / v.att) * 100 : 0);
        report.scopeMapPng = completudeByZS.size > 0 ? await renderZSMap(completudeByZS) : null;
      } catch {
        report.scopeMapPng = null;
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
            "Page de garde (image polio + logo PEV)",
            "Plan de présentation",
            "Points saillants",
            "Complétude des rapports par ZS (jauge + tableau J1..Jn)",
            "Spatialisation de la complétude (carte RDC des Zones de Santé)",
            "Couvertures vaccinales nVPO2 par jour (cible / vaccinés / couverture)",
            "Couvertures vaccinales VPOb par jour",
            "Récupération PEV de routine — enfants vaccinés par antigène (EV)",
            "Gestion du vaccin nVPO2 (flacons reçus / utilisés / rendus / perdus)",
            "Gestion du vaccin VPOb",
            "Surveillance des MPV par ZS (PFA, Rougeole, FJ, TNN)",
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

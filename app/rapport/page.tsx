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

/** Concatène une liste d'unités (AS/ZS) en limitant la longueur affichée. */
function joinUnits(names: string[], max = 8): string {
  if (names.length === 0) return "—";
  if (names.length <= max) return names.join(", ");
  return `${names.slice(0, max).join(", ")} … (+${names.length - max})`;
}

/**
 * Déduit dynamiquement les problèmes à partir des analyses, au niveau de
 * désagrégation courant (`unitLabel` = Aire de Santé, Zone de Santé, etc.).
 * Chaque problème liste les unités réellement concernées et propose une action
 * correctrice cohérente. La saisie des données est assurée par les gestionnaires
 * de données (DATA) — il n'y a pas de collecteurs pour cette campagne.
 */
function computeProblemes(byUnit: UnitAgg[], t: Totals, unitLabel: string): ProblemeRow[] {
  const out: ProblemeRow[] = [];
  const NV_SEUIL = 11; // seuil acceptable de perte nVPO2 (%)
  const VP_SEUIL = 10; // seuil acceptable de perte VPOb (%)
  const perimetre = `Ensemble du périmètre (${unitLabel.toLowerCase()})`;

  // 1. Complétude des rapports de vaccination insuffisante (< 95 %).
  const complBas = byUnit
    .filter((u) => u.vaccAttendus > 0 && (u.vaccRecus / u.vaccAttendus) * 100 < 95)
    .sort((a, b) => a.vaccRecus / a.vaccAttendus - b.vaccRecus / b.vaccAttendus)
    .map((u) => u.unit);
  if (complBas.length > 0) {
    out.push({
      probleme: "Complétude des rapports de vaccination insuffisante (< 95 %)",
      causes: "Transmission tardive ou incomplète des rapports journaliers ; couverture réseau limitée pour la saisie ; indisponibilité ponctuelle des gestionnaires de données (DATA)",
      zs: joinUnits(complBas),
      solutions: "Relancer les DATA et les BCZ concernés, sécuriser la chaîne de transmission des rapports et consolider les rapports manquants au plus tard à J+1",
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
      causes: "Enfants absents ou non atteints, refus parentaux, sites d'accès difficile, déploiement tardif des équipes",
      zs: joinUnits(nvBas),
      solutions: "Planifier des passages de ratissage ciblés et renforcer la mobilisation sociale et le porte-à-porte dans les aires sous-performantes",
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
      causes: "Co-administration incomplète, ruptures ponctuelles d'approvisionnement en VPOb, enfants déjà partis du site",
      zs: joinUnits(vpBas),
      solutions: "Garantir la co-administration systématique nVPO2 + VPOb, réapprovisionner les sites en rupture et planifier le rattrapage",
    });
  }

  // 4. Incohérence de couverture nVPO2 vs VPOb (co-administration).
  // Les deux antigènes sont co-administrés à la même cible le même jour : leurs
  // couvertures doivent être identiques ; un écart traduit une erreur de saisie.
  const coAdminEcart = byUnit
    .filter((u) => {
      if (u.nvpo2Cible <= 0 || u.vpobCible <= 0) return false;
      const cvN = (u.nvpo2Vacc / u.nvpo2Cible) * 100;
      const cvV = (u.vpobVacc / u.vpobCible) * 100;
      return Math.abs(cvN - cvV) >= 0.1;
    })
    .sort((a, b) => {
      const da = Math.abs(a.nvpo2Vacc / a.nvpo2Cible - a.vpobVacc / a.vpobCible);
      const db = Math.abs(b.nvpo2Vacc / b.nvpo2Cible - b.vpobVacc / b.vpobCible);
      return db - da;
    })
    .map((u) => u.unit);
  if (coAdminEcart.length > 0) {
    out.push({
      probleme: "Incohérence de couverture entre nVPO2 et VPOb (co-administration)",
      causes: "nVPO2 et VPOb étant co-administrés à la même cible le même jour, leurs couvertures devraient être strictement égales ; l'écart observé traduit une erreur de saisie des effectifs vaccinés",
      zs: joinUnits(coAdminEcart),
      solutions: "Faire recroiser par les DATA les effectifs vaccinés nVPO2 et VPOb et corriger la saisie afin d'aligner les deux couvertures",
    });
  }

  // 5. Taux de perte nVPO2 hors seuil (> 11 % ou négatif).
  const nvPerte = byUnit
    .filter((u) => {
      const x = tauxPerte(u.nvpo2Vacc, u.nvpo2FlaconsUtil, NVPO2_DOSES_PAR_FLACON);
      return x != null && (x > NV_SEUIL || x < 0);
    })
    .map((u) => u.unit);
  if (nvPerte.length > 0) {
    out.push({
      probleme: `Taux de perte nVPO2 hors seuil (> ${NV_SEUIL} % ou négatif)`,
      causes: "Saisie incohérente des flacons (reçus / utilisés / restitués), maîtrise insuffisante de la chaîne du froid, gestion non conforme des flacons entamés",
      zs: joinUnits(nvPerte),
      solutions: "Faire fiabiliser par les DATA la saisie des mouvements de flacons, renforcer la gestion de la chaîne du froid et appliquer la politique des flacons entamés",
    });
  }

  // 6. Taux de perte VPOb hors seuil (> 10 % ou négatif).
  const vpPerte = byUnit
    .filter((u) => {
      const x = tauxPerte(u.vpobVacc, u.vpobFlaconsUtil, VPOB_DOSES_PAR_FLACON);
      return x != null && (x > VP_SEUIL || x < 0);
    })
    .map((u) => u.unit);
  if (vpPerte.length > 0) {
    out.push({
      probleme: `Taux de perte VPOb hors seuil (> ${VP_SEUIL} % ou négatif)`,
      causes: "Saisie incohérente des flacons, flacons multidoses partiellement utilisés, conditions de conservation inadéquates",
      zs: joinUnits(vpPerte),
      solutions: "Faire fiabiliser par les DATA la saisie des flacons VPOb, respecter le délai d'utilisation après ouverture et la chaîne du froid",
    });
  }

  // 7. MAPI graves notifiées.
  if (t.mapiGraves > 0) {
    out.push({
      probleme: `MAPI graves notifiées (${t.mapiGraves})`,
      causes: "Manifestations adverses post-immunisation graves nécessitant une investigation immédiate",
      zs: "Voir diapositive « Surveillance des MAPI »",
      solutions: "Investiguer chaque cas sous 48 h, notifier au niveau hiérarchique supérieur, assurer la prise en charge médicale et la communication de crise",
    });
  }

  const vaccTotal = t.nvpo2Vacc + t.vpobVacc;

  // 8. Non-rapportage des enfants récupérés en PEV de routine.
  const recupTotal = t.antigenesEV.reduce((a, b) => a + b, 0);
  if (vaccTotal > 0 && t.recup === 0 && recupTotal === 0) {
    out.push({
      probleme: "Non-rapportage de la récupération PEV de routine",
      causes: "Volet « récupération PEV de routine » non renseigné dans le masque alors que la vaccination a été réalisée (co-administration non documentée)",
      zs: perimetre,
      solutions: "Sensibiliser les équipes et les DATA à l'enregistrement systématique des antigènes de routine co-administrés pendant la campagne",
    });
  }

  // Récupération PEV de routine — analyse par entité (enfants ZD = zéro dose,
  // SV = sous-vaccinés), à partir des enfants identifiés / récupérés par antigène.
  const identTotalU = (u: UnitAgg) => (u.antigenesIdentifies ?? []).reduce((a, b) => a + b, 0);
  const recupTotalU = (u: UnitAgg) => (u.antigenesEV ?? []).reduce((a, b) => a + b, 0);

  // 8 bis. Non identification des enfants ZD ou SV : aucune donnée d'identification
  // par antigène dans la feuille « Donnees de base » du masque pour ces entités.
  // On alerte dès qu'une entité n'a aucun enfant identifié — y compris lorsque tout
  // le périmètre est concerné (ex. Kinshasa sans bloc identification saisi).
  const sansIdent = byUnit.filter((u) => identTotalU(u) === 0).map((u) => u.unit);
  if (sansIdent.length > 0) {
    out.push({
      probleme: "Non identification des enfants ZD ou SV",
      causes: "Non encodage des listes des enfants à conflit vaccinal dans le masque de saisie, dénombrement non réalisé, non identification par les RECO",
      zs: joinUnits(sansIdent),
      solutions: "Encoder les données des enfants ZD et SV identifiés dans le masque de saisie de la campagne",
    });
  }

  // 8 ter. Non récupération des enfants ZD ou SV : enfants identifiés mais aucune
  // récupération enregistrée (nombre récupéré = 0 pour tous les antigènes).
  const sansRecup = byUnit
    .filter((u) => identTotalU(u) > 0 && recupTotalU(u) === 0)
    .map((u) => u.unit);
  if (sansRecup.length > 0) {
    out.push({
      probleme: "Non récupération des enfants ZD ou SV",
      causes: "Non organisation de site de récupération en routine",
      zs: joinUnits(sansRecup),
      solutions: "Organiser le site de vaccination de routine",
    });
  }

  // 8 quater. Faible récupération des enfants ZD ou SV : taux de récupération
  // (récupérés ÷ identifiés, tous antigènes) strictement compris entre 0 et 80 %.
  const faibleRecup = byUnit
    .filter((u) => {
      const id = identTotalU(u);
      if (id <= 0) return false;
      const taux = (recupTotalU(u) / id) * 100;
      return taux > 0 && taux < 80;
    })
    .sort((a, b) => recupTotalU(a) / identTotalU(a) - recupTotalU(b) / identTotalU(b))
    .map((u) => u.unit);
  if (faibleRecup.length > 0) {
    out.push({
      probleme: "Faible récupération des enfants ZD ou SV (< 80 %)",
      causes: "Faible sensibilisation, nombre insuffisant de sites de vaccination de routine",
      zs: joinUnits(faibleRecup),
      solutions: "Intensifier la récupération des enfants ZD et SV en routine",
    });
  }

  // 9. Non-notification de la surveillance des MEV (MPV).
  const survTotal = t.survPFA + t.survRougeole + t.survFJ + t.survTNN;
  if (vaccTotal > 0 && survTotal === 0) {
    out.push({
      probleme: "Non-notification de la surveillance des MEV (MPV)",
      causes: "Aucun cas PFA / Rougeole / Fièvre Jaune / TNN notifié : recherche active des cas non documentée pendant la campagne",
      zs: perimetre,
      solutions: "Renforcer la recherche active des cas de MEV et documenter la surveillance, y compris les notifications « zéro cas »",
    });
  }

  // 10. Non-notification des MAPI.
  if (vaccTotal > 0 && t.mapiMineures === 0 && t.mapiGraves === 0) {
    out.push({
      probleme: "Non-notification des MAPI",
      causes: "Aucune MAPI (mineure ou grave) notifiée malgré le volume de doses administrées : sous-notification probable",
      zs: perimetre,
      solutions: "Rappeler aux équipes la notification systématique des MAPI, y compris mineures, et documenter le « zéro cas »",
    });
  }

  // 11. Incohérences des données saisies (taux négatif / reçus > attendus).
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
      causes: "Effectifs vaccinés supérieurs aux doses disponibles (taux de perte négatif) ou rapports reçus supérieurs aux attendus : erreurs de saisie des flacons ou des cibles",
      zs: joinUnits(incoh),
      solutions: "Faire vérifier et corriger par les DATA la saisie des flacons et des cibles dans le masque, par recoupement avec les fiches de pointage",
    });
  }

  return out;
}

export default function RapportPage() {
  const { data: localData, filters, setFilter, toggleProvince, resetFilter, resetFilters } = useApp();
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
  const provinceCount = opts?.provinces.length ?? 0;
  const drill = useMemo(() => resolveDrillLevel(filters, provinceCount), [filters, provinceCount]);
  const t = useMemo(() => totals(filtered), [filtered]);
  const byUnit = useMemo(() => aggregateByUnit(filtered, drill.level), [filtered, drill.level]);

  // Problèmes déduits automatiquement des analyses (recalculés à chaque changement
  // de périmètre). L'utilisateur peut ensuite les ajuster manuellement.
  const autoProblemes = useMemo(() => computeProblemes(byUnit, t, drill.label), [byUnit, t, drill.label]);
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
      filters.provinces.length === 1 ? `Province : ${filters.provinces[0]}` :
      filters.provinces.length > 1 ? `Provinces : ${filters.provinces.join(", ")}` :
      `Niveau national — toutes les provinces`;

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
      // Complétude journalière = rapports reçus ce jour ÷ rapports ATTENDUS ce jour
      // (chaque feuille JourN porte son propre attendu), et non l'attendu total
      // de la campagne — sinon le % est faussé (ex. 8/24 au lieu de 8/8).
      const daily = Array.from({ length: nbJours }, (_, i) => {
        const recus = a.rapportsRecusDaily[i] ?? 0;
        const attendus = a.rapportsAttendusDaily[i] ?? 0;
        return { recus, attendus, couv: attendus > 0 ? (recus / attendus) * 100 : null };
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
    const recupAntigenByUnit = byUnit.map((a) => ({ unit: a.unit, ev: a.antigenesEV, ident: a.antigenesIdentifies }));
    const survByUnit = byUnit.map((a) => ({
      unit: a.unit, pfa: a.survPFA, rougeole: a.survRougeole, fj: a.survFJ, tnn: a.survTNN,
    }));

    return {
      province: filters.provinces.length === 1 ? filters.provinces[0] : data!.meta.province,
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
      recupAntigenIdentTotals: ANTIGENES.map((_, j) => t.antigenesIdentifies[j] ?? 0),
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
          {Array.from(new Set(nationalData.records.map((r) => r.province))).length} province(s). Sélectionnez une ou
          plusieurs provinces pour cibler, ou laissez vide pour télécharger la situation de toutes les provinces à la
          fois (agrégation par province par défaut).
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
          <MultiSelect
            label="Province"
            values={filters.provinces}
            options={opts!.provinces}
            onToggle={toggleProvince}
            onReset={() => resetFilter("provinces")}
          />
          <Select label="Antenne" value={filters.antenne} options={opts!.antennes} onChange={(v) => setFilter("antenne", v)} onReset={() => resetFilter("antenne")} />
          <Select label="Zone de Santé" value={filters.zs} options={opts!.zones} onChange={(v) => setFilter("zs", v)} onReset={() => resetFilter("zs")} />
          <Select label="Aire de Santé" value={filters.as} options={opts!.aires} onChange={(v) => setFilter("as", v)} onReset={() => resetFilter("as")} />
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
            "Récupération PEV de routine — identifiés / récupérés / % par antigène (réparti sur plusieurs diapos)",
            "Gestion du vaccin nVPO2 (flacons reçus / utilisés / rendus / perdus)",
            "Gestion du vaccin VPOb",
            "Surveillance des MPV par ZS (PFA, Rougeole, FJ, TNN)",
            "Synthèse Surveillance MAPI et récupération en routine",
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
              <Field placeholder={`${drill.label}(s) concernée(s)`} value={p.zs} onChange={(v) => editRow(setProblemes, i, "zs", v)} />
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

function FilterHeader({ label, active, onReset }: { label: string; active: boolean; onReset: () => void }) {
  return (
    <div className="mb-1 flex items-center justify-between gap-2">
      <span className="text-[11px] font-semibold uppercase tracking-wide text-surface-400">{label}</span>
      {active && (
        <button
          type="button"
          onClick={onReset}
          title={`Réinitialiser : ${label}`}
          aria-label={`Réinitialiser le filtre ${label}`}
          className="inline-flex h-5 w-5 items-center justify-center rounded-full text-sm leading-none text-surface-400 transition hover:bg-accent-50 hover:text-accent-600"
        >
          ↺
        </button>
      )}
    </div>
  );
}

function Select({
  label,
  value,
  options,
  onChange,
  onReset,
}: {
  label: string;
  value: string | null;
  options: string[];
  onChange: (v: string | null) => void;
  onReset: () => void;
}) {
  return (
    <div className="block">
      <FilterHeader label={label} active={value != null} onReset={onReset} />
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
    </div>
  );
}

function MultiSelect({
  label,
  values,
  options,
  onToggle,
  onReset,
}: {
  label: string;
  values: string[];
  options: string[];
  onToggle: (v: string) => void;
  onReset: () => void;
}) {
  const [open, setOpen] = useState(false);
  const summary =
    values.length === 0 ? "Toutes" : values.length === 1 ? values[0] : `${values.length} provinces sélectionnées`;
  return (
    <div className="relative block">
      <FilterHeader label={label} active={values.length > 0} onReset={onReset} />
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="flex w-full items-center justify-between gap-2 rounded-lg border border-surface-300 bg-white px-3 py-2.5 text-left text-sm font-medium text-navy-700 focus:border-navy-500 focus:outline-none focus:ring-2 focus:ring-navy-200"
      >
        <span className="truncate">{summary}</span>
        <span className="ml-1 text-surface-400">▾</span>
      </button>
      {open && (
        <>
          <div className="fixed inset-0 z-10" onClick={() => setOpen(false)} aria-hidden />
          <div className="absolute z-20 mt-1 max-h-64 w-full overflow-auto rounded-lg border border-surface-200 bg-white p-1 shadow-card">
            {options.length === 0 && (
              <div className="px-3 py-2 text-xs text-surface-400">Aucune province</div>
            )}
            {options.map((o) => {
              const checked = values.includes(o);
              return (
                <button
                  key={o}
                  type="button"
                  onClick={() => onToggle(o)}
                  className="flex w-full items-center gap-2 rounded-md px-2.5 py-2 text-left text-sm text-navy-700 hover:bg-navy-50"
                >
                  <span
                    className={`flex h-4 w-4 shrink-0 items-center justify-center rounded border text-[10px] font-bold ${
                      checked ? "border-navy-700 bg-navy-700 text-white" : "border-surface-300 text-transparent"
                    }`}
                  >
                    ✓
                  </span>
                  <span className="truncate">{o}</span>
                </button>
              );
            })}
          </div>
        </>
      )}
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

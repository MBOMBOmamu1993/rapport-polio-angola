/**
 * Assemblage des données du rapport PowerPoint (contrat `ReportData`) à partir du
 * masque importé, des filtres et des supervisions ODK. Module pur (sans DOM) :
 * utilisable côté navigateur comme dans un script Node de test.
 */

import {
  ANTIGENES,
  DOSES_PAR_FLACON,
  SEUIL_PERTE,
  VACCINE_LABELS,
  type ASRecord,
  type MasqueData,
  type VaccineKey,
} from "./parse-masque";
import { aggregateByUnit, applyFilters, resolveDrillLevel, scopeLabel, totals, type UnitAgg } from "./analytics";
import {
  aggregateSupervisionByZS,
  indicatorConformity,
  type SupervisionPayload,
  type ZSRef,
  type ZSSupervision,
} from "./odk-supervision";
import type { Filters } from "./store";

export interface ActionPCRow {
  activite: string;
  responsable: string;
  statut: string;
  echeance: string;
}
export interface ProblemeRow {
  probleme: string;
  causes: string;
  zs: string;
  solutions: string;
}

export interface SupervisionData {
  available: boolean;
  reason?: string;
  fetchedAt: string;
  dateMin: string;
  formTitle: string;
  total: number;
  byZS: ZSSupervision[];
  conformity: ReturnType<typeof indicatorConformity>;
  points: { lat: number; lon: number }[];
  /** Cartes (PNG data URL) — remplies côté navigateur avant export. */
  mapPng?: string | null;
  pointsMapPng?: string | null;
}

export interface ReportData {
  titre: string;
  province: string;
  provinceLabel: string;
  scopeLabel: string;
  coverEntity: string;
  dateMaj: string;
  dateLancement: string;
  periode: string;
  jourLabels: string[];
  nbJours: number;
  byUnitLabel: string;
  /** Unités du niveau de désagrégation principal (ZS, ou AS si une ZS est filtrée). */
  units: UnitAgg[];
  /** Vue par Antenne (vide si une seule antenne dans le périmètre). */
  antennes: UnitAgg[];
  /** Vue par Zone de Santé (toujours, pour la jointure supervision). */
  zones: UnitAgg[];
  /** Vue par Aire de Santé (tableaux MAPI / MPV « par AS »). */
  aires: UnitAgg[];
  total: UnitAgg;
  supervision: SupervisionData;
  actionsPC: ActionPCRow[];
  problemes: ProblemeRow[];
}

export const REPORT_TITLE = "Campagne intégrée RR-POLIO Kasaï Central";

/** Libellé de province propre (« KASAI-CENTRAL » → « Kasaï Central »). */
export function prettyProvince(p: string): string {
  const s = (p || "").trim();
  if (/kasa[iï]/i.test(s) && /central/i.test(s)) return "Kasaï Central";
  return s
    .toLowerCase()
    .replace(/[-_]+/g, " ")
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

export function fmtDateFR(iso: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
  if (!m) return iso || "";
  return `${m[3]}/${m[2]}/${m[1]}`;
}

function joinUnits(names: string[], max = 8): string {
  if (names.length === 0) return "—";
  if (names.length <= max) return names.join(", ");
  return `${names.slice(0, max).join(", ")} … (+${names.length - max})`;
}

/**
 * Problèmes déduits automatiquement des analyses (au niveau de désagrégation
 * courant) — l'utilisateur peut ensuite les ajuster dans l'interface.
 */
export function computeProblemes(units: UnitAgg[], t: UnitAgg, unitLabel: string, sup: SupervisionData | null): ProblemeRow[] {
  const out: ProblemeRow[] = [];
  const perimetre = `Ensemble du périmètre (${unitLabel.toLowerCase()})`;
  const started = units.some((u) => u.vaccRecus > 0 || u.rr.vacc > 0 || u.nvpo2.vacc > 0);

  const complBas = units
    .filter((u) => u.vaccAttendus > 0 && (u.completude ?? 0) < 95)
    .sort((a, b) => (a.completude ?? 0) - (b.completude ?? 0))
    .map((u) => u.unit);
  if (started && complBas.length > 0) {
    out.push({
      probleme: "Faible complétude des données de vaccination (< 95 %)",
      causes: "Faible remontée des données dans les ZS non couvertes en réseau Internet ; transmission tardive des rapports journaliers ; indisponibilité ponctuelle des gestionnaires de données",
      zs: joinUnits(complBas),
      solutions: "Revue journalière des données par le poste de commandement avec feedback aux ZS ; relance des DATA/BCZ concernés et consolidation des rapports manquants à J+1",
    });
  }

  const vaccKeys: VaccineKey[] = ["rr", "nvpo2", "vpob"];
  for (const k of vaccKeys) {
    const bas = units
      .filter((u) => u[k].cible > 0 && u[k].vacc > 0 && (u[k].cv ?? 0) < 95)
      .sort((a, b) => (a[k].cv ?? 0) - (b[k].cv ?? 0))
      .map((u) => u.unit);
    if (bas.length > 0) {
      out.push({
        probleme: `Couverture vaccinale ${VACCINE_LABELS[k]} sous l'objectif (< 95 %)`,
        causes: k === "rr"
          ? "Enfants absents ou non atteints, refus parentaux, sites d'accès difficile, déploiement tardif des équipes"
          : "Enfants absents ou non atteints, refus, ruptures ponctuelles, co‑administration incomplète",
        zs: joinUnits(bas),
        solutions: "Planifier des passages de ratissage ciblés, renforcer la mobilisation sociale et le porte‑à‑porte dans les aires sous‑performantes",
      });
    }
  }

  const coAdmin = units
    .filter((u) => u.nvpo2.cible > 0 && u.vpob.cible > 0 && (u.nvpo2.vacc > 0 || u.vpob.vacc > 0) && Math.abs((u.nvpo2.cv ?? 0) - (u.vpob.cv ?? 0)) >= 0.5)
    .map((u) => u.unit);
  if (coAdmin.length > 0) {
    out.push({
      probleme: "Écart entre vaccinés nVPO2 et VPOb (co‑administration)",
      causes: "nVPO2 et VPOb sont co‑administrés à la même cible : un écart traduit une erreur de saisie ou une rupture d'un des deux vaccins",
      zs: joinUnits(coAdmin),
      solutions: "Faire recroiser par les DATA les effectifs vaccinés nVPO2 et VPOb et corriger la saisie ; vérifier la disponibilité des deux vaccins sur les sites",
    });
  }

  for (const k of vaccKeys) {
    const seuil = SEUIL_PERTE[k];
    const hors = units
      .filter((u) => u[k].tauxPerte != null && ((u[k].tauxPerte as number) > seuil || (u[k].tauxPerte as number) < 0))
      .map((u) => u.unit);
    if (hors.length > 0) {
      out.push({
        probleme: `Taux de perte ${VACCINE_LABELS[k]} hors seuil (> ${seuil} % ou négatif)`,
        causes: "Écart entre données de vaccination et gestion des vaccins et autres intrants ; saisie incohérente des flacons (reçus / utilisés / rendus) ; gestion des flacons entamés",
        zs: joinUnits(hors),
        solutions: "Supervision du personnel des sites et des coordonnateurs des AS pour améliorer la qualité de travail et la compilation journalière des données sur les fiches synthèses",
      });
    }
  }

  if (t.mapiGraves > 0) {
    out.push({
      probleme: `MAPI graves notifiées (${t.mapiGraves})`,
      causes: "Manifestations adverses post‑immunisation graves nécessitant une investigation immédiate",
      zs: joinUnits(units.filter((u) => u.mapiGraves > 0).map((u) => u.unit)),
      solutions: "Suivi des cas par le poste de commandement avec feedback aux ZS pour clarifier la situation et obtenir les dossiers complets d'investigation si confirmation",
    });
  }
  const doses = t.rr.vacc + t.nvpo2.vacc + t.vpob.vacc;
  if (doses > 0 && t.mapiNonGraves === 0 && t.mapiGraves === 0) {
    out.push({
      probleme: "Non‑notification des MAPI",
      causes: "Aucune MAPI (non grave ou grave) notifiée malgré le volume de doses administrées : sous‑notification probable",
      zs: perimetre,
      solutions: "Rappeler aux équipes la notification systématique des MAPI, y compris non graves, et documenter le « zéro cas »",
    });
  }
  const survTotal = t.survPFA + t.survRougeole + t.survFJ + t.survTNN;
  if (doses > 0 && survTotal === 0) {
    out.push({
      probleme: "Non‑notification de la surveillance des MPV",
      causes: "Aucun cas PFA / Rougeole / Fièvre Jaune / TNN notifié : recherche active des cas non documentée pendant la campagne",
      zs: perimetre,
      solutions: "Renforcer la recherche active des cas de MPV et documenter la surveillance, y compris les notifications « zéro cas »",
    });
  }
  const identTotal = t.pevIdent.reduce((a, b) => a + b, 0);
  const recupTotal = t.pevRecup.reduce((a, b) => a + b, 0);
  if (doses > 0 && identTotal === 0 && recupTotal === 0) {
    out.push({
      probleme: "Non‑rapportage de la récupération en PEV de routine",
      causes: "Volet « enfants et femmes enceintes identifiés et récupérés en PEV systématique » non renseigné dans le masque",
      zs: perimetre,
      solutions: "Sensibiliser les équipes et les DATA à l'enregistrement systématique des récupérations en routine pendant la campagne",
    });
  } else if (identTotal > 0 && recupTotal / identTotal < 0.8) {
    out.push({
      probleme: "Faible récupération des enfants identifiés en PEV de routine (< 80 %)",
      causes: "Faible sensibilisation, nombre insuffisant de sites de vaccination de routine",
      zs: joinUnits(units.filter((u) => u.pevIdent.reduce((a, b) => a + b, 0) > 0 && u.pevRecup.reduce((a, b) => a + b, 0) / u.pevIdent.reduce((a, b) => a + b, 0) < 0.8).map((u) => u.unit)),
      solutions: "Intensifier la récupération des enfants identifiés en routine (sites de récupération, RECO)",
    });
  }

  if (sup && sup.available) {
    const faibles = sup.byZS.filter((z) => z.pctASVisitees != null && z.pctASVisitees < 80).sort((a, b) => (a.pctASVisitees ?? 0) - (b.pctASVisitees ?? 0));
    if (faibles.length > 0) {
      out.push({
        probleme: `Faible couverture de la supervision des équipes (< 80 % d'AS visitées) — ${faibles.length} ZS`,
        causes: "Déploiement incomplet des superviseurs, distances / accessibilité, non‑saisie ODK des supervisions réalisées",
        zs: joinUnits(faibles.map((z) => z.zs)),
        solutions: "Redéployer les superviseurs vers les AS non visitées et exiger la saisie ODK de chaque supervision le jour même",
      });
    }
    const weak = sup.conformity.filter((c) => c.n >= 3 && c.pct != null && c.pct < 70).sort((a, b) => (a.pct ?? 0) - (b.pct ?? 0)).slice(0, 5);
    if (weak.length > 0) {
      out.push({
        probleme: "Non‑conformités relevées lors des supervisions des équipes",
        causes: weak.map((w) => `${w.label} (${Math.round(w.pct ?? 0)} %)`).join(" ; "),
        zs: perimetre,
        solutions: "Briefing correctif des équipes concernées par les superviseurs de proximité et suivi des actions correctrices le lendemain",
      });
    }
  }
  return out;
}

export interface BuildReportInput {
  data: MasqueData;
  filters: Filters;
  supervision: SupervisionPayload | null;
  supervisionReason?: string;
  actionsPC: ActionPCRow[];
  problemes?: ProblemeRow[];
  dateLancement?: string;
  /** Titre du rapport (par défaut : titre Kasaï Central). */
  titre?: string;
  now?: Date;
}

export function buildReportData(inp: BuildReportInput): ReportData {
  const { data, filters } = inp;
  const filtered = applyFilters(data, filters);
  const provincesInScope = Array.from(new Set(filtered.map((r) => r.province))).sort((a, b) =>
    a.localeCompare(b, "fr")
  );
  const drill = resolveDrillLevel(filters, provincesInScope.length > 1);
  const units = aggregateByUnit(filtered, drill.level);
  const antennesAll = aggregateByUnit(filtered, "antenne");
  const antennes = antennesAll.length > 1 ? antennesAll : [];
  const zones = aggregateByUnit(filtered, "zs");
  const aires = aggregateByUnit(filtered, "as");
  const total = totals(filtered);

  // Supervision ODK : jointure sur les ZS du périmètre.
  const zsRefs: ZSRef[] = zones.map((z) => ({
    zs: z.unit,
    antenne: z.antenne,
    aires: filtered.filter((r: ASRecord) => r.zs === z.unit).map((r) => r.as),
  }));
  const zsSet = new Set(zsRefs.map((z) => normKey(z.zs)));
  const supRecords = (inp.supervision?.records ?? []).filter((r) => zsSet.has(normKey(r.zs)));
  const supervision: SupervisionData = {
    available: Boolean(inp.supervision?.ok),
    reason: inp.supervision?.ok ? undefined : inp.supervisionReason ?? inp.supervision?.reason,
    fetchedAt: inp.supervision?.fetchedAt ?? "",
    dateMin: inp.supervision?.dateMin ?? "",
    formTitle: inp.supervision?.formTitle ?? "",
    total: supRecords.length,
    byZS: aggregateSupervisionByZS(supRecords, zsRefs),
    conformity: indicatorConformity(supRecords),
    points: supRecords.filter((r) => r.lat != null && r.lon != null).map((r) => ({ lat: r.lat as number, lon: r.lon as number })),
  };

  const now = inp.now ?? new Date();
  const dateMaj = `${String(now.getDate()).padStart(2, "0")}/${String(now.getMonth() + 1).padStart(2, "0")}/${now.getFullYear()}`;
  const multiProvince = provincesInScope.length > 1;
  const provinceLabel = multiProvince
    ? "RD Congo"
    : prettyProvince(provincesInScope[0] || filters.provinces[0] || data.meta.province);
  const provincesList = provincesInScope.map(prettyProvince);
  const coverEntity =
    filters.as ? `Aire de Santé : ${filters.as}` :
    filters.zs ? `Zone de Santé : ${filters.zs}` :
    filters.antenne ? `Antenne : ${filters.antenne}` :
    multiProvince
      ? (provincesList.length <= 5
          ? `Provinces : ${provincesList.join(", ")}`
          : `RD Congo — ${provincesList.length} provinces`)
      : `Province du ${provinceLabel}`;
  const dateLancement = inp.dateLancement || data.meta.dateDebut || "";
  const jl = data.meta.jourLabels;
  const periodeJours = jl.length ? (jl.length === 1 ? jl[0] : `${jl[0]} → ${jl[jl.length - 1]}`) : "";

  return {
    titre: inp.titre ?? REPORT_TITLE,
    province: provincesInScope[0] || filters.provinces[0] || data.meta.province,
    provinceLabel,
    scopeLabel: scopeLabel(filters, provinceLabel),
    coverEntity,
    dateMaj,
    dateLancement: fmtDateFR(dateLancement),
    periode: periodeJours,
    jourLabels: jl,
    nbJours: data.meta.nbJours,
    byUnitLabel: drill.label,
    units,
    antennes,
    zones,
    aires,
    total,
    supervision,
    actionsPC: inp.actionsPC,
    problemes: inp.problemes ?? computeProblemes(units, total, drill.label, supervision),
  };
}

function normKey(s: string): string {
  return (s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

export { ANTIGENES, DOSES_PAR_FLACON };

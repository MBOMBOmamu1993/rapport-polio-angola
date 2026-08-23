/**
 * Supervision des équipes de vaccination (formulaire ODK « POLIO‑RR SUPERVISION
 * EQUIPES INTEGREES BLOC 3 », id 17559 sur api.whonghub.org).
 *
 * Ce module ne contient que les types et les calculs (utilisables côté client) ;
 * l'appel réseau authentifié est fait côté serveur dans `odk-server.ts` via la
 * route `/api/supervision`.
 */

/** Indicateurs clés extraits du questionnaire — `good` = réponse conforme. */
export const SUPERVISION_INDICATORS: { key: string; field: string; label: string; good: "oui" | "non"; group: string }[] = [
  { key: "membre_present", field: "group_offre/membre_present", label: "Équipe complète et en poste", good: "oui", group: "Offre de service" },
  { key: "membre_forme", field: "group_offre/membre_forme", label: "Membres de l'équipe formés", good: "oui", group: "Offre de service" },
  { key: "membre_localite", field: "group_offre/membre_localite", label: "Au moins un membre de la localité", good: "oui", group: "Offre de service" },
  { key: "respect_circuit", field: "group_offre/respect_circuit", label: "Circuit des personnes respecté", good: "oui", group: "Offre de service" },
  { key: "vaccin_suffisant", field: "group_stock/vaccin_suffisant", label: "Vaccins et consommables suffisants", good: "oui", group: "Stocks" },
  { key: "manque_rr", field: "group_stock/manque_rr", label: "Pas de manque de RR / diluant", good: "non", group: "Stocks" },
  { key: "manque_nvpo2", field: "group_stock/manque_nvpo2", label: "Pas de manque de nVPO2", good: "non", group: "Stocks" },
  { key: "manque_vpob", field: "group_stock/manque_vpob", label: "Pas de manque de VPOb", good: "non", group: "Stocks" },
  { key: "manque_sab", field: "group_stock/manque_sab", label: "Pas de manque de SAB", good: "non", group: "Stocks" },
  { key: "presence_coclip", field: "group_stock/presence_coclip", label: "Utilisation du co‑clip (nVPO2 / VPOb)", good: "oui", group: "Stocks" },
  { key: "porte_vaccin", field: "group_chaine_froid/disponibilite_porte_vaccin", label: "≥ 2 porte‑vaccins en bon état", good: "oui", group: "Chaîne du froid" },
  { key: "vaccins_porte_vaccin", field: "group_chaine_froid/RR_porte_vaccin", label: "Vaccins dans le porte‑vaccin avec accumulateurs", good: "oui", group: "Chaîne du froid" },
  { key: "flacon_etat", field: "group_chaine_froid/flacon_etat", label: "Aucun flacon en mauvais état", good: "non", group: "Chaîne du froid" },
  { key: "precharge", field: "group_chaine_froid/precharge_vaccin", label: "Pas de préchargement des seringues", good: "non", group: "Chaîne du froid" },
  { key: "compte_gouttes_change", field: "group_chaine_froid/change_compte_gouttes", label: "Compte‑gouttes changé à chaque flacon", good: "oui", group: "Chaîne du froid" },
  { key: "verification_age", field: "group_securite/verification_age", label: "Vérification de l'âge", good: "oui", group: "Sécurité & MAPI" },
  { key: "maitrise_technique", field: "group_securite/maitrise_technique", label: "Maîtrise de la technique vaccinale", good: "oui", group: "Sécurité & MAPI" },
  { key: "flacon_6h", field: "group_securite/flacon_ouvert_plus_six_h", label: "Pas de flacon ouvert > 6 h", good: "non", group: "Sécurité & MAPI" },
  { key: "compte_goutte", field: "group_securite/compte_goutte", label: "Nombre correct de gouttes", good: "oui", group: "Sécurité & MAPI" },
  { key: "lavage_mains", field: "group_securite/lavage_mains", label: "Lavage régulier des mains", good: "oui", group: "Sécurité & MAPI" },
  { key: "recapuchon", field: "group_securite/recapuchon_seringue", label: "Pas de recapuchonnage", good: "non", group: "Sécurité & MAPI" },
  { key: "securite_boite", field: "group_securite/securite_boite", label: "Boîtes de sécurité en lieu protégé", good: "oui", group: "Sécurité & MAPI" },
  { key: "kit_mapi", field: "group_securite/kit_mapi", label: "Kit MAPI disponible", good: "oui", group: "Sécurité & MAPI" },
  { key: "outils_collecte", field: "group_donnees/disponibilite_outil_collecte", label: "Outils de collecte disponibles", good: "oui", group: "Données" },
  { key: "fiche_pointage", field: "group_donnees/fiche_pointage", label: "Pas de manque de fiches de pointage", good: "non", group: "Données" },
  { key: "cartes", field: "group_donnees/disponibilite_carte_vaccination", label: "Pas de manque de cartes de vaccination", good: "non", group: "Données" },
  { key: "fiche_mapi", field: "group_donnees/fiche_notification_MAPI", label: "Pas de manque de fiches MAPI", good: "non", group: "Données" },
  { key: "carte_ok", field: "group_donnees/remplissage_correct_carte", label: "Cartes correctement remplies", good: "oui", group: "Données" },
  { key: "fiche_ok", field: "group_donnees/remplissage_correct_fiche", label: "Fiche de pointage correctement remplie", good: "oui", group: "Données" },
  { key: "conformite_rr", field: "group_donnees/conformite_donnees", label: "Doses RR conformes aux vaccinés", good: "oui", group: "Données" },
  { key: "conformite_nvpo2", field: "group_donnees/conformite_nVPO2", label: "Doses nVPO2 conformes aux vaccinés", good: "oui", group: "Données" },
  { key: "conformite_vpob", field: "group_donnees/conformite_VPOb", label: "Doses VPOb conformes aux vaccinés", good: "oui", group: "Données" },
  { key: "supports", field: "group_communication/supports_affiches", label: "Supports de visibilité présents", good: "oui", group: "Communication" },
  { key: "megaphones", field: "group_communication/megaphones", label: "Mégaphones avec piles", good: "oui", group: "Communication" },
  { key: "date_informee", field: "group_communication/date_informee", label: "Communauté informée des dates", good: "oui", group: "Communication" },
  { key: "communication_equipe", field: "group_communication/communication_equipe", label: "Messages bien communiqués", good: "oui", group: "Communication" },
  { key: "plan_deploiement", field: "group_conclusion/plan_deploiement", label: "Plan de déploiement disponible", good: "oui", group: "Conclusion" },
  { key: "visite_superviseur", field: "group_conclusion/visite_superviseur", label: "Visite du superviseur d'équipe du jour", good: "oui", group: "Conclusion" },
];

export interface SupervisionRecord {
  id: number;
  date: string;
  province: string;
  antenne: string;
  zs: string;
  as: string;
  site: string;
  superviseur: string;
  profil: string;
  typeEquipe: string;
  lat: number | null;
  lon: number | null;
  /** Réponses aux indicateurs clés (clé = SUPERVISION_INDICATORS.key) : "oui" | "non" | null. */
  answers: Record<string, "oui" | "non" | null>;
  actions: string;
  recommandations: string;
}

export interface SupervisionPayload {
  ok: boolean;
  reason?: string;
  fetchedAt: string;
  dateMin: string;
  province: string;
  formId: number;
  formTitle: string;
  total: number;
  records: SupervisionRecord[];
  /** Vrai si les données proviennent d'une extraction antérieure (serveur ODK indisponible). */
  stale?: boolean;
  /** Curseur incrémental : plus grand `_submission_time` déjà récupéré (ISO, sans fuseau). */
  lastSubmissionTime?: string;
  /** Vrai si l'extraction s'est arrêtée au budget de temps — la suivante reprend au curseur. */
  partial?: boolean;
}

/** Normalise un nom d'unité (ODK slug « bena_leka » ↔ masque « BENA LEKA »). */
export function normUnit(s: string): string {
  return (s || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

/** Score de conformité d'une supervision (% d'indicateurs conformes parmi les renseignés). */
export function supervisionScore(r: SupervisionRecord): number | null {
  let n = 0;
  let ok = 0;
  for (const ind of SUPERVISION_INDICATORS) {
    const v = r.answers[ind.key];
    if (v !== "oui" && v !== "non") continue;
    n++;
    if (v === ind.good) ok++;
  }
  return n ? (ok / n) * 100 : null;
}

export interface ZSSupervision {
  zs: string;
  antenne: string;
  nbSupervisions: number;
  nbASTotal: number;
  nbASVisitees: number;
  pctASVisitees: number | null;
  score: number | null;
  /** Aires de Santé du masque non encore visitées (noms du masque). */
  asNonVisitees: string[];
  /** Aires de Santé visitées (noms du masque quand reconnues, sinon libellé ODK). */
  asVisitees: string[];
}

/** Zones de Santé du périmètre (masque) : nom, antenne et liste des AS. */
export interface ZSRef { zs: string; antenne: string; aires: string[] }

function matchAS(odkAs: string, aires: string[]): string | null {
  const t = normUnit(odkAs);
  if (!t) return null;
  const normed = aires.map((a) => normUnit(a));
  const exact = normed.indexOf(t);
  if (exact >= 0) return aires[exact];
  // Tolérance : « masuika2 » ↔ « MASUIKA II », « MASUIKA 2 » ; préfixe commun long.
  const t2 = t.replace(/(II|2)$/, "");
  for (let i = 0; i < normed.length; i++) {
    const n = normed[i];
    const n2 = n.replace(/(II|2)$/, "");
    if (t2 && n2 && t2 === n2 && (t.endsWith("2") || t.endsWith("II")) === (n.endsWith("2") || n.endsWith("II"))) return aires[i];
  }
  let best = -1;
  let bestLen = 0;
  for (let i = 0; i < normed.length; i++) {
    const n = normed[i];
    if (n.length >= 4 && (n.startsWith(t) || t.startsWith(n))) {
      const l = Math.min(n.length, t.length);
      if (l > bestLen) { bestLen = l; best = i; }
    }
  }
  return best >= 0 ? aires[best] : null;
}

/** Agrège les supervisions par Zone de Santé du périmètre. */
export function aggregateSupervisionByZS(records: SupervisionRecord[], zsRefs: ZSRef[]): ZSSupervision[] {
  const byZS = new Map<string, SupervisionRecord[]>();
  for (const r of records) {
    const k = normUnit(r.zs);
    const arr = byZS.get(k);
    if (arr) arr.push(r);
    else byZS.set(k, [r]);
  }
  return zsRefs.map((ref) => {
    const recs = byZS.get(normUnit(ref.zs)) ?? [];
    const visited = new Set<string>();
    const unmatched = new Set<string>();
    for (const r of recs) {
      const m = matchAS(r.as, ref.aires);
      if (m) visited.add(m);
      else if (r.as) unmatched.add(r.as);
    }
    const nbVis = Math.min(ref.aires.length, visited.size + unmatched.size);
    const scores = recs.map(supervisionScore).filter((s): s is number => s != null);
    return {
      zs: ref.zs,
      antenne: ref.antenne,
      nbSupervisions: recs.length,
      nbASTotal: ref.aires.length,
      nbASVisitees: nbVis,
      pctASVisitees: ref.aires.length ? (nbVis / ref.aires.length) * 100 : null,
      score: scores.length ? scores.reduce((a, b) => a + b, 0) / scores.length : null,
      asNonVisitees: ref.aires.filter((a) => !visited.has(a)),
      asVisitees: [...Array.from(visited), ...Array.from(unmatched)],
    };
  });
}

/** % de conformité par indicateur sur un ensemble de supervisions. */
export function indicatorConformity(records: SupervisionRecord[]): { key: string; label: string; group: string; n: number; ok: number; pct: number | null }[] {
  return SUPERVISION_INDICATORS.map((ind) => {
    let n = 0;
    let ok = 0;
    for (const r of records) {
      const v = r.answers[ind.key];
      if (v !== "oui" && v !== "non") continue;
      n++;
      if (v === ind.good) ok++;
    }
    return { key: ind.key, label: ind.label, group: ind.group, n, ok, pct: n ? (ok / n) * 100 : null };
  });
}

/** Couleur du taux d'AS visitées / score de supervision (seuils du modèle : < 50, 50‑80, ≥ 80). */
export function supervisionColor(v: number | null): "red" | "orange" | "green" | "none" {
  if (v == null || !Number.isFinite(v)) return "none";
  if (v >= 80) return "green";
  if (v >= 50) return "orange";
  return "red";
}

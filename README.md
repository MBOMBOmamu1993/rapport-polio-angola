# Campagne intégrée RR‑POLIO — Kasaï Central

Application web (Next.js) à **deux pages** pour la province du Kasaï Central, seule
province menant la campagne intégrée Rougeole‑Rubéole + Polio (bloc 3, août 2026) :

1. **Importer le masque de saisie** — le masque Excel « Masque_Saisie_Intégration
   Polio‑RR » (province, antenne ou zone de santé). Les analyses sont calculées
   instantanément dans le navigateur.
2. **Télécharger le rapport** — génération d'un PowerPoint (.pptx) qui reproduit le
   modèle officiel « Résultats partiels de la campagne intégrée Bloc 3 » (design et
   analyses du début à la fin), avec **commentaires générés automatiquement** et
   **analyses de supervision lues directement dans ODK**.

Cibles : polio (nVPO2 / VPOb) 0‑9 ans (≈ 33,3 % de la population) et RR 6 mois‑14 ans
(≈ 46 %), telles que présentes dans le masque (colonnes « Cible Polio » / « Cible RR »).
La campagne dure 5 jours + ratissage : les feuilles `Jour1`…`Jour5` sont les jours J1..J5,
`Jour6` est le **ratissage**.

## Contenu du rapport généré

| # | Diapositive(s) | Source |
|---|---|---|
| 1 | Page de garde (logos MinSanté / PEV, photos, partenaires, date de mise à jour) | — |
| 2 | Plan de présentation | — |
| 3 | Suivi des points d'action des précédents PC (saisis dans l'application) | UI |
| 4 | Points saillants (complétude, CV RR / nVPO2 / VPOb, vaccinés, flacons, pertes, MAPI) | masque |
| 5 | Synthèse des principaux indicateurs par Antenne puis par ZS (cellules colorées) | masque |
| 6 | Complétude journalière (J1..J5, ratissage) et globale : tableaux + graphiques | masque |
| 7‑9 | Couverture vaccinale et taux de perte RR, nVPO2, VPOb (par Antenne, par ZS) | masque |
| 10 | Notification des MAPI (par Antenne / ZS / AS) + proportion pour 100 000 doses | masque |
| 11 | Surveillance MPV (PFA, rougeole, fièvre jaune, TNN) | masque |
| 12 | Récupération PEV de routine par antigène (22 antigènes) | masque |
| 13 | Proportion des vaccinés par tranche d'âge (RR, nVPO2, VPOb) | masque |
| 14 | Proportion des vaccinés par sexe | masque |
| 15 | Supervision des équipes : % d'AS visitées par ZS, cartes (choroplèthe + points GPS), scores, conformité par indicateur | **ODK** |
| 16 | Problèmes rencontrés / Actions correctrices (générés automatiquement, modifiables) | analyses |
| 17 | Merci | — |

Filtres en cascade Antenne → Zone de Santé → Aire de Santé : le rapport se désagrège
au niveau immédiatement inférieur (ZS par défaut, AS si une ZS est sélectionnée).

## Supervision des équipes (ODK)

Les analyses de supervision proviennent du formulaire ODK **« POLIO‑RR SUPERVISION
EQUIPES INTEGREES BLOC 3 »** (`id 17559`, `https://api.whonghub.org/api/v1/data/17559`),
filtré sur `Province = Kasai_Central` et `date_supervision ≥ 2026‑08‑17`. La route
serveur `GET /api/supervision` interroge l'API (identifiants côté serveur, cache 5 min)
et renvoie les enregistrements utiles ; le navigateur calcule :

- % d'aires de santé visitées par ZS (numérateur : AS distinctes supervisées dans ODK,
  dénominateur : AS du masque) et liste des AS non encore visitées ;
- score de conformité par supervision / par ZS (≈ 38 indicateurs oui/non du
  questionnaire, voir `lib/odk-supervision.ts`) et % de conformité par indicateur ;
- cartes des ZS du Kasaï Central (fond `public/geo/rdc_zs.topojson`) : choroplèthe
  < 50 % / 50‑80 % / ≥ 80 % et carte des points GPS des supervisions.

Variables d'environnement optionnelles (Vercel → Settings → Environment Variables) :
`ODK_BASE_URL`, `ODK_USERNAME`, `ODK_PASSWORD`, `ODK_SUPERVISION_FORM_ID` (17559),
`ODK_PROVINCE` (Kasai_Central), `ODK_DATE_MIN` (2026‑08‑17). Sans ces variables, la
configuration par défaut de la campagne (identique au dépôt
`rr-polio-independent-monitoring-dashboard`) est utilisée. La date minimale est aussi
modifiable dans l'interface.

## Développement

```bash
npm install
npm run dev        # http://localhost:3000
npm run build
npm run typecheck
# Test hors navigateur (masque → PPTX) :
npx tsx scripts/test-report.ts <masque.xlsx> <sortie.pptx> [--no-odk] [--zs=NOM]
```

## Structure

- `lib/parse-masque.ts` — lecture du masque intégré (Synthèse, Jour1..Jour6, Donnees de base) : cibles, vaccinés par tranche d'âge et sexe, gestion des flacons, MPV, MAPI, PEV systématique.
- `lib/analytics.ts` — filtres en cascade + agrégations par unité (RR / nVPO2 / VPOb, complétude journalière, MAPI pour 100 000 doses…).
- `lib/odk-supervision.ts` — indicateurs du questionnaire, agrégation par ZS, scores ; `lib/odk-server.ts` — accès API ODK (serveur) ; `app/api/supervision/route.ts`.
- `lib/report-data.ts` — assemblage des données du rapport + problèmes automatiques.
- `lib/export-report-pptx.ts` — génération PowerPoint (design du modèle Bloc 3).
- `lib/zs-map.ts` — cartes des ZS (SVG → PNG).
- `public/cover/` — images de la page de garde (logos, photos, partenaires, « Merci »).

## Compilation partagée (toutes les antennes / ZS)

Chaque import est envoyé au stockage partagé **Vercel KV** (`POST /api/import`) et la page
Rapport propose « Compilation provinciale » (`GET /api/national`). Activation : Vercel →
Storage → KV / Upstash → Connect → Redeploy. La réinitialisation admin est protégée par
`ADMIN_RESET_CODE`. Sans KV, l'application fonctionne en mode local (navigateur).

## Codes couleur

| Couverture / complétude | | Taux de perte | |
|---|---|---|---|
| < 80 % | rouge | > 10 % | rouge |
| 80 – 95 % | jaune | 5 – 10 % | jaune |
| 95 – 100 % | vert | 0 – 5 % | vert |
| > 100 % | bleu | < 0 % | bleu |

Supervision (% AS visitées, score) : < 50 % rouge, 50 – 80 % orange, ≥ 80 % vert.

Programme Élargi de Vaccination — RD Congo. Usage interne.

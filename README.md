# Rapport — Campagne de vaccination polio synchronisée avec l'Angola

Application web à **deux pages** pour les 5 provinces de la RDC organisant la
campagne polio synchronisée avec l'Angola (co-administration **nVPO2** + **VPOb**).

1. **Importer le masque de saisie** — chaque niveau (province / antenne / zone de
   santé) importe son masque Excel. Les analyses sont calculées instantanément
   dans le navigateur. Réimporter remplace automatiquement l'ancienne version.
2. **Télécharger le rapport** — filtres en cascade (Province → Antenne → ZS →
   Aire de Santé) et génération d'un rapport **PowerPoint (.pptx)** reproduisant
   fidèlement le modèle officiel pour la composante polio (nVPO2 et VPOb).

> Aucune donnée n'est envoyée à un serveur : tout le traitement se fait côté
> navigateur et le masque importé est conservé localement (localStorage).

## ✨ Nouveautés de cette version (v2)

- **Palette navy / marine** alignée sur le modèle Power BI officiel
  (bandeaux `#1F3864`, accent `#2563EB`).
- **Tableau jour-par-jour** par Zone de Santé pour chaque vaccin (nVPO2, VPOb) :
  *Cible Campagne | Cible Polio Journalière | Vaccinés J1 | Couvert. J1 | … |
  Couvert. Globale*, avec coloration des cellules selon les seuils du modèle
  (rouge < 80 %, jaune 80–95 %, vert 95–100 %, bleu > 100 %).
- **Tableau Gestion des vaccins** complet (flacons reçus / utilisés / rendus /
  perdus / enfants vaccinés / % de perte) + bar-chart horizontal du taux de
  perte par ZS, comme dans le modèle Kwango.
- **Slide « Complétude »** : KPI géant + cards par jour + tableau journalier.
- **Lecture des feuilles `Jour1` / `Jour2` / … du masque** pour produire les
  colonnes journalières (les feuilles vides ne génèrent pas de colonne).
- **Composante RR retirée** — le rapport ne traite plus que la partie polio
  (nVPO2 et VPOb) puisque la campagne synchronisée avec l'Angola est une
  campagne polio + co-administration.

## Développement

```bash
npm install
npm run dev      # http://localhost:3000
npm run build    # build de production
npm run typecheck
```

## Structure

- `app/import` — page d'import du masque de saisie.
- `app/rapport` — page de filtres + génération du rapport PPTX.
- `lib/parse-masque.ts` — lecture de la feuille « Synthèse » + des feuilles
  journalières (`Jour1`, `Jour2`, …).
- `lib/analytics.ts` — filtres en cascade + agrégations (totaux, par unité, par
  jour) avec helper `coverageByDay()` reproduisant le tableau Power BI.
- `lib/export-report-pptx.ts` — génération du rapport PowerPoint (12 slides
  reproduisant le modèle officiel pour nVPO2 et VPOb uniquement).
- `lib/zs-map.ts` — rendu choroplèthe optionnel des ZS pour la slide
  « Spatialisation ».
- `public/cover-polio.png` — image de la page de garde ; `public/logo/pev.png` —
  logo PEV.

## Slides générés (modèle nVPO2 + VPOb)

1. Page de garde (image polio + logos)
2. Plan de présentation
3. Points saillants
4. Complétude des rapports par ZS (jauge + tableau J1..Jn)
5. Couvertures vaccinales **nVPO2** par jour (cible / vaccinés / couverture)
6. Couvertures vaccinales **VPOb** par jour
7. Récupération PEV de routine (co-administration)
8. Gestion du vaccin **nVPO2** (flacons reçus / utilisés / rendus / perdus + % perte)
9. Gestion du vaccin **VPOb**
10. Surveillance des MAPI
11. Problèmes / Actions correctrices (éditables dans l'UI)
12. Merci pour votre attention

## Déploiement Vercel

1. Pousser ce dossier dans un dépôt GitHub.
2. Sur https://vercel.com → **Add New Project** → importer le dépôt.
3. Framework **Next.js** détecté automatiquement.
4. **Deploy**. Vercel fournit l'URL publique à partager à tous les niveaux.

## Compilation nationale (toutes les provinces)

Pour que les 5 provinces / antennes / ZS alimentent une **vue pays commune**,
l'application utilise un stockage partagé **Vercel KV**. Chaque import est
découpé et stocké **par Zone de Santé** ; réimporter une entité remplace
uniquement ses ZS. Le niveau pays peut alors télécharger la situation de toutes
les provinces à la fois.

### Activation (une seule fois)

1. Vercel → votre projet → onglet **Storage** → **Create Database** → **KV**
   (ou **Upstash for Redis** via le Marketplace).
2. **Connect** la base au projet : Vercel injecte automatiquement les variables
   `KV_REST_API_URL` et `KV_REST_API_TOKEN`.
3. **Redeploy** le projet.

Tant que KV n'est pas activé, l'application fonctionne en **mode local** :
chaque import reste disponible dans le navigateur qui l'a importé.

### Endpoints

- `POST /api/import` — enregistre un import (corps = données parsées du masque).
- `GET /api/national` — renvoie la compilation consolidée + la liste des entités.

## Codes couleur (seuils officiels)

| Couverture | Couleur | Code     |
|------------|---------|----------|
| < 80 %     | Rouge   | `#E23636` |
| 80 – 95 %  | Jaune   | `#F1C40F` |
| 95 – 100 % | Vert    | `#22B457` |
| > 100 %    | Bleu    | `#1D4ED8` |

## Licence

Programme Élargi de Vaccination — RD Congo. Usage interne.

# Rapport — Campagne de vaccination polio synchronisée avec l'Angola

Application web à **deux pages** pour les 5 provinces de la RDC organisant la campagne
polio synchronisée avec l'Angola (co-administration **nVPO2** + **VPOb**).

1. **Importer le masque de saisie** — chaque niveau (province / antenne / zone de santé)
   importe son masque Excel. Les analyses sont calculées instantanément dans le
   navigateur. Réimporter remplace automatiquement l'ancienne version.
2. **Télécharger le rapport** — filtres en cascade (Province → Antenne → ZS → Aire de
   Santé) et génération d'un rapport **PowerPoint (.pptx)** reproduisant le modèle
   officiel (composante polio uniquement).

> Aucune donnée n'est envoyée à un serveur : tout le traitement se fait côté navigateur
> et le masque importé est conservé localement (localStorage).

## Développement

```bash
npm install
npm run dev      # http://localhost:3000
npm run build    # build de production
```

## Structure

- `app/import` — page d'import du masque de saisie.
- `app/rapport` — page de filtres + génération du rapport PPTX.
- `lib/parse-masque.ts` — lecture de la feuille « Synthèse » du masque (polio only).
- `lib/analytics.ts` — filtres en cascade + agrégation des indicateurs.
- `lib/export-report-pptx.ts` — génération du rapport PowerPoint.
- `public/cover-polio.png` — image de la page de garde ; `public/logo/pev.png` — logo PEV.

## Déploiement Vercel

1. Pousser ce dossier dans un dépôt GitHub.
2. Sur https://vercel.com → **Add New Project** → importer le dépôt.
3. Framework **Next.js** détecté automatiquement. Si l'application est dans un
   sous-dossier, définir **Root Directory** = `polio-angola`.
4. **Deploy**. Vercel fournit l'URL publique à partager à tous les niveaux.

## Compilation nationale (toutes les provinces)

Pour que les 5 provinces / antennes / ZS alimentent une **vue pays commune** (et non
chacune isolée dans son navigateur), l'application utilise un stockage partagé
**Vercel KV**. Chaque import est découpé et stocké **par Zone de Santé** ; réimporter
une entité remplace uniquement ses ZS. Le niveau pays peut alors télécharger la
situation de toutes les provinces à la fois.

### Activation (une seule fois)

1. Vercel → votre projet → onglet **Storage** → **Create Database** → **KV**
   (ou **Upstash for Redis** via le Marketplace).
2. **Connect** la base au projet : Vercel injecte automatiquement les variables
   `KV_REST_API_URL` et `KV_REST_API_TOKEN`.
3. **Redeploy** le projet.

Tant que KV n'est pas activé, l'application fonctionne en **mode local** : chaque
import reste disponible dans le navigateur qui l'a importé.

### Endpoints

- `POST /api/import` — enregistre un import (corps = données parsées du masque).
- `GET /api/national` — renvoie la compilation consolidée + la liste des entités.

# Handoff — Mise à jour du rapport polio Angola (v2)

Ce package contient les fichiers à intégrer dans le dépôt
[`MBOMBOmamu1993/rapport-polio-angola`](https://github.com/MBOMBOmamu1993/rapport-polio-angola)
afin de reproduire fidèlement les analyses et le design du **modèle Power BI
Kwango** pour la composante polio (nVPO2 + VPOb), tout en améliorant le design
général de l'application avec une **palette navy / marine**.

## Fichiers modifiés

Remplacer les fichiers existants par les versions fournies dans ce ZIP. La
structure du projet est identique — seuls les contenus changent.

| Fichier | Rôle | Type |
|--|--|--|
| `tailwind.config.ts` | Palette `navy` + `accent` + seuils `threshold` | **Remplacer** |
| `app/globals.css` | Bandeau navy réutilisable + arrière-plan léger | **Remplacer** |
| `components/Nav.tsx` | Header navy + ring blanc autour du logo | **Remplacer** |
| `app/import/page.tsx` | Hero navy + dropzone + KPI tone par seuil | **Remplacer** |
| `app/rapport/page.tsx` | Filtres + slides list + boutons navy | **Remplacer** |
| `lib/parse-masque.ts` | Lit aussi les feuilles `Jour1..Jour4` + métadonnées `nbJours` / `jourLabels` | **Remplacer** |
| `lib/analytics.ts` | Ajoute `coverageByDay()`, `gestion` avec flacons reçus/rendus, agrégats journaliers | **Remplacer** |
| `lib/export-report-pptx.ts` | Générateur PPTX entièrement réécrit (palette navy + tableaux jour-par-jour conformes au modèle) | **Remplacer** |
| `lib/kv-store.ts` | Ajoute `nbJours` / `jourLabels` à la compilation nationale | **Remplacer** |
| `README.md` | Documentation mise à jour | **Remplacer** |

> ⚠️ **Aucun fichier supprimé**, aucune dépendance ajoutée ou retirée.
> `package.json` et `package-lock.json` restent **inchangés**.

## Points-clés de la nouvelle version

### 1. Palette navy alignée sur le modèle Power BI

- Bandeau principal `#1F3864` (navy-700).
- Bandeau secondaire `#162A4D` (navy-800).
- Accent CTA `#2563EB` (accent-500).
- Seuils de coloration : rouge `#E23636` (< 80 %), jaune `#F1C40F` (80–95 %),
  vert `#22B457` (95–100 %), bleu `#1D4ED8` (> 100 %).

### 2. Reproduction fidèle du modèle Kwango pour nVPO2 et VPOb

Slides 5 et 6 — **Couvertures par ZS × par jour** :
- Colonnes : `ZS | Cible Campagne | Cible Polio Journalière | Vaccinés J1 |
  Couvert. J1 | … | Couvert. Globale`.
- Lignes triées par couverture globale décroissante (comme le modèle).
- Cellules de couverture colorées selon les seuils, en blanc/gras.
- Total en bas en navy foncé + texte blanc.

Slides 8 et 9 — **Gestion vaccin** :
- Tableau gauche : `ZS | Flacons reçus | utilisés | rendus | perdus | Enfants
  vaccinés | % Perte`.
- Bar chart horizontal à droite : « Répartition du taux de perte (%) par ZS »,
  trié décroissant.
- % de perte affiché avec coloration sémantique.

Slide 4 — **Complétude** :
- KPI géant (donut coloré selon seuil) au lieu de la jauge cassée.
- Tableau par ZS avec colonnes `Attendus | Reçus J1 | Compl. J1 | … | Compl. Globale`.

### 3. Lecture des données journalières dans le masque

Le parseur lit désormais les feuilles `Jour1`, `Jour2`, `Jour3`, `Jour4` du
masque pour extraire les vaccinés nVPO2 / VPOb et rapports reçus par jour. Si
le masque ne contient que la feuille Synthèse, le rapport tombe en mode
« agrégé » sans colonnes journalières.

### 4. Co-administration uniquement

La composante RR a été entièrement retirée du rapport. Les références à
« Rougeole-Rubéole » dans les commentaires et titres ont été supprimées —
seules `nVPO2` et `VPOb` apparaissent.

## Étapes pour Claude Code

```bash
# 1. Cloner le repo
git clone https://github.com/MBOMBOmamu1993/rapport-polio-angola
cd rapport-polio-angola

# 2. Dézipper le handoff par-dessus (sauf node_modules / .next)
unzip -o /chemin/vers/handoff-polio-v2.zip -d .

# 3. Vérifier le build
npm install        # pas de nouvelle dépendance, mais par sécurité
npm run typecheck  # doit passer sans erreur
npm run build      # doit passer sans erreur

# 4. Tester en local
npm run dev
# → http://localhost:3000 — Importer un masque puis Télécharger le rapport

# 5. Commit + push
git add -A
git commit -m "feat(v2): palette navy + reproduction fidèle du modèle Kwango (nVPO2 + VPOb)"
git push origin main
```

Vercel redéploie automatiquement à la réception du push (CI/CD déjà en place).

## Vérifications post-déploiement

- [ ] Page `/import` : hero navy avec bandeau dégradé, dropzone propre.
- [ ] Import du masque : toast vert « Masque de saisie importé avec succès ».
- [ ] Page `/rapport` : filtres en cascade fonctionnels (Province → Antenne → ZS
      → Aire de Santé).
- [ ] Télécharger le rapport : PPTX généré contient 12 slides, **aucune
      référence à RR**, tableaux jour-par-jour pour nVPO2 et VPOb, coloration
      des cellules.
- [ ] Page de garde : image polio + logos PEV / OMS + bandeau navy à gauche.

## Compatibilité

- TypeScript 5.6.2 — strict.
- Next.js 14.2.13.
- React 18.3.1.
- pptxgenjs 4.0.1.
- xlsx 0.18.5.

---

**Auteur** : Claude Design (Anthropic) — mai 2026.
**Modèle de référence** : `PRESENTATION DES RESULTATS PARTIELS AU J6 RR POLIO_vf.pptx`
(Kwango) — composante polio uniquement.

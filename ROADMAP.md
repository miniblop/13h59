# 13H59 SHOP — Feuille de route

> Source de vérité du code = le dépôt GitHub `miniblop/13h59`.
> Ce fichier tient lieu de suivi léger (pas de gestionnaire de tâches dans les Projets Claude).
> À mettre à jour à la fin de chaque palier.

## État actuel (fait ✅)
- Caisse Google Sheets (Apps Script) : `Code.gs` + `caisse.html`.
- Site web statique (GitHub Pages) : `index.html`, branché sur le backend Apps Script `Api.gs`.
- 3 rôles par mot de passe (caisse / vendeur / gestion), Sheet privé possible.
- Dashboards : onglet Vendeur, onglet Gestion (tableau transposé, filtres, CSV, graphes).
- Commission recalculée **par mois × vendeur, figée sur toutes les données** (logique à préserver).

## Palier 1 — Sortir de Google Sheets 🎯 (prochaine étape)
- [ ] Schéma **Postgres** : vendeurs, articles (stock par vendeur), ventes, paniers, remises, taxes.
- [ ] Backend **FastAPI** minimal répliquant l'existant : `data`, `caisse_data`, `caisse_save`.
- [ ] Réimplémenter la logique métier en Python **à l'identique** :
      prix_client, taxe (1,75 % CB / 0 espèces), prime (remises « magasin » = prix plein),
      commission mois × vendeur figée (barème 8/12 % avant 2026, 10/15 % après ; seuils 100/250 €).
- [ ] Script de migration : import de l'historique `All data` → Postgres.
- [ ] Brancher le front actuel (`index.html`) sur FastAPI au lieu d'Apps Script.
- [ ] Déploiement backend (Render / Railway / Fly.io).

## Palier 2 — Comptes & droits
- [ ] Vraie **authentification** (comptes vendeurs, sessions) — remplace les mots de passe partagés.
- [ ] **Gestion des articles par vendeur** + suivi de **stock** (entrées, sorties, seuils).
- [ ] Droits par page / rôle propres.

## Palier 3 — Vente en ligne
- [ ] Catalogue public + panier client.
- [ ] **Paiement Stripe** (probablement **Stripe Connect** pour reverser aux vendeurs).

## ⚠️ Points en attente / à valider avant d'avancer
- **Juridique (bloquant pour le Palier 3)** : encaisser pour le compte de tiers (primes vendeurs)
  est réglementé en France (statut proche marketplace / agent de paiement). À faire valider par un
  professionnel AVANT de brancher les paiements réels. (Claude n'est ni juriste ni conseiller financier.)
- Décision **maison vs Shopify/Woo/Medusa** selon l'ambition e-commerce réelle.
- `numero_panier` incomplet dans l'historique `All data` → à compléter pour fiabiliser les métriques « par panier ».

## Réflexe de maintenance
- Le code fait référence vit sur **GitHub**. Si le connecteur GitHub est actif, Claude lit
  directement la dernière version → plus de fichiers périmés à re-déposer à la main.

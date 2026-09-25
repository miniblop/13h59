# 13H59 SHOP — Feuille de route

> Source de vérité du code = le dépôt GitHub `miniblop/13h59`.
> Ce fichier tient lieu de suivi léger (pas de gestionnaire de tâches dans les Projets Claude).
> À mettre à jour à la fin de chaque palier.

## État actuel (fait ✅)
- Caisse Google Sheets (Apps Script) : `apps-script/Code.js` + `apps-script/Caisse.html`.
- Site web statique (GitHub Pages) : `index.html`, branché sur le backend Apps Script `apps-script/Api.js`.
- 3 rôles par mot de passe (caisse / créateur / gestion), Sheet privé possible.
- Dashboards : onglet Créateur, onglet Gestion (tableau transposé, filtres, CSV, graphes).
- Commission recalculée **par mois × créateur, figée sur toutes les données** (logique à préserver).
- Backoffice v1 : le Google Sheet est structuré comme la future base (un onglet = une table) : `ventes` (clé `id_vente`, clé étrangère `id_createur`), `createurs`, `categories`, `remises`, `paiements`, `journal`. Plus de formules ni d'onglet `All data` : la caisse écrit des valeurs.
- Caisse : **anti-doublon** par identifiant de transaction (idempotence serveur) + récap « Ventes du jour » (totaux CB / espèces, détection des paniers identiques).

## Palier 1 — Sortir de Google Sheets 🎯 (prochaine étape)
- [ ] Schéma **Postgres** : createurs, articles (stock par créateur), ventes, paniers, remises, taxes.
- [ ] Backend **FastAPI** minimal répliquant l'existant : `data`, `caisse_data`, `caisse_save`.
- [ ] Réimplémenter la logique métier en Python **à l'identique** :
      prix_client, frais (1,75 % CB / 0 espèces), prime = prix client − frais (la remise est toujours à la charge du créateur),
      commission mois × créateur figée (barème 8/12 % avant 2026, 10/15 % après ; seuils 100/250 €).
- [ ] Script de migration : import des onglets du Sheet (une table chacun) → Postgres.
- [ ] Brancher le front actuel (`index.html`) sur FastAPI au lieu d'Apps Script.
- [ ] Déploiement backend (Render / Railway / Fly.io).

## Palier 2 — Comptes & droits
- [ ] Vraie **authentification** (comptes créateurs, sessions) — remplace les mots de passe partagés.
- [ ] **Gestion des articles par créateur** + suivi de **stock** (entrées, sorties, seuils).
- [ ] Droits par page / rôle propres.

## Palier 3 — Vente en ligne
- [ ] Catalogue public + panier client.
- [ ] **Paiement Stripe** (probablement **Stripe Connect** pour reverser aux créateurs).

## ⚠️ Points en attente / à valider avant d'avancer
- **Juridique (bloquant pour le Palier 3)** : encaisser pour le compte de tiers (primes créateurs)
  est réglementé en France (statut proche marketplace / agent de paiement). À faire valider par un
  professionnel AVANT de brancher les paiements réels. (Claude n'est ni juriste ni conseiller financier.)
- Décision **maison vs Shopify/Woo/Medusa** selon l'ambition e-commerce réelle.
- `id_panier` vide pour les ventes antérieures à la caisse web (historique) : métriques « par panier » partielles sur cette période.

## Réflexe de maintenance
- Le code fait référence vit sur **GitHub**. Si le connecteur GitHub est actif, Claude lit
  directement la dernière version → plus de fichiers périmés à re-déposer à la main.

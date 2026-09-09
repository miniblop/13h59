# Dashboard 13H59 SHOP

Tableau de bord des ventes du magasin (dépôt-vente multi-vendeurs), branché sur le Google Sheet de gestion.
Site statique hébergé sur GitHub Pages.

## Accès

- **Onglet Vendeur** : suivi d'un vendeur (KPIs, graphe, récap par période, détail des ventes). Mot de passe *vendeur*.
- **Onglet Gestion** : vue multi-vendeurs (tableau transposé par période avec Δ%, groupement, filtres, monitoring, export CSV). Mot de passe *gestion*.

Les mots de passe ne sont **pas** dans le code : ils sont vérifiés côté serveur par un backend Google Apps Script.

## Architecture

- `index.html` — tout le front (HTML + CSS + JS, sans dépendance externe).
  - `DEMO_MODE = true` : données de démonstration, aucun backend requis (utile pour tester le rendu).
  - `DEMO_MODE = false` + `WEB_APP_URL` renseignée : données réelles via le backend.
- Backend : Web App Apps Script (`Api.gs`, dans le projet Apps Script du Google Sheet), déployé en « Exécuter en tant que moi / Accès : tout le monde ». Le Sheet peut donc rester **privé**.

## Mettre à jour le site

Modifier `index.html`, puis committer / pousser (ou *Add file ▸ Upload files* sur GitHub). GitHub Pages se met à jour tout seul en ~1 min.

## Calculs

La logique métier (barème de commission avant/après 01/01/2026, prime nette, ratios par panier) est isolée dans les modules `CALC` et `GEST_METRIQUES` en haut du `<script>` de `index.html`.

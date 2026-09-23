"""
Retrouve les ventes supprimées de l'onglet `ventes` en comparant les révisions
exportées par apps-script/Revisions.js (fichiers rev_<date>_<id>_<auteur>.xlsx).

Usage :  python tools/analyse_revisions.py <dossier_des_xlsx>
Dépendance : pip install openpyxl

Pour chaque ligne qui existe dans une révision puis disparaît dans la suivante,
le script indique quand, par qui, et si son contenu était un doublon probable
(même créateur / référence / prix / remise / paiement qu'une ligne d'un panier
voisin enregistré à moins de 10 minutes).
"""
import sys
from collections import defaultdict
from datetime import timedelta
from pathlib import Path

import openpyxl

FENETRE = timedelta(minutes=10)


def texte(v):
    """Même représentation quel que soit le typage de l'export (145 vs 145.0)."""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v if v is not None else "").strip()


def nombre(v):
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None


def lire_ventes(path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)["ventes"]
    it = ws.iter_rows(values_only=True)
    entetes = [str(h).strip().lower() if h is not None else "" for h in next(it)]
    idx = {h: i for i, h in enumerate(entetes) if h}
    lignes = {}
    for r in it:
        v = r[idx["numero_vente"]] if "numero_vente" in idx else None
        if v is None or r[idx["date"]] is None:
            continue
        g = lambda k: r[idx[k]] if k in idx else None
        lignes[int(v)] = {
            "vente": int(v),
            "panier": nombre(g("numero_panier")),
            "date": g("date"),
            "createur": str(g("vendeur") or "").strip(),      # colonne du Sheet, nom historique
            "reference": texte(g("reference_produit")),
            "remise": str(g("type_de_remise") or "").strip(),
            "paiement": str(g("type_de_paiement") or "").strip(),
            "prix": nombre(g("prix")),
            "prix_client": g("prix_client"),
        }
    return lignes


def contenu(l):
    return (l["createur"].lower(), l["reference"].lower(), l["prix"], l["remise"], l["paiement"])


def cle(l):
    """Identité d'une ligne : numéro + contenu (un numéro peut être réutilisé
    si le dernier panier a été supprimé avant la vente suivante)."""
    return (l["vente"],) + contenu(l)


def main(dossier):
    fichiers = sorted(Path(dossier).glob("rev_*.xlsx"))
    if len(fichiers) < 2:
        sys.exit("Il faut au moins deux révisions.")
    print(f"{len(fichiers)} révisions de {fichiers[0].name[4:23]} à {fichiers[-1].name[4:23]}\n")

    precedent, nom_prec = None, None
    supprimees = []                                  # (fichier, ligne, présentes_avant)
    for f in fichiers:
        try:
            cur = lire_ventes(f)
        except Exception as e:                       # révision illisible : on la saute
            print(f"  ! {f.name} ignorée ({e})")
            continue
        if precedent is not None:
            cles_cur = {cle(l) for l in cur.values()}
            for l in precedent.values():
                if cle(l) not in cles_cur:
                    supprimees.append((f.name, l, precedent))
        precedent, nom_prec = cur, f.name

    if not supprimees:
        print("Aucune ligne de `ventes` supprimée entre ces révisions.")
        return

    par_rev = defaultdict(list)
    for nom, l, avant in supprimees:
        par_rev[nom].append((l, avant))

    total_dbl = 0
    for nom, items in par_rev.items():
        _, date, rid, *auteur = nom[:-5].split("_")
        print(f"=== Révision {date}  par {' '.join(auteur) or '?'} : {len(items)} ligne(s) supprimée(s)")
        for l, avant in sorted(items, key=lambda x: x[0]["vente"]):
            jumelles = [
                o for o in avant.values()
                if o["vente"] != l["vente"] and o["panier"] != l["panier"]
                and contenu(o) == contenu(l)
                and hasattr(o["date"], "year") and hasattr(l["date"], "year")
                and abs(o["date"] - l["date"]) <= FENETRE
            ]
            tag = ""
            if jumelles:
                total_dbl += 1
                j = jumelles[0]
                ecart = abs((l["date"] - j["date"]).total_seconds())
                tag = f"  ← DOUBLON de la vente {j['vente']} (panier {j['panier']}, {ecart:.0f} s d'écart)"
            print(f"  vente {l['vente']} panier {l['panier']} {l['date']}  {l['createur']} | "
                  f"{l['reference']} | {l['prix']} € {l['paiement']}{tag}")
        print()
    print(f"Total : {len(supprimees)} ligne(s) supprimée(s), dont {total_dbl} doublon(s) probable(s).")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "revisions")

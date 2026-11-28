# Boutique de modèles Excel / Google Sheets

Objectif : vendre des modèles de tableurs sur Etsy et ajouter des produits régulièrement.
Chaque produit en ligne continue de se vendre, donc les ventes s'additionnent au fil des mois.

## Produits

| Produit | Fichier à vendre | Prix conseillé | État |
|---|---|---|---|
| Planificateur de budget 2027 | `budget/Budget_2027.xlsx` | 5,90 € | Prêt |
| Menus de la semaine & liste de courses | `menus/Menus_et_courses.xlsx` | 4,90 € | Prêt |
| Objectifs d'épargne & défi 52 semaines | `epargne/Epargne_defi_52_semaines.xlsx` | 3,90 € | Prêt |
| Comptes à deux | `comptes/Comptes_a_deux.xlsx` | 4,90 € | Prêt |
| Pack des 4 modèles | les 4 fichiers ci-dessus | 12,90 € | Prêt |

- Les textes des annonces (titres, descriptions, mots-clés) sont dans [`annonces.md`](annonces.md).
- Chaque dossier contient aussi une version `_exemple` remplie, pour les photos de l'annonce.
- `commun.py` : couleurs, police et mode d'emploi partagés, pour que tous les produits se ressemblent.
- Pour régénérer un produit : `python3 <dossier>/build_<produit>.py` (le budget accepte une année : `python3 budget/build_budget.py 2028`).

## Prochaines étapes

1. Ouvrir la boutique Etsy et publier les annonces (avant janvier, la meilleure période).
2. Faire les images des annonces (captures des versions exemple).
3. Idées de produits suivants : organisation de mariage, suivi des dettes, planning ménage,
   suivi d'habitudes, puis versions en anglais pour toucher plus d'acheteurs.

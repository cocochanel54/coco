# Coco Tableur

Petite appli web pour ouvrir, modifier et réenregistrer des fichiers Excel, directement dans le navigateur.

Ouvre `index.html` dans un navigateur, rien à installer.

- Ouvrir un `.xlsx`, `.xls`, `.csv` ou `.ods` (toutes les feuilles)
- Modifier les cases, ajouter ou supprimer des lignes et des colonnes, renommer les en-têtes
- Rechercher et trier
- Totaux et moyennes calculés automatiquement pour les colonnes de nombres
- Télécharger le résultat en `.xlsx`
- Sauvegarde automatique dans le navigateur

# Journée agent (`journee-agent/`)

Appli de saisie pour les techniciens, faite à partir du classeur `Journee_agent.xlsm`.

- Le technicien choisit son nom une fois, puis pour chaque intervention : la date, le projet, la prestation (seulement si le projet en a plusieurs) et la quantité. Le prix vient du BPU.
- Mêmes contrôles que l'onglet « Justificatifs » : 2 projets maximum par jour, prix sur devis à saisir, alerte week-end ou jour férié.
- Onglet « Équipe » : production et classement des techniciens par jour, semaine ou mois, et téléchargement d'un Excel au format des colonnes A à L de « Justificatifs ».
- **Pour les techniciens, sans compte** : version Google Sheets + Apps Script dans `journee-agent/apps-script/`. Les techniciens ne voient aucun prix (calculés côté Google uniquement). Les saisies arrivent dans un Google Sheet, les prix et les techniciens se modifient dans ses onglets, et sans réseau la saisie attend sur le téléphone. La macro `journee-agent/excel/Saisie_macros.bas` synchronise automatiquement ces saisies dans l'onglet « Justificatifs » du classeur Excel. Installation : [`journee-agent/INSTALLATION.md`](journee-agent/INSTALLATION.md).
- La même page fonctionne aussi en artifact claude.ai (saisies partagées entre comptes claude.ai) ou ouverte directement (saisies gardées sur l'appareil).

Les sources sont `app.src.html`, `apps-script/Code.src.gs` et `data.json`. Après modification, lancer `python3 journee-agent/build.py` pour régénérer `index.html`, `apps-script/Index.html` et `apps-script/Code.gs`.

# Devis BPU Axione (`devis-axione/`)

`build_devis.py` fabrique `Devis_BPU_Axione_RHD57.xlsx`, au format du devis Circet (modèle `Exemple_devis_LOSANGE.xlsx`) pour le client Axione, avec les articles des deux BPU Moselle Fibre RHD57 (Maintenance et VDR Optique, 116 articles + majorations MC-07 / MC-08).

- On choisit l'article dans une liste (code · désignation) et on saisit la quantité : unité, P.U.H.T, montant et sous-total HT se calculent.
- Articles « Sur devis » ou hors BPU : le P.U.H.T s'affiche en jaune, on tape le prix dans la case.
- Majoration MC-07 / MC-08 : ligne placée sous l'intervention, P.U.H.T de la ligne au-dessus × (coef − 1).

```
python3 devis-axione/build_devis.py <BPU Maintenance.xlsx> <BPU VDR Optique.xlsx> <Exemple_devis_LOSANGE.xlsx> Devis_BPU_Axione_RHD57.xlsx
```

Le classeur produit (logo, signature, prix) n'est pas versionné : ce dépôt est public.

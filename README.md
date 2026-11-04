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

# Devis Circet (`devis/`)

Devis au format du devis Circet (modèle `Exemple_devis_LOSANGE.xlsx` : en-tête, tableau, logo, signature), avec les articles et les prix du BPU du client. Partie commune : `modele_circet.py`.

- On choisit l'article dans une liste (code · désignation) et on saisit la quantité : unité, P.U.H.T, montant et sous-total HT se calculent.
- Articles « Sur devis » ou hors BPU : le P.U.H.T s'affiche en jaune, on tape le prix dans la case.
- Majorations : ligne placée sous l'intervention, P.U.H.T de la ligne au-dessus × (coef − 1).

| Client | Script | BPU | Majorations |
|---|---|---|---|
| Axione, à l'attention de Nicolas FRUTOS | `axione.py` | Moselle Fibre RHD57, Maintenance + VDR Optique (116 articles) | MC-07, MC-08 |
| Celeste | `celeste.py` | BPU FO Celeste, prix v3 du 15/05/2025 (62 articles) | nuit +30 % |

```
python3 devis/axione.py <BPU Maintenance.xlsx> <BPU VDR Optique.xlsx> <Exemple_devis_LOSANGE.xlsx> Devis_BPU_Axione_RHD57.xlsx
python3 devis/celeste.py <BPU_FO_CELESTE.xlsx> <Exemple_devis_LOSANGE.xlsx> Devis_BPU_Celeste.xlsx
```

Le BPU Celeste arrive en `.xls` : l'enregistrer d'abord en `.xlsx`. Les classeurs produits (logo, signature, prix) ne sont pas versionnés : ce dépôt est public.

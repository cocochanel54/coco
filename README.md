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

# Feuille de prod (`feuille-prod/`)

La production des techniciens comptée en journées et non plus en euros, pour qu'un technicien ne soit plus jugé sur le marché où il est envoyé. Fichier : [`feuille-prod/Feuille_prod.xlsx`](feuille-prod/Feuille_prod.xlsx).

- Onglet « Règles » : combien un technicien fait dans une journée normale sur chaque marché (4 tickets LOSANGE ou ALTITUDE, 2 tickets COVAGE ou IELO, 2 SAV CELESTE, 4 City Fast, 7 heures AXIONE). Une unité vaut alors 1/ce nombre de journée. Les autres prestations COVAGE et les malfaçons sont comptées au prix, par rapport à une journée de tickets du même marché.
- Onglet « Saisies » : on y colle les lignes de l'onglet Saisies du Google Sheet de Journée agent (mêmes colonnes A à N). Les lignes à vérifier sont signalées en colonne W.
- Chantier sur plusieurs jours (câble tiré avant l'activation, forfait payé à l'activation) : les jours « Chantier en cours » et la prestation finale portent le même n° de référence, et la valeur de la prestation finale est partagée en parts égales entre tous ces jours, même avec plusieurs techniciens.
- Onglet « Production » : pour les dates choisies, jours travaillés, journées produites, taux (100 % = une journée normale par jour travaillé, congés exclus), classement, et le CA en euros à côté. Le tableau par marché montre ce que rapporte une journée de travail sur chaque marché.
- Le fichier est généré par `python3 feuille-prod/build_feuille_prod.py` à partir de `journee-agent/data.json`, puis recalculé (LibreOffice) pour que les valeurs s'affichent partout.

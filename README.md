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

# Duel des Royaumes (`duel/`)

Jeu de cartes à collectionner façon Hearthstone, en français, qui marche sur ordinateur et sur téléphone. Ouvre `duel/index.html` dans un navigateur.

- 4 héros (Mage, Guerrier, Prêtresse, Rôdeuse) avec chacun un pouvoir et un deck de départ, 50 cartes en 4 raretés (commune, rare, épique, légendaire).
- Contre l'ordinateur, ou en ligne contre un ami avec un code de partie à 4 caractères (sur claude.ai par le salon partagé de la page, ailleurs par PeerJS).
- Boutique : l'or gagné en jouant achète des orbes de 5 cartes (Orbe du Royaume, Orbe épique, Orbe légendaire avec une légendaire garantie) et des dos de cartes ; cartes premium holographiques ; collection et éditeur de decks.
- Animations : cartes en 3D qui suivent le doigt, reflets holographiques, attaques avec impact, sorts lancés, entrée des légendaires.

Les sources sont `duel/jeu.html` (aussi la version publiée sur claude.ai) et les images de `duel/art/`. Après modification, lancer `python3 duel/build.py` pour régénérer `duel/index.html`.

Illustrations : composées pour le jeu par `duel/outils/illustrations.py` (décors peints générés par programme) à partir des Fluent Emoji 3D de Microsoft (© Microsoft Corporation, licence MIT).

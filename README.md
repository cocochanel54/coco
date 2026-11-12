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

# Tickets Praxedo (`tickets-praxedo/`)

Analyse de l'export « Interventions » de Praxedo, surtout les tickets non planifiés :

```
python3 tickets-praxedo/analyse_tickets.py export_interventions.csv
```

- **Échéances GTR** calculées depuis la création du ticket : niveau 4 = 10 jours ouvrés, niveau 3 = 5 jours ouvrés, niveau 2 hors RDV client = 1 jour ouvré, niveau 1 = 4 h, [4H HO] / [8H HO] = 4 ou 8 heures ouvrées (8 h-18 h). L'échéance « À faire avant » de Praxedo ne sert que pour les tickets sans niveau.
- **Nouveaux** : les tickets créés depuis moins de 24 h sont en violet sur la carte, reliés à l'intervention déjà planifiée la plus proche pour les programmer dans la foulée.
- **Justifiés / en attente** : un ticket dont une note Praxedo (colonne Notes, Commentaire, Remarque… de l'export) ou la description dit « GTR justifiée », « poteau à remplacer », « attente Orange », « attente nacelle »… passe en gris et ne compte plus comme en retard. Les mots reconnus sont dans `RE_JUSTIF`, en haut du script. Chaque ticket justifié a un suivi : qui relancer (Orange, Enedis, mairie, syndic…), quoi vérifier, et une relance tous les 10 jours ouvrés depuis la dernière date écrite dans la note (« relance Orange 10/11 ») ou depuis la création. Une date de travaux annoncée (« remplacement prévu le 25/11 ») devient la date où vérifier si les travaux sont faits. Les relances dues ont un contour orange sur la carte.
- **À planifier** : tickets « Qualifiée » et « Pré-planifiée », triés par priorité (échéance GTR dépassée ou proche, client coupé, déjà déplanifié…).
- **Tickets proches** : même immeuble, même boîte, même PTO ou même référence client, « suite à » un autre ticket, tickets à quelques mètres, tournées possibles, et l'intervention déjà planifiée la plus proche où caser un ticket.
- **Incohérences** : échéance dépassée ou qui ne colle pas à la GTR, déplanifiés, planifiées après l'échéance ou un jour férié, interventions passées pas clôturées, adresses fausses ou vagues…
- Résultat : un Excel (`Analyse_tickets_AAAA-MM-JJ.xlsx`) et une carte (`Carte_tickets_AAAA-MM-JJ.html`) à côté de l'export. Seules la rue, le code postal et la ville sont envoyés au service d'adresses de l'IGN pour placer les tickets sur la carte.
- **Carte en direct** (`tickets-praxedo/apps-script/`) : Google Sheet + Apps Script, comme Journée agent. Chaque export Praxedo importé met la carte à jour (page web) : un ticket au statut de fin (Terminée, Réalisée…) ou sorti du filtre depuis l'import précédent passe en vert. Un mail « [FIN INTERVENTION] » qui arrive dans Gmail (transfert automatique ou à la main) fait de même dans les 5 minutes, avec le lien du rapport. Un mail « [CREATION INTERVENTION] » ajoute le ticket (violet, avec sa GTR) avant même le prochain export, et un mail avec l'export Praxedo en pièce jointe (.csv) est importé tout seul. Les tickets viennent de l'export Praxedo importé par le menu « Carte tickets ». Mêmes règles que l'analyse Python (GTR, nouveaux, justifiés, relances). Installation : [`tickets-praxedo/apps-script/INSTALLATION.md`](tickets-praxedo/apps-script/INSTALLATION.md).

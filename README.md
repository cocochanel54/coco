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

# Attachement sous-traitants (`attachement-sous-traitant/`)

**`Attachement_sous-traitant.xlsx`** : le fichier à remplir chaque mois, un par sous-traitant, au format du suivi maintenance (titre rouge, colonnes N° à CERFA/ASSURANCE, « NB Ticket » en vert, cases jaunes, TOTAL HT). Il fonctionne sans macro.

- **Chaque mois** : efface l'onglet **Export**, colle en A1 l'export des interventions du sous-traitant (avec sa ligne de titres). L'onglet Attachement est protégé (sans mot de passe) pour qu'on n'y colle pas l'export par erreur, ce qui effacerait ses formules ; une alerte rouge s'affiche si c'est arrivé quand même. L'onglet **Attachement** se remplit et calcule tout seul. Le nom du sous-traitant se tape une fois dans **Réglages**.
- **Un forfait d'équipes par projet**, le projet étant trouvé automatiquement à partir de la colonne Agence, par un mot qu'elle contient : NANCY → P2200154 ; CIRCET 54 → P2100474 ; GLAUCOS ou RESOPTIC → P1800267. Espaces, parenthèses, tirets et majuscules ne comptent pas (« Circet Nancy RES », « CIRCET NANCY (RES) » et « CIRCET - NANCY » donnent tous P2200154). Une agence non reconnue est signalée en rouge (« Non classés ») : il suffit d'ajouter un mot dans le tableau « Agences de l'export » des Réglages.
- Pour chaque projet : **nombre d'équipes** = tickets retenus du projet ÷ jours ouvrés du mois ÷ 4 (arrondi au centième) ; **montant forfait** = équipes × 12 000 €. Base et tickets demandés par jour se règlent par projet.
- **Forfait** : tickets classiques et travaux commencés entre 8h et 18h, du lundi au vendredi hors jours fériés. **Hors forfait** (avant 8h, après 18h, week-end, jour férié) : sur devis, la plus-value se saisit à la main colonne T et s'ajoute au TOTAL HT du projet.
- **Bonif ticket** : un ticket au forfait de 2h30 ou plus (Clôturée − Date de début) compte un ticket de plus par tranche de 2h30 (3h = 2 tickets, 5h = 3 tickets). La bonif est proposée et la case BONIF TICKET (à côté de NOMBRE DE TICKETS) reste orange tant qu'elle n'est pas vérifiée : on y saisit le nombre de tickets en plus. Tickets retenus = nombre de tickets + bonif, et c'est eux qui comptent pour le nombre d'équipes.
- Les colonnes sont retrouvées dans l'export par leur titre (ligne 12 de l'Attachement), sinon par d'autres titres acceptés (« *début* », « *debut* », « *clôtur* »… modifiables dans Réglages). L'onglet Réglages montre quelle colonne de l'export est utilisée. Si la date de début ou de clôture manque, une alerte rouge en haut de l'attachement l'explique (sans date, pas de mois, donc jours ouvrés et équipes à 0).
- Un même N° de ticket n'est compté qu'une fois ; un ticket d'un autre mois n'est pas compté (affiché en gris). Mois et jours ouvrés sont lus dans l'export et modifiables (jours fériés jusqu'en 2032).
- Impression : paysage, une page de large, limitée aux tickets présents. 300 tickets maximum par mois.
- Le fichier est généré par `build_modele.py` (projets en tête du script ; `python3 attachement-sous-traitant/build_modele.py "CET" "LOR"` crée un fichier par sous-traitant).

`index.html` : première version, une page qui calcule dans le navigateur à partir de l'export et télécharge un attachement Excel. Elle ne fait pas de forfait par projet, et les tickets hors forfait y sont payés à un prix fixe réglable, pas sur devis.

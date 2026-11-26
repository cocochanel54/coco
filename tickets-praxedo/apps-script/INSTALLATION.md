# Carte des tickets Praxedo : mode opératoire

La carte se met à jour à chaque export Praxedo que tu importes et à chaque mail de création ou de fin d'intervention : nouveaux tickets, plannings, interventions terminées en vert. Ça tourne dans ton compte Google, comme Journée agent, même quand ton PC est éteint.

## Tu as déjà la carte : passer à la nouvelle version (10 minutes)

1. Ouvre ton Sheet **Carte tickets**, puis **Extensions → Apps Script** (ou **script.google.com → Mes projets**, le projet dont la page **Exécutions** montre « lireMails » toutes les 5 minutes).
2. Clique **Code.gs**, **Ctrl+A**, colle le nouveau **Code**. Pareil pour **Analyse** et **Carte**. **Import** ne change pas.
3. **Ctrl+S**. Les envois automatiques (rapports, mails du matin) prennent le nouveau code dès cet enregistrement.
4. **Déployer → Gérer les déploiements** → crayon → **Version : Nouvelle version** → **Déployer**. Le lien de la carte ne change pas. (Pas « Nouveau déploiement » : ça créerait un autre lien.)
5. Dans le Sheet, **F5**, puis **Carte tickets → Mettre à jour les onglets de contrôle**. De nouveaux onglets apparaissent : **Qualité**, **Graphiques**, **Envoi du matin**, **Annulés**, **Corrélés** (et **Backlog** après le premier rapport du soir).
6. **Onglet Accès**, ligne **Annulations et corrélations** : écris en **colonne D** l'adresse mail de la personne qui gère les annulations et les corrélations (et la tienne si tu veux le faire depuis Outlook, séparée par une virgule). Donne-lui le code de cette ligne.
7. **Onglet Envoi du matin** : les adresses qui recevront l'Excel et la carte chaque jour ouvré vers 7 h (une par ligne, en colonne A). **oui** en colonne B : elle reçoit aussi le rapport du soir et les bilans.
8. Pour voir les mails tout de suite : **Carte tickets → Rapports par mail** et **Carte tickets → Envoi du matin : envoyer maintenant**. Ensuite tout part seul : rapport chaque soir vers 17 h, bilan de la semaine le dimanche, bilan du mois le 1er.
9. Facultatif : **l'analyse des photos des rapports par l'IA** (OK ou NO OK — à vérifier). Elle demande une clé payante : voir « Analyse des photos par l'IA » plus bas. Sans clé, rien ne change.

Les rapports des 30 derniers jours sont relus une fois dans les heures qui suivent (câbles, positions au PM, références client), pour l'analyse qualité.

# Première installation (une seule fois, environ 15 minutes, sur ordinateur)

## 1. Régler le filtre Praxedo

Ajoute à ton filtre d'export :

- **La colonne « Créée »** (date de création), **au moins une fois**. Les GTR se comptent depuis cette date. La carte garde ensuite les dates d'un import à l'autre, et les nouveaux tickets prennent la leur dans le mail de création. Sans elle, la date est estimée d'après le n° du ticket (BB-20261102T… = 02/11/2026) : c'est juste 9 fois sur 10, et un bandeau orange le signale sur la carte.
- **« Desc. site »** si Praxedo la propose : elle contient les coordonnées GPS (« Lat/Long: … »), plus justes que l'adresse.
- La colonne « À faire avant » ne sert pas : la GTR est calculée avec les règles par niveau, et un ticket sans niveau n'a pas de GTR.
- **Les statuts de fin** (Terminée, Réalisée, Validée… selon ce que Praxedo propose), sur les **7 derniers jours** par exemple. Ces tickets passent en vert sur la carte. S'il existe une colonne de date de fin (« Date de fin », « Réalisée le »…), ajoute-la aussi.
- **La colonne des notes**, pour les tickets justifiés, si Praxedo la propose. Sinon, tape tes notes dans l'onglet **Notes** du Sheet (voir « Au quotidien »).

Sans les statuts de fin, la carte se débrouille quand même : un ticket présent dans l'import précédent et absent du nouveau passe en vert « sorti de l'export : terminé ou annulé » jusqu'au soir.

**Les mails « [FIN INTERVENTION] »** : dès qu'ils arrivent dans ton Gmail (transfert automatique quand la DSI l'autorise, ou à la main, un mail à la fois et pas plusieurs en pièce jointe), le script les lit toutes les 5 minutes et le ticket passe en vert avec le lien du rapport. Rien d'autre à régler : le n° du ticket est lu dans l'objet, même avec « TR: » devant.

## 2. Créer le Google Sheet et coller le code

1. Va sur <https://sheets.new> et nomme le fichier **Carte tickets**.
2. Menu **Extensions → Apps Script**.
3. Fichier **Code.gs** : efface tout, colle le contenu de [`Code.gs`](Code.gs).
4. Clique **+** à côté de « Fichiers » → **HTML**, nomme-le **Carte** (sans `.html`), efface tout, colle [`Carte.html`](Carte.html).
5. Pareil pour **Analyse** ([`Analyse.html`](Analyse.html)) et **Import** ([`Import.html`](Import.html)).
6. Roue dentée **Paramètres du projet** : fuseau horaire **(GMT+01:00) Paris**.
7. Enregistre (Ctrl+S).

## 3. Lancer l'installation

1. En haut de l'éditeur, choisis la fonction **installer** puis **Exécuter**.
2. Google demande une autorisation : **Examiner les autorisations** → ton compte → **Paramètres avancés** → **Accéder à Carte tickets** → **Autoriser**. Le script lit les mails dont l'objet contient « FIN INTERVENTION » (ceux que tu transfères), écrit dans ce Sheet et place les adresses sur la carte.
3. Les onglets **Tickets**, **Fins d'intervention**, **Créations**, **Sortis de l'export**, **Notes**, **Accès** et **Adresses** sont créés (les onglets de contrôle et de bilan apparaissent au premier import), et Gmail sera lu toutes les 5 minutes.

## 4. Importer un export Praxedo

1. Recharge le Sheet : un menu **Carte tickets** apparaît.
2. **Carte tickets → Importer un export Praxedo…** → choisis le CSV exporté de Praxedo (ton filtre habituel, avec la colonne des notes si tu l'as ajoutée).
3. La première fois, compte une minute : les adresses sont placées sur la carte.

Refais un import aussi souvent que tu veux (matin, midi, soir…) : nouveaux tickets, plannings et fins d'intervention se mettent à jour à chaque fois.

**Mails de création** : chaque mail « [CREATION INTERVENTION] » qui arrive dans Gmail ajoute le ticket à la carte dans les 5 minutes, en violet « nouveau », avec sa GTR (lue dans la description : [04h HO], [Niveau 3]…) et l'adresse du client final. Le prochain export le remplace par sa ligne complète ; s'il n'y est pas, il disparaît (sorti du filtre).

**Import automatique par mail** : tout mail qui arrive dans ton Gmail avec l'export Praxedo en pièce jointe (fichier .csv) est importé tout seul dans les 5 minutes. Le script prend le plus récent, ignore les autres CSV (ceux qui n'ont pas de colonne « N° ») et n'importe jamais deux fois le même mail.

- Si Praxedo peut **envoyer l'export par mail à heure fixe** (export programmé ou planifié), fais-le envoyer à ton Gmail (ou à ton Outlook, avec une règle de transfert vers Gmail pour les mails qui ont un .csv) : la carte est alors 100 % automatique.
- Sinon, exporte depuis Praxedo et **envoie-toi le CSV par mail** : même depuis le téléphone, la carte se met à jour sans ouvrir le Sheet.

## 5. Mettre la carte en ligne

1. Dans l'éditeur : **Déployer → Nouveau déploiement** → roue dentée → **Application Web**.
2. **Exécuter en tant que : Moi**. **Qui a accès : Tout le monde**.
3. **Déployer**, puis copie l'**URL de l'application Web** (elle finit par `/exec`). Ouvre-la sur ordinateur ou téléphone.

La page se recharge toute seule toutes les 2 minutes.

## 6. Donner l'accès aux techniciens et à la conduite d'activité

La carte demande un **code** à la première ouverture (gardé ensuite sur le téléphone ou le PC). Il y en a trois, dans l'onglet **Accès** du Sheet (ou menu **Carte tickets → Codes d'accès à la carte**) :

- **Conduite d'activité** : la carte, les chiffres clés, le **Bilan en graphiques** et les **Listes de contrôle** (Incohérences, Tickets liés, PM chargés, Suites à prévoir, Clôtures, Boîtes, Villes, Qualité, Annulés, Corrélés ; un clic sur une ligne montre le ticket sur la carte).
- **Techniciens** : la carte (tickets, priorités, nouveaux, terminés). Chacun choisit son nom dans **Mes interventions** : ses interventions ont un contour noir, et la liste montre aussi les tickets à côté des siennes.
- **Annulations et corrélations** : tout ce que voit la conduite d'activité, plus des boutons pour modifier (voir « Ce qu'elle peut modifier depuis la carte » plus bas). À ne donner qu'à la personne qui gère la planification, les annulations et les corrélations.

Envoie l'adresse `/exec` et le bon code par SMS ou WhatsApp. Personne ne peut rien modifier depuis la carte, sauf la personne qui a le code **Annulations et corrélations** (troisième ligne de l'onglet Accès). L'import et les menus du Sheet ne marchent que pour toi. Pour couper l'accès à quelqu'un, change le code dans l'onglet Accès et donne le nouveau aux autres.

**Ce qu'elle peut modifier depuis la carte, avec son code :**

- **retirer un ticket annulé** dans Praxedo, ou **regrouper des tickets** sous un ticket principal (voir plus bas), et les remettre ;
- **ajouter une note** sur un ticket (fiche du ticket → « Ajouter une note ») : « GTR justifiée… », « attente Orange », « relance Orange 12/11 », « prévu le 25/11 »… Elle part dans l'onglet **Notes** et compte comme une note tapée dans le Sheet (le ticket passe en gris s'il est justifié) ;
- **noter OK ou NO OK une intervention** après avoir vérifié ses photos (le lien « rapport » ouvre le PDF), avec un commentaire : dans la fiche d'un ticket terminé (« Photos vérifiées ? »), ou dans **Listes de contrôle → Clôtures** (toutes les clôtures des 90 derniers jours) et **→ Qualité**. C'est écrit dans la colonne **Photos vérifiées** de l'onglet Fins d'intervention, avec la date. Ça marche aussi sans l'analyse par l'IA.

Elle reçoit aussi l'Excel et la carte du matin si son adresse est dans l'onglet **Envoi du matin**, et le rapport du soir et les bilans si tu écris **oui** en colonne B à côté.

**Lui donner aussi le Google Sheet** (quand Circet a débloqué Google Sheets sur son PC) :

1. Il lui faut un compte Google à son adresse : son adresse Circet si Circet est d'accord (<https://accounts.google.com/signup>, « Utiliser mon adresse e-mail actuelle »), sinon une adresse Gmail.
2. Dans le Sheet : **Partager** → son adresse → **Éditeur** → **Envoyer**.
3. Facultatif : clic droit sur l'onglet **Accès** → **Protéger la feuille** → **Définir les autorisations** → **Seulement vous**, pour que les codes ne changent pas par erreur.

Dans le Sheet, elle voit tout et peut modifier les onglets : **Notes**, **Annulés**, **Corrélés**, la colonne **Photos vérifiées** de l'onglet Fins d'intervention… Ce qu'elle y écrit compte comme si tu l'avais écrit (carte, listes et rapports dans les 5 minutes, onglets de contrôle dans l'heure).

Le menu **Carte tickets** reste à toi seul. Dis-lui de ne pas s'en servir : Google lui demanderait d'abord des autorisations sur son compte, puis le menu refuserait. Pour importer un export Praxedo, elle l'envoie en pièce jointe (.csv) à ton adresse Gmail : il est importé tout seul dans les 5 minutes. Une personne qui peut modifier le Sheet voit aussi les codes de l'onglet Accès et peut ouvrir le code du script : ne le partage qu'à elle.

## Là où Google est bloqué : la carte et l'Excel du jour

Sur les PC où la sécurité de l'entreprise bloque le lien de la carte ou Google Sheets, utilise les **fichiers du jour** :

1. Dans le Sheet : **Carte tickets → Fichiers du jour : carte + Excel (pour Teams)…**
2. Clique **Télécharger Carte_du_jour_….html** et **Télécharger Tickets_du_jour_….xlsx**, ou **M'envoyer les deux par mail** : ils arrivent en pièces jointes dans ton Gmail.
3. Dépose-les dans Teams. Tes collègues les **téléchargent puis les ouvrent** : la carte dans leur navigateur, l'Excel dans Excel.

- **La carte** : la carte complète (zoom, fiches des tickets, couleurs, listes de contrôle).
- **L'Excel** : un onglet **Résumé** avec les chiffres du jour, **Graphiques** (clôtures par semaine, communes et boîtes où on intervient le plus, comme le bilan de la carte ; leurs chiffres sont dans l'onglet « Données graphiques »), **À planifier** (par priorité, avec la GTR, l'intervention planifiée à côté et un lien **Google Maps** vers chaque ticket), **Justifiés** (qui relancer et quand), puis les mêmes onglets que le Sheet : Suites à prévoir, Incohérences, Tickets liés, PM chargés, Clôtures, Stats boîtes, villes et semaines. Couleurs, filtres et liens « rapport » cliquables compris.

Le Résumé de l'Excel donne aussi le lien de la carte en direct (il apparaît dès que quelqu'un a ouvert la carte avec la dernière version).

Pas de lien Google ni de code, mais ce sont des **photos** : les données sont celles du moment où tu as fabriqué les fichiers. Refais-les quand tu veux une version à jour. Les numéros de téléphone y sont masqués. Ils contiennent les adresses et les descriptions des tickets : ne les dépose que dans l'équipe.

## Annuler un ticket

Quand un ticket est annulé dans Praxedo, aucun mail n'arrive : la carte ne le sait qu'au prochain export. Pour le retirer tout de suite, c'est **la personne qui gère les annulations** qui s'en occupe :

1. **Une fois** : dans l'onglet **Accès**, écris son adresse mail en **colonne D** de la ligne **Annulations et corrélations**, et donne-lui le code de cette ligne.
2. **Par mail, depuis Outlook** (même quand Google est bloqué) : elle transfère à ton Gmail le mail de création du ticket en ajoutant **ANNULÉ** au début de l'objet. Elle peut aussi écrire un mail avec **ANNULÉ** dans l'objet et le n° du ticket (`BB-20261111T0296#0`) dans l'objet ou le texte. Dans les 5 minutes, le ticket est retiré. Les mails des autres adresses sont ignorés.
3. **Sur la carte, avec son code** : dans la fiche du ticket, **« Annulé dans Praxedo ? Retirer ce ticket »**. En cas d'erreur : **Listes de contrôle → Annulés → Remettre**.

Toi, tu peux aussi le faire depuis le Sheet : **Carte tickets → Annuler un ticket…**, ou écrire le n° en colonne A de l'onglet **Annulés** (pour le rétablir : efface sa ligne).

Le ticket disparaît de la carte, des listes, de l'Excel et des mails du matin, et il ne passe pas en vert « terminé » quand il sort de l'export. Avec `#0`, `#1`… seule cette fois-là est annulée (un ticket rouvert reste) ; sans `#`, toutes les fois du ticket. L'onglet **Annulés** garde qui l'a fait et quand.

## Regrouper des tickets (corrélation)

Quand plusieurs tickets ont la même cause (un câble coupé pour plusieurs clients…), on les regroupe sous un **ticket principal**. Les autres **disparaissent** de la carte, des listes, de l'Excel et des mails ; la fiche du principal les liste (« Regroupe aussi … ») et il prend l'**échéance GTR la plus proche** du groupe. Leurs fins d'intervention ne comptent pas comme des interventions de plus.

C'est la même personne que pour les annulations, avec le même code et la même adresse :

- **Par mail, depuis Outlook** : un mail avec **CORRÉLATION** dans l'objet, le **n° du ticket principal en premier**, puis les autres (dans l'objet ou dans le texte). Par exemple, objet : `CORRÉLATION BB-20261111T0296#0 BB-20261111T0297#0 BB-20261111T0301#0`. Les n° d'un mail transféré ou cité en dessous ne sont pas pris.
- **Sur la carte, avec son code** : dans la fiche du ticket principal, **« Regrouper des tickets sous celui-ci »**. Les tickets liés (même adresse, même boîte, même PTO…) sont proposés à cocher, et on peut écrire d'autres n°. En cas d'erreur : **Listes de contrôle → Corrélés → Dégrouper**.

Toi, tu peux aussi le faire depuis le Sheet : **Carte tickets → Regrouper des tickets (corrélation)…**, ou écrire dans l'onglet **Corrélés** (n° regroupé en colonne A, ticket principal en colonne B ; pour dégrouper, efface la ligne).

L'analyse qualité **propose** aussi des regroupements : des tickets ouverts sur la même boîte, le même câble ou la même ligne du PM, créés à moins de 3 jours d'écart (« À regrouper sous un seul ticket ? »), dans l'onglet Qualité et le rapport du soir.

## Envoi automatique chaque matin

Pour que l'équipe reçoive l'Excel et la carte du jour sans que tu aies à les refaire :

1. Dans le Sheet, onglet **Envoi du matin** (il apparaît après la mise à jour des onglets), écris les adresses mail en **colonne A**, une par ligne. Écris **oui** en **colonne B** pour qu'une adresse reçoive aussi le rapport du soir et les bilans de la semaine et du mois (par exemple la personne qui fait la planification).
2. Pour essayer tout de suite : **Carte tickets → Envoi du matin : envoyer maintenant**.
3. Ensuite, c'est automatique : chaque **jour ouvré vers 7 h**, deux mails partent à ces adresses, d'abord l'**Excel du jour** (avec les chiffres du jour dans le texte), puis la **carte du jour**. Si la messagerie de l'entreprise bloque la carte (fichier .html), l'Excel arrive quand même.

Rien ne part tant que la colonne A est vide ; pour arrêter, vide-la. Gmail limite les envois à environ 100 destinataires par jour : jusqu'à 40 adresses, c'est large. Ces fichiers contiennent les adresses des tickets : n'y mets que l'équipe.

## Rapports par mail : chaque soir, chaque dimanche, chaque 1er du mois

Sans rien régler, tu reçois sur ton Gmail (et les adresses avec **oui** en colonne B de l'onglet **Envoi du matin** aussi) :

- **Chaque soir vers 17 h**, les fins d'intervention reçues depuis le rapport de la veille :
  - **Backlog**, en résumé sur trois périodes : **depuis la veille 17 h**, **cette semaine** (depuis lundi) et **ce mois** (depuis le 1er). Pour chacune : combien de tickets **ouverts** (créés), combien **clôturés**, les autres sorties (sortis de l'export, annulés, regroupés), le **backlog au début** et **maintenant**, et l'**évolution**. Puis le détail des tickets ouverts (à planifier dont GTR KO et à moins de 48 h, planifiés dont GTR KO, justifiés, ouverts depuis plus de 30 jours, le plus ancien) et les niveaux des tickets créés ;
  - **Peut-être causé par nos interventions** : un nouveau défaut apparu dans les 15 jours après le passage d'un de nos techniciens au même endroit (position voisine au même PM, même câble, même boîte), plus probable si la cause notée est un brassage, une malfaçon ou une mauvaise manipulation. C'est à vérifier avec les deux rapports, pas une certitude ;
  - **Redondances** : même client (PTO ou référence client) dépanné deux fois en 30 jours, ticket rouvert après notre clôture, plusieurs interventions le même jour dans la même boîte, câble / boîte / ligne du PM touchés 3 fois ou plus en 90 jours, déplacement peut-être inutile (fin d'impact avant l'arrivée du technicien) ;
  - **Corrections à faire** : créer le ticket de suite, mettre à jour NETGEO, signaler à l'opérateur (jarretière, brassage), faire compléter un rapport (boîte, cause, position au PM ou réponse « défaut corrigé » manquante) ;
  - les autres points : GTR non respectée, rapport illisible, clôture d'un ticket inconnu, durée sur place inhabituelle, Cerfa ou plainte, fini un autre jour que prévu ;
  - **Photos à vérifier ou NO OK** : les interventions notées NO OK après vérification (carte ou Sheet), et les photos que l'IA juge NO OK si l'analyse est activée (voir plus bas) ;
  - les interventions **prévues aujourd'hui sans fin d'intervention**, puis le **détail de chaque clôture** (horaires, GTR, résultat, boîte, câble, position au PM, cause, compte-rendu). Chaque n° de ticket ouvre le rapport Praxedo.
- **Le dimanche**, en plus, le **bilan de la semaine** : le backlog jour par jour (créés, clôturés, ouverts le soir) avec le stock du début et de la fin de semaine, les chiffres comparés à la semaine d'avant, par jour, par technicien (avec les défauts peut-être liés à ses passages), les communes et les boîtes, les reprises toujours sans ticket de suite, et tous les constats de la semaine.
- **Le 1er du mois**, le **bilan du mois d'avant**, de la même façon, par semaine (backlog compris).

Chaque soir, une ligne s'ajoute à l'onglet **Backlog** du Sheet (créés, clôturés, tickets ouverts…), même s'il n'y a pas de mail : c'est l'historique qui donne le backlog au début de chaque période. Tant qu'il n'existe pas encore, le backlog du début est **estimé** avec les entrées et les sorties, et c'est écrit. L'historique et le même résumé (depuis le rapport de la veille, la semaine, le mois) sont aussi dans l'Excel du jour.

Un soir sans rien, aucun mail. Pour voir un rapport tout de suite : **Carte tickets → Rapports par mail** (envoyé à toi seul).

Pour repérer les câbles et les positions au PM, le script lit dans le rapport du technicien le n° du câble (par exemple `GNY-CA-0002728`), la **Position PM** (`T1L6C3` : tiroir 1, ligne 6, connecteur 3), la référence client et la PTO. Plus les techniciens les écrivent, plus la détection est complète. Après la mise à jour, les rapports des 30 derniers jours sont relus une fois (en quelques heures), pour y trouver ces informations.

Tous les constats des 90 derniers jours sont aussi dans l'onglet **Qualité** du Sheet et de l'Excel du jour (rouge : fort, jaune : moyen), et dans **Listes de contrôle → Qualité** sur la carte.

## Analyse des photos par l'IA (facultatif)

L'IA (Claude, d'Anthropic) regarde les photos de chaque rapport d'intervention lu et donne son avis : **OK**, l'intervention est validée, ou **NO OK — à vérifier**, avec ce qui ne va pas : fibre ou surlongueur qui traîne hors de la cassette, lovage en vrac, fibre pincée, boîtier ou cassette mal refermés, jarretières en vrac ou qui pendent au PM, étiquette absente, dommage visible, photo qui ne correspond pas à sa rubrique… Elle sait sous quelle rubrique chaque photo est rangée (« Photo de la réparation mise en place », « Photo de l'étiquette »…) et lit le compte-rendu du technicien. Le script contrôle aussi la date et le lieu imprimés sur les photos : des photos datées d'un autre jour que l'intervention, ou prises à plus de 5 km de l'adresse, passent en NO OK.

C'est un avis, pas une certitude : c'est toi qui tranches.

**L'activer (5 minutes, une seule fois) :**

1. Crée un compte sur la console d'Anthropic : <https://platform.claude.com>. Dans **Billing**, ajoute du crédit (carte bancaire, 10 € suffisent pour commencer). Dans **API keys**, **Create key**, puis copie la clé (elle commence par `sk-ant-`).
2. Dans le Sheet : **Carte tickets → Analyse des photos (IA) → Activer, voir le coût, changer la clé ou arrêter…**, colle la clé, **OK**. Le script fait un essai de quelques secondes et dit si la clé marche.

Ensuite tout est automatique : les photos de chaque nouveau rapport lu partent à l'analyse (et, au début, celles des rapports des 3 derniers jours). Les verdicts arrivent en général en moins d'une heure. La clé est gardée dans ton compte Google, à part : les personnes qui ont accès au Sheet ne la voient pas.

**Où voir les verdicts :**

- onglet **Fins d'intervention**, colonnes **Photos (IA)**, **Photos analysées le** et **Photos vérifiées** ;
- onglet **Qualité** (Sheet et Excel du jour) : les NO OK, en rouge (important) ou en jaune ;
- **rapport du soir** : la partie **Photos à vérifier ou NO OK**, et « photos OK (IA) » ou « photos NO OK (IA) » dans le détail de chaque clôture ; dans les bilans de la semaine et du mois, une colonne **Photos NO OK** par technicien ;
- sur la carte, avec le code de la conduite d'activité seulement : dans la fiche du ticket terminé.

**Après avoir regardé les photos** (le n° de ticket ouvre le rapport Praxedo) : note **OK** ou **NO OK**, dans la fiche du ticket sur la carte avec le code **Annulations et corrélations** (avec un commentaire), ou dans la colonne **Photos vérifiées** de l'onglet Fins d'intervention. OK : l'alerte disparaît. NO OK : le problème est confirmé, il reste dans l'onglet Qualité (« faire reprendre le travail ou en parler au technicien »), passe dans le rapport du soir et compte dans les bilans. On peut aussi noter NO OK une intervention que l'IA n'a pas regardée.

**Ce que ça coûte** : l'analyse utilise Claude Opus 5.5, un des modèles les plus capables d'Anthropic, et passe par lots, à moitié prix. Compte **environ 2 à 5 centimes par rapport** (une dizaine de photos), soit **environ 10 à 25 € par mois** (facturés en dollars) pour 25 rapports par jour ouvré. Le menu **Activer, voir le coût…** affiche le coût estimé du mois ; le montant exact est dans la console Anthropic (**Usage**). Garde-fou : 80 rapports par jour au plus. Pour arrêter : même menu, tape **STOP**.

**Ce qui part chez Anthropic** : les photos du rapport (avec l'adresse et les coordonnées GPS imprimées dessus), les horaires, la commune, la boîte et le PM, les réponses du rapport et le compte-rendu du technicien, **sans n° de téléphone, sans adresse mail, sans référence client ni PTO**. Ces données quittent ton compte Google : vérifie avec ton responsable que c'est accepté chez Circet et Altitude avant d'activer.

Si l'analyse tombe en panne (clé refusée, plus de crédit), le rapport du soir le dit en rouge, et l'analyse reprend toute seule une fois le problème réglé. La colonne **Photos (IA)** peut aussi valoir **Sans photo** (rapport sans photo), **À vérifier à la main** (l'IA n'a pas pu les analyser) ou **erreur 1 à 3** (nouvel essai au passage suivant).

## Les rapports d'intervention, le bilan et les graphiques

Chaque mail « [FIN INTERVENTION] » contient un lien vers le rapport Praxedo (un PDF). Le script le lit tout seul, 4 rapports au plus toutes les 5 minutes, et remplit les colonnes de l'onglet **Fins d'intervention** : heure de début et de fin, durée sur place, défaut corrigé ou non, réparation définitive ou non, dommage réseau signalé, compte-rendu du technicien, boîte où il est intervenu, PM, nature et cause du défaut, technicien, et les autres réponses du rapport dans une colonne à part. Il n'ouvre que les liens de Praxedo. « Rapport lu » vaut **oui**, **illisible**, **autre ticket**, **lien refusé** ou **erreur 1 à 3** (3 essais).

Avec ces rapports, le Sheet tient à jour (sur les 90 derniers jours) :

- **Suites à prévoir** : les clôtures en provisoire, non corrigées ou avec un dommage signalé, et leur ticket de suite quand il existe (en jaune : pas encore de ticket de suite). Un défaut non corrigé dont le reste est **chez l'opérateur** n'en fait pas partie : jarretière absente au PM à l'arrivée, ou compte-rendu qui parle de jarretière ou de brassage et de l'opérateur (OC, FAI, Free, SFR, Bouygues, Orange…).
- **Clôtures** : une ligne par intervention clôturée, avec la durée, la GTR respectée ou non (de la date de la demande à la fin de l'intervention ; la fin d'impact client, côté opérateur, est notée pour information) et le résultat (en gris : « reste côté opérateur (jarretière, brassage) », ce n'est pas une reprise pour nous). La colonne « À vérifier » signale un retour dans une boîte sous 30 jours, ou une fin d'impact notée avant l'arrivée du technicien.
- **Graphiques** : les clôtures par semaine, les communes et les boîtes où on intervient le plus (sans suite à donner en bleu, à reprendre en orange), refaits tout seuls ; leurs chiffres sont à droite, à partir de la colonne R.
- **Stats boîtes**, **Stats villes**, **Stats semaines** : où on intervient le plus, les **retours sous 30 jours** (une intervention dans une boîte déjà visitée moins de 30 jours avant), le temps passé, les provisoires, la part de GTR respectées. Sans le nombre de prises, ce sont des nombres d'interventions, pas des taux de panne.

Sur la carte (conduite d'activité) : les **chiffres clés** en haut, les **boîtes à répétition** (un anneau avec le nombre d'interventions, orange si la dernière réparation est à reprendre ; un clic donne l'historique et les rapports), le bouton **Bilan en graphiques** (clôtures par semaine, communes et boîtes où on intervient le plus, avec le choix de la période) et, dans **Listes de contrôle**, les onglets Suites à prévoir, Clôtures, Boîtes et Villes. La fiche d'un ticket donne ses **tickets liés** (un trait les relie quand la fiche est ouverte), l'intervention d'avant avec son rapport, et ce qu'on sait déjà de sa boîte.

Tout ça se remplit au fil des mails de fin : au début, il y a peu de chiffres.

## Au quotidien

**Les onglets de contrôle** du Sheet remplacent l'Excel d'analyse et se refont tout seuls (à chaque import, à chaque mail de création ou de fin, et au moins une fois par heure) :

- **Incohérences** : échéance dépassée, planifiée après l'échéance, déplanifié, date passée pas clôturée, jour férié, doublon probable, adresse fausse… colorées par gravité. Filtre par gravité ou par problème avec les flèches de l'en-tête.
- **Tickets liés** : tickets d'un même immeuble, d'une même boîte, d'une même PTO ou référence client, « suite de » un autre ticket, ou au même endroit.
- **PM chargés** : PM avec au moins 3 tickets ouverts (défaut collectif possible).

Pour un vrai fichier Excel à envoyer : **Carte tickets → Fichiers du jour** (voir plus haut). Pour refaire les onglets tout de suite : **Carte tickets → Mettre à jour les onglets de contrôle**.

**Les couleurs de la carte :**

- **Vert** : terminé aujourd'hui (mail de fin, statut de fin dans l'export, ou ticket sorti de l'export). Il disparaît de la carte le lendemain. La liste « Fins d'intervention aujourd'hui » est dans la légende.
- **Violet** : nouveau depuis moins de 24 h (export ou mail de création), avec un trait vers l'intervention déjà planifiée la plus proche.
- **Gris** : justifié / en attente. Contour orange : à relancer ou à vérifier.
- **Rouge** : GTR KO (échéance dépassée), à planifier. **Orange** : GTR dans moins de 48 h. **Jaune** : à planifier, GTR encore bonne ou sans GTR. **Bleu** : déjà planifié.
- **Contour rouge** : GTR KO sur un ticket déjà planifié (bleu) ou nouveau (violet).
- **GTR KO ou bientôt KO à proximité** : c'est ce qui ressort en premier partout. Un trait pointillé rouge ou orange relie le ticket à l'intervention planifiée la plus proche (moins de 1 km). La liste « GTR KO ou bientôt KO à côté d'une intervention planifiée » les donne avec le technicien et la distance. La fiche d'une intervention planifiée liste les urgents autour d'elle. Un technicien qui a choisi son nom voit « GTR KO ou bientôt KO à côté de mes interventions ».
- **Cercle vert de 1 km** : un technicien vient de clôturer à cet endroit (mail de fin reçu depuis moins d'1 heure). La liste « Viennent de clôturer » et la fiche du ticket donnent les tickets à planifier dans le cercle, GTR KO d'abord ; un clic sur un n° montre le ticket. La carte ne sait pas quand une intervention commence, seulement quand elle est clôturée.
- **Bandeau orange** en haut de la légende (conduite d'activité) : ce qui manque dans l'export, par exemple des dates de création estimées.
- Un mail de fin pour un ticket absent de l'export apparaît dans la liste « Mails de fin pour des tickets absents de l'export ».
- Un ticket annulé, ou regroupé sous un autre ticket, n'apparaît pas sur la carte.

**Tes notes** : écris-les dans l'onglet **Notes** (colonne A : le n° du ticket, avec ou sans « #0 » ; colonne B : la note), jamais dans l'onglet Tickets, qui est remplacé à chaque import. Elles restent d'un import à l'autre, apparaissent dans la fiche du ticket sur la carte (dans les 2 minutes), et plusieurs lignes pour un même ticket s'ajoutent.

- Une note avec « GTR justifiée… », « poteau à remplacer », « attente Orange / Enedis / nacelle / mairie… » passe le ticket en **gris** : il ne compte plus comme en retard.
- « relance Orange 12/11 » : la prochaine relance est 10 jours ouvrés après cette date.
- « remplacement prévu le 25/11 » : rien à faire avant, puis « vérifier si les travaux sont faits ».
- Toute autre note (« client absent, repasser après 17h ») est juste affichée.
- **Carte tickets → Lire les mails maintenant** : pour ne pas attendre les 5 minutes.

## Mettre à jour le code plus tard

Le Sheet et son code sont dans ton compte Google, pas sur un PC : depuis n'importe quel ordinateur, connecte-toi à ce compte (celui qui reçoit les mails FIN INTERVENTION). Le code se retrouve dans le Sheet (**Extensions → Apps Script**) ou directement sur **script.google.com → Mes projets**. Le bon projet est celui dont la page **Exécutions** montre « lireMails » toutes les 5 minutes : renomme-le **Carte tickets** pour le retrouver facilement.

Colle le nouveau code, **Ctrl+S**, puis **Déployer → Gérer les déploiements** → crayon → **Version : Nouvelle version** → **Déployer**. L'URL ne change pas. Dans le Sheet : **F5**, puis **Carte tickets → Mettre à jour les onglets de contrôle**.

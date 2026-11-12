# Installer la carte des tickets en direct (une seule fois, environ 15 minutes, sur ordinateur)

La carte se met à jour à chaque export Praxedo que tu importes : nouveaux tickets, plannings, et interventions terminées en vert. Plus tu importes souvent, plus elle est à jour.

Ça tourne dans ton compte Google, comme Journée agent.

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
3. Les onglets **Tickets**, **Fins d'intervention**, **Créations**, **Sortis de l'export**, **Notes**, **Accès** et **Adresses** sont créés, et Gmail sera lu toutes les 5 minutes.

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

La carte demande un **code** à la première ouverture (gardé ensuite sur le téléphone ou le PC). Il y en a deux, dans l'onglet **Accès** du Sheet (ou menu **Carte tickets → Codes d'accès à la carte**) :

- **Conduite d'activité** : la carte, et le bouton **Listes de contrôle** (Incohérences, Tickets liés, PM chargés, filtrables par gravité ; un clic sur une ligne montre le ticket sur la carte).
- **Techniciens** : la carte (tickets, priorités, nouveaux, terminés). Chacun choisit son nom dans **Mes interventions** : ses interventions ont un contour noir, et la liste montre aussi les tickets à côté des siennes.

Envoie l'adresse `/exec` et le bon code par SMS ou WhatsApp. Personne ne peut rien modifier depuis la carte : l'import et les mises à jour ne marchent que pour toi, depuis le Sheet. Pour couper l'accès à quelqu'un, change le code dans l'onglet Accès et donne le nouveau aux autres.

## Au quotidien

**Les onglets de contrôle** du Sheet remplacent l'Excel d'analyse et se refont tout seuls (à chaque import, à chaque mail de création ou de fin, et au moins une fois par heure) :

- **Incohérences** : échéance dépassée, planifiée après l'échéance, déplanifié, date passée pas clôturée, jour férié, doublon probable, adresse fausse… colorées par gravité. Filtre par gravité ou par problème avec les flèches de l'en-tête.
- **Tickets liés** : tickets d'un même immeuble, d'une même boîte, d'une même PTO ou référence client, « suite de » un autre ticket, ou au même endroit.
- **PM chargés** : PM avec au moins 3 tickets ouverts (défaut collectif possible).

Pour un vrai fichier Excel (à envoyer, par exemple) : **Fichier → Télécharger → Microsoft Excel (.xlsx)**. Pour les refaire tout de suite : **Carte tickets → Mettre à jour les onglets de contrôle**.


- **Vert** : terminé aujourd'hui (mail de fin, statut de fin dans l'export, ou ticket sorti de l'export). Il disparaît de la carte le lendemain. La liste « Fins d'intervention aujourd'hui » est dans la légende.
- **Violet** : nouveau depuis moins de 24 h (export ou mail de création), avec un trait vers l'intervention déjà planifiée la plus proche.
- **Gris** : justifié / en attente. Contour orange : à relancer ou à vérifier.
- **Rouge** : GTR KO (échéance dépassée), à planifier. **Orange** : GTR dans moins de 48 h. **Jaune** : à planifier, GTR encore bonne ou sans GTR. **Bleu** : déjà planifié.
- **Contour rouge** : GTR KO sur un ticket déjà planifié (bleu) ou nouveau (violet).
- **Cercle vert de 1 km** : un technicien vient de clôturer à cet endroit (mail de fin reçu depuis moins d'1 heure). La liste « Viennent de clôturer » et la fiche du ticket donnent les tickets à planifier dans le cercle, GTR KO d'abord ; un clic sur un n° montre le ticket. La carte ne sait pas quand une intervention commence, seulement quand elle est clôturée.
- **Bandeau orange** en haut de la légende (conduite d'activité) : ce qui manque dans l'export, par exemple des dates de création estimées.
- Un mail de fin pour un ticket absent de l'export apparaît dans la liste « Mails de fin pour des tickets absents de l'export ».
- Un ticket annulé n'apparaît pas sur la carte.

**Tes notes** : écris-les dans l'onglet **Notes** (colonne A : le n° du ticket, avec ou sans « #0 » ; colonne B : la note), jamais dans l'onglet Tickets, qui est remplacé à chaque import. Elles restent d'un import à l'autre, apparaissent dans la fiche du ticket sur la carte (dans les 2 minutes), et plusieurs lignes pour un même ticket s'ajoutent.

- Une note avec « GTR justifiée… », « poteau à remplacer », « attente Orange / Enedis / nacelle / mairie… » passe le ticket en **gris** : il ne compte plus comme en retard.
- « relance Orange 12/11 » : la prochaine relance est 10 jours ouvrés après cette date.
- « remplacement prévu le 25/11 » : rien à faire avant, puis « vérifier si les travaux sont faits ».
- Toute autre note (« client absent, repasser après 17h ») est juste affichée.
- **Carte tickets → Lire les mails maintenant** : pour ne pas attendre les 5 minutes.

## Mettre à jour le code plus tard

Colle le nouveau code, puis **Déployer → Gérer les déploiements** → crayon → **Version : Nouvelle version** → **Déployer**. L'URL ne change pas.

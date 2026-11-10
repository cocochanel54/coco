# Installer la carte des tickets en direct (une seule fois, environ 15 minutes, sur ordinateur)

La carte se met à jour toute seule : chaque mail « [CREATION INTERVENTION] » ajoute le ticket en violet (au centre de sa commune, en attendant le prochain export), et chaque mail « [FIN INTERVENTION] » passe le ticket en vert, avec le lien du rapport. L'adresse exacte et la GTR arrivent quand tu importes un export Praxedo.

Ça tourne dans ton compte Google, comme Journée agent. Utilise **le même compte Google** pour le Sheet et pour recevoir les mails.

## 1. Relier ta boîte Outlook à ton Gmail

Une seule règle pour les deux sortes de mails.

**Nouvel Outlook ou Outlook sur le web**
1. Engrenage **Paramètres** → **Courrier** → **Règles** → **+ Ajouter une nouvelle règle**.
2. Nom : `Praxedo vers Gmail`.
3. Condition : **L'objet inclut** → tape `FIN INTERVENTION`, puis ajoute une deuxième valeur `CREATION INTERVENTION` (sans accent, comme dans l'objet). Deux valeurs dans la même condition = l'une **ou** l'autre.
4. Action : **Transférer et rediriger** → **Rediriger vers** → ton adresse Gmail.
5. Laisse « Exécuter la règle maintenant » décoché, puis **Enregistrer**.

**Outlook classique (logiciel Windows)**
1. **Accueil** → **Règles** → **Gérer les règles et les alertes** → **Nouvelle règle…** → **Appliquer la règle aux messages que je reçois** → **Suivant**.
2. Coche **comprenant des mots spécifiques dans l'objet**, clique sur le lien bleu, ajoute `FIN INTERVENTION` puis `CREATION INTERVENTION` (**Ajouter** après chaque) → **OK** → **Suivant**.
3. Coche **le rediriger vers des personnes ou une liste publique**, clique sur le lien bleu, tape ton Gmail dans **À** → **OK** → **Suivant** → **Suivant**.
4. Nomme la règle `Praxedo vers Gmail` → **Terminer**.

**Rediriger** plutôt que transférer : le mail arrive dans Gmail tel quel (même objet, même expéditeur). Vérifie que le prochain mail de création ou de fin arrive bien dans Gmail (regarde aussi les **Spams** : le script les lit quand même, mais tu peux cliquer « Non spam »).

Si le mail n'arrive jamais dans Gmail, ou si tu reçois une erreur « transfert externe non autorisé » / « 550 5.7.520 », la DSI bloque le transfert vers l'extérieur : dis-le, une petite macro Outlook prendra le relais.

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
2. Google demande une autorisation : **Examiner les autorisations** → ton compte → **Paramètres avancés** → **Accéder à Carte tickets** → **Autoriser**. Le script lit les mails dont l'objet contient « CREATION INTERVENTION » ou « FIN INTERVENTION », écrit dans ce Sheet et place les adresses sur la carte.
3. Les onglets **Tickets**, **Mails Praxedo** et **Adresses** sont créés, et Gmail sera lu toutes les 5 minutes.

## 4. Importer un export Praxedo

1. Recharge le Sheet : un menu **Carte tickets** apparaît.
2. **Carte tickets → Importer un export Praxedo…** → choisis le CSV exporté de Praxedo (ton filtre habituel, avec la colonne des notes si tu l'as ajoutée).
3. La première fois, compte une minute : les adresses sont placées sur la carte.

Les créations et les fins d'intervention arrivent toutes seules par mail. Refais un import de temps en temps (le matin, par exemple) pour avoir l'adresse exacte, la GTR et les tickets planifiés.

## 5. Mettre la carte en ligne

1. Dans l'éditeur : **Déployer → Nouveau déploiement** → roue dentée → **Application Web**.
2. **Exécuter en tant que : Moi**. **Qui a accès : Tout le monde**.
3. **Déployer**, puis copie l'**URL de l'application Web** (elle finit par `/exec`). Ouvre-la sur ordinateur ou téléphone.

La page se recharge toute seule toutes les 2 minutes. Toute personne qui a l'adresse voit la carte (adresses et descriptions des tickets) : ne la donne qu'à l'équipe.

## Au quotidien

- **Vert** : mail de fin reçu. La liste « Fins d'intervention aujourd'hui » est dans la légende.
- **Violet** : nouveau depuis moins de 24 h, avec un trait vers l'intervention déjà planifiée la plus proche.
- **Violet à contour pointillé blanc** : ticket reçu par mail de création, pas encore dans l'export, placé au centre de sa commune. Liste « Nouveaux reçus par mail » dans la légende (si l'objet dit « - NA », la commune est inconnue : le ticket est seulement dans la liste).
- **Gris** : justifié / en attente. Contour orange : à relancer ou à vérifier.
- **Rouge / orange** : à planifier (rouge = échéance GTR dépassée ou dans moins de 48 h). **Bleu** : déjà planifié.
- Un mail de fin pour un ticket absent de l'export apparaît dans la liste « Mails de fin pour des tickets absents de l'export ».
- **Carte tickets → Lire les mails maintenant** : pour ne pas attendre les 5 minutes.

## Mettre à jour le code plus tard

Colle le nouveau code, puis **Déployer → Gérer les déploiements** → crayon → **Version : Nouvelle version** → **Déployer**. L'URL ne change pas.

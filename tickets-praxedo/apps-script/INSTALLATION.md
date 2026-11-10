# Installer la carte des tickets en direct (une seule fois, environ 15 minutes, sur ordinateur)

La carte se met à jour toute seule : chaque mail « [CREATION INTERVENTION] » ajoute le ticket en violet (au centre de sa commune, en attendant le prochain export), et chaque mail « [FIN INTERVENTION] » passe le ticket en vert, avec le lien du rapport. L'adresse exacte et la GTR arrivent quand tu importes un export Praxedo.

Ça tourne dans ton compte Google, comme Journée agent. Utilise **le même compte Google** pour le Sheet et pour recevoir les mails.

## 1. Envoyer les mails Praxedo vers ton Gmail

Chez Circet, la DSI bloque les règles qui transfèrent automatiquement vers l'extérieur : on transfère donc à la main.

1. Si tu as créé la règle « Praxedo vers Gmail », supprime-la ou désactive-la (**Paramètres → Courrier → Règles**) : sinon tu reçois un mail d'erreur à chaque fois.
2. Pour chaque mail **[CREATION INTERVENTION]** ou **[FIN INTERVENTION]** : **Transférer** → ton adresse Gmail → **Envoyer**. Rien à écrire, et laisse l'objet tel quel (le « TR: » ajouté ne gêne pas).
3. **Un mail à la fois** : si tu sélectionnes plusieurs mails et les transfères d'un coup, Outlook les met en pièces jointes et le script ne peut pas les lire.
4. Pas besoin de transférer à la seconde : la carte affiche l'heure du mail d'origine (lue dans la ligne « Envoyé : » du transfert).

Pour aller plus vite, si ton Outlook le propose : une **étape rapide** « Transférer à » avec ton Gmail déjà rempli.

Dans Gmail, regarde aussi les **Spams** la première fois : le script les lit quand même, mais tu peux cliquer « Non spam ».

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

Les créations et les fins d'intervention arrivent par les mails que tu transfères. Refais un import de temps en temps (le matin, par exemple) pour avoir l'adresse exacte, la GTR et les tickets planifiés.

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

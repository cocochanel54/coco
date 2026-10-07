# Installer Journée agent sur Google (une seule fois, environ 10 minutes, sur ordinateur)

Les techniciens saisissent sur leur téléphone, sans compte ni mot de passe. Chaque saisie arrive dans un Google Sheet à toi.

## 1. Créer le Google Sheet

1. Va sur <https://sheets.new> (connecté à ton compte Google) et nomme le fichier **Journée agent**.
2. Menu **Extensions → Apps Script**. Un éditeur s'ouvre.

## 2. Coller le code

1. Dans l'éditeur, fichier **Code.gs** : efface tout, puis colle le contenu de [`apps-script/Code.gs`](apps-script/Code.gs).
2. Clique sur **+** à côté de « Fichiers » → **HTML**, nomme-le **Index** (sans `.html`). Efface tout, puis colle le contenu de [`apps-script/Index.html`](apps-script/Index.html).
3. Roue dentée **Paramètres du projet** : fuseau horaire **(GMT+01:00) Paris**.
4. Enregistre (icône disquette ou Ctrl+S).

## 3. Créer les onglets

1. En haut de l'éditeur, choisis la fonction **installer** puis clique **Exécuter**.
2. Google demande une autorisation : **Examiner les autorisations** → ton compte → « Google n'a pas validé cette application » → **Paramètres avancés** → **Accéder à Journée agent** → **Autoriser**. (C'est ton propre script, il n'accède qu'à ce Sheet.)
3. Retourne dans le Sheet : les onglets **Saisies**, **BPU**, **Techniciens**, **Projets** et **Jours fériés** sont créés et remplis à partir de ton Excel.

## 4. Mettre l'appli en ligne

1. Dans l'éditeur : **Déployer → Nouveau déploiement** → roue dentée → **Application Web**.
2. **Exécuter en tant que : Moi**. **Qui a accès : Tout le monde**.
3. **Déployer**, puis copie l'**URL de l'application Web** (elle finit par `/exec`).

## 5. Donner l'appli aux techniciens

1. Envoie l'URL par SMS ou WhatsApp.
2. Sur le téléphone, ils l'ouvrent puis l'ajoutent à l'écran d'accueil, comme une appli :
   - **iPhone (Safari)** : bouton Partager → **Sur l'écran d'accueil**.
   - **Android (Chrome)** : menu ⋮ → **Ajouter à l'écran d'accueil**.
3. Au premier lancement, chacun choisit son nom. Le téléphone s'en souvient.

Sans réseau, la saisie reste sur le téléphone et part toute seule au retour du réseau (« en attente » s'affiche à côté de la ligne).

## Au quotidien

- **Voir les saisies** : onglet **Saisies** du Google Sheet, en direct. Les colonnes A à L sont celles de « Justificatifs ».
- **Les mettre dans ton Excel** : sélectionne les nouvelles lignes, colonnes A à L, copie, puis colle à la suite dans l'onglet « Justificatifs » de `Journee_agent.xlsm`. Tu peux aussi télécharger le Sheet : Fichier → Télécharger → Microsoft Excel.
- **Changer un prix** : onglet **BPU**, colonne « Prix unitaire ». Laisse vide pour un prix sur devis (le technicien le tape).
- **Ajouter ou retirer un technicien** : onglet **Techniciens**. Ton Excel n'a que 6 colonnes de techniciens dans « Journée agent » : un nouveau nom doit aussi y être ajouté.
- Les changements du Sheet apparaissent dans l'appli à la prochaine ouverture.

## Mettre à jour le code plus tard

Colle le nouveau code, puis **Déployer → Gérer les déploiements** → crayon → **Version : Nouvelle version** → **Déployer**. L'URL ne change pas.

## À savoir

- Toute personne qui a l'URL peut saisir : ne la donne qu'aux techniciens. L'adresse est longue et impossible à deviner.
- Si une ligne est fausse, le technicien peut la supprimer dans l'appli, ou tu la supprimes dans l'onglet Saisies (ligne entière).
- Google affiche un bandeau « Cette application a été créée par un utilisateur de Google Apps Script » en haut de l'appli : c'est normal.

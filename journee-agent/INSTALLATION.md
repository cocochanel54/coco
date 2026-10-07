# Installer Journée agent (une seule fois, environ 15 minutes, sur ordinateur)

Les techniciens saisissent sur une page web, depuis leur téléphone, sans compte ni mot de passe. Chaque saisie arrive dans un Google Sheet à toi, puis ton Excel `Journee_agent.xlsm` la récupère tout seul dans l'onglet « Justificatifs ».

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

## 5. Donner la page aux techniciens

Envoie l'URL par SMS ou WhatsApp. Ils l'ouvrent dans le navigateur du téléphone et choisissent leur nom au premier passage. (Facultatif : « Ajouter à l'écran d'accueil » pour la retrouver en un geste.)

Sans réseau, la saisie reste sur le téléphone et part toute seule au retour du réseau (« en attente » s'affiche à côté de la ligne).

## 6. Brancher ton Excel (Excel sous Windows)

1. Ouvre `Journee_agent.xlsm` et appuie sur **Alt+F11**.
2. À gauche, clic droit sur le module **Saisie_macros** → **Supprimer Saisie_macros** → **Non** (pas besoin de l'exporter).
3. Menu **Fichier → Importer un fichier…** → choisis [`excel/Saisie_macros.bas`](excel/Saisie_macros.bas). C'est ton module d'origine, avec la synchronisation en plus.
4. Ferme l'éditeur, enregistre, ferme le fichier et rouvre-le (clique **Activer le contenu**).
5. Sur la feuille « Saisie », clique le nouveau bouton **Synchroniser le web** (ou **Ctrl+Maj+S**). La première fois, colle l'URL de l'étape 4.

Ensuite, c'est automatique : à chaque ouverture du fichier, puis toutes les 5 minutes tant qu'il est ouvert, les nouvelles saisies arrivent dans « Justificatifs ». Le message sous les boutons de la feuille « Saisie » indique la dernière synchro.

- Seules les colonnes Date, Technicien, Projet, Prestation, Quantité, N° ticket et Commentaire sont remplies : le reste se calcule avec tes formules, et « Journée agent », « Classement », « Synthèse »… se mettent à jour.
- Une ligne supprimée sur le web est retirée d'Excel à la synchro suivante. Les lignes que tu tapes toi-même dans Excel ne sont jamais touchées.
- La colonne **U** de « Justificatifs » contient l'identifiant des lignes venues du web : ne la modifie pas.
- Prix sur devis : le prix tapé sur le web est écrit en orange dans PU et Montant.
- Pour changer le lien : Alt+F8 → **ChangerLienWeb**.

## Au quotidien

- **Voir les saisies** : onglet **Saisies** du Google Sheet, en direct. Les colonnes A à L sont celles de « Justificatifs ».
- **Ton Excel** se remplit tout seul (étape 6). Sur Mac, la synchro automatique ne marche pas : copie les colonnes A à L de l'onglet Saisies et colle-les à la suite dans « Justificatifs ».
- **Changer un prix** : dans l'onglet **BPU** du Google Sheet (c'est ce que voient les techniciens) et aussi dans l'onglet BPU de ton Excel (c'est ce qui calcule les montants). Garde les mêmes libellés des deux côtés. Laisse vide pour un prix sur devis.
- **Ajouter ou retirer un technicien** : onglet **Techniciens**. Ton Excel n'a que 6 colonnes de techniciens dans « Journée agent » : un nouveau nom doit aussi y être ajouté.
- Les changements du Sheet apparaissent dans l'appli à la prochaine ouverture.

## Mettre à jour le code plus tard

Colle le nouveau code, puis **Déployer → Gérer les déploiements** → crayon → **Version : Nouvelle version** → **Déployer**. L'URL ne change pas.

## À savoir

- Toute personne qui a l'URL peut saisir : ne la donne qu'aux techniciens. L'adresse est longue et impossible à deviner.
- Si une ligne est fausse, le technicien peut la supprimer dans l'appli, ou tu la supprimes dans l'onglet Saisies (ligne entière).
- Google affiche un bandeau « Cette application a été créée par un utilisateur de Google Apps Script » en haut de l'appli : c'est normal.

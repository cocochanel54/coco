# Installer la carte des tickets en direct (une seule fois, environ 15 minutes, sur ordinateur)

La carte se met à jour à chaque export Praxedo que tu importes : nouveaux tickets, plannings, et interventions terminées en vert. Plus tu importes souvent, plus elle est à jour.

Ça tourne dans ton compte Google, comme Journée agent.

## 1. Régler le filtre Praxedo

Ajoute à ton filtre d'export :

- **Les statuts de fin** (Terminée, Réalisée, Validée… selon ce que Praxedo propose), sur les **7 derniers jours** par exemple. Ces tickets passent en vert sur la carte. S'il existe une colonne de date de fin (« Date de fin », « Réalisée le »…), ajoute-la aussi.
- **La colonne des notes**, pour les tickets justifiés.

Sans les statuts de fin, la carte se débrouille quand même : un ticket présent dans l'import précédent et absent du nouveau passe en vert « sorti de l'export : terminé ou annulé ».

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
3. Les onglets **Tickets**, **Fins d'intervention**, **Sortis de l'export** et **Adresses** sont créés, et Gmail sera lu toutes les 5 minutes.

## 4. Importer un export Praxedo

1. Recharge le Sheet : un menu **Carte tickets** apparaît.
2. **Carte tickets → Importer un export Praxedo…** → choisis le CSV exporté de Praxedo (ton filtre habituel, avec la colonne des notes si tu l'as ajoutée).
3. La première fois, compte une minute : les adresses sont placées sur la carte.

Refais un import aussi souvent que tu veux (matin, midi, soir…) : nouveaux tickets, plannings et fins d'intervention se mettent à jour à chaque fois.

**Import automatique par mail** : tout mail qui arrive dans ton Gmail avec l'export Praxedo en pièce jointe (fichier .csv) est importé tout seul dans les 5 minutes. Le script prend le plus récent, ignore les autres CSV (ceux qui n'ont pas de colonne « N° ») et n'importe jamais deux fois le même mail.

- Si Praxedo peut **envoyer l'export par mail à heure fixe** (export programmé ou planifié), fais-le envoyer à ton Gmail (ou à ton Outlook, avec une règle de transfert vers Gmail pour les mails qui ont un .csv) : la carte est alors 100 % automatique.
- Sinon, exporte depuis Praxedo et **envoie-toi le CSV par mail** : même depuis le téléphone, la carte se met à jour sans ouvrir le Sheet.

## 5. Mettre la carte en ligne

1. Dans l'éditeur : **Déployer → Nouveau déploiement** → roue dentée → **Application Web**.
2. **Exécuter en tant que : Moi**. **Qui a accès : Tout le monde**.
3. **Déployer**, puis copie l'**URL de l'application Web** (elle finit par `/exec`). Ouvre-la sur ordinateur ou téléphone.

La page se recharge toute seule toutes les 2 minutes. Toute personne qui a l'adresse voit la carte (adresses et descriptions des tickets) : ne la donne qu'à l'équipe.

## Au quotidien

- **Vert** : terminé (statut de fin dans l'export, ticket sorti de l'export, ou mail de fin transféré). La liste « Fins d'intervention aujourd'hui » est dans la légende.
- **Violet** : nouveau depuis moins de 24 h, avec un trait vers l'intervention déjà planifiée la plus proche.
- **Gris** : justifié / en attente. Contour orange : à relancer ou à vérifier.
- **Rouge / orange** : à planifier (rouge = échéance GTR dépassée ou dans moins de 48 h). **Bleu** : déjà planifié.
- Un mail de fin pour un ticket absent de l'export apparaît dans la liste « Mails de fin pour des tickets absents de l'export ».
- Un ticket annulé n'apparaît pas sur la carte.
- **Carte tickets → Lire les mails maintenant** : pour ne pas attendre les 5 minutes.

## Mettre à jour le code plus tard

Colle le nouveau code, puis **Déployer → Gérer les déploiements** → crayon → **Version : Nouvelle version** → **Déployer**. L'URL ne change pas.

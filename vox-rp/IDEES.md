# Vox RP : les idées

Tout ce dont on a parlé, rangé par thème. ✅ = fait, 🔜 = prochaine étape, le reste = plus tard.

## Ce qui nous rend différents

- **Une ville qui vote** : les joueurs élisent un maire. Il fixe les impôts, les amendes, le salaire des
  fonctionnaires, et ses décisions changent vraiment le jeu.
- **Une économie gérée par les joueurs** : vraies entreprises (stocks, salaires, factures), prix qui bougent
  selon l'offre et la demande.
- **Un bon accueil des nouveaux** : un tutoriel sous forme d'histoire à l'arrivée en ville.
- **Des médias** : journal et chaîne TV faits par les joueurs, visibles dans le téléphone et sur des écrans en ville.
  Le journal raconte les événements de la soirée.
- **Un serveur qui ne lag pas** : chaque script est surveillé pour rester léger.

## Événements (directeur d'événements)

- ✅ Colis mystère
- ✅ Accident de la route
- ✅ Panne de courant
- ✅ Tempête (plus d'accidents, la foudre peut couper le courant)
- ✅ Convoi de fonds à braquer (fourgon conduit par des PNJ, police prévenue)
- Course sauvage (rendez-vous secret envoyé aux pilotes, la police l'apprend après)
- Largage d'une caisse dans la nature
- Marchand mystère qui apparaît avec des objets rares
- Évasion d'un prisonnier PNJ
- Soirées bonus (livreurs payés double, happy hour au bar…)
- Événements lancés par le staff : concerts, tournois de boxe, courses officielles, élections, procès publics

## Illégal

- ✅ **Plantations** : graines, arrosage, engrais, les plants poussent même hors ligne. Risque de vol et de descente.
- **Labo de synthèse** : réaliste dans le gameplay (matériel, étapes, mini-jeux, fumées, risque d'explosion,
  odeur qui attire la police, la qualité fait le prix). Produits et étapes inventés, pas de vraie recette.
- **Lieux cachés sous terre** : labos et plantations dans les intérieurs secrets de GTA (ferme, labo, entrepôt),
  avec une entrée cachée (garage, ruelle, égout) et une clé ou un code.
- **Territoires et guerres de la drogue** : un groupe revendique un quartier. Si un autre groupe y vend beaucoup,
  le propriétaire reçoit une notification. Contrôler un quartier donne de meilleurs prix.
- **Trafic d'armes** : fournisseurs, livraisons, pièces à assembler.
- **Plusieurs gangs** avec leurs quartiers, leur hiérarchie et leurs alliances.

## Braquages

| Cible | Limite | Préparation | Policiers en service |
|---|---|---|---|
| Supérette (PNJ) | 3 par jour sur le serveur | non | au moins 2 |
| Bijouterie (PNJ) | 1 par jour | oui | oui |
| Banque (PNJ) | 1 par semaine | oui | oui |
| Commerce d'un joueur (concession, garage…) | 1 par semaine | oui | oui |
| Commissariat | 1 par mois | oui | oui, beaucoup |

- 🔜 Commerces PNJ : supérette, bijouterie, banque…
- Entreprises de joueurs : concession, garage des mécanos…
- Prises d'otages bien encadrées : négociation avec la police, revendications, fin qui dépend des joueurs.
- La panne de courant coupe les caméras : `exports.vox_events:EnPanne()` est déjà prêt pour les braquages.

## Des PNJ utiles : témoins et clients

- ✅ **Vente de drogue aux PNJ** : on peut vendre dans la rue. Si un autre PNJ voit la vente, il peut appeler la police.
- ✅ **Les PNJ appellent la police** quand ils voient quelque chose de suspect : coups de feu, enlèvement…
- ✅ **Témoignages** : un témoin donne le type de véhicule et sa couleur, sans plus (pas la plaque, pas le visage).
- **Enlèvements** : possibles, et sans alerte si personne n'a rien vu (les témoins sont prêts, il reste à coder l'enlèvement lui-même).

## Police, justice et prison

- **Descente au commissariat** : seulement quand il y a beaucoup de policiers en service, pour libérer un
  gardé à vue ou récupérer des saisies (en quantité limitée). Préparation obligatoire, rare, sanctions lourdes
  si ça échoue. Voir les règles proposées plus bas.
- **Garde à vue** pour les petits délits : quelques minutes en cellule au commissariat, puis libéré.
- **Jugement puis prison** dès que c'est grave.
- **Prison RP** : vraie vie en prison (travaux, cantine, bagarres, évasions préparées).
- Avocats, juges, procès, casier judiciaire.

### Règles proposées pour la descente au commissariat

- Possible seulement avec au moins 6 policiers en service (à régler).
- Une fois par mois maximum sur tout le serveur.
- Il faut préparer le coup avant : plan du bâtiment, matériel, et une information sur ce qui est dans la salle des saisies.
- On ne récupère qu'une partie des saisies, jamais tout.
- Le staff est prévenu automatiquement pour surveiller que ça reste RP.
- En cas d'échec : peines de prison lourdes et casier judiciaire.

## Lancement

- Développement sur PC avec la clé gratuite (48 joueurs maximum), serveur loué seulement pour la bêta.
- Discord à lancer tôt pour réunir la communauté.
- Bêta fermée avec les potes, puis ouverture et streamers.
- **Bande-annonce** juste avant l'ouverture : scénario plan par plan, voix off, et un « mode cinéma » pour
  filmer (caméra libre, ralenti, HUD caché, bandes noires). Montage avec DaVinci Resolve ou CapCut.
- Boutique uniquement via Tebex, sans pay-to-win (file prioritaire, cosmétiques, VIP sans avantage).

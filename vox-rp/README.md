# Vox RP

Serveur GTA RP français (FiveM), construit sur **Qbox** (la base RP : argent, métiers, inventaire, téléphone…).
Ce dossier contient **nos propres scripts**, ceux qui rendent Vox RP différent. Ils s'ajoutent à Qbox.

Toutes les idées prévues sont dans [IDEES.md](IDEES.md).

## Ce qui est déjà fait

### `vox_events` : le directeur d'événements

La ville lance toute seule des petits événements, toutes les 15 à 30 minutes, selon le nombre de joueurs
connectés et les métiers en service.

| Événement | Ce qui se passe | Qui en profite |
|---|---|---|
| **Colis mystère** | Un cercle jaune apparaît sur la carte. Le colis est caché dedans, le premier qui l'ouvre gagne de l'argent ou un objet. | Tout le monde |
| **Accident de la route** | Épaves qui fument et blessés au sol. Les ambulanciers soignent, les mécanos dégagent les épaves, la police sécurise. Chacun touche une prime. Sans ambulancier en service, les civils peuvent faire les premiers secours. | Police, ambulanciers, mécanos |
| **Convoi de fonds** | Un fourgon blindé conduit par deux gardes armés traverse la ville. Une rumeur prévient les joueurs. Il faut l'immobiliser, faire sauter les portes arrière, puis prendre les sacs d'argent sale. La police est prévenue dès l'attaque. Il faut au moins 3 policiers en service. | Braqueurs, police |
| **Panne de courant** | Toute la ville s'éteint quelques minutes (les phares des voitures marchent encore). | Les malins… |
| **Tempête** | Orage et vent. Les accidents arrivent 3 fois plus souvent, et la foudre peut provoquer une panne. | Tout le monde |

Tout se règle dans `resources/[vox]/vox_events/config.lua` : fréquence, primes, récompenses,
nombre de joueurs nécessaire, lieux…

**Commandes du staff** (admins seulement) :

| Commande | Effet |
|---|---|
| `/event` | Liste des événements et leur état |
| `/event colis` | Lance le colis mystère tout de suite (aussi : `accident`, `convoi`, `panne`, `tempete`) |
| `/eventstop colis` | Arrête un événement en cours |
| `/eventauto` | Active / coupe le lancement automatique |

### `vox_plantations` : les plantations

- On utilise une graine dehors, en pleine terre (pas sur la route, pas dans un bâtiment, pas près du commissariat).
- Il faut **arroser** (bouteille d'eau), sinon le plant ne pousse pas. Sans eau trop longtemps, il meurt.
- **L'engrais** améliore la qualité, donc la quantité récoltée.
- Les plants poussent **en temps réel, même quand tu n'es pas connecté** (2 h par défaut).
- N'importe qui peut récolter ou arracher un plant : il faut bien le cacher ! Le propriétaire est prévenu.
- La police en service qui saisit un plant touche une prime.
- 8 plants maximum par personnage. Tout se règle dans `resources/[vox]/vox_plantations/config.lua`.

Les objets doivent exister dans l'inventaire. Ajoute ces lignes dans `ox_inventory/data/items.lua`
(dans la liste, avant la dernière accolade `}`) :

```lua
['graine_cannabis'] = { label = 'Graine de cannabis', weight = 10, stack = true, close = true, description = 'À planter dans un coin discret, en pleine terre.' },
['engrais'] = { label = 'Engrais', weight = 500, stack = true, close = true },
['cannabis'] = { label = 'Cannabis', weight = 20, stack = true, close = true },
```

### `vox_temoins` : les PNJ témoins

Les passants ne sont plus des décors : s'ils voient ou entendent quelque chose de suspect, ils peuvent appeler la police.

- **Coups de feu** : détectés automatiquement. Avec un silencieux, il faut être tout près pour être entendu.
  Les policiers en service qui tirent ne déclenchent rien, ni le stand de tir d'Ammu-Nation.
- **Le témoin met 10 à 25 secondes à appeler** (il sort son téléphone). S'il est neutralisé avant, la police ne saura rien.
- **Personne autour = pas d'alerte** (enlèvement discret possible, par exemple).
- **Description vague** : homme ou femme, à pied ou le type de véhicule et sa couleur (parfois seulement l'un des deux).
  Jamais la plaque ni le visage. Le point envoyé à la police est approximatif.
- Plusieurs appels pour la même scène n'envoient qu'une alerte.
- Les autres scripts peuvent s'en servir : `exports.vox_temoins:ActeSuspect('braquage')` (côté joueur)
  ou `exports.vox_temoins:Alerter('braquage', { position, zone, description })` (côté serveur).

### `vox_deal` : la vente de rue

- Avec du cannabis sur toi, approche un passant à pied et appuie sur **E**.
- Il achète (1 à 3 unités, payées en argent sale), refuse, ou… **appelle la police** lui-même.
- Un autre passant qui voit l'échange peut aussi appeler.
- Un même passant n'achète qu'une fois toutes les 10 minutes. Interdit près du commissariat.
- Il faut au moins 1 policier en service (sinon « c'est trop calme ce soir »).
- Les ventes sont comptées par quartier : ça servira aux guerres de territoires.

## Installer le serveur sur ton PC (pour tester, gratuit)

Il te faut GTA V sur PC et [FiveM](https://fivem.net) installés.

1. **Clé de licence** : va sur [portal.cfx.re](https://portal.cfx.re), connecte-toi avec ton compte Cfx.re,
   puis *Server Registration Keys* → *Generate key*. Garde la clé, ne la partage pas.
2. **Base de données** : Qbox enregistre les personnages dans une base de données.
   Installe [MariaDB](https://mariadb.org/download/) (gratuit). Pendant l'installation, choisis un mot de passe
   pour l'utilisateur `root` et note-le.
3. **FXServer** (le programme du serveur) : télécharge la version *Recommended* de `server.7z` sur
   [runtime.fivem.net/artifacts/fivem/build_server_windows/master](https://runtime.fivem.net/artifacts/fivem/build_server_windows/master/),
   et décompresse-la dans un dossier, par exemple `C:\VoxRP\server`.
4. Lance `FXServer.exe`. **txAdmin** (le panneau de contrôle) s'ouvre dans ton navigateur.
   Relie ton compte Cfx.re, crée ton mot de passe (le code PIN est affiché dans la fenêtre noire).
5. Dans l'assistant de txAdmin :
   - nom du serveur : `Vox RP` ;
   - type d'installation : *Popular Recipes* → **Qbox** (si Qbox n'est pas dans la liste, suis le guide
     d'installation sur [docs.qbox.re](https://docs.qbox.re)) ;
   - colle ta clé de licence, puis les infos de la base de données (hôte `localhost`, utilisateur `root`,
     le mot de passe choisi à l'étape 2).

   txAdmin télécharge et installe tout. Ça prend quelques minutes. Toi, tu es automatiquement admin.
6. Ouvre FiveM, appuie sur **F8** et tape `connect localhost`. Tu arrives en ville 🎉

## Ajouter les scripts Vox RP

1. Copie le dossier `vox-rp/resources/[vox]` dans le dossier `resources` de ton serveur
   (txAdmin te montre où il est : *Settings* → *FXServer* → *Server Data Folder*).
2. Ouvre le `server.cfg` du serveur (txAdmin → *CFG Editor*) et ajoute **tout à la fin** :
   ```
   ensure [vox]
   ```
3. Redémarre le serveur depuis txAdmin.

## Tester les événements

En jeu, dans le chat (touche **T**) :

1. `/event` : tu dois voir la liste des 4 événements.
2. `/event colis` : un cercle jaune apparaît sur la carte. Va dedans, trouve le carton, appuie sur **E**.
3. `/event tempete` puis `/event panne` : la météo change, la ville s'éteint.
4. `/event convoi` : la commande du staff le lance même sans policiers, pour tester (tout seul, il attend
   3 policiers en service et 6 joueurs). Va au point vert : le fourgon apparaît quand tu arrives. Tire dessus pour
   l'arrêter, puis **E** derrière le fourgon.
5. **Plantations** : donne-toi des graines, de l'eau et de l'engrais avec la commande admin d'ox_inventory
   (par exemple `/giveitem [ton id] graine_cannabis 3`), utilise une graine dans l'herbe, puis **E** devant le plant.
   Pour tester vite, mets `Duree = 5` dans la config des plantations.
6. **Témoins et deal** (mets un deuxième joueur en policier en service, ou toi-même) : tire en l'air en pleine rue,
   l'alerte arrive chez la police 10 à 25 secondes après. Avec du cannabis, approche un passant et appuie sur **E**.
7. `/event accident` : un point rouge apparaît. Pour tester les métiers, donne-toi un métier avec la commande
   admin de Qbox (par exemple `/setjob [ton id] ambulance 0`), passe en service, et soigne les blessés.

**Si quelque chose ne marche pas**, envoie-moi :
- les lignes rouges de la console **F8** (côté jeu) ;
- les lignes rouges de la *Live Console* de txAdmin (côté serveur) ;
- et dis-moi si ton serveur a une ressource nommée `qb-weathersync` ou `Renewed-Weathersync`
  (txAdmin → *Resources*). C'est elle qui gère la météo, et la tempête doit s'y brancher.

## Ce qui n'a pas encore pu être vérifié

Les scripts ont été testés dans une simulation (logique du serveur, récompenses, conditions, sécurité),
mais pas encore dans le vrai jeu. À vérifier au premier essai :
- le bon placement du colis et de l'accident (trottoir et route les plus proches) ;
- la tempête et la panne avec la ressource météo installée par Qbox ;
- les noms de métiers : `police`, `ambulance`, `mechanic` par défaut (modifiables dans `Config.Metiers`) ;
- les objets donnés par le colis (`lockpick`, `armour`) et l'argent sale du convoi (`black_money`) doivent exister dans l'inventaire du serveur ;
- la conduite du fourgon par les PNJ et la réaction des gardes ;
- les modèles des plants (`bkr_prop_weed_...`) et leur hauteur sur le sol ;
- le témoin qui sort son téléphone, et la vente aux passants (ils doivent s'arrêter et se tourner vers toi).

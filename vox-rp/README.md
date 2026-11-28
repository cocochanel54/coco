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
| **Panne de courant** | Toute la ville s'éteint quelques minutes (les phares des voitures marchent encore). | Les malins… |
| **Tempête** | Orage et vent. Les accidents arrivent 3 fois plus souvent, et la foudre peut provoquer une panne. | Tout le monde |

Tout se règle dans `resources/[vox]/vox_events/config.lua` : fréquence, primes, récompenses,
nombre de joueurs nécessaire, lieux…

**Commandes du staff** (admins seulement) :

| Commande | Effet |
|---|---|
| `/event` | Liste des événements et leur état |
| `/event colis` | Lance le colis mystère tout de suite (aussi : `accident`, `panne`, `tempete`) |
| `/eventstop colis` | Arrête un événement en cours |
| `/eventauto` | Active / coupe le lancement automatique |

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
4. `/event accident` : un point rouge apparaît. Pour tester les métiers, donne-toi un métier avec la commande
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
- les objets donnés par le colis (`lockpick`, `armour`) doivent exister dans l'inventaire du serveur.

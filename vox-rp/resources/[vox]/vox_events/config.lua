Config = {}

---------------------------------------------------------------------------
-- Le directeur : il choisit tout seul quand lancer un événement
---------------------------------------------------------------------------
Config.Directeur = {
    Auto = true,          -- false = seulement quand le staff lance un événement (/event)
    PremierDelai = 10,    -- minutes après le démarrage du serveur avant le premier événement
    IntervalleMin = 15,   -- minutes minimum entre deux tentatives
    IntervalleMax = 30,   -- minutes maximum entre deux tentatives
    MaxEnMemeTemps = 2,   -- événements en même temps au maximum
}

---------------------------------------------------------------------------
-- Qui compte comme policier, ambulancier ou mécano (type ou nom du métier Qbox)
---------------------------------------------------------------------------
Config.Metiers = {
    police = { types = { 'leo' }, noms = { 'police', 'sheriff' } },
    ems = { types = { 'ems' }, noms = { 'ambulance' } },
    mecano = { types = { 'mechanic' }, noms = { 'mechanic', 'bennys' } },
}

-- Compte où l'argent des récompenses est versé ('cash' = liquide, 'bank' = banque)
Config.TypeArgent = 'cash'

-- Ressource qui gère la météo du serveur :
--   'auto'           = détection automatique
--   'qb-weathersync' = utilise ses fonctions pour la tempête et la panne
--   'vox'            = vox_events force lui-même la météo et la panne chez chaque joueur
Config.Meteo = { Ressource = 'auto' }

---------------------------------------------------------------------------
-- Lieux de départ des événements. Le jeu cherche ensuite la rue ou le
-- trottoir le plus proche. Ajoute les tiens avec /coords en jeu.
---------------------------------------------------------------------------
Config.Lieux = {
    { nom = 'Legion Square', coords = vector3(195.17, -933.77, 30.69) },
    { nom = 'Alta', coords = vector3(-269.40, -955.30, 31.22) },
    { nom = 'Pillbox Hill', coords = vector3(298.62, -584.54, 43.26) },
    { nom = 'Strawberry', coords = vector3(340.97, -1396.33, 32.51) },
    { nom = 'l\'aéroport', coords = vector3(-1037.74, -2738.04, 20.17) },
    { nom = 'Sandy Shores', coords = vector3(1839.44, 3672.71, 34.28) },
    { nom = 'Paleto Bay', coords = vector3(-247.76, 6331.23, 32.43) },
}

---------------------------------------------------------------------------
-- Les événements (durées et recharges en minutes)
--   Poids    : plus il est grand, plus l'événement tombe souvent
--   Recharge : temps minimum avant que le même événement revienne
---------------------------------------------------------------------------
Config.Evenements = {
    colis = {
        Actif = true,
        Poids = 3,
        Duree = 20,
        Recharge = 30,
        MinJoueurs = 3,
        RayonZone = 200.0,    -- taille du cercle affiché sur la carte
        Ecart = 120.0,        -- le colis est caché jusqu'à cette distance du lieu
        DureeOuverture = 5,   -- secondes pour ouvrir le colis
        Modele = 'prop_cs_cardbox_01',
        -- Une récompense est tirée au sort (les objets doivent exister dans ox_inventory)
        Recompenses = {
            { poids = 6, argent = { 1500, 4000 } },
            { poids = 2, objet = 'lockpick', quantite = 3 },
            { poids = 2, objet = 'armour', quantite = 1 },
        },
    },

    accident = {
        Actif = true,
        Poids = 2,
        Duree = 25,
        Recharge = 40,
        MinJoueurs = 4,
        MinSecours = 1,       -- au moins 1 policier, ambulancier ou mécano en service
        Victimes = { 1, 3 },  -- nombre de blessés (minimum, maximum)
        Vehicules = { 'blista', 'asea', 'premier', 'sultan', 'bison', 'minivan' },
        ModelesVictimes = { 'a_m_y_business_01', 'a_f_y_tourist_01', 'a_m_m_farmer_01', 'a_f_m_bevhills_01' },
        DureeSoin = 8,        -- secondes
        DureeDegagement = 10, -- secondes
        Primes = { soin = 750, degagement = 600, police = 400 },
        -- Sans ambulancier en service, les civils peuvent faire les premiers secours
        SoinsCivilsSansEms = true,
        PrimeCivile = 300,
    },

    convoi = {
        Actif = true,
        Poids = 1,
        Duree = 30,
        Recharge = 120,
        MinJoueurs = 6,
        MinPolice = 3,            -- pas de convoi sans assez de policiers pour le défendre
        DistanceMin = 1500.0,     -- distance minimum entre le départ et l'arrivée (mètres)
        DistanceMax = 4500.0,
        DistanceArrivee = 60.0,
        Vehicule = 'stockade',
        ModeleGarde = 's_m_m_armoured_01',
        Arme = 'WEAPON_PISTOL',
        -- ObjetCharge = 'thermite', -- décommenter pour exiger un objet pour faire sauter les portes
        DureeCharge = 8,          -- secondes pour poser la charge
        DelaiExplosion = 30,      -- secondes avant que les portes sautent
        Sacs = 4,
        DureeSac = 6,             -- secondes pour prendre un sac
        ArgentParSac = { 2500, 4000 },
        ObjetArgent = 'black_money', -- argent sale (si l'objet n'existe pas, c'est du liquide)
    },

    panne = {
        Actif = true,
        Poids = 1,
        Duree = 6,
        Recharge = 60,
        MinJoueurs = 5,
    },

    tempete = {
        Actif = true,
        Poids = 1,
        Duree = 20,
        Recharge = 90,
        MinJoueurs = 1,
        Meteo = 'THUNDER',       -- météo pendant la tempête
        MeteoApres = 'CLEARING', -- météo quand elle s'éloigne
        ChancePanne = 35,        -- % de chance que la foudre provoque une panne de courant
        BonusAccident = 3,       -- les accidents tombent 3 fois plus souvent pendant la tempête
    },
}

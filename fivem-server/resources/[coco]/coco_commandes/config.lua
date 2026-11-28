Config = {}

-- Mettre false pour désactiver une commande (par exemple /reparer sur un serveur RP sérieux)
Config.Commandes = {
    aide = true,     -- /aide : liste des commandes
    voiture = true,  -- /voiture [modèle]
    dv = true,       -- /dv : supprimer le véhicule
    reparer = true,  -- /reparer
    laver = true,    -- /laver
    soigner = true,  -- /soigner
    armure = false,  -- /armure
    tpm = true,      -- /tpm : aller au point posé sur la carte
    coords = true,   -- /coords : afficher sa position
    me = true,       -- /me [action]
    ['do'] = true,   -- /do [description]
    joueurs = true,  -- /joueurs : liste des joueurs connectés
}

-- Plaque des véhicules apparus avec /voiture (8 caractères maximum)
Config.Plaque = 'COCO'

-- true = l'ancienne voiture disparaît quand on en fait apparaître une nouvelle
Config.UneSeuleVoiture = true

-- true = la voiture d'un joueur disparaît quand il se déconnecte
Config.SupprimerAuDepart = true

-- Véhicules que /voiture refuse (tanks, avions de chasse, motos volantes...)
Config.VehiculesInterdits = {
    'rhino', 'khanjali', 'apc', 'chernobog',
    'hydra', 'lazer', 'strikeforce', 'starling', 'molotok', 'pyro', 'rogue', 'nokota',
    'hunter', 'savage', 'akula', 'annihilator2', 'valkyrie', 'valkyrie2', 'buzzard',
    'oppressor', 'oppressor2', 'deluxo', 'vigilante', 'ruiner2', 'scramjet', 'thruster',
    'stromberg', 'toreador', 'tampa3', 'insurgent', 'insurgent3', 'menacer', 'technical',
}

-- Distance (en mètres) à laquelle les autres voient tes /me et /do
Config.DistanceRP = 20.0

-- Durée d'affichage (en secondes) du texte au-dessus de la tête
Config.DureeRP = 7

-- Lignes ajoutées à la fin de /aide (commandes des autres ressources)
Config.AideEnPlus = {
    '/skin [modèle] : changer de personnage',
    '/monid : voir ton ID et ta licence',
}

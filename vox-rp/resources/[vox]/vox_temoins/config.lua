Config = {}

-- Les actes que les PNJ peuvent signaler à la police
--   Rayon       : distance (mètres) à laquelle un PNJ peut être témoin
--   Vue         : true = il faut que le PNJ voie la scène ; false = il suffit d'entendre (coups de feu)
--   Chance      : % de chance que le témoin appelle vraiment la police
--   Recharge    : secondes avant qu'un même joueur puisse déclencher un nouveau signalement de ce type
--   Imprecision : le point envoyé à la police est décalé jusqu'à cette distance (le témoin n'est pas un GPS)
Config.Actes = {
    tirs = { Titre = 'Coups de feu', Rayon = 80.0, Vue = false, Chance = 75, Recharge = 30, Imprecision = 60.0 },
    vente_drogue = { Titre = 'Deal de rue', Rayon = 25.0, Vue = true, Chance = 40, Recharge = 20, Imprecision = 40.0 },
    enlevement = { Titre = 'Enlèvement', Rayon = 40.0, Vue = true, Chance = 80, Recharge = 30, Imprecision = 50.0 },
    braquage = { Titre = 'Braquage', Rayon = 40.0, Vue = true, Chance = 90, Recharge = 60, Imprecision = 20.0 },
}

-- Secondes que met le témoin à appeler (s'il est neutralisé avant, l'appel n'arrive jamais)
Config.DelaiAppel = { 10, 25 }

-- Durée (secondes) pendant laquelle l'alerte reste sur la carte des policiers
Config.DureeAlerte = 180

-- Deux alertes du même type à moins de cette distance et de ce délai n'en font qu'une
Config.Regroupement = { Distance = 150.0, Secondes = 60 }

-- Les policiers en service qui tirent ne déclenchent pas d'appel
Config.Police = { types = { 'leo' }, noms = { 'police', 'sheriff' } }

-- Endroits où les coups de feu sont normaux
Config.ZonesSansAlerte = {
    { nom = 'Stand de tir d\'Ammu-Nation', centre = vector3(13.6, -1097.5, 29.8), rayon = 40.0 },
}

-- Ce que dit le témoin du véhicule (classes du jeu). Volontairement vague.
Config.Vehicules = {
    [0] = 'une citadine', [1] = 'une berline', [2] = 'un SUV', [3] = 'un coupé', [4] = 'une muscle car',
    [5] = 'une voiture de collection', [6] = 'une voiture de sport', [7] = 'une supercar', [8] = 'une moto',
    [9] = 'un tout-terrain', [10] = 'un véhicule industriel', [11] = 'un utilitaire', [12] = 'une camionnette',
    [13] = 'un vélo', [14] = 'un bateau', [15] = 'un hélicoptère', [16] = 'un avion', [17] = 'un véhicule de service',
    [18] = 'un véhicule d\'urgence', [19] = 'un véhicule militaire', [20] = 'un camion',
}

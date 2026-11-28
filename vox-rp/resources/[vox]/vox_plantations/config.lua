Config = {}

-- Ce qu'on peut planter. Les objets (graine, récolte) doivent exister dans ox_inventory (voir README).
Config.Plantes = {
    cannabis = {
        Label = 'Plant de cannabis',
        Graine = 'graine_cannabis',
        Recolte = 'cannabis',
        Quantite = { 3, 8 },  -- récolte selon la qualité (minimum, maximum)
        Duree = 120,          -- minutes pour pousser, en temps réel (même hors ligne)
        -- Un modèle par stade : petite pousse, plant moyen, plant prêt à récolter
        Modeles = { 'bkr_prop_weed_01_small_01c', 'bkr_prop_weed_med_01a', 'bkr_prop_weed_lrg_01a' },
    },
}

-- Eau (de 0 à 100 %). Sans eau, la plante arrête de pousser, puis meurt.
Config.Eau = {
    Objet = 'water',     -- une bouteille d'eau = un arrosage
    Gain = 50,           -- % ajoutés à chaque arrosage
    PerteParHeure = 25,  -- % perdus par heure
    DelaiMort = 90,      -- minutes sans eau avant que la plante meure
}

-- Engrais : améliore la qualité, donc la quantité récoltée
Config.Engrais = { Objet = 'engrais', Gain = 20 }
Config.QualiteDepart = 40

Config.MaxParJoueur = 8         -- plants en même temps par personnage
Config.DistanceEntrePlantes = 1.2
Config.DistanceAffichage = 120.0

-- La police en service gagne une prime quand elle détruit un plant
Config.Police = { types = { 'leo' }, noms = { 'police', 'sheriff' }, Prime = 150 }

-- true = n'importe qui peut arracher un plant (sabotage entre gangs)
Config.DestructionLibre = true

-- Endroits où on ne peut pas planter
Config.ZonesInterdites = {
    { nom = 'le commissariat de Mission Row', centre = vector3(441.0, -982.0, 30.7), rayon = 150.0 },
    { nom = 'Legion Square', centre = vector3(195.17, -933.77, 30.69), rayon = 200.0 },
}

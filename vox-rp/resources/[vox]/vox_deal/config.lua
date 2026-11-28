Config = {}

-- Pas de vente de rue s'il n'y a pas assez de policiers en service
Config.MinPolice = 1
Config.Police = { types = { 'leo' }, noms = { 'police', 'sheriff' } }

-- Ce qu'on peut vendre (objets ox_inventory). Prix par unité, en argent sale.
Config.Produits = {
    { Objet = 'cannabis', Label = 'cannabis', Prix = { 60, 110 }, Quantite = { 1, 3 } },
}

-- Objet donné en paiement (argent sale). S'il n'existe pas, le joueur reçoit du liquide.
Config.Paiement = 'black_money'

-- Réaction du PNJ, en % : il refuse, ou il appelle la police. Le reste du temps, il achète.
Config.Chances = { Refuse = 30, Appelle = 10 }

Config.Recharge = 8     -- secondes entre deux propositions
Config.MemeClient = 600 -- secondes avant de pouvoir revendre au même PNJ

Config.ZonesInterdites = {
    { nom = 'du commissariat', centre = vector3(441.0, -982.0, 30.7), rayon = 120.0 },
}

Config.Refus = {
    'Non merci, je ne touche pas à ça.',
    'Dégage, ça ne m\'intéresse pas.',
    'Pas aujourd\'hui.',
    'T\'es sérieux ? En pleine rue ?',
}

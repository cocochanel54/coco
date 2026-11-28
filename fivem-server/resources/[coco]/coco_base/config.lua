Config = {}

-- Nom affiché dans le HUD et dans les messages du chat
Config.NomServeur = 'Coco RP'

-- Où les joueurs apparaissent en se connectant (un point au hasard dans la liste).
-- Pour ajouter un point : va à l'endroit voulu en jeu, tape /coords et colle le résultat ici.
Config.PointsDeDepart = {
    { nom = 'Legion Square', coords = vector4(195.17, -933.77, 30.69, 144.0) },
    { nom = 'Aéroport de Los Santos', coords = vector4(-1037.74, -2738.04, 20.17, 330.0) },
}

-- Après une mort, on réapparaît à l'hôpital le plus proche
Config.Hopitaux = {
    { nom = 'Pillbox Hill', coords = vector4(298.62, -584.54, 43.26, 70.0) },
    { nom = 'Strawberry', coords = vector4(340.97, -1396.33, 32.51, 50.0) },
    { nom = 'Sandy Shores', coords = vector4(1839.44, 3672.71, 34.28, 210.0) },
    { nom = 'Paleto Bay', coords = vector4(-247.76, 6331.23, 32.43, 225.0) },
}

-- Secondes à attendre avant de réapparaître après une mort
Config.DelaiReapparition = 10

-- Personnages possibles à la première connexion (un au hasard).
-- Chaque joueur peut ensuite changer avec /skin, son choix est retenu.
Config.Modeles = {
    'a_m_y_business_01',
    'a_f_y_business_01',
    'a_m_y_hipster_01',
    'a_f_y_hipster_01',
    'a_m_y_skater_01',
    'a_f_y_tourist_01',
}

-- true = tout le monde peut utiliser /skin
Config.SkinLibre = true

-- Petit texte en haut à droite (nom du serveur, ton ID, joueurs, rue)
Config.HUD = true

-- false = pas d'étoiles de recherche ni de police contrôlée par le jeu
Config.PoliceIA = false

-- true = les joueurs peuvent se blesser entre eux
Config.PvP = true

-- Messages « X a rejoint / quitté la ville » dans le chat
Config.MessagesConnexion = true

-- Coco Admin (côté serveur)
-- Toutes ces commandes sont réservées : seuls les joueurs ajoutés au groupe admin
-- (ligne « add_principal ... group.admin » dans secrets.cfg) peuvent les utiliser.

local function repondre(src, texte, erreur)
    if src == 0 then
        print('[coco_admin] ' .. texte)
        return
    end
    TriggerClientEvent('chat:addMessage', src, {
        color = erreur and { 255, 90, 90 } or { 255, 170, 0 },
        multiline = true,
        args = { 'Admin', texte },
    })
end

local function joueurExiste(id)
    if not id then
        return false
    end
    for _, idTexte in ipairs(GetPlayers()) do
        if tonumber(idTexte) == id then
            return true
        end
    end
    return false
end

local function joueurCible(src, valeur)
    local id = tonumber(valeur)
    if not joueurExiste(id) then
        repondre(src, ('Aucun joueur avec l\'ID %s. Tape /joueurs pour voir les ID.'):format(valeur or '?'), true)
        return nil
    end
    return id
end

local function enJeuSeulement(src)
    if src == 0 then
        repondre(src, 'Cette commande se tape en jeu.', true)
        return false
    end
    return true
end

local commandesAdmin = {}

local function commandeAdmin(nom, description, parametres, fonction)
    -- true = commande réservée (permission « command.<nom> », donnée au groupe admin dans server.cfg)
    RegisterCommand(nom, function(src, args)
        fonction(src, args)
    end, true)
    commandesAdmin[#commandesAdmin + 1] = { nom = nom, description = description, parametres = parametres }
end

---------------------------------------------------------------------------
-- Météo et heure, partagées par tous les joueurs
---------------------------------------------------------------------------

local minutes = (Config.HeureDepart % 24) * 60 -- minutes écoulées depuis minuit dans le jeu
local heureFigee = false

GlobalState.meteo = Config.MeteoDepart
GlobalState.heure = minutes

CreateThread(function()
    while true do
        Wait(Config.DureeMinuteJeu)
        if not heureFigee then
            minutes = (minutes + 1) % 1440
            GlobalState.heure = minutes
        end
    end
end)

commandeAdmin('meteo', 'Changer la météo pour tout le monde', {
    { name = 'type', help = 'soleil, clair, nuageux, couvert, pluie, orage, brouillard, neige, noel...' },
}, function(src, args)
    local demande = (args[1] or ''):lower()
    local meteo = Config.Meteos[demande]
    if not meteo then
        for _, nomJeu in pairs(Config.Meteos) do
            if nomJeu == demande:upper() then
                meteo = nomJeu
            end
        end
    end
    if not meteo then
        local noms = {}
        for nom in pairs(Config.Meteos) do
            noms[#noms + 1] = nom
        end
        table.sort(noms)
        repondre(src, 'Météo inconnue. Choix possibles : ' .. table.concat(noms, ', '), true)
        return
    end
    GlobalState.meteo = meteo
    repondre(src, ('Météo changée : %s (transition de 15 secondes).'):format(demande))
end)

commandeAdmin('heure', 'Changer l\'heure pour tout le monde', {
    { name = 'heure', help = '0 à 23' },
    { name = 'minutes', help = '0 à 59 (facultatif)' },
}, function(src, args)
    local h, m = tonumber(args[1]), tonumber(args[2] or 0)
    if not h or not m or h < 0 or h > 23 or m < 0 or m > 59 then
        repondre(src, 'Utilise /heure [0-23] [0-59], par ex. /heure 21 30', true)
        return
    end
    h, m = math.floor(h), math.floor(m)
    minutes = h * 60 + m
    GlobalState.heure = minutes
    repondre(src, ('Il est maintenant %02d:%02d.'):format(h, m))
end)

commandeAdmin('figerheure', 'Arrêter ou relancer l\'écoulement du temps', nil, function(src)
    heureFigee = not heureFigee
    repondre(src, heureFigee and 'Le temps est figé.' or 'Le temps s\'écoule de nouveau.')
end)

---------------------------------------------------------------------------
-- Annonces et joueurs
---------------------------------------------------------------------------

commandeAdmin('annonce', 'Envoyer une annonce à tout le serveur', {
    { name = 'texte', help = 'le message' },
}, function(src, args)
    local texte = table.concat(args, ' ')
    if texte == '' or not utf8.len(texte) then
        repondre(src, 'Utilise /annonce [texte]', true)
        return
    end
    TriggerClientEvent('chat:addMessage', -1, {
        color = { 255, 80, 80 },
        multiline = true,
        args = { 'ANNONCE', texte },
    })
    TriggerClientEvent('coco_admin:annonce', -1, (texte:gsub('~', '')))
end)

commandeAdmin('kick', 'Expulser un joueur du serveur', {
    { name = 'id', help = 'ID du joueur (voir /joueurs)' },
    { name = 'raison', help = 'facultatif' },
}, function(src, args)
    local id = joueurCible(src, args[1])
    if not id then
        return
    end
    local raison = table.concat(args, ' ', 2)
    if raison == '' then
        raison = 'Expulsé par un admin.'
    end
    local nom = GetPlayerName(id)
    DropPlayer(id, raison)
    repondre(src, ('%s a été expulsé (%s).'):format(nom, raison))
end)

commandeAdmin('tp', 'Aller vers un joueur', {
    { name = 'id', help = 'ID du joueur (voir /joueurs)' },
}, function(src, args)
    if not enJeuSeulement(src) then
        return
    end
    local id = joueurCible(src, args[1])
    if not id then
        return
    end
    TriggerClientEvent('coco_admin:teleporter', src, GetEntityCoords(GetPlayerPed(id)))
end)

commandeAdmin('amener', 'Faire venir un joueur à côté de toi', {
    { name = 'id', help = 'ID du joueur (voir /joueurs)' },
}, function(src, args)
    if not enJeuSeulement(src) then
        return
    end
    local id = joueurCible(src, args[1])
    if not id then
        return
    end
    TriggerClientEvent('coco_admin:teleporter', id, GetEntityCoords(GetPlayerPed(src)))
    repondre(src, ('%s arrive.'):format(GetPlayerName(id)))
end)

commandeAdmin('tpcoords', 'Aller à des coordonnées', {
    { name = 'x', help = 'ex : 195.17' },
    { name = 'y', help = 'ex : -933.77' },
    { name = 'z', help = 'ex : 30.69' },
}, function(src, args)
    if not enJeuSeulement(src) then
        return
    end
    -- Accepte aussi un copier-coller de /coords : vector4(x, y, z, cap)
    local texte = table.concat(args, ' '):gsub('vector%d', '')
    local valeurs = {}
    for nombre in texte:gmatch('-?%d+%.?%d*') do
        valeurs[#valeurs + 1] = tonumber(nombre)
    end
    if #valeurs < 3 then
        repondre(src, 'Utilise /tpcoords [x] [y] [z]', true)
        return
    end
    TriggerClientEvent('coco_admin:teleporter', src, vector3(valeurs[1], valeurs[2], valeurs[3]))
end)

commandeAdmin('noclip', 'Voler à travers les murs (retape pour arrêter)', nil, function(src)
    if not enJeuSeulement(src) then
        return
    end
    TriggerClientEvent('coco_admin:noclip', src)
end)

commandeAdmin('reanimer', 'Réanimer et soigner un joueur (toi si pas d\'ID)', {
    { name = 'id', help = 'facultatif' },
}, function(src, args)
    local id = src
    if args[1] then
        id = joueurCible(src, args[1])
        if not id then
            return
        end
    elseif not enJeuSeulement(src) then
        return
    end
    TriggerClientEvent('coco_admin:reanimer', id)
    repondre(src, ('%s a été réanimé.'):format(GetPlayerName(id)))
end)

commandeAdmin('admin', 'Liste des commandes admin', nil, function(src)
    local lignes = {}
    for _, commande in ipairs(commandesAdmin) do
        local parametres = ''
        for _, parametre in ipairs(commande.parametres or {}) do
            parametres = parametres .. ' [' .. parametre.name .. ']'
        end
        lignes[#lignes + 1] = ('/%s%s : %s'):format(commande.nom, parametres, commande.description)
    end
    repondre(src, table.concat(lignes, '\n'))
end)

---------------------------------------------------------------------------
-- Suggestions du chat : seulement pour ceux qui ont le droit
---------------------------------------------------------------------------

RegisterNetEvent('coco_admin:demanderSuggestions', function()
    local src = source
    local liste = {}
    for _, commande in ipairs(commandesAdmin) do
        if IsPlayerAceAllowed(src, 'command.' .. commande.nom) then
            liste[#liste + 1] = commande
        end
    end
    if #liste > 0 then
        TriggerClientEvent('coco_admin:suggestions', src, liste)
    end
end)

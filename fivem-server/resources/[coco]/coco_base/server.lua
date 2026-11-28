-- Coco Base (côté serveur) : messages de connexion, nombre de joueurs, /monid

local function majNombreJoueurs()
    GlobalState.nbJoueurs = #GetPlayers()
end

GlobalState.maxJoueurs = GetConvarInt('sv_maxclients', 48)
majNombreJoueurs()

local function licence(src)
    for i = 0, GetNumPlayerIdentifiers(src) - 1 do
        local id = GetPlayerIdentifier(src, i)
        if id and id:sub(1, 8) == 'license:' then
            return id
        end
    end
    return nil
end

local function message(cible, auteur, texte, couleur)
    TriggerClientEvent('chat:addMessage', cible, {
        color = couleur,
        multiline = true,
        args = { auteur, texte },
    })
end

AddEventHandler('playerJoining', function()
    local src = source
    local nom = GetPlayerName(src) or ('Joueur ' .. src)
    print(('[coco_base] %s (ID %s) a rejoint le serveur - %s'):format(nom, src, licence(src) or 'licence inconnue'))
    if Config.MessagesConnexion then
        message(-1, Config.NomServeur, ('%s a rejoint la ville.'):format(nom), { 120, 200, 120 })
    end
    majNombreJoueurs()
end)

AddEventHandler('playerDropped', function(raison)
    local src = source
    local nom = GetPlayerName(src) or ('Joueur ' .. src)
    print(('[coco_base] %s (ID %s) est parti : %s'):format(nom, src, raison or '?'))
    if Config.MessagesConnexion then
        message(-1, Config.NomServeur, ('%s a quitté la ville.'):format(nom), { 200, 120, 120 })
    end
    -- Le joueur compte encore pendant cet événement : on recompte un peu après
    SetTimeout(1000, majNombreJoueurs)
end)

-- Recomptage régulier, au cas où
CreateThread(function()
    while true do
        Wait(30000)
        majNombreJoueurs()
    end
end)

-- /monid : donne au joueur la ligne à coller dans secrets.cfg pour devenir admin
RegisterCommand('monid', function(src)
    if src == 0 then
        print('[coco_base] /monid se tape en jeu, dans le chat.')
        return
    end
    local lic = licence(src)
    if not lic then
        message(src, 'Mon ID', ('Ton ID en jeu : %d (licence introuvable)'):format(src), { 255, 170, 0 })
        return
    end
    message(src, 'Mon ID', (
        'Ton ID en jeu : %d\nTa licence : %s\n' ..
        'Pour devenir admin, ajoute cette ligne dans secrets.cfg :\n' ..
        'add_principal identifier.%s group.admin'
    ):format(src, lic, lic), { 255, 170, 0 })
    print(('[coco_base] %s : add_principal identifier.%s group.admin'):format(GetPlayerName(src), lic))
end, false)

-- Coco Commandes (côté serveur) : véhicules, /me, /do, /joueurs

---------------------------------------------------------------------------
-- Véhicules apparus avec /voiture
---------------------------------------------------------------------------

local vehicules = {} -- [ID du joueur] = ID réseau de son véhicule

local function supprimerVehiculeDuJoueur(src)
    local netId = vehicules[src]
    vehicules[src] = nil
    if not netId then
        return
    end
    local entite = NetworkGetEntityFromNetworkId(netId)
    -- On vérifie que c'est bien encore son véhicule (les ID réseau sont réutilisés)
    if entite ~= 0 and DoesEntityExist(entite) and Entity(entite).state.cocoProprietaire == src then
        DeleteEntity(entite)
    end
end

RegisterNetEvent('coco_commandes:nouveauVehicule', function(netId)
    local src = source
    if not Config.Commandes.voiture or type(netId) ~= 'number' then
        return
    end
    -- Le message peut arriver avant que le serveur connaisse le véhicule : on patiente un peu
    local entite = 0
    for _ = 1, 20 do
        entite = NetworkGetEntityFromNetworkId(netId)
        if entite ~= 0 and DoesEntityExist(entite) then
            break
        end
        Wait(100)
    end
    if entite == 0 or not DoesEntityExist(entite) or GetEntityType(entite) ~= 2 then
        return
    end
    -- Seul celui qui vient de créer le véhicule peut le déclarer comme le sien
    if NetworkGetEntityOwner(entite) ~= src then
        return
    end
    if Config.UneSeuleVoiture and vehicules[src] ~= netId then
        supprimerVehiculeDuJoueur(src)
    end
    Entity(entite).state:set('cocoProprietaire', src, false)
    vehicules[src] = netId
end)

RegisterNetEvent('coco_commandes:supprimerVehicule', function(netId)
    local src = source
    if not (Config.Commandes.dv or Config.Commandes.voiture) or type(netId) ~= 'number' then
        return
    end
    local entite = NetworkGetEntityFromNetworkId(netId)
    if entite == 0 or not DoesEntityExist(entite) or GetEntityType(entite) ~= 2 then
        return
    end
    -- Protection : on ne supprime que les véhicules à côté du joueur
    local distance = #(GetEntityCoords(GetPlayerPed(src)) - GetEntityCoords(entite))
    if distance > 10.0 then
        return
    end
    if vehicules[src] == netId then
        vehicules[src] = nil
    end
    DeleteEntity(entite)
end)

AddEventHandler('playerDropped', function()
    local src = source
    if Config.SupprimerAuDepart then
        supprimerVehiculeDuJoueur(src)
    end
    vehicules[src] = nil
end)

---------------------------------------------------------------------------
-- /me et /do : visibles seulement par les joueurs proches
---------------------------------------------------------------------------

local MODELES_CHAT = {
    me = '<span style="color:#c27aff"><b>* {0}</b> {1}</span>',
    ['do'] = '<span style="color:#ffc45c"><b>** {0} :</b> {1}</span>',
}

local function messageRP(src, args, genre)
    if src == 0 then
        print('[coco_commandes] /' .. genre .. ' se tape en jeu.')
        return
    end
    -- Les ~ servent aux codes couleur du jeu : on les retire
    local texte = table.concat(args, ' '):gsub('~', ''):gsub('^%s+', ''):gsub('%s+$', '')
    local longueur = utf8.len(texte)
    if texte == '' or not longueur then
        TriggerClientEvent('chat:addMessage', src, { args = { 'RP', ('Utilise /%s [texte]'):format(genre) } })
        return
    end
    if longueur > 200 then
        texte = texte:sub(1, utf8.offset(texte, 201) - 1) .. '...'
    end

    local nom = GetPlayerName(src)
    local position = GetEntityCoords(GetPlayerPed(src))
    for _, idTexte in ipairs(GetPlayers()) do
        local id = tonumber(idTexte)
        local ped = GetPlayerPed(id)
        if ped ~= 0 and #(position - GetEntityCoords(ped)) <= Config.DistanceRP then
            TriggerClientEvent('chat:addMessage', id, { template = MODELES_CHAT[genre], args = { nom, texte } })
            TriggerClientEvent('coco_commandes:texteRP', id, src, texte, genre)
        end
    end
end

if Config.Commandes.me then
    RegisterCommand('me', function(src, args)
        messageRP(src, args, 'me')
    end, false)
end

if Config.Commandes['do'] then
    RegisterCommand('do', function(src, args)
        messageRP(src, args, 'do')
    end, false)
end

---------------------------------------------------------------------------
-- /joueurs
---------------------------------------------------------------------------

if Config.Commandes.joueurs then
    RegisterCommand('joueurs', function(src)
        local ids = {}
        for _, idTexte in ipairs(GetPlayers()) do
            ids[#ids + 1] = tonumber(idTexte)
        end
        table.sort(ids)
        local lignes = {}
        for _, id in ipairs(ids) do
            lignes[#lignes + 1] = ('%d : %s'):format(id, GetPlayerName(id))
        end
        local texte = ('%d joueur(s) connecté(s)\n%s'):format(#ids, table.concat(lignes, '\n'))
        if src == 0 then
            print(texte)
        else
            TriggerClientEvent('chat:addMessage', src, {
                color = { 120, 180, 255 },
                multiline = true,
                args = { 'Joueurs', texte },
            })
        end
    end, false)
end

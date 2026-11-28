-- Coco Commandes (côté joueur)

local function notifier(message)
    BeginTextCommandThefeedPost('STRING')
    AddTextComponentSubstringPlayerName(message)
    EndTextCommandThefeedPostTicker(false, true)
end

-- Liste utilisée pour /aide et pour les suggestions du chat
local listeAide = {}

local function ajouterCommande(nom, description, parametres, fonction)
    if not Config.Commandes[nom] then
        return
    end
    -- Sans fonction : la commande est gérée par le serveur (server.lua)
    if fonction then
        RegisterCommand(nom, function(_, args)
            fonction(args)
        end, false)
    end
    listeAide[#listeAide + 1] = { nom = nom, description = description, parametres = parametres }
end

---------------------------------------------------------------------------
-- Outils
---------------------------------------------------------------------------

local function chargerModele(hash)
    RequestModel(hash)
    local limite = GetGameTimer() + 5000
    while not HasModelLoaded(hash) do
        if GetGameTimer() > limite then
            return false
        end
        Wait(0)
    end
    return true
end

-- Le véhicule dans lequel on est, sinon le plus proche dans le rayon donné
local function vehiculeProche(rayon)
    local ped = PlayerPedId()
    local dedans = GetVehiclePedIsIn(ped, false)
    if dedans ~= 0 then
        return dedans
    end
    local position = GetEntityCoords(ped)
    local choix, meilleureDistance = nil, rayon
    for _, vehicule in ipairs(GetGamePool('CVehicle')) do
        local distance = #(position - GetEntityCoords(vehicule))
        if distance < meilleureDistance then
            choix, meilleureDistance = vehicule, distance
        end
    end
    return choix
end

-- Il faut « contrôler » un véhicule pour le modifier (sinon le jeu annule le changement)
local function prendreControle(entite)
    if not NetworkGetEntityIsNetworked(entite) then
        return true
    end
    local limite = GetGameTimer() + 1500
    while not NetworkHasControlOfEntity(entite) do
        if GetGameTimer() > limite then
            return false
        end
        NetworkRequestControlOfEntity(entite)
        Wait(0)
    end
    return true
end

local function supprimerVehicule(vehicule)
    if NetworkGetEntityIsNetworked(vehicule) then
        -- Le serveur supprime le véhicule pour tout le monde
        TriggerServerEvent('coco_commandes:supprimerVehicule', NetworkGetNetworkIdFromEntity(vehicule))
    else
        SetEntityAsMissionEntity(vehicule, true, true)
        DeleteVehicle(vehicule)
    end
    local limite = GetGameTimer() + 1500
    while DoesEntityExist(vehicule) and GetGameTimer() < limite do
        Wait(0)
    end
    return not DoesEntityExist(vehicule)
end

local vehiculesInterdits = {}
for _, nom in ipairs(Config.VehiculesInterdits) do
    vehiculesInterdits[GetHashKey(nom)] = true
end

---------------------------------------------------------------------------
-- Véhicules
---------------------------------------------------------------------------

ajouterCommande('voiture', 'Faire apparaître un véhicule', {
    { name = 'modele', help = 'ex : sultan, adder, faggio, sanchez, bati' },
}, function(args)
    local nom = args[1] and args[1]:lower()
    if not nom then
        notifier('Utilise /voiture [modèle], par ex. /voiture sultan')
        return
    end
    local hash = GetHashKey(nom)
    if not IsModelInCdimage(hash) or not IsModelAVehicle(hash) then
        notifier(('~r~Véhicule inconnu :~s~ %s'):format(nom))
        return
    end
    if vehiculesInterdits[hash] then
        notifier('~r~Ce véhicule est interdit sur le serveur.')
        return
    end
    local ped = PlayerPedId()
    if IsEntityDead(ped) then
        notifier('~r~Impossible pendant que tu es mort.')
        return
    end
    if not chargerModele(hash) then
        notifier('~r~Le véhicule met trop de temps à charger, réessaie.')
        return
    end

    -- Si on conduit déjà, l'ancien véhicule est remplacé
    local actuel = GetVehiclePedIsIn(ped, false)
    if actuel ~= 0 and GetPedInVehicleSeat(actuel, -1) == ped then
        supprimerVehicule(actuel)
    end

    local position = GetEntityCoords(ped)
    local vehicule = CreateVehicle(hash, position.x, position.y, position.z, GetEntityHeading(ped), true, false)
    SetModelAsNoLongerNeeded(hash)
    if vehicule == 0 then
        notifier('~r~Impossible de créer le véhicule ici.')
        return
    end

    SetPedIntoVehicle(ped, vehicule, -1)
    SetVehicleOnGroundProperly(vehicule)
    SetVehicleNumberPlateText(vehicule, Config.Plaque)
    SetVehicleEngineOn(vehicule, true, true, false)
    SetVehRadioStation(vehicule, 'OFF')

    -- On prévient le serveur, qui supprimera l'ancienne voiture du joueur
    local limite = GetGameTimer() + 2000
    while not NetworkGetEntityIsNetworked(vehicule) and GetGameTimer() < limite do
        Wait(0)
    end
    if NetworkGetEntityIsNetworked(vehicule) then
        TriggerServerEvent('coco_commandes:nouveauVehicule', NetworkGetNetworkIdFromEntity(vehicule))
    end

    notifier(('~g~%s~s~ prêt, bonne route !'):format(GetLabelText(GetDisplayNameFromVehicleModel(hash))))
end)

ajouterCommande('dv', 'Supprimer le véhicule où tu es (ou le plus proche)', nil, function()
    local vehicule = vehiculeProche(5.0)
    if not vehicule then
        notifier('Aucun véhicule à moins de 5 mètres.')
        return
    end
    if supprimerVehicule(vehicule) then
        notifier('Véhicule supprimé.')
    else
        notifier('~r~Impossible de supprimer ce véhicule.')
    end
end)

ajouterCommande('reparer', 'Réparer le véhicule où tu es (ou le plus proche)', nil, function()
    local vehicule = vehiculeProche(5.0)
    if not vehicule then
        notifier('Aucun véhicule à moins de 5 mètres.')
        return
    end
    if not prendreControle(vehicule) then
        notifier('~r~Quelqu\'un d\'autre conduit ce véhicule.')
        return
    end
    SetVehicleFixed(vehicule)
    SetVehicleDeformationFixed(vehicule)
    SetVehicleEngineHealth(vehicule, 1000.0)
    SetVehicleBodyHealth(vehicule, 1000.0)
    SetVehiclePetrolTankHealth(vehicule, 1000.0)
    SetVehicleUndriveable(vehicule, false)
    notifier('Véhicule ~g~réparé~s~.')
end)

ajouterCommande('laver', 'Laver le véhicule où tu es (ou le plus proche)', nil, function()
    local vehicule = vehiculeProche(5.0)
    if not vehicule then
        notifier('Aucun véhicule à moins de 5 mètres.')
        return
    end
    if not prendreControle(vehicule) then
        notifier('~r~Quelqu\'un d\'autre conduit ce véhicule.')
        return
    end
    SetVehicleDirtLevel(vehicule, 0.0)
    WashDecalsFromVehicle(vehicule, 1.0)
    notifier('Véhicule tout ~b~propre~s~ !')
end)

---------------------------------------------------------------------------
-- Personnage
---------------------------------------------------------------------------

ajouterCommande('soigner', 'Remettre ta santé au maximum', nil, function()
    local ped = PlayerPedId()
    if IsEntityDead(ped) then
        notifier('~r~Impossible pendant que tu es mort.')
        return
    end
    SetEntityHealth(ped, GetEntityMaxHealth(ped))
    ClearPedBloodDamage(ped)
    ResetPedVisibleDamage(ped)
    notifier('Tu es en pleine forme.')
end)

ajouterCommande('armure', 'Mettre un gilet pare-balles', nil, function()
    local ped = PlayerPedId()
    if IsEntityDead(ped) then
        notifier('~r~Impossible pendant que tu es mort.')
        return
    end
    SetPedArmour(ped, 100)
    notifier('Gilet pare-balles équipé.')
end)

---------------------------------------------------------------------------
-- Position
---------------------------------------------------------------------------

ajouterCommande('tpm', 'Aller au point posé sur la carte', nil, function()
    local blip = GetFirstBlipInfoId(8)
    if not DoesBlipExist(blip) then
        notifier('Pose d\'abord un point sur la carte (clic droit dans la carte du menu Échap).')
        return
    end
    local cible = GetBlipInfoIdCoord(blip)
    local ped = PlayerPedId()
    local entite = ped
    local vehicule = GetVehiclePedIsIn(ped, false)
    if vehicule ~= 0 and GetPedInVehicleSeat(vehicule, -1) == ped then
        entite = vehicule
    end
    local depart = GetEntityCoords(entite)

    DoScreenFadeOut(250)
    while not IsScreenFadedOut() do
        Wait(0)
    end
    FreezeEntityPosition(entite, true)

    -- On cherche le sol à plusieurs hauteurs, le temps que le décor se charge
    local trouve, solZ = false, 0.0
    for _, hauteur in ipairs({ 100.0, 150.0, 50.0, 0.0, 200.0, 250.0, 300.0, 400.0, 500.0, 600.0, 700.0, 800.0, 900.0 }) do
        SetEntityCoordsNoOffset(entite, cible.x, cible.y, hauteur, false, false, false)
        RequestCollisionAtCoord(cible.x, cible.y, hauteur)
        Wait(50)
        trouve, solZ = GetGroundZFor_3dCoord(cible.x, cible.y, hauteur, false)
        if trouve then
            break
        end
    end

    if trouve then
        SetEntityCoordsNoOffset(entite, cible.x, cible.y, solZ + 1.0, false, false, false)
    elseif entite == ped then
        SetEntityCoordsNoOffset(ped, cible.x, cible.y, 600.0, false, false, false)
        GiveWeaponToPed(ped, GetHashKey('GADGET_PARACHUTE'), 1, false, true)
        notifier('Sol introuvable : voici un ~b~parachute~s~ !')
    else
        SetEntityCoordsNoOffset(entite, depart.x, depart.y, depart.z, false, false, false)
        notifier('~r~Sol introuvable à cet endroit.')
    end

    FreezeEntityPosition(entite, false)
    DoScreenFadeIn(250)
end)

ajouterCommande('coords', 'Afficher ta position (pour remplir les fichiers config)', nil, function()
    local ped = PlayerPedId()
    local p = GetEntityCoords(ped)
    local texte = ('vector4(%.2f, %.2f, %.2f, %.1f)'):format(p.x, p.y, p.z, GetEntityHeading(ped))
    TriggerEvent('chat:addMessage', { color = { 120, 180, 255 }, args = { 'Position', texte } })
    print(texte) -- aussi dans la console F8, pour copier-coller
end)

---------------------------------------------------------------------------
-- Roleplay : /me et /do (gérés par le serveur, affichés ici)
---------------------------------------------------------------------------

ajouterCommande('me', 'Décrire une action de ton personnage', { { name = 'action', help = 'ex : ouvre la portière' } })
ajouterCommande('do', 'Décrire ce qu\'on voit autour de toi', { { name = 'description', help = 'ex : la porte est fermée à clé' } })
ajouterCommande('joueurs', 'Liste des joueurs connectés et de leur ID')

local textesRP = {} -- [ID serveur du joueur] = { texte, genre, fin }

RegisterNetEvent('coco_commandes:texteRP', function(serveurId, texte, genre)
    if type(serveurId) ~= 'number' or type(texte) ~= 'string' then
        return
    end
    local longueur = utf8.len(texte)
    if not longueur then
        return
    end
    -- Au-dessus de la tête on garde 90 caractères, le chat affiche le texte complet
    if longueur > 90 then
        texte = texte:sub(1, utf8.offset(texte, 91) - 1) .. '...'
    end
    textesRP[serveurId] = { texte = texte, genre = genre, fin = GetGameTimer() + Config.DureeRP * 1000 }
end)

-- Le jeu n'accepte que 99 octets par morceau de texte : on découpe sans casser les accents
local function ajouterTexteLong(texte)
    local morceau, taille = {}, 0
    for _, code in utf8.codes(texte) do
        local caractere = utf8.char(code)
        if taille + #caractere > 99 then
            AddTextComponentSubstringPlayerName(table.concat(morceau))
            morceau, taille = {}, 0
        end
        morceau[#morceau + 1] = caractere
        taille = taille + #caractere
    end
    if #morceau > 0 then
        AddTextComponentSubstringPlayerName(table.concat(morceau))
    end
end

local function dessiner3D(position, texte, r, g, b)
    SetDrawOrigin(position.x, position.y, position.z, 0)
    SetTextFont(4)
    SetTextScale(0.35, 0.35)
    SetTextColour(r, g, b, 255)
    SetTextOutline()
    SetTextCentre(true)
    BeginTextCommandDisplayText('CELL_EMAIL_BCON')
    ajouterTexteLong(texte)
    EndTextCommandDisplayText(0.0, 0.0)
    ClearDrawOrigin()
end

CreateThread(function()
    while true do
        local attente = 500
        local maintenant = GetGameTimer()
        for serveurId, info in pairs(textesRP) do
            if maintenant > info.fin then
                textesRP[serveurId] = nil
            else
                local joueur = GetPlayerFromServerId(serveurId)
                local ped = joueur ~= -1 and GetPlayerPed(joueur) or 0
                if ped ~= 0 and DoesEntityExist(ped) then
                    attente = 0
                    local tete = GetPedBoneCoords(ped, 31086, 0.0, 0.0, 0.0) + vector3(0.0, 0.0, 0.45)
                    if info.genre == 'me' then
                        dessiner3D(tete, '* ' .. info.texte .. ' *', 194, 122, 255)
                    else
                        dessiner3D(tete, info.texte, 255, 196, 92)
                    end
                end
            end
        end
        Wait(attente)
    end
end)

---------------------------------------------------------------------------
-- /aide et suggestions du chat
---------------------------------------------------------------------------

ajouterCommande('aide', 'Voir la liste des commandes', nil, function()
    local lignes = {}
    for _, commande in ipairs(listeAide) do
        local parametres = ''
        for _, parametre in ipairs(commande.parametres or {}) do
            parametres = parametres .. ' [' .. parametre.name .. ']'
        end
        lignes[#lignes + 1] = ('/%s%s : %s'):format(commande.nom, parametres, commande.description)
    end
    for _, ligne in ipairs(Config.AideEnPlus) do
        lignes[#lignes + 1] = ligne
    end
    TriggerEvent('chat:addMessage', {
        color = { 255, 200, 90 },
        multiline = true,
        args = { 'Commandes', table.concat(lignes, '\n') },
    })
end)

CreateThread(function()
    Wait(1000)
    for _, commande in ipairs(listeAide) do
        TriggerEvent('chat:addSuggestion', '/' .. commande.nom, commande.description, commande.parametres)
    end
end)

AddEventHandler('onResourceStop', function(ressource)
    if ressource == GetCurrentResourceName() then
        for _, commande in ipairs(listeAide) do
            TriggerEvent('chat:removeSuggestion', '/' .. commande.nom)
        end
    end
end)

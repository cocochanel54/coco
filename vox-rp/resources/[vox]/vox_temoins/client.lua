-- Témoins (côté joueur) : quand le joueur fait quelque chose de suspect, on cherche un PNJ
-- qui a vu ou entendu. Personne autour = pas d'alerte.

-- Couleurs du jeu regroupées en grandes familles (le témoin ne donne pas la teinte exacte)
local FAMILLES = {
    { m = 'noir', f = 'noire', ids = { 0, 1, 2, 11, 12, 15, 16, 21, 22, 141, 142, 143, 147 } },
    { m = 'gris', f = 'grise', ids = { 3, 4, 5, 6, 7, 8, 9, 10, 13, 14, 17, 18, 19, 20, 23, 24, 25, 26, 117, 118, 119, 120, 156 } },
    { m = 'rouge', f = 'rouge', ids = { 27, 28, 29, 30, 31, 32, 33, 34, 35, 39, 40, 43, 44, 45, 46, 47, 48, 150 } },
    { m = 'orange', f = 'orange', ids = { 36, 38, 41, 104, 123, 124, 130, 138 } },
    { m = 'jaune', f = 'jaune', ids = { 37, 42, 88, 89, 91, 126, 158, 159 } },
    { m = 'vert', f = 'verte', ids = { 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 92, 125, 128, 133, 139, 144, 151, 152, 155 } },
    { m = 'bleu', f = 'bleue', ids = { 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 127, 140, 146, 157 } },
    { m = 'violet', f = 'violette', ids = { 71, 72, 145, 148, 149 } },
    { m = 'rose', f = 'rose', ids = { 135, 136, 137 } },
    { m = 'marron', f = 'marron', ids = { 90, 94, 95, 96, 97, 98, 100, 101, 102, 103, 108, 109, 110, 114, 115, 129, 153 } },
    { m = 'beige', f = 'beige', ids = { 93, 99, 105, 106, 107, 113, 116, 154 } },
    { m = 'blanc', f = 'blanche', ids = { 111, 112, 121, 122, 131, 132, 134 } },
}

local COULEURS = {}
for _, famille in ipairs(FAMILLES) do
    for _, id in ipairs(famille.ids) do
        COULEURS[id] = famille
    end
end

-- Policiers, SWAT et militaires PNJ ne « témoignent » pas : ils interviennent directement
local TYPES_FORCES_ORDRE = { [6] = true, [27] = true, [29] = true }

local function decrire(ped, aVu)
    if not aVu then
        return 'Le témoin a seulement entendu, il n\'a rien vu.'
    end
    local personne = IsPedMale(ped) and 'Un homme' or 'Une femme'
    local vehicule = GetVehiclePedIsIn(ped, false)
    if vehicule == 0 then
        return personne .. ' à pied.'
    end
    local genre = Config.Vehicules[GetVehicleClass(vehicule)] or 'un véhicule'
    local feminin = genre:sub(1, 4) == 'une '
    local famille = COULEURS[(GetVehicleColours(vehicule))]
    local tirage = math.random()
    -- Un témoin n'est pas parfait : parfois il ne retient que le type, ou que la couleur
    if not famille or tirage < 0.2 then
        return ('%s à bord d\'%s.'):format(personne, genre)
    elseif tirage < 0.35 then
        return ('%s à bord d\'un véhicule %s.'):format(personne, famille.m)
    end
    return ('%s à bord d\'%s %s.'):format(personne, genre, feminin and famille.f or famille.m)
end

local function temoinPossible(ped, moi, exclus)
    return ped ~= moi
        and not exclus[ped]
        and not IsPedAPlayer(ped)
        and IsPedHuman(ped)
        and not IsPedDeadOrDying(ped, true)
        and not TYPES_FORCES_ORDRE[GetPedType(ped)]
end

-- Renvoie le meilleur témoin (celui qui a vu, sinon le plus proche) et s'il a vu la scène
local function chercherTemoin(regle, rayon, exclus)
    local moi = PlayerPedId()
    local position = GetEntityCoords(moi)
    local choix, choixVoit, meilleureDistance = nil, false, rayon
    for _, ped in ipairs(GetGamePool('CPed')) do
        if temoinPossible(ped, moi, exclus) then
            local distance = #(position - GetEntityCoords(ped))
            if distance < rayon then
                local voit = HasEntityClearLosToEntityInFront(ped, moi)
                    or (distance < 8.0 and HasEntityClearLosToEntity(ped, moi, 17))
                if voit or not regle.Vue then
                    if (voit and not choixVoit) or (voit == choixVoit and distance < meilleureDistance) then
                        choix, choixVoit, meilleureDistance = ped, voit, distance
                    end
                end
            end
        end
    end
    return choix, choixVoit
end

local function zoneSansAlerte(position)
    for _, zone in ipairs(Config.ZonesSansAlerte) do
        if #(position - zone.centre) < zone.rayon then
            return true
        end
    end
    return false
end

local recharges = {}

-- À utiliser depuis les autres scripts :
--   exports.vox_temoins:ActeSuspect('braquage')
--   exports.vox_temoins:ActeSuspect('vente_drogue', { exclure = { pedAcheteur } })
--   exports.vox_temoins:ActeSuspect('vente_drogue', { temoin = ped, certain = true }) -- ce PNJ appelle à coup sûr
-- Renvoie true si un témoin a été trouvé
local function acteSuspect(acte, options)
    options = options or {}
    local regle = Config.Actes[acte]
    if not regle then
        return false
    end
    local maintenant = GetGameTimer()
    if not options.certain and (recharges[acte] or 0) > maintenant then
        return false
    end
    local moi = PlayerPedId()
    local position = GetEntityCoords(moi)
    if zoneSansAlerte(position) then
        return false
    end

    local temoin, voit
    if options.temoin and DoesEntityExist(options.temoin) then
        temoin, voit = options.temoin, true
    else
        local exclus = {}
        for _, ped in ipairs(options.exclure or {}) do
            exclus[ped] = true
        end
        temoin, voit = chercherTemoin(regle, options.rayon or regle.Rayon, exclus)
    end
    if not temoin then
        return false -- personne n'a rien vu : pas d'alerte
    end

    recharges[acte] = maintenant + regle.Recharge * 1000
    TriggerServerEvent('vox_temoins:signaler', acte, {
        position = position,
        zone = GetLabelText(GetNameOfZone(position.x, position.y, position.z)),
        description = decrire(moi, voit),
        temoin = NetworkGetEntityIsNetworked(temoin) and NetworkGetNetworkIdFromEntity(temoin) or nil,
        certain = options.certain == true,
    })
    return true
end

exports('ActeSuspect', acteSuspect)

---------------------------------------------------------------------------
-- Coups de feu : détectés automatiquement
---------------------------------------------------------------------------

CreateThread(function()
    while true do
        local ped = PlayerPedId()
        if IsPedArmed(ped, 4) then -- 4 = armes à feu
            if IsPedShooting(ped) then
                local regle = Config.Actes.tirs
                -- Avec un silencieux, il faut être tout près pour entendre
                local rayon = IsPedCurrentWeaponSilenced(ped) and regle.Rayon * 0.25 or regle.Rayon
                acteSuspect('tirs', { rayon = rayon })
            end
            Wait(0)
        else
            Wait(500)
        end
    end
end)

---------------------------------------------------------------------------
-- Le témoin téléphone (chez le joueur qui « possède » ce PNJ dans le jeu)
---------------------------------------------------------------------------

RegisterNetEvent('vox_temoins:telephoner', function(idAppel, netId, delai)
    if not NetworkDoesNetworkIdExist(netId) then
        return
    end
    local ped = NetworkGetEntityFromNetworkId(netId)
    if ped == 0 or not DoesEntityExist(ped) then
        return
    end
    if NetworkHasControlOfEntity(ped) then
        ClearPedTasks(ped)
        TaskStartScenarioInPlace(ped, 'WORLD_HUMAN_STAND_MOBILE', 0, true)
    end
    local fin = GetGameTimer() + delai * 1000
    while GetGameTimer() < fin do
        if not DoesEntityExist(ped) or IsPedDeadOrDying(ped, true) then
            -- Témoin neutralisé avant la fin de l'appel : la police ne saura rien
            TriggerServerEvent('vox_temoins:annuler', idAppel)
            return
        end
        Wait(500)
    end
    if DoesEntityExist(ped) and NetworkHasControlOfEntity(ped) then
        ClearPedTasks(ped)
        TaskWanderStandard(ped, 10.0, 10)
    end
end)

---------------------------------------------------------------------------
-- Alertes reçues par la police
---------------------------------------------------------------------------

local alertes = {}

local function supprimer(id)
    for _, blip in ipairs(alertes[id] or {}) do
        if DoesBlipExist(blip) then
            RemoveBlip(blip)
        end
    end
    alertes[id] = nil
end

RegisterNetEvent('vox_temoins:alerte', function(alerte)
    if type(alerte) ~= 'table' or not alerte.position then
        return
    end
    lib.notify({
        title = alerte.titre,
        description = alerte.texte,
        type = 'warning',
        icon = 'phone-volume',
        duration = 15000,
        position = 'top',
    })
    PlaySoundFrontend(-1, 'Event_Start_Text', 'GTAO_FM_Events_Soundset', false)
    local p = alerte.position
    local zone = AddBlipForRadius(p.x, p.y, p.z, alerte.rayon)
    SetBlipColour(zone, 1)
    SetBlipAlpha(zone, 100)
    local point = AddBlipForCoord(p.x, p.y, p.z)
    SetBlipColour(point, 1)
    SetBlipFlashes(point, true)
    BeginTextCommandSetBlipName('STRING')
    AddTextComponentSubstringPlayerName(alerte.titre)
    EndTextCommandSetBlipName(point)
    alertes[alerte.id] = { zone, point }
    SetTimeout(Config.DureeAlerte * 1000, function()
        supprimer(alerte.id)
    end)
end)

AddEventHandler('onResourceStop', function(ressource)
    if ressource == GetCurrentResourceName() then
        for id in pairs(alertes) do
            supprimer(id)
        end
    end
end)

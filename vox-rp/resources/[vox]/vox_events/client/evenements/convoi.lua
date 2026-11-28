-- Convoi de fonds (côté joueur)
-- Le joueur qui « possède » le fourgon dans le jeu fait conduire le chauffeur PNJ.
-- Si un autre joueur en devient propriétaire, c'est lui qui reprend la conduite.

local cfg = Config.Evenements.convoi
local convoi = nil
local STYLE_CONDUITE = 786603 -- conduite normale, respecte la route

local function entiteReseau(netId)
    if netId and NetworkDoesNetworkIdExist(netId) then
        local entite = NetworkGetEntityFromNetworkId(netId)
        if entite ~= 0 and DoesEntityExist(entite) then
            return entite
        end
    end
    return nil
end

local function nettoyer()
    if not convoi then
        return
    end
    Outils.supprimerBlips(convoi.blips)
    Outils.cacherAide('convoi')
    convoi = nil
end

RegisterNetEvent('vox_events:convoi:debut', function(donnees)
    nettoyer()
    convoi = donnees
    convoi.taches = {}
    local point = donnees.position or donnees.approx
    convoi.blips = Outils.blipPoint(point, 2, 'Convoi de fonds', false)
end)

RegisterNetEvent('vox_events:convoi:vehicule', function(id, netVehicule, netPeds)
    if convoi and convoi.id == id then
        convoi.netVehicule, convoi.netPeds = netVehicule, netPeds
    end
end)

RegisterNetEvent('vox_events:convoi:position', function(id, position)
    if convoi and convoi.id == id and convoi.blips and convoi.blips[1] then
        SetBlipCoords(convoi.blips[1], position.x, position.y, position.z)
    end
end)

RegisterNetEvent('vox_events:convoi:etat', function(id, etat, sacs, delai)
    if not convoi or convoi.id ~= id then
        return
    end
    convoi.etat, convoi.sacs = etat, sacs
    if etat == 'charge' and delai then
        convoi.finCharge = GetGameTimer() + delai * 1000
    end
end)

RegisterNetEvent('vox_events:convoi:explosion', function(id, position)
    if not convoi or convoi.id ~= id then
        return
    end
    AddExplosion(position.x, position.y, position.z, 2, 0.3, true, false, 1.0)
    local vehicule = entiteReseau(convoi.netVehicule)
    if vehicule then
        SetVehicleDoorOpen(vehicule, 2, false, false)
        SetVehicleDoorOpen(vehicule, 3, false, false)
    end
end)

RegisterNetEvent('vox_events:convoi:fin', function(id)
    if convoi and convoi.id == id then
        nettoyer()
    end
end)

---------------------------------------------------------------------------
-- Conduite des PNJ (seulement chez le joueur qui en est propriétaire)
---------------------------------------------------------------------------

local function gererPnj()
    local vehicule = entiteReseau(convoi.netVehicule)
    if not vehicule then
        return
    end
    for index, netId in ipairs(convoi.netPeds or {}) do
        local ped = entiteReseau(netId)
        if ped then
            local proprietaire = NetworkGetEntityOwner(ped) == PlayerId()
            if not proprietaire then
                convoi.taches[netId] = nil
            elseif convoi.taches[netId] ~= convoi.etat and not IsPedDeadOrDying(ped, true) then
                convoi.taches[netId] = convoi.etat
                SetPedCombatAttributes(ped, 46, true) -- se bat jusqu'au bout
                SetPedFleeAttributes(ped, 0, false)
                if index == 1 and convoi.etat == 'route' then
                    -- Le chauffeur ignore ce qui se passe autour et roule vers la destination
                    SetBlockingOfNonTemporaryEvents(ped, true)
                    SetDriverAbility(ped, 1.0)
                    SetDriverAggressiveness(ped, 0.0)
                    SetPedKeepTask(ped, true)
                    local d = convoi.destination
                    TaskVehicleDriveToCoordLongrange(ped, vehicule, d.x, d.y, d.z, 18.0, STYLE_CONDUITE, 20.0)
                else
                    -- Fourgon arrêté : les gardes réagissent et se défendent
                    SetBlockingOfNonTemporaryEvents(ped, false)
                end
            end
        end
    end
end

---------------------------------------------------------------------------
-- Boucle principale : placement, conduite, interactions
---------------------------------------------------------------------------

local function dessinerTexte3D(position, texte)
    SetDrawOrigin(position.x, position.y, position.z, 0)
    SetTextFont(4)
    SetTextScale(0.45, 0.45)
    SetTextColour(255, 80, 80, 255)
    SetTextOutline()
    SetTextCentre(true)
    BeginTextCommandDisplayText('STRING')
    AddTextComponentSubstringPlayerName(texte)
    EndTextCommandDisplayText(0.0, 0.0)
    ClearDrawOrigin()
end

local function interagir(action, duree, label, anim)
    convoi.occupe = true
    Outils.cacherAide('convoi')
    local id = convoi.id
    local reussi = lib.progressBar({
        duration = duree * 1000,
        label = label,
        canCancel = true,
        disable = { move = true, car = true, combat = true },
        anim = anim,
    })
    if reussi and convoi and convoi.id == id then
        TriggerServerEvent('vox_events:convoi:' .. action, id)
    end
    if convoi then
        convoi.occupe = false
    end
end

CreateThread(function()
    local prochaineConduite = 0
    while true do
        local attente = 1000
        if convoi then
            local ped = PlayerPedId()
            local joueur = GetEntityCoords(ped)

            if convoi.etat == 'attente' and not convoi.placementDemande and Outils.distance2D(joueur, convoi.approx) < 300.0 then
                local a = convoi.approx
                local trouve, position, cap = GetClosestVehicleNodeWithHeading(a.x, a.y, a.z, 0, 3.0, 0)
                if trouve then
                    convoi.placementDemande = true
                    TriggerServerEvent('vox_events:convoi:placer', convoi.id, vector3(position.x, position.y, position.z), cap)
                end
            end

            if convoi.netVehicule and GetGameTimer() > prochaineConduite then
                prochaineConduite = GetGameTimer() + 1000
                gererPnj()
            end

            local vehicule = entiteReseau(convoi.netVehicule)
            if vehicule then
                local portes = GetOffsetFromEntityInWorldCoords(vehicule, 0.0, -3.8, 0.0)
                local distance = #(joueur - portes)
                if distance < 30.0 then
                    attente = 0
                end
                if convoi.etat == 'charge' and convoi.finCharge and distance < 40.0 then
                    local reste = math.max(0, math.ceil((convoi.finCharge - GetGameTimer()) / 1000))
                    dessinerTexte3D(portes + vector3(0.0, 0.0, 1.2), ('Explosion dans %d s, recule !'):format(reste))
                end
                local action
                if distance < 2.5 and not convoi.occupe and not IsPedInAnyVehicle(ped, false) then
                    if convoi.etat == 'immobilise' then
                        action = 'charge'
                        Outils.afficherAide('convoi', '[E] Poser la charge sur les portes')
                    elseif convoi.etat == 'ouvert' and (convoi.sacs or 0) > 0 then
                        action = 'sac'
                        Outils.afficherAide('convoi', ('[E] Prendre un sac (%d restant%s)'):format(convoi.sacs, convoi.sacs > 1 and 's' or ''))
                    end
                end
                if action then
                    if IsControlJustReleased(0, 38) then
                        if action == 'charge' then
                            interagir('charge', cfg.DureeCharge, 'Pose de la charge…', { dict = 'anim@heists@ornate_bank@thermal_charge', clip = 'thermal_charge' })
                        else
                            interagir('sac', cfg.DureeSac, 'Récupération du sac…', { dict = 'anim@heists@ornate_bank@grab_cash', clip = 'grab' })
                        end
                    end
                else
                    Outils.cacherAide('convoi')
                end
            end
        end
        Wait(attente)
    end
end)

AddEventHandler('onResourceStop', function(ressource)
    if ressource == GetCurrentResourceName() then
        nettoyer()
    end
end)

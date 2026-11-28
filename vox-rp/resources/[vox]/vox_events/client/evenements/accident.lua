-- Accident de la route (côté joueur)
-- Les épaves et les blessés sont créés chez chaque joueur proche, aux mêmes endroits.
-- C'est le serveur qui décide de ce qui est soigné ou dégagé.

local cfg = Config.Evenements.accident
local accident = nil

local function supprimerEntites()
    if not accident or not accident.entites then
        return
    end
    for _, entite in pairs(accident.entites.vehicules) do
        if DoesEntityExist(entite) then
            DeleteVehicle(entite)
        end
    end
    for _, entite in pairs(accident.entites.victimes) do
        if DoesEntityExist(entite) then
            DeletePed(entite)
        end
    end
    accident.entites = nil
end

local function nettoyer()
    if not accident then
        return
    end
    supprimerEntites()
    Outils.supprimerBlips(accident.blips)
    Outils.cacherAide('accident')
    accident = nil
end

RegisterNetEvent('vox_events:accident:debut', function(donnees)
    nettoyer()
    accident = donnees
    accident.blips = Outils.blipPoint(donnees.approx, 1, 'Accident de la route', false)
end)

RegisterNetEvent('vox_events:accident:site', function(id, site)
    if accident and accident.id == id then
        accident.site = site
    end
end)

RegisterNetEvent('vox_events:accident:maj', function(id, genre, index, etat)
    if not accident or accident.id ~= id then
        return
    end
    local liste = genre == 'victime' and accident.victimes or accident.vehicules
    if not liste[index] then
        return
    end
    liste[index].etat = etat
    local entites = accident.entites
    if not entites then
        return
    end
    if genre == 'victime' and entites.victimes[index] then
        local ped = entites.victimes[index]
        entites.victimes[index] = nil
        -- Le blessé est évacué : il disparaît doucement
        CreateThread(function()
            for alpha = 255, 0, -15 do
                if not DoesEntityExist(ped) then
                    return
                end
                SetEntityAlpha(ped, alpha, false)
                Wait(50)
            end
            DeletePed(ped)
        end)
    elseif genre == 'vehicule' and entites.vehicules[index] then
        DeleteVehicle(entites.vehicules[index])
        entites.vehicules[index] = nil
    end
end)

RegisterNetEvent('vox_events:accident:fin', function(id)
    if accident and accident.id == id then
        nettoyer()
    end
end)

---------------------------------------------------------------------------
-- Création de la scène
---------------------------------------------------------------------------

local function creerEpave(place, modele)
    local hash = Outils.chargerModele(modele) or Outils.chargerModele('blista')
    if not hash then
        return nil
    end
    local p = place.position
    local z = Outils.sol(p.x, p.y, p.z)
    local vehicule = CreateVehicle(hash, p.x, p.y, z + 0.5, place.cap, false, false)
    SetModelAsNoLongerNeeded(hash)
    SetVehicleOnGroundProperly(vehicule)
    SetVehicleDamage(vehicule, 0.0, 1.2, 0.2, 900.0, 900.0, true)
    SetVehicleDamage(vehicule, 0.7, 0.9, 0.3, 600.0, 700.0, true)
    SmashVehicleWindow(vehicule, 0)
    SmashVehicleWindow(vehicule, 1)
    SetVehicleDoorOpen(vehicule, 0, false, true)
    SetVehicleEngineHealth(vehicule, 120.0) -- moteur qui fume
    SetVehicleIndicatorLights(vehicule, 0, true)
    SetVehicleIndicatorLights(vehicule, 1, true)
    SetVehicleDoorsLocked(vehicule, 2)
    SetEntityInvincible(vehicule, true)
    FreezeEntityPosition(vehicule, true)
    return vehicule
end

local function creerBlesse(place, modele)
    local hash = Outils.chargerModele(modele) or Outils.chargerModele('a_m_y_business_01')
    if not hash then
        return nil
    end
    local p = place.position
    local z = Outils.sol(p.x, p.y, p.z)
    local ped = CreatePed(4, hash, p.x, p.y, z + 1.0, place.cap, false, false)
    SetModelAsNoLongerNeeded(hash)
    SetEntityInvincible(ped, true)
    SetBlockingOfNonTemporaryEvents(ped, true)
    SetPedCanRagdoll(ped, false)
    ApplyPedDamagePack(ped, 'BigHitByVehicle', 0.0, 1.0)
    if Outils.chargerAnim('combat@damage@writhe') then
        TaskPlayAnim(ped, 'combat@damage@writhe', 'writhe_loop', 8.0, -8.0, -1, 1, 0, false, false, false)
    end
    SetTimeout(1500, function()
        if DoesEntityExist(ped) then
            FreezeEntityPosition(ped, true)
        end
    end)
    return ped
end

local function creerScene()
    accident.entites = { vehicules = {}, victimes = {} }
    for i, vehicule in ipairs(accident.vehicules) do
        local place = accident.site.vehicules[i]
        if vehicule.etat == 'present' and place then
            accident.entites.vehicules[i] = creerEpave(place, vehicule.modele)
        end
    end
    for i, victime in ipairs(accident.victimes) do
        local place = accident.site.victimes[i]
        if victime.etat == 'blessee' and place then
            accident.entites.victimes[i] = creerBlesse(place, victime.modele)
        end
    end
end

---------------------------------------------------------------------------
-- Interactions
---------------------------------------------------------------------------

local function actionProche(joueur)
    for i, victime in ipairs(accident.victimes) do
        local place = accident.site.victimes[i]
        if victime.etat == 'blessee' and place and Outils.distance2D(joueur, place.position) < 2.5 then
            return 'victime', i
        end
    end
    for i, vehicule in ipairs(accident.vehicules) do
        local place = accident.site.vehicules[i]
        if vehicule.etat == 'present' and place and Outils.distance2D(joueur, place.position) < 4.5 then
            return 'vehicule', i
        end
    end
    return nil
end

local function agir(genre, index)
    accident.occupe = true
    Outils.cacherAide('accident')
    local id = accident.id
    local options
    if genre == 'victime' then
        options = {
            duration = cfg.DureeSoin * 1000,
            label = 'Premiers soins…',
            anim = { dict = 'amb@medic@standing@tendtodead@idle_a', clip = 'idle_a' },
        }
    else
        options = {
            duration = cfg.DureeDegagement * 1000,
            label = 'Dégagement de l\'épave…',
            anim = { dict = 'mini@repair', clip = 'fixing_a_ped' },
        }
    end
    options.canCancel = true
    options.disable = { move = true, car = true, combat = true }
    local reussi = lib.progressBar(options)
    if reussi and accident and accident.id == id then
        TriggerServerEvent(genre == 'victime' and 'vox_events:accident:soigner' or 'vox_events:accident:degager', id, index)
    end
    if accident then
        accident.occupe = false
    end
end

CreateThread(function()
    while true do
        local attente = 1000
        if accident then
            local ped = PlayerPedId()
            local joueur = GetEntityCoords(ped)
            if not accident.site then
                -- Le premier joueur proche trouve la route la plus proche du lieu
                if not accident.placementDemande and Outils.distance2D(joueur, accident.approx) < 200.0 then
                    local a = accident.approx
                    local trouve, position, cap = GetClosestVehicleNodeWithHeading(a.x, a.y, a.z, 0, 3.0, 0)
                    if trouve then
                        accident.placementDemande = true
                        TriggerServerEvent('vox_events:accident:placer', accident.id, vector3(position.x, position.y, position.z), cap)
                    end
                end
            else
                local distance = Outils.distance2D(joueur, accident.site.position)
                if distance < 150.0 and not accident.entites then
                    creerScene()
                elseif distance > 200.0 and accident.entites then
                    supprimerEntites()
                end
                if distance < 30.0 then
                    attente = 0
                    local genre, index = actionProche(joueur)
                    if genre and not accident.occupe and not IsPedInAnyVehicle(ped, false) then
                        Outils.afficherAide('accident', genre == 'victime' and '[E] Prodiguer les premiers soins' or '[E] Dégager l\'épave')
                        if IsControlJustReleased(0, 38) then
                            agir(genre, index)
                        end
                    else
                        Outils.cacherAide('accident')
                    end
                else
                    Outils.cacherAide('accident')
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

-- Colis mystère (côté joueur)

local cfg = Config.Evenements.colis
local colis = nil -- événement en cours, vu par ce joueur

local function supprimerObjet()
    if colis and colis.objet and DoesEntityExist(colis.objet) then
        DeleteObject(colis.objet)
    end
    if colis then
        colis.objet = nil
    end
end

local function nettoyer()
    if not colis then
        return
    end
    supprimerObjet()
    Outils.supprimerBlips(colis.blips)
    Outils.cacherAide('colis')
    colis = nil
end

RegisterNetEvent('vox_events:colis:debut', function(donnees)
    nettoyer()
    colis = donnees
    colis.blips = Outils.blipZone(donnees.zone, donnees.rayon, 5, 'Colis mystère')
end)

RegisterNetEvent('vox_events:colis:position', function(id, position)
    if colis and colis.id == id then
        colis.position = position
    end
end)

RegisterNetEvent('vox_events:colis:fin', function(id)
    if colis and colis.id == id then
        nettoyer()
    end
end)

-- Le trottoir le plus proche du point approximatif (le décor est chargé car on est à côté)
local function trouverEmplacement(approx)
    local z = Outils.sol(approx.x, approx.y, approx.z)
    local trouve, position = GetSafeCoordForPed(approx.x, approx.y, z, true, 16)
    if not trouve then
        trouve, position = GetSafeCoordForPed(approx.x, approx.y, z, false, 0)
    end
    if not trouve then
        return vector3(colis.zone.x, colis.zone.y, colis.zone.z)
    end
    return vector3(position.x, position.y, position.z)
end

local function creerObjet(position)
    local hash = Outils.chargerModele(cfg.Modele) or Outils.chargerModele('prop_cs_cardbox_01')
    if not hash then
        return nil
    end
    local objet = CreateObject(hash, position.x, position.y, position.z, false, false, false)
    PlaceObjectOnGroundProperly(objet)
    FreezeEntityPosition(objet, true)
    SetEntityInvincible(objet, true)
    SetModelAsNoLongerNeeded(hash)
    return objet
end

local function ouvrir()
    colis.ouverture = true
    Outils.cacherAide('colis')
    local id = colis.id
    local reussi = lib.progressBar({
        duration = cfg.DureeOuverture * 1000,
        label = 'Ouverture du colis…',
        canCancel = true,
        disable = { move = true, car = true, combat = true },
        anim = { dict = 'random@domestic', clip = 'pickup_low' },
    })
    if reussi and colis and colis.id == id then
        TriggerServerEvent('vox_events:colis:ouvrir', id)
    end
    if colis then
        colis.ouverture = false
    end
end

CreateThread(function()
    while true do
        local attente = 1000
        if colis then
            local joueur = GetEntityCoords(PlayerPedId())
            if not colis.position then
                -- Personne n'a encore posé le colis : le premier qui passe à côté le fait
                if not colis.placementDemande and Outils.distance2D(joueur, colis.approx) < 150.0 then
                    colis.placementDemande = true
                    TriggerServerEvent('vox_events:colis:placer', colis.id, trouverEmplacement(colis.approx))
                end
            else
                local distance = #(joueur - colis.position)
                if distance < 120.0 and not colis.objet then
                    colis.objet = creerObjet(colis.position)
                elseif distance > 150.0 and colis.objet then
                    supprimerObjet()
                end
                if distance < 25.0 then
                    attente = 0
                end
                if distance < 2.0 and not colis.ouverture and not IsPedInAnyVehicle(PlayerPedId(), false) then
                    Outils.afficherAide('colis', '[E] Ouvrir le colis')
                    if IsControlJustReleased(0, 38) then
                        ouvrir()
                    end
                else
                    Outils.cacherAide('colis')
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

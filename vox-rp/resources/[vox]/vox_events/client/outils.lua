-- Outils côté joueur partagés par les événements

Outils = {}

function Outils.distance2D(a, b)
    local dx, dy = a.x - b.x, a.y - b.y
    return math.sqrt(dx * dx + dy * dy)
end

function Outils.chargerModele(modele)
    local hash = type(modele) == 'number' and modele or GetHashKey(modele)
    if not IsModelInCdimage(hash) then
        return nil
    end
    RequestModel(hash)
    local limite = GetGameTimer() + 5000
    while not HasModelLoaded(hash) do
        if GetGameTimer() > limite then
            return nil
        end
        Wait(0)
    end
    return hash
end

function Outils.chargerAnim(dictionnaire)
    RequestAnimDict(dictionnaire)
    local limite = GetGameTimer() + 3000
    while not HasAnimDictLoaded(dictionnaire) do
        if GetGameTimer() > limite then
            return false
        end
        Wait(0)
    end
    return true
end

-- Hauteur du sol sous un point (le décor doit être chargé, donc le joueur pas trop loin)
function Outils.sol(x, y, zDepart)
    for _, decalage in ipairs({ 5.0, 20.0, 50.0 }) do
        local trouve, z = GetGroundZFor_3dCoord(x, y, zDepart + decalage, false)
        if trouve then
            return z
        end
    end
    return zDepart
end

local function nommerBlip(blip, nom)
    BeginTextCommandSetBlipName('STRING')
    AddTextComponentSubstringPlayerName(nom)
    EndTextCommandSetBlipName(blip)
end

-- Cercle sur la carte + point au centre. Renvoie une table à donner à Outils.supprimerBlips
function Outils.blipZone(centre, rayon, couleur, nom)
    local zone = AddBlipForRadius(centre.x, centre.y, centre.z, rayon)
    SetBlipColour(zone, couleur)
    SetBlipAlpha(zone, 110)
    local point = AddBlipForCoord(centre.x, centre.y, centre.z)
    SetBlipColour(point, couleur)
    SetBlipScale(point, 0.9)
    nommerBlip(point, nom)
    return { zone, point }
end

function Outils.blipPoint(position, couleur, nom, itineraire)
    local blip = AddBlipForCoord(position.x, position.y, position.z)
    SetBlipColour(blip, couleur)
    SetBlipScale(blip, 1.0)
    SetBlipFlashes(blip, true)
    nommerBlip(blip, nom)
    if itineraire then
        SetBlipRoute(blip, true)
        SetBlipRouteColour(blip, couleur)
    end
    return { blip }
end

function Outils.supprimerBlips(blips)
    for _, blip in ipairs(blips or {}) do
        if DoesBlipExist(blip) then
            RemoveBlip(blip)
        end
    end
end

-- Texte d'aide « [E] ... » : un seul affiché à la fois
local aideActuelle = nil

function Outils.afficherAide(cle, texte)
    if aideActuelle ~= cle then
        aideActuelle = cle
        lib.showTextUI(texte, { position = 'right-center' })
    end
end

function Outils.cacherAide(cle)
    if aideActuelle == cle then
        aideActuelle = nil
        lib.hideTextUI()
    end
end

---------------------------------------------------------------------------
-- Alertes pour les métiers en service (point clignotant + itinéraire)
---------------------------------------------------------------------------

local alertes = {}

RegisterNetEvent('vox_events:alerte', function(alerte)
    if type(alerte) ~= 'table' or not alerte.coords then
        return
    end
    Outils.supprimerBlips(alertes[alerte.id])
    alertes[alerte.id] = Outils.blipPoint(alerte.coords, alerte.couleur or 1, alerte.titre or 'Alerte', true)
    PlaySoundFrontend(-1, 'Event_Start_Text', 'GTAO_FM_Events_Soundset', false)
    local id = alerte.id
    SetTimeout((alerte.duree or 300) * 1000, function()
        Outils.supprimerBlips(alertes[id])
        alertes[id] = nil
    end)
end)

RegisterNetEvent('vox_events:alerteFin', function(id)
    Outils.supprimerBlips(alertes[id])
    alertes[id] = nil
end)

-- À l'arrivée sur le serveur, on demande les événements déjà en cours
CreateThread(function()
    while not NetworkIsPlayerActive(PlayerId()) do
        Wait(500)
    end
    Wait(5000)
    TriggerServerEvent('vox_events:demanderEtat')
end)

AddEventHandler('onResourceStop', function(ressource)
    if ressource == GetCurrentResourceName() then
        for _, blips in pairs(alertes) do
            Outils.supprimerBlips(blips)
        end
        if aideActuelle then
            lib.hideTextUI()
        end
    end
end)

-- Coco Admin (côté joueur) : météo et heure du serveur, téléportation, noclip, annonces

local function notifier(message)
    BeginTextCommandThefeedPost('STRING')
    AddTextComponentSubstringPlayerName(message)
    EndTextCommandThefeedPostTicker(false, true)
end

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

---------------------------------------------------------------------------
-- Météo et heure : le serveur décide, tout le monde voit la même chose
---------------------------------------------------------------------------

local meteoActuelle = nil

local function appliquerMeteo(meteo, transition)
    if type(meteo) ~= 'string' or meteo == meteoActuelle then
        return
    end
    meteoActuelle = meteo
    ClearOverrideWeather()
    ClearWeatherTypePersist()
    if transition then
        SetWeatherTypeOvertimePersist(meteo, 15.0)
    else
        SetWeatherTypePersist(meteo)
        SetWeatherTypeNow(meteo)
        SetWeatherTypeNowPersist(meteo)
    end
    -- Traces dans la neige pour Noël
    local neige = meteo == 'XMAS'
    SetForceVehicleTrails(neige)
    SetForcePedFootstepsTracks(neige)
end

local function appliquerHeure(total)
    if type(total) ~= 'number' then
        return
    end
    total = math.floor(total)
    NetworkOverrideClockTime(total // 60, total % 60, 0)
end

AddStateBagChangeHandler('meteo', 'global', function(_, _, valeur)
    appliquerMeteo(valeur, true)
end)

AddStateBagChangeHandler('heure', 'global', function(_, _, valeur)
    appliquerHeure(valeur)
end)

CreateThread(function()
    while true do
        if not meteoActuelle then
            appliquerMeteo(GlobalState.meteo, false)
        end
        appliquerHeure(GlobalState.heure)
        Wait(1000)
    end
end)

---------------------------------------------------------------------------
-- Téléportation et réanimation
---------------------------------------------------------------------------

RegisterNetEvent('coco_admin:teleporter', function(coords)
    if not coords or not coords.x then
        return
    end
    local ped = PlayerPedId()
    DoScreenFadeOut(200)
    while not IsScreenFadedOut() do
        Wait(0)
    end
    RequestCollisionAtCoord(coords.x, coords.y, coords.z)
    SetPedCoordsKeepVehicle(ped, coords.x, coords.y, coords.z + 0.5)
    local limite = GetGameTimer() + 3000
    while not HasCollisionLoadedAroundEntity(ped) and GetGameTimer() < limite do
        Wait(0)
    end
    DoScreenFadeIn(200)
end)

RegisterNetEvent('coco_admin:reanimer', function()
    local ped = PlayerPedId()
    if IsEntityDead(ped) then
        local position = GetEntityCoords(ped)
        NetworkResurrectLocalPlayer(position.x, position.y, position.z, GetEntityHeading(ped), 0, false)
        ped = PlayerPedId()
        ClearPedTasksImmediately(ped)
    end
    SetEntityHealth(ped, GetEntityMaxHealth(ped))
    ClearPedBloodDamage(ped)
    notifier('Tu as été ~g~soigné~s~ par un admin.')
end)

---------------------------------------------------------------------------
-- Noclip
---------------------------------------------------------------------------

local noclip = false

-- Touches : déplacements habituels, Espace (monter), Ctrl (descendre), Maj (rapide)
local TOUCHES = { avant = 32, arriere = 33, gauche = 34, droite = 35, monter = 22, descendre = 36, rapide = 21 }

local function boucleNoclip()
    local ped = PlayerPedId()
    SetEntityCollision(ped, false, false)
    FreezeEntityPosition(ped, true)
    SetEntityInvincible(ped, true)
    SetEntityVisible(ped, false, false)

    while noclip and DoesEntityExist(ped) and not IsEntityDead(ped) do
        for _, touche in pairs(TOUCHES) do
            DisableControlAction(0, touche, true)
        end
        SetLocalPlayerVisibleLocally(true)

        local rotation = GetGameplayCamRot(2)
        local cap, inclinaison = math.rad(rotation.z), math.rad(rotation.x)
        local horizontal = math.abs(math.cos(inclinaison))
        local avant = vector3(-math.sin(cap) * horizontal, math.cos(cap) * horizontal, math.sin(inclinaison))
        local droite = vector3(math.cos(cap), math.sin(cap), 0.0)
        local haut = vector3(0.0, 0.0, 1.0)

        local direction = vector3(0.0, 0.0, 0.0)
        if IsDisabledControlPressed(0, TOUCHES.avant) then direction = direction + avant end
        if IsDisabledControlPressed(0, TOUCHES.arriere) then direction = direction - avant end
        if IsDisabledControlPressed(0, TOUCHES.droite) then direction = direction + droite end
        if IsDisabledControlPressed(0, TOUCHES.gauche) then direction = direction - droite end
        if IsDisabledControlPressed(0, TOUCHES.monter) then direction = direction + haut end
        if IsDisabledControlPressed(0, TOUCHES.descendre) then direction = direction - haut end

        local vitesse = IsDisabledControlPressed(0, TOUCHES.rapide) and 100.0 or 20.0 -- mètres par seconde
        local position = GetEntityCoords(ped) + direction * (vitesse * GetFrameTime())
        SetEntityCoordsNoOffset(ped, position.x, position.y, position.z, true, true, true)
        SetEntityHeading(ped, rotation.z)
        Wait(0)
    end

    noclip = false
    SetEntityCollision(ped, true, true)
    FreezeEntityPosition(ped, false)
    SetEntityInvincible(ped, false)
    SetEntityVisible(ped, true, false)
end

RegisterNetEvent('coco_admin:noclip', function()
    if noclip then
        noclip = false
        notifier('Noclip ~r~désactivé~s~.')
        return
    end
    if IsPedInAnyVehicle(PlayerPedId(), false) then
        notifier('~r~Sors du véhicule pour utiliser le noclip.')
        return
    end
    noclip = true
    notifier('Noclip ~g~activé~s~ : touches de déplacement, Espace pour monter, Ctrl pour descendre, Maj pour aller vite.')
    CreateThread(boucleNoclip)
end)

---------------------------------------------------------------------------
-- Annonce en haut de l'écran
---------------------------------------------------------------------------

local annonceActuelle = 0

RegisterNetEvent('coco_admin:annonce', function(texte)
    if type(texte) ~= 'string' or not utf8.len(texte) then
        return
    end
    annonceActuelle = annonceActuelle + 1
    local numero = annonceActuelle
    local fin = GetGameTimer() + 10000
    PlaySoundFrontend(-1, 'Event_Message_Purple', 'GTAO_FM_Events_Soundset', false)

    CreateThread(function()
        while numero == annonceActuelle and GetGameTimer() < fin do
            DrawRect(0.5, 0.11, 0.44, 0.12, 0, 0, 0, 190)

            SetTextFont(4)
            SetTextScale(0.65, 0.65)
            SetTextColour(255, 80, 80, 255)
            SetTextCentre(true)
            SetTextOutline()
            BeginTextCommandDisplayText('STRING')
            AddTextComponentSubstringPlayerName('ANNONCE')
            EndTextCommandDisplayText(0.5, 0.055)

            SetTextFont(4)
            SetTextScale(0.42, 0.42)
            SetTextColour(255, 255, 255, 255)
            SetTextCentre(true)
            SetTextWrap(0.29, 0.71)
            SetTextOutline()
            BeginTextCommandDisplayText('CELL_EMAIL_BCON')
            ajouterTexteLong(texte)
            EndTextCommandDisplayText(0.5, 0.1)

            Wait(0)
        end
    end)
end)

---------------------------------------------------------------------------
-- Suggestions du chat pour les admins
---------------------------------------------------------------------------

local suggestionsAjoutees = {}

RegisterNetEvent('coco_admin:suggestions', function(liste)
    for _, commande in ipairs(liste) do
        TriggerEvent('chat:addSuggestion', '/' .. commande.nom, commande.description, commande.parametres)
        suggestionsAjoutees[#suggestionsAjoutees + 1] = '/' .. commande.nom
    end
end)

CreateThread(function()
    Wait(2000)
    TriggerServerEvent('coco_admin:demanderSuggestions')
end)

AddEventHandler('onResourceStop', function(ressource)
    if ressource ~= GetCurrentResourceName() then
        return
    end
    for _, nom in ipairs(suggestionsAjoutees) do
        TriggerEvent('chat:removeSuggestion', nom)
    end
    if noclip then
        noclip = false
        local ped = PlayerPedId()
        SetEntityCollision(ped, true, true)
        FreezeEntityPosition(ped, false)
        SetEntityInvincible(ped, false)
        SetEntityVisible(ped, true, false)
    end
end)

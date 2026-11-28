-- Panne de courant et tempête, quand vox_events gère lui-même la météo (Config.Meteo = 'vox').
-- Avec qb-weathersync, c'est lui qui applique tout et ce fichier ne fait rien.

local panne = false
local meteoForcee = nil
local finTransition = 0

local function majPanne(valeur)
    local nouvelle = valeur == true
    if nouvelle == panne then
        return
    end
    panne = nouvelle
    SetArtificialLightsState(panne)
    SetArtificialLightsStateAffectsVehicles(false) -- les phares des voitures restent allumés
end

local function majMeteo(valeur)
    if type(valeur) ~= 'string' then
        valeur = nil
    end
    if valeur == meteoForcee then
        return
    end
    meteoForcee = valeur
    ClearOverrideWeather()
    ClearWeatherTypePersist()
    if meteoForcee then
        SetWeatherTypeOvertimePersist(meteoForcee, 30.0)
        finTransition = GetGameTimer() + 31000
    else
        -- La tempête s'éloigne, puis la météo normale du serveur reprend la main
        SetWeatherTypeOvertimePersist(Config.Evenements.tempete.MeteoApres, 30.0)
        SetTimeout(31000, function()
            if not meteoForcee then
                ClearWeatherTypePersist()
            end
        end)
    end
end

AddStateBagChangeHandler('voxPanne', 'global', function(_, _, valeur)
    majPanne(valeur)
end)

AddStateBagChangeHandler('voxMeteo', 'global', function(_, _, valeur)
    majMeteo(valeur)
end)

CreateThread(function()
    Wait(2000)
    majPanne(GlobalState.voxPanne)
    majMeteo(GlobalState.voxMeteo)
    while true do
        -- On réapplique régulièrement, au cas où un autre script changerait la lumière ou la météo
        if panne then
            SetArtificialLightsState(true)
        end
        if meteoForcee and GetGameTimer() > finTransition then
            SetWeatherTypeNowPersist(meteoForcee)
        end
        Wait((panne or meteoForcee) and 500 or 2000)
    end
end)

AddEventHandler('onResourceStop', function(ressource)
    if ressource == GetCurrentResourceName() then
        if panne then
            SetArtificialLightsState(false)
        end
        if meteoForcee then
            ClearOverrideWeather()
            ClearWeatherTypePersist()
        end
    end
end)

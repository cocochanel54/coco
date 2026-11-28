-- Tout ce qui dépend de Qbox, ox_lib, ox_inventory ou de la météo est ici.
-- Si une de ces ressources change, c'est le seul fichier à adapter.

Bridge = {}

local function joueurQbox(src)
    return exports.qbx_core:GetPlayer(src)
end

---------------------------------------------------------------------------
-- Métiers
---------------------------------------------------------------------------

-- role : 'police', 'ems' ou 'mecano' (voir Config.Metiers)
function Bridge.estEnService(src, role)
    local joueur = joueurQbox(src)
    local job = joueur and joueur.PlayerData and joueur.PlayerData.job
    if not job or not job.onduty then
        return false
    end
    local regle = Config.Metiers[role]
    if not regle then
        return false
    end
    for _, typeMetier in ipairs(regle.types or {}) do
        if job.type == typeMetier then
            return true
        end
    end
    for _, nom in ipairs(regle.noms or {}) do
        if job.name == nom then
            return true
        end
    end
    return false
end

function Bridge.joueursEnService(role)
    local liste = {}
    for _, idTexte in ipairs(GetPlayers()) do
        local id = tonumber(idTexte)
        if Bridge.estEnService(id, role) then
            liste[#liste + 1] = id
        end
    end
    return liste
end

function Bridge.compterEnService(role)
    return #Bridge.joueursEnService(role)
end

---------------------------------------------------------------------------
-- Récompenses
---------------------------------------------------------------------------

function Bridge.donnerArgent(src, montant, raison)
    local joueur = joueurQbox(src)
    if not joueur then
        return false
    end
    joueur.Functions.AddMoney(Config.TypeArgent, montant, raison)
    return true
end

-- Renvoie false si l'objet n'existe pas ou si l'inventaire est plein
function Bridge.donnerObjet(src, nom, quantite)
    if GetResourceState('ox_inventory') ~= 'started' then
        return false
    end
    local ok, peutPorter = pcall(function()
        return exports.ox_inventory:CanCarryItem(src, nom, quantite)
    end)
    if not ok or not peutPorter then
        return false
    end
    local ok2, reussi = pcall(function()
        return exports.ox_inventory:AddItem(src, nom, quantite)
    end)
    return ok2 and reussi and true or false
end

-- Renvoie true seulement si l'objet a bien été retiré
function Bridge.retirerObjet(src, nom, quantite)
    if GetResourceState('ox_inventory') ~= 'started' then
        return false
    end
    local ok, retire = pcall(function()
        return exports.ox_inventory:RemoveItem(src, nom, quantite)
    end)
    return ok and retire and true or false
end

---------------------------------------------------------------------------
-- Messages
---------------------------------------------------------------------------

-- cible : ID d'un joueur, ou -1 pour tout le monde. donnees : options de lib.notify
function Bridge.notifier(cible, donnees)
    donnees.position = donnees.position or 'top'
    TriggerClientEvent('ox_lib:notify', cible, donnees)
end

-- Réponse aux commandes du staff (console ou chat)
function Bridge.repondre(src, texte)
    if src == 0 then
        print('[vox_events] ' .. texte)
        return
    end
    TriggerClientEvent('chat:addMessage', src, {
        color = { 120, 180, 255 },
        multiline = true,
        args = { 'Événements', texte },
    })
end

-- Alerte envoyée aux joueurs en service (point sur la carte + itinéraire)
function Bridge.alerter(roles, alerte)
    local dejaPrevenus = {}
    for _, role in ipairs(roles) do
        for _, id in ipairs(Bridge.joueursEnService(role)) do
            if not dejaPrevenus[id] then
                dejaPrevenus[id] = true
                Bridge.notifier(id, {
                    title = alerte.titre,
                    description = alerte.texte,
                    type = 'warning',
                    icon = alerte.icone,
                    duration = 12000,
                })
                TriggerClientEvent('vox_events:alerte', id, alerte)
            end
        end
    end
end

---------------------------------------------------------------------------
-- Météo et électricité
---------------------------------------------------------------------------

Bridge.meteo = {}

local function ressourceMeteo()
    local choix = Config.Meteo.Ressource
    if choix ~= 'auto' then
        return choix
    end
    if GetResourceState('qb-weathersync') == 'started' then
        return 'qb-weathersync'
    end
    return 'vox'
end

function Bridge.meteo.panne(actif)
    if ressourceMeteo() == 'qb-weathersync' then
        local ok = pcall(function()
            exports['qb-weathersync']:setBlackout(actif)
        end)
        if ok then
            return
        end
        print('[vox_events] qb-weathersync ne répond pas, vox_events gère la panne lui-même.')
    end
    GlobalState.voxPanne = actif
end

-- actif = true : on force la météo donnée ; false : on rend la main avec la météo donnée
function Bridge.meteo.tempete(actif, meteo)
    if ressourceMeteo() == 'qb-weathersync' then
        local ok = pcall(function()
            exports['qb-weathersync']:setDynamicWeather(not actif)
            exports['qb-weathersync']:setWeather(meteo)
        end)
        if ok then
            return
        end
        print('[vox_events] qb-weathersync ne répond pas, vox_events gère la météo lui-même.')
    end
    GlobalState.voxMeteo = actif and meteo or false
end

-- Si la ressource redémarre, on repart d'un état propre
GlobalState.voxPanne = false
GlobalState.voxMeteo = false

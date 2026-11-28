-- Coco Base (côté joueur) : première apparition, mort et hôpital, HUD, /skin

-- Si la ressource redémarre en cours de partie, le joueur est déjà en jeu
local aDejaSpawn = LocalPlayer.state.cocoSpawn == true
local apparitionEnCours = false

local function notifier(message)
    BeginTextCommandThefeedPost('STRING')
    AddTextComponentSubstringPlayerName(message)
    EndTextCommandThefeedPostTicker(false, true)
end

local function dessinerTexte(texte, x, y, taille, options)
    options = options or {}
    SetTextFont(4)
    SetTextScale(taille, taille)
    SetTextColour(255, 255, 255, 235)
    SetTextOutline()
    if options.centre then
        SetTextCentre(true)
    end
    if options.droite then
        SetTextRightJustify(true)
        SetTextWrap(0.0, x)
    end
    BeginTextCommandDisplayText('STRING')
    AddTextComponentSubstringPlayerName(texte)
    EndTextCommandDisplayText(x, y)
end

-- Le personnage choisi avec /skin, sinon un au hasard dans la config (puis retenu)
local function modeleDuJoueur()
    local nom = GetResourceKvpString('modele')
    if nom and nom ~= '' and IsModelInCdimage(GetHashKey(nom)) then
        return nom
    end
    nom = Config.Modeles[GetRandomIntInRange(1, #Config.Modeles + 1)]
    SetResourceKvp('modele', nom)
    return nom
end

local function versPointSpawn(point, avecModele)
    local c = point.coords
    return {
        x = c.x,
        y = c.y,
        z = c.z,
        heading = c.w,
        model = avecModele and modeleDuJoueur() or nil,
    }
end

local function appliquerReglagesPed()
    if Config.PvP then
        SetCanAttackFriendly(PlayerPedId(), true, false)
        NetworkSetFriendlyFireOption(true)
    end
end

-- Déclenché par spawnmanager à chaque apparition
AddEventHandler('playerSpawned', function()
    aDejaSpawn = true
    apparitionEnCours = false
    LocalPlayer.state:set('cocoSpawn', true, false)
    appliquerReglagesPed()
end)

---------------------------------------------------------------------------
-- Première apparition
---------------------------------------------------------------------------

local function premiereApparition()
    if aDejaSpawn or apparitionEnCours then
        return
    end
    apparitionEnCours = true

    local point = Config.PointsDeDepart[GetRandomIntInRange(1, #Config.PointsDeDepart + 1)]
    exports.spawnmanager:spawnPlayer(versPointSpawn(point, true), function()
        apparitionEnCours = false
        TriggerEvent('chat:addMessage', {
            color = { 255, 120, 190 },
            multiline = true,
            args = {
                Config.NomServeur,
                ('Bienvenue %s ! Tape /aide pour voir les commandes.'):format(GetPlayerName(PlayerId())),
            },
        })
        notifier(('Bienvenue sur ~p~%s~s~ !'):format(Config.NomServeur))
    end)
end

-- spawnmanager appelle cette fonction quand il faut faire apparaître le joueur.
-- Après la première fois, c'est la boucle « mort » plus bas qui s'en occupe.
exports.spawnmanager:setAutoSpawnCallback(function()
    if not aDejaSpawn then
        premiereApparition()
    end
end)

AddEventHandler('onClientGameTypeStart', function()
    exports.spawnmanager:setAutoSpawn(true)
    exports.spawnmanager:forceRespawn()
end)

-- Filet de sécurité : si l'événement de départ a été manqué, on fait quand même apparaître le joueur
CreateThread(function()
    while not NetworkIsPlayerActive(PlayerId()) do
        Wait(500)
    end
    Wait(10000)
    if not aDejaSpawn and not apparitionEnCours then
        premiereApparition()
    end
end)

---------------------------------------------------------------------------
-- Mort : compte à rebours puis hôpital le plus proche
---------------------------------------------------------------------------

local function hopitalLePlusProche(position)
    local choix, meilleureDistance = Config.Hopitaux[1], math.huge
    for _, hopital in ipairs(Config.Hopitaux) do
        local distance = #(position - hopital.coords.xyz)
        if distance < meilleureDistance then
            choix, meilleureDistance = hopital, distance
        end
    end
    return choix
end

local function gererMort()
    apparitionEnCours = true
    local positionMort = GetEntityCoords(PlayerPedId())
    local fin = GetGameTimer() + Config.DelaiReapparition * 1000
    AnimpostfxPlay('DeathFailOut', 0, true)

    while GetGameTimer() < fin do
        if not IsEntityDead(PlayerPedId()) then
            -- Quelqu'un nous a réanimé (un admin par exemple) : pas besoin d'hôpital
            AnimpostfxStop('DeathFailOut')
            apparitionEnCours = false
            return
        end
        local restant = math.ceil((fin - GetGameTimer()) / 1000)
        dessinerTexte('~r~Tu es mort', 0.5, 0.38, 0.9, { centre = true })
        dessinerTexte(('Réapparition à l\'hôpital dans %d s'):format(restant), 0.5, 0.46, 0.5, { centre = true })
        Wait(0)
    end

    AnimpostfxStop('DeathFailOut')
    local hopital = hopitalLePlusProche(positionMort)
    exports.spawnmanager:spawnPlayer(versPointSpawn(hopital, false), function()
        apparitionEnCours = false
        notifier(('Tu as été soigné à l\'hôpital de ~g~%s~s~.'):format(hopital.nom))
    end)
end

CreateThread(function()
    while true do
        Wait(250)
        if aDejaSpawn and not apparitionEnCours and IsEntityDead(PlayerPedId()) then
            gererMort()
        end
    end
end)

---------------------------------------------------------------------------
-- Police du jeu désactivée
---------------------------------------------------------------------------

if not Config.PoliceIA then
    CreateThread(function()
        SetMaxWantedLevel(0)
        for service = 1, 15 do
            EnableDispatchService(service, false)
        end
        while true do
            if GetPlayerWantedLevel(PlayerId()) > 0 then
                ClearPlayerWantedLevel(PlayerId())
            end
            Wait(1000)
        end
    end)
end

---------------------------------------------------------------------------
-- HUD en haut à droite
---------------------------------------------------------------------------

if Config.HUD then
    CreateThread(function()
        local ligne1, ligne2 = '', ''
        local prochaineMiseAJour = 0
        while true do
            local maintenant = GetGameTimer()
            if maintenant >= prochaineMiseAJour then
                prochaineMiseAJour = maintenant + 500
                ligne1 = ('%s  |  ID %d  |  %d/%d joueurs'):format(
                    Config.NomServeur,
                    GetPlayerServerId(PlayerId()),
                    GlobalState.nbJoueurs or 1,
                    GlobalState.maxJoueurs or 48
                )
                local pos = GetEntityCoords(PlayerPedId())
                local rue = GetStreetNameFromHashKey((GetStreetNameAtCoord(pos.x, pos.y, pos.z)))
                local zone = GetLabelText(GetNameOfZone(pos.x, pos.y, pos.z))
                ligne2 = (rue ~= '' and (rue .. '  |  ') or '') .. zone
            end
            if not IsPauseMenuActive() and not IsHudHidden() then
                dessinerTexte(ligne1, 0.995, 0.005, 0.33, { droite = true })
                dessinerTexte(ligne2, 0.995, 0.03, 0.30, { droite = true })
            end
            Wait(0)
        end
    end)
end

---------------------------------------------------------------------------
-- /skin [modèle]
---------------------------------------------------------------------------

RegisterCommand('skin', function(_, args)
    if not Config.SkinLibre then
        notifier('~r~/skin est désactivé sur ce serveur.')
        return
    end
    local nom = args[1] and args[1]:lower()
    if not nom then
        notifier('Utilise /skin [modèle], par ex. /skin a_f_y_hipster_01')
        return
    end
    if IsEntityDead(PlayerPedId()) then
        notifier('~r~Impossible pendant que tu es mort.')
        return
    end
    local hash = GetHashKey(nom)
    if not IsModelInCdimage(hash) or not IsModelAPed(hash) then
        notifier(('~r~Personnage inconnu :~s~ %s'):format(nom))
        return
    end

    RequestModel(hash)
    local limite = GetGameTimer() + 5000
    while not HasModelLoaded(hash) do
        if GetGameTimer() > limite then
            notifier('~r~Le personnage met trop de temps à charger, réessaie.')
            return
        end
        Wait(0)
    end

    SetPlayerModel(PlayerId(), hash)
    SetModelAsNoLongerNeeded(hash)
    SetPedDefaultComponentVariation(PlayerPedId())
    SetResourceKvp('modele', nom)
    appliquerReglagesPed()
    notifier('Nouveau personnage enregistré !')
end, false)

CreateThread(function()
    Wait(1000)
    TriggerEvent('chat:addSuggestion', '/skin', 'Changer de personnage (ton choix est retenu)', {
        { name = 'modele', help = 'ex : a_m_y_hipster_01, a_f_y_beach_01, s_m_y_cop_01' },
    })
    TriggerEvent('chat:addSuggestion', '/monid', 'Voir ton ID et ta licence (utile pour devenir admin)')
end)

AddEventHandler('onResourceStop', function(ressource)
    if ressource == GetCurrentResourceName() then
        TriggerEvent('chat:removeSuggestion', '/skin')
        TriggerEvent('chat:removeSuggestion', '/monid')
        AnimpostfxStop('DeathFailOut')
    end
end)

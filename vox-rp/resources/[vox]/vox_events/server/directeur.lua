-- Le directeur d'événements : il tire au sort un événement de temps en temps,
-- en tenant compte des joueurs connectés et des métiers en service.

Directeur = {
    evenements = {},     -- [nom] = définition (voir server/evenements/*.lua)
    enCours = {},        -- [nom] = instance de l'événement en cours
    rechargeJusqua = {}, -- [nom] = heure (os.time) avant laquelle il ne peut pas revenir
    auto = Config.Directeur.Auto,
}

local compteur = 0

-- def = { nom, titre, config, demarrer(instance, ctx), arreter(instance, raison),
--         peutDemarrer(ctx) -> ok, raison, poids(ctx) -> nombre, envoyerEtat(instance, src) }
function Directeur.enregistrer(def)
    if def.config and def.config.Actif == false then
        return
    end
    Directeur.evenements[def.nom] = def
end

function Directeur.actif(nom)
    return Directeur.enCours[nom]
end

function Directeur.contexte()
    return {
        joueurs = #GetPlayers(),
        police = Bridge.compterEnService('police'),
        ems = Bridge.compterEnService('ems'),
        mecanos = Bridge.compterEnService('mecano'),
    }
end

-- Les conditions pour qu'un événement puisse se lancer tout seul
function Directeur.conditions(def, ctx)
    local cfg = def.config or {}
    if ctx.joueurs < (cfg.MinJoueurs or 1) then
        return false, ('il faut %d joueurs connectés'):format(cfg.MinJoueurs or 1)
    end
    if def.peutDemarrer then
        return def.peutDemarrer(ctx)
    end
    return true
end

local function poidsDe(def, ctx)
    if def.poids then
        return def.poids(ctx)
    end
    return (def.config and def.config.Poids) or 1
end

-- options : { force = true pour ignorer les conditions, cause = texte libre }
function Directeur.lancer(nom, options)
    options = options or {}
    local def = Directeur.evenements[nom]
    if not def then
        return false, 'événement inconnu'
    end
    if Directeur.enCours[nom] then
        return false, 'déjà en cours'
    end
    local ctx = Directeur.contexte()
    if not options.force then
        local ok, raison = Directeur.conditions(def, ctx)
        if not ok then
            return false, raison
        end
    end

    compteur = compteur + 1
    local instance = { id = compteur, nom = nom, debut = os.time(), cause = options.cause }
    Directeur.enCours[nom] = instance

    local ok, erreur = pcall(def.demarrer, instance, ctx)
    if not ok or instance.annule then
        Directeur.enCours[nom] = nil
        if not ok then
            print(('[vox_events] Erreur au lancement de « %s » : %s'):format(nom, erreur))
        end
        return false, ok and (instance.annule or 'annulé') or 'erreur, voir la console du serveur'
    end

    print(('[vox_events] Événement lancé : %s (n°%d)'):format(def.titre or nom, instance.id))
    local duree = ((def.config and def.config.Duree) or 15) * 60000
    SetTimeout(duree, function()
        if Directeur.enCours[nom] == instance then
            Directeur.terminer(nom, 'temps')
        end
    end)
    return true
end

-- raison : 'temps', 'staff', 'fini'... (passée à def.arreter)
function Directeur.terminer(nom, raison)
    local instance = Directeur.enCours[nom]
    if not instance then
        return false
    end
    Directeur.enCours[nom] = nil
    local def = Directeur.evenements[nom]
    Directeur.rechargeJusqua[nom] = os.time() + ((def.config and def.config.Recharge) or 30) * 60
    if def.arreter then
        local ok, erreur = pcall(def.arreter, instance, raison)
        if not ok then
            print(('[vox_events] Erreur à la fin de « %s » : %s'):format(nom, erreur))
        end
    end
    print(('[vox_events] Événement terminé : %s (%s)'):format(def.titre or nom, raison or '?'))
    return true
end

function Directeur.nombreEnCours()
    local nombre = 0
    for _ in pairs(Directeur.enCours) do
        nombre = nombre + 1
    end
    return nombre
end

-- Tirage au sort pondéré parmi les événements qui peuvent se lancer
function Directeur.tirerAuSort(ctx)
    local candidats, total = {}, 0
    for nom, def in pairs(Directeur.evenements) do
        local libre = not Directeur.enCours[nom] and os.time() >= (Directeur.rechargeJusqua[nom] or 0)
        if libre and Directeur.conditions(def, ctx) then
            local poids = poidsDe(def, ctx)
            if poids > 0 then
                candidats[#candidats + 1] = { nom = nom, poids = poids }
                total = total + poids
            end
        end
    end
    if total == 0 then
        return nil
    end
    table.sort(candidats, function(a, b) return a.nom < b.nom end)
    local tirage = math.random() * total
    for _, candidat in ipairs(candidats) do
        tirage = tirage - candidat.poids
        if tirage <= 0 then
            return candidat.nom
        end
    end
    return candidats[#candidats].nom
end

CreateThread(function()
    Wait(Config.Directeur.PremierDelai * 60000)
    while true do
        if Directeur.auto and Directeur.nombreEnCours() < Config.Directeur.MaxEnMemeTemps then
            local nom = Directeur.tirerAuSort(Directeur.contexte())
            if nom then
                Directeur.lancer(nom)
            end
        end
        Wait(math.random(Config.Directeur.IntervalleMin, Config.Directeur.IntervalleMax) * 60000)
    end
end)

-- Un joueur qui arrive en cours de route reçoit les événements déjà lancés
RegisterNetEvent('vox_events:demanderEtat', function()
    local src = source
    for nom, instance in pairs(Directeur.enCours) do
        local def = Directeur.evenements[nom]
        if def.envoyerEtat then
            def.envoyerEtat(instance, src)
        end
    end
end)

-- Pour les autres ressources : exports.vox_events:EvenementEnCours('panne')
exports('EvenementEnCours', function(nom)
    return Directeur.enCours[nom] ~= nil
end)

---------------------------------------------------------------------------
-- Petits outils partagés par les événements
---------------------------------------------------------------------------

Outils = {}

function Outils.distance2D(a, b)
    local dx, dy = a.x - b.x, a.y - b.y
    return math.sqrt(dx * dx + dy * dy)
end

-- Vérifie une position envoyée par un joueur (on ne fait pas confiance au client)
function Outils.vecteur(v)
    if type(v) ~= 'vector3' and type(v) ~= 'table' then
        return nil
    end
    local x, y, z = v.x, v.y, v.z
    if type(x) ~= 'number' or type(y) ~= 'number' or type(z) ~= 'number' then
        return nil
    end
    if x ~= x or y ~= y or z ~= z then -- NaN
        return nil
    end
    return vector3(x, y, z)
end

function Outils.positionJoueur(src)
    local ped = GetPlayerPed(src)
    if not ped or ped == 0 then
        return nil
    end
    return GetEntityCoords(ped)
end

-- Vecteurs « devant » et « à droite » pour un cap GTA (0 = nord, sens inverse des aiguilles)
function Outils.directions(cap)
    local r = math.rad(cap)
    return vector3(-math.sin(r), math.cos(r), 0.0), vector3(math.cos(r), math.sin(r), 0.0)
end

function Outils.argent(montant)
    local texte = tostring(math.floor(montant))
    local avecEspaces = texte:reverse():gsub('(%d%d%d)', '%1 '):reverse():gsub('^ ', '')
    return avecEspaces .. ' $'
end

-- Tirage pondéré dans une liste d'éléments qui ont un champ « poids »
function Outils.tirage(liste)
    local total = 0
    for _, element in ipairs(liste) do
        total = total + (element.poids or 1)
    end
    local tirage = math.random() * total
    for _, element in ipairs(liste) do
        tirage = tirage - (element.poids or 1)
        if tirage <= 0 then
            return element
        end
    end
    return liste[#liste]
end

-- Arrêt de la ressource : on termine proprement (rallumer la ville, rendre la météo...)
AddEventHandler('onResourceStop', function(ressource)
    if ressource ~= GetCurrentResourceName() then
        return
    end
    for nom in pairs(Directeur.enCours) do
        Directeur.terminer(nom, 'arret')
    end
end)

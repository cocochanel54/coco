-- Plantations (côté serveur). L'état de chaque plant est calculé à partir de la dernière
-- sauvegarde et de l'heure : pas besoin de tout réécrire en base de données chaque minute.

local plantes = {} -- [id] = plant (ligne de la base de données)
local vus = {}     -- [joueur] = { [id] = stade déjà envoyé }

---------------------------------------------------------------------------
-- Outils
---------------------------------------------------------------------------

local function joueurQbox(src)
    return exports.qbx_core:GetPlayer(src)
end

local function citizenid(src)
    local joueur = joueurQbox(src)
    return joueur and joueur.PlayerData and joueur.PlayerData.citizenid
end

local function estPolice(src)
    local joueur = joueurQbox(src)
    local job = joueur and joueur.PlayerData and joueur.PlayerData.job
    if not job or not job.onduty then
        return false
    end
    for _, t in ipairs(Config.Police.types) do
        if job.type == t then
            return true
        end
    end
    for _, n in ipairs(Config.Police.noms) do
        if job.name == n then
            return true
        end
    end
    return false
end

local function notifier(src, texte, genre)
    TriggerClientEvent('ox_lib:notify', src, {
        title = 'Plantation',
        description = texte,
        type = genre or 'inform',
        icon = 'seedling',
        position = 'top',
    })
end

local function distance2D(ax, ay, bx, by)
    local dx, dy = ax - bx, ay - by
    return math.sqrt(dx * dx + dy * dy)
end

local function positionJoueur(src)
    local ped = GetPlayerPed(src)
    if not ped or ped == 0 then
        return nil
    end
    return GetEntityCoords(ped)
end

local function vecteur(v)
    if type(v) ~= 'vector3' and type(v) ~= 'table' then
        return nil
    end
    local x, y, z = v.x, v.y, v.z
    if type(x) ~= 'number' or type(y) ~= 'number' or type(z) ~= 'number' or x ~= x or y ~= y or z ~= z then
        return nil
    end
    return vector3(x, y, z)
end

local function retirerObjet(src, nom)
    local ok, retire = pcall(function()
        return exports.ox_inventory:RemoveItem(src, nom, 1)
    end)
    return ok and retire and true or false
end

local function donnerObjet(src, nom, quantite)
    local ok, peut = pcall(function()
        return exports.ox_inventory:CanCarryItem(src, nom, quantite)
    end)
    if not ok or not peut then
        return false
    end
    local ok2, reussi = pcall(function()
        return exports.ox_inventory:AddItem(src, nom, quantite)
    end)
    return ok2 and reussi and true or false
end

---------------------------------------------------------------------------
-- Croissance
---------------------------------------------------------------------------

-- État d'un plant à l'heure t (en secondes), à partir de sa dernière sauvegarde
local function etatActuel(p, t)
    local reglage = Config.Plantes[p.type]
    local minutes = math.max(0, (t - p.maj) / 60)
    local perteParMinute = Config.Eau.PerteParHeure / 60
    local eau, croissance, sansEau, mort = p.eau, p.croissance, p.sans_eau, p.mort

    if not mort then
        -- Nombre de minutes pendant lesquelles il restait de l'eau
        local avecEau = minutes
        if perteParMinute > 0 then
            avecEau = math.min(minutes, eau / perteParMinute)
        end
        croissance = math.min(1.0, croissance + avecEau / reglage.Duree)
        eau = math.max(0, eau - perteParMinute * minutes)
        if avecEau > 0 then
            sansEau = minutes - avecEau
        else
            sansEau = sansEau + minutes
        end
        mort = croissance < 1.0 and sansEau >= Config.Eau.DelaiMort
    end

    local stade = 1
    if croissance >= 1.0 then
        stade = 3
    elseif croissance >= 0.4 then
        stade = 2
    end
    return { eau = eau, croissance = croissance, sansEau = sansEau, mort = mort, stade = stade, qualite = p.qualite }
end

local function rebaser(p, t)
    local e = etatActuel(p, t)
    p.eau, p.croissance, p.sans_eau, p.mort, p.maj = e.eau, e.croissance, e.sansEau, e.mort, t
    return e
end

local function sauvegarder(p)
    MySQL.update.await(
        'UPDATE vox_plantations SET eau = ?, croissance = ?, sans_eau = ?, qualite = ?, mort = ?, maj = ? WHERE id = ?',
        { p.eau, p.croissance, p.sans_eau, p.qualite, p.mort and 1 or 0, p.maj, p.id }
    )
end

local function supprimer(p)
    plantes[p.id] = nil
    MySQL.query.await('DELETE FROM vox_plantations WHERE id = ?', { p.id })
end

---------------------------------------------------------------------------
-- Affichage : chaque joueur ne reçoit que les plants autour de lui
---------------------------------------------------------------------------

local function majVisibilite()
    local t = os.time()
    local etats = {}
    local rayon2 = Config.DistanceAffichage * Config.DistanceAffichage
    for _, idTexte in ipairs(GetPlayers()) do
        local src = tonumber(idTexte)
        local position = positionJoueur(src)
        if position then
            local dejaVu = vus[src] or {}
            local visibles = {}
            for id, p in pairs(plantes) do
                local dx, dy = p.x - position.x, p.y - position.y
                if dx * dx + dy * dy < rayon2 then
                    local e = etats[id] or etatActuel(p, t)
                    etats[id] = e
                    local cle = e.stade .. (e.mort and 'm' or '')
                    visibles[id] = cle
                    if dejaVu[id] ~= cle then
                        TriggerClientEvent('vox_plantations:afficher', src, {
                            id = id, type = p.type, x = p.x, y = p.y, z = p.z, stade = e.stade, mort = e.mort,
                        })
                    end
                end
            end
            for id in pairs(dejaVu) do
                if not visibles[id] then
                    TriggerClientEvent('vox_plantations:retirer', src, id)
                end
            end
            vus[src] = visibles
        end
    end
end

AddEventHandler('playerDropped', function()
    vus[source] = nil
end)

---------------------------------------------------------------------------
-- Démarrage : table de la base de données et chargement des plants
---------------------------------------------------------------------------

CreateThread(function()
    MySQL.query.await([[
        CREATE TABLE IF NOT EXISTS vox_plantations (
            id INT AUTO_INCREMENT PRIMARY KEY,
            type VARCHAR(32) NOT NULL,
            proprietaire VARCHAR(64) NOT NULL,
            x FLOAT NOT NULL,
            y FLOAT NOT NULL,
            z FLOAT NOT NULL,
            eau FLOAT NOT NULL DEFAULT 0,
            croissance FLOAT NOT NULL DEFAULT 0,
            sans_eau FLOAT NOT NULL DEFAULT 0,
            qualite FLOAT NOT NULL DEFAULT 40,
            mort TINYINT(1) NOT NULL DEFAULT 0,
            maj INT NOT NULL,
            plantee INT NOT NULL
        )
    ]])
    local lignes = MySQL.query.await('SELECT * FROM vox_plantations') or {}
    for _, ligne in ipairs(lignes) do
        if Config.Plantes[ligne.type] then
            ligne.mort = ligne.mort == 1 or ligne.mort == true
            plantes[ligne.id] = ligne
        end
    end
    print(('[vox_plantations] %d plant(s) chargé(s).'):format(#lignes))

    while true do
        majVisibilite()
        Wait(3000)
    end
end)

---------------------------------------------------------------------------
-- Planter
---------------------------------------------------------------------------

for nomType, reglage in pairs(Config.Plantes) do
    exports.qbx_core:CreateUseableItem(reglage.Graine, function(source)
        TriggerClientEvent('vox_plantations:planter', source, nomType)
    end)
end

local function zoneInterdite(position)
    for _, zone in ipairs(Config.ZonesInterdites) do
        if distance2D(position.x, position.y, zone.centre.x, zone.centre.y) < zone.rayon then
            return zone
        end
    end
    return nil
end

RegisterNetEvent('vox_plantations:planter', function(nomType, position)
    local src = source
    local reglage = Config.Plantes[nomType]
    position = vecteur(position)
    local proprietaire = citizenid(src)
    local joueur = positionJoueur(src)
    if not reglage or not position or not proprietaire or not joueur then
        return
    end
    if #(joueur - position) > 3.0 then
        return
    end
    local zone = zoneInterdite(position)
    if zone then
        notifier(src, ('Impossible de planter près de %s.'):format(zone.nom), 'error')
        return
    end
    local mesPlants = 0
    for _, p in pairs(plantes) do
        if p.proprietaire == proprietaire then
            mesPlants = mesPlants + 1
        end
        if distance2D(p.x, p.y, position.x, position.y) < Config.DistanceEntrePlantes and math.abs(p.z - position.z) < 2.0 then
            notifier(src, 'Il y a déjà un plant juste ici.', 'error')
            return
        end
    end
    if mesPlants >= Config.MaxParJoueur then
        notifier(src, ('Tu as déjà %d plants, c\'est le maximum.'):format(Config.MaxParJoueur), 'error')
        return
    end
    if not retirerObjet(src, reglage.Graine) then
        notifier(src, 'Il te faut une graine.', 'error')
        return
    end

    local t = os.time()
    local p = {
        type = nomType, proprietaire = proprietaire, x = position.x, y = position.y, z = position.z,
        eau = 0, croissance = 0, sans_eau = 0, qualite = Config.QualiteDepart, mort = false, maj = t, plantee = t,
    }
    p.id = MySQL.insert.await(
        'INSERT INTO vox_plantations (type, proprietaire, x, y, z, eau, croissance, sans_eau, qualite, mort, maj, plantee) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
        { p.type, p.proprietaire, p.x, p.y, p.z, p.eau, p.croissance, p.sans_eau, p.qualite, 0, p.maj, p.plantee }
    )
    if not p.id then
        notifier(src, 'Erreur de sauvegarde, réessaie.', 'error')
        return
    end
    plantes[p.id] = p
    notifier(src, 'Graine plantée. Pense à l\'arroser, sinon elle ne poussera pas !', 'success')
    majVisibilite()
end)

---------------------------------------------------------------------------
-- Infos et actions sur un plant
---------------------------------------------------------------------------

local function plantProche(src, id)
    local p = plantes[id]
    local joueur = positionJoueur(src)
    if not p or not joueur or #(joueur - vector3(p.x, p.y, p.z)) > 3.0 then
        return nil
    end
    return p
end

lib.callback.register('vox_plantations:infos', function(source, id)
    local p = plantProche(source, id)
    if not p then
        return nil
    end
    local e = etatActuel(p, os.time())
    local reglage = Config.Plantes[p.type]
    local reste = math.ceil((1.0 - e.croissance) * reglage.Duree)
    return {
        label = reglage.Label,
        eau = math.floor(e.eau),
        qualite = math.floor(e.qualite),
        croissance = math.floor(e.croissance * 100),
        reste = reste,
        mort = e.mort,
        prete = e.stade == 3 and not e.mort,
        police = estPolice(source),
        proprietaire = p.proprietaire == citizenid(source),
    }
end)

local function prevenirProprietaire(p, texte)
    for _, idTexte in ipairs(GetPlayers()) do
        local id = tonumber(idTexte)
        if citizenid(id) == p.proprietaire then
            notifier(id, texte, 'warning')
        end
    end
end

local actions = {}

function actions.arroser(src, p, e)
    if e.mort then
        return notifier(src, 'Ce plant est mort, il n\'y a plus rien à faire.', 'error')
    end
    if not retirerObjet(src, Config.Eau.Objet) then
        return notifier(src, 'Il te faut une bouteille d\'eau.', 'error')
    end
    p.eau = math.min(100, p.eau + Config.Eau.Gain)
    p.sans_eau = 0
    sauvegarder(p)
    notifier(src, ('Plant arrosé (eau : %d %%).'):format(math.floor(p.eau)), 'success')
end

function actions.engrais(src, p, e)
    if e.mort then
        return notifier(src, 'Ce plant est mort, il n\'y a plus rien à faire.', 'error')
    end
    if not retirerObjet(src, Config.Engrais.Objet) then
        return notifier(src, 'Il te faut de l\'engrais.', 'error')
    end
    p.qualite = math.min(100, p.qualite + Config.Engrais.Gain)
    sauvegarder(p)
    notifier(src, ('Engrais ajouté (qualité : %d %%).'):format(math.floor(p.qualite)), 'success')
end

function actions.recolter(src, p, e)
    if e.mort or e.stade ~= 3 then
        return notifier(src, 'Ce plant n\'est pas encore prêt.', 'error')
    end
    local reglage = Config.Plantes[p.type]
    local quantite = math.floor(reglage.Quantite[1] + (reglage.Quantite[2] - reglage.Quantite[1]) * p.qualite / 100 + 0.5)
    if not donnerObjet(src, reglage.Recolte, quantite) then
        return notifier(src, 'Ton inventaire est plein.', 'error')
    end
    supprimer(p)
    notifier(src, ('Récolte : %d x %s.'):format(quantite, reglage.Recolte), 'success')
    if p.proprietaire ~= citizenid(src) then
        prevenirProprietaire(p, 'Quelqu\'un a récolté un de tes plants à ta place…')
    end
end

function actions.detruire(src, p)
    local police = estPolice(src)
    local proprietaire = p.proprietaire == citizenid(src)
    if not police and not proprietaire and not Config.DestructionLibre then
        return notifier(src, 'Ce n\'est pas ton plant.', 'error')
    end
    supprimer(p)
    if police and not proprietaire then
        exports.qbx_core:GetPlayer(src).Functions.AddMoney('bank', Config.Police.Prime, 'vox_plantations:saisie')
        notifier(src, ('Plant saisi et détruit. Prime : %d $'):format(Config.Police.Prime), 'success')
    else
        notifier(src, 'Plant arraché.', 'success')
    end
    if not proprietaire then
        prevenirProprietaire(p, 'Un de tes plants a été arraché…')
    end
end

RegisterNetEvent('vox_plantations:action', function(id, nomAction)
    local src = source
    local action = actions[nomAction]
    local p = type(id) == 'number' and plantProche(src, id)
    if not action or not p then
        return
    end
    local e = rebaser(p, os.time())
    action(src, p, e)
    majVisibilite()
end)

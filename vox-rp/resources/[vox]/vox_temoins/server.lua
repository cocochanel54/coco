-- Témoins (côté serveur) : tirage au sort, délai d'appel, regroupement et envoi à la police

local appels = {}      -- [id] = appel en cours (le témoin est au téléphone)
local compteur = 0
local recentes = {}    -- alertes récentes, pour ne pas spammer la police
local limites = {}     -- [joueur .. acte] = heure avant laquelle on ignore ses signalements

local function joueurQbox(src)
    return exports.qbx_core:GetPlayer(src)
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

local function policiers()
    local liste = {}
    for _, idTexte in ipairs(GetPlayers()) do
        local id = tonumber(idTexte)
        if estPolice(id) then
            liste[#liste + 1] = id
        end
    end
    return liste
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

local function texteSur(valeur, longueur, defaut)
    if type(valeur) ~= 'string' or valeur == '' then
        return defaut
    end
    return (valeur:gsub('[~<>]', ''):sub(1, longueur))
end

-- Envoie l'alerte aux policiers en service. Renvoie false si une alerte proche vient d'être envoyée.
-- Utilisable par les autres ressources serveur : exports.vox_temoins:Alerter('braquage', { position, zone, description })
local function alerter(acte, donnees)
    local regle = Config.Actes[acte]
    local position = vecteur(donnees.position)
    if not regle or not position then
        return false
    end
    local maintenant = os.time()
    for i = #recentes, 1, -1 do
        local r = recentes[i]
        if maintenant - r.heure > Config.Regroupement.Secondes then
            table.remove(recentes, i)
        elseif r.acte == acte and #(r.position - position) < Config.Regroupement.Distance then
            return false
        end
    end
    recentes[#recentes + 1] = { acte = acte, position = position, heure = maintenant }

    -- Le témoin donne un endroit approximatif
    local angle = math.random() * 2 * math.pi
    local ecart = math.random() * regle.Imprecision
    compteur = compteur + 1
    local alerte = {
        id = 'temoin' .. compteur,
        titre = 'Appel d\'un témoin : ' .. regle.Titre:lower(),
        texte = ('%s près de %s. %s'):format(regle.Titre, texteSur(donnees.zone, 40, 'Los Santos'), texteSur(donnees.description, 120, '')),
        position = position + vector3(math.cos(angle) * ecart, math.sin(angle) * ecart, 0.0),
        rayon = regle.Imprecision + 20.0,
    }
    for _, id in ipairs(policiers()) do
        TriggerClientEvent('vox_temoins:alerte', id, alerte)
    end
    return true
end

exports('Alerter', alerter)

RegisterNetEvent('vox_temoins:signaler', function(acte, donnees)
    local src = source
    local regle = Config.Actes[acte]
    if not regle or type(donnees) ~= 'table' then
        return
    end
    local position = vecteur(donnees.position)
    local joueur = positionJoueur(src)
    if not position or not joueur or #(joueur - position) > 50.0 then
        return
    end
    if acte == 'tirs' and estPolice(src) then
        return
    end
    local cle = src .. ':' .. acte
    if (limites[cle] or 0) > os.time() then
        return
    end
    limites[cle] = os.time() + math.floor(regle.Recharge / 2)
    if #policiers() == 0 then
        return -- personne pour répondre à l'appel
    end
    if not donnees.certain and math.random(100) > regle.Chance then
        return -- le témoin a préféré ne pas s'en mêler
    end

    compteur = compteur + 1
    local idAppel = compteur
    local appel = {
        acte = acte,
        donnees = { position = position, zone = donnees.zone, description = donnees.description },
        annule = false,
    }
    appels[idAppel] = appel
    local delai = math.random(Config.DelaiAppel[1], Config.DelaiAppel[2])

    -- Le joueur qui « possède » le PNJ le fait téléphoner, et prévient si le témoin est neutralisé
    if type(donnees.temoin) == 'number' then
        local entite = NetworkGetEntityFromNetworkId(donnees.temoin)
        if entite ~= 0 and DoesEntityExist(entite) then
            local proprietaire = NetworkGetEntityOwner(entite)
            if proprietaire and proprietaire > 0 then
                appel.proprietaire = proprietaire
                TriggerClientEvent('vox_temoins:telephoner', proprietaire, idAppel, donnees.temoin, delai)
            end
        end
    end

    SetTimeout(delai * 1000, function()
        appels[idAppel] = nil
        if not appel.annule then
            alerter(acte, appel.donnees)
        end
    end)
end)

RegisterNetEvent('vox_temoins:annuler', function(idAppel)
    local appel = appels[idAppel]
    if appel and appel.proprietaire == source then
        appel.annule = true
    end
end)

AddEventHandler('playerDropped', function()
    local prefixe = source .. ':'
    for cle in pairs(limites) do
        if cle:sub(1, #prefixe) == prefixe then
            limites[cle] = nil
        end
    end
end)

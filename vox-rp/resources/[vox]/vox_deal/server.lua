-- Vente de rue (côté serveur) : c'est ici qu'on décide si le PNJ achète, refuse ou appelle la police

local recharges = {}    -- [joueur] = heure avant la prochaine proposition
local dejaServis = {}   -- [ID réseau du PNJ] = heure avant laquelle il n'achète plus
local ventesParZone = {} -- [zone du jeu] = quantité vendue (pour les futures guerres de territoires)

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

local function nombrePoliciers()
    local nombre = 0
    for _, idTexte in ipairs(GetPlayers()) do
        if estPolice(tonumber(idTexte)) then
            nombre = nombre + 1
        end
    end
    return nombre
end

local function compter(src, objet)
    local ok, nombre = pcall(function()
        return exports.ox_inventory:Search(src, 'count', objet)
    end)
    return ok and type(nombre) == 'number' and nombre or 0
end

local function retirer(src, objet, quantite)
    local ok, retire = pcall(function()
        return exports.ox_inventory:RemoveItem(src, objet, quantite)
    end)
    return ok and retire and true or false
end

local function payer(src, montant)
    local ok, paye = pcall(function()
        if not exports.ox_inventory:CanCarryItem(src, Config.Paiement, montant) then
            return false
        end
        return exports.ox_inventory:AddItem(src, Config.Paiement, montant)
    end)
    if not (ok and paye) then
        joueurQbox(src).Functions.AddMoney('cash', montant, 'vox_deal')
    end
end

local function erreur(message)
    return { resultat = 'erreur', message = message }
end

lib.callback.register('vox_deal:proposer', function(src, netId, infos)
    local maintenant = os.time()
    if (recharges[src] or 0) > maintenant then
        return erreur('Doucement, tu vas te faire repérer…')
    end
    recharges[src] = maintenant + Config.Recharge

    local ped = GetPlayerPed(src)
    local joueur = ped ~= 0 and GetEntityCoords(ped)
    if not joueur then
        return nil
    end
    if nombrePoliciers() < Config.MinPolice then
        return erreur('C\'est trop calme ce soir, personne n\'achète.')
    end
    for _, zone in ipairs(Config.ZonesInterdites) do
        if #(joueur - zone.centre) < zone.rayon then
            return erreur(('T\'es fou ? Pas à côté %s !'):format(zone.nom))
        end
    end
    if type(netId) == 'number' then
        local client = NetworkGetEntityFromNetworkId(netId)
        if client == 0 or not DoesEntityExist(client) or #(GetEntityCoords(client) - joueur) > 5.0 then
            return erreur('Il est parti.')
        end
        if (dejaServis[netId] or 0) > maintenant then
            return { resultat = 'refuse', message = 'Il t\'a déjà acheté quelque chose, il n\'en veut plus.' }
        end
    end

    local produit, enStock
    for _, p in ipairs(Config.Produits) do
        local nombre = compter(src, p.Objet)
        if nombre > 0 then
            produit, enStock = p, nombre
            break
        end
    end
    if not produit then
        return erreur('Tu n\'as rien à vendre.')
    end

    local tirage = math.random(100)
    if tirage <= Config.Chances.Appelle then
        if netId then
            dejaServis[netId] = maintenant + Config.MemeClient
        end
        return { resultat = 'appelle' }
    elseif tirage <= Config.Chances.Appelle + Config.Chances.Refuse then
        return { resultat = 'refuse', message = Config.Refus[math.random(#Config.Refus)] }
    end

    local quantite = math.min(enStock, math.random(produit.Quantite[1], produit.Quantite[2]))
    if not retirer(src, produit.Objet, quantite) then
        return erreur('Tu n\'as rien à vendre.')
    end
    local montant = 0
    for _ = 1, quantite do
        montant = montant + math.random(produit.Prix[1], produit.Prix[2])
    end
    payer(src, montant)
    if netId then
        dejaServis[netId] = maintenant + Config.MemeClient
    end
    local zone = type(infos) == 'table' and type(infos.zone) == 'string' and infos.zone:sub(1, 16) or '?'
    ventesParZone[zone] = (ventesParZone[zone] or 0) + quantite
    return { resultat = 'accepte', quantite = quantite, montant = montant, label = produit.Label }
end)

-- Pour les futures guerres de territoires : exports.vox_deal:VentesParZone()
exports('VentesParZone', function()
    return ventesParZone
end)

-- Ménage régulier
CreateThread(function()
    while true do
        Wait(600000)
        local maintenant = os.time()
        for netId, fin in pairs(dejaServis) do
            if fin < maintenant then
                dejaServis[netId] = nil
            end
        end
    end
end)

AddEventHandler('playerDropped', function()
    recharges[source] = nil
end)

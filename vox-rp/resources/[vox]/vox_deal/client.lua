-- Vente de rue (côté joueur) : approche un passant à pied avec de la marchandise et appuie sur E

local TYPES_FORCES_ORDRE = { [6] = true, [27] = true, [29] = true }
local produitsEnPoche = {}
local occupe = false
local aideAffichee = false

local function notifier(texte, genre)
    lib.notify({ title = 'Deal', description = texte, type = genre or 'inform', icon = 'cannabis', position = 'top' })
end

local function cacherAide()
    if aideAffichee then
        aideAffichee = false
        lib.hideTextUI()
    end
end

local function compterProduits()
    local liste = {}
    for _, produit in ipairs(Config.Produits) do
        local ok, nombre = pcall(function()
            return exports.ox_inventory:Search('count', produit.Objet)
        end)
        if ok and type(nombre) == 'number' and nombre > 0 then
            liste[#liste + 1] = produit
        end
    end
    return liste
end

-- Le passant le plus proche qui peut être un client (pas un joueur, pas un flic, pas en train de fuir)
local function clientProche()
    local moi = PlayerPedId()
    local position = GetEntityCoords(moi)
    local choix, distanceMin = nil, 2.0
    for _, ped in ipairs(GetGamePool('CPed')) do
        if ped ~= moi and not IsPedAPlayer(ped) and IsPedHuman(ped) and not IsPedDeadOrDying(ped, true)
            and not IsPedInAnyVehicle(ped, false) and not TYPES_FORCES_ORDRE[GetPedType(ped)]
            and not IsPedFleeing(ped) and not IsPedInCombat(ped, 0) then
            local distance = #(position - GetEntityCoords(ped))
            if distance < distanceMin then
                choix, distanceMin = ped, distance
            end
        end
    end
    return choix
end

local function prendreControle(ped)
    if not NetworkGetEntityIsNetworked(ped) then
        return true
    end
    local limite = GetGameTimer() + 500
    while not NetworkHasControlOfEntity(ped) and GetGameTimer() < limite do
        NetworkRequestControlOfEntity(ped)
        Wait(0)
    end
    return NetworkHasControlOfEntity(ped)
end

local function chargerAnim(dictionnaire)
    RequestAnimDict(dictionnaire)
    local limite = GetGameTimer() + 2000
    while not HasAnimDictLoaded(dictionnaire) and GetGameTimer() < limite do
        Wait(0)
    end
    return HasAnimDictLoaded(dictionnaire)
end

local function signaler(options)
    pcall(function()
        exports.vox_temoins:ActeSuspect('vente_drogue', options)
    end)
end

local function proposer(ped)
    occupe = true
    cacherAide()
    local moi = PlayerPedId()
    local controle = prendreControle(ped)
    if controle then
        ClearPedTasks(ped)
        TaskTurnPedToFaceEntity(ped, moi, 3000)
    end
    local netId = NetworkGetEntityIsNetworked(ped) and NetworkGetNetworkIdFromEntity(ped) or nil
    local position = GetEntityCoords(moi)
    local reponse = lib.callback.await('vox_deal:proposer', false, netId, {
        zone = GetNameOfZone(position.x, position.y, position.z),
    })

    if not reponse then
        occupe = false
        return
    end
    if reponse.resultat == 'accepte' then
        if chargerAnim('mp_common') then
            TaskPlayAnim(moi, 'mp_common', 'givetake1_a', 8.0, -8.0, 2000, 0, 0, false, false, false)
            if controle then
                TaskPlayAnim(ped, 'mp_common', 'givetake1_b', 8.0, -8.0, 2000, 0, 0, false, false, false)
            end
        end
        Wait(2000)
        notifier(('Vendu : %d x %s pour %d $ d\'argent sale.'):format(reponse.quantite, reponse.label, reponse.montant), 'success')
        -- Un autre passant a peut-être vu l'échange
        signaler({ exclure = { ped } })
    elseif reponse.resultat == 'appelle' then
        notifier('Il sort son téléphone… il va appeler les flics !', 'error')
        signaler({ temoin = ped, certain = true })
    elseif reponse.resultat == 'refuse' then
        notifier(reponse.message or 'Pas intéressé.')
    else
        notifier(reponse.message or 'Impossible.', 'error')
    end
    if controle and DoesEntityExist(ped) and reponse.resultat ~= 'appelle' then
        TaskWanderStandard(ped, 10.0, 10)
    end
    occupe = false
end

CreateThread(function()
    local prochaineVerification = 0
    while true do
        local attente = 1000
        local maintenant = GetGameTimer()
        if maintenant > prochaineVerification then
            produitsEnPoche = compterProduits()
            prochaineVerification = maintenant + 3000
        end
        local moi = PlayerPedId()
        local cible = nil
        if #produitsEnPoche > 0 and not occupe and not IsPedInAnyVehicle(moi, false) and not IsEntityDead(moi) then
            attente = 250
            cible = clientProche()
        end
        if cible then
            if not aideAffichee then
                aideAffichee = true
                lib.showTextUI(('[E] Proposer du %s'):format(produitsEnPoche[1].Label), { position = 'right-center' })
            end
            -- On guette la touche pendant un court moment, sans rechercher les passants à chaque image
            local fin = GetGameTimer() + 250
            while GetGameTimer() < fin do
                if IsControlJustReleased(0, 38) then
                    proposer(cible)
                    break
                end
                Wait(0)
            end
            attente = 0
        else
            cacherAide()
        end
        Wait(attente)
    end
end)

AddEventHandler('onResourceStop', function(ressource)
    if ressource == GetCurrentResourceName() then
        cacherAide()
    end
end)

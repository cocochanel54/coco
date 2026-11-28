-- Plantations (côté joueur)

local plantes = {} -- [id] = { donnees, objet }
local occupe = false
local aideAffichee = false

local function notifier(texte, genre)
    lib.notify({ title = 'Plantation', description = texte, type = genre or 'inform', icon = 'seedling', position = 'top' })
end

local function chargerModele(nom)
    local hash = GetHashKey(nom)
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

local function sol(x, y, z)
    for _, decalage in ipairs({ 1.0, 3.0, 10.0 }) do
        local trouve, solZ = GetGroundZFor_3dCoord(x, y, z + decalage, false)
        if trouve then
            return solZ
        end
    end
    return z
end

local function supprimerObjet(plante)
    if plante.objet and DoesEntityExist(plante.objet) then
        DeleteObject(plante.objet)
    end
    plante.objet = nil
end

local function creerObjet(plante)
    local d = plante.donnees
    local modeles = Config.Plantes[d.type] and Config.Plantes[d.type].Modeles
    if not modeles then
        return
    end
    -- Un plant mort reste petit
    local modele = modeles[d.mort and 1 or d.stade] or modeles[1]
    local hash = chargerModele(modele) or chargerModele('prop_weed_01')
    if not hash then
        return
    end
    local objet = CreateObject(hash, d.x, d.y, d.z, false, false, false)
    SetModelAsNoLongerNeeded(hash)
    FreezeEntityPosition(objet, true)
    SetEntityInvincible(objet, true)
    if d.mort then
        SetEntityAlpha(objet, 150, false)
    end
    plante.objet = objet
end

RegisterNetEvent('vox_plantations:afficher', function(donnees)
    local plante = plantes[donnees.id]
    if plante then
        supprimerObjet(plante)
    else
        plante = {}
        plantes[donnees.id] = plante
    end
    plante.donnees = donnees
    creerObjet(plante)
end)

RegisterNetEvent('vox_plantations:retirer', function(id)
    local plante = plantes[id]
    if plante then
        supprimerObjet(plante)
        plantes[id] = nil
    end
end)

---------------------------------------------------------------------------
-- Planter (quand on utilise une graine)
---------------------------------------------------------------------------

RegisterNetEvent('vox_plantations:planter', function(nomType)
    if occupe then
        return
    end
    local ped = PlayerPedId()
    if IsPedInAnyVehicle(ped, false) or IsEntityDead(ped) then
        return notifier('Impossible maintenant.', 'error')
    end
    if GetInteriorFromEntity(ped) ~= 0 then
        return notifier('Il faut planter dehors, en pleine terre.', 'error')
    end
    if IsEntityInWater(ped) then
        return notifier('Pas dans l\'eau !', 'error')
    end
    local devant = GetOffsetFromEntityInWorldCoords(ped, 0.0, 0.8, 0.0)
    local z = sol(devant.x, devant.y, devant.z)
    if IsPointOnRoad(devant.x, devant.y, z, 0) then
        return notifier('Pas sur la route, trouve un coin de terre discret.', 'error')
    end

    occupe = true
    local reussi = lib.progressBar({
        duration = 6000,
        label = 'Plantation de la graine…',
        canCancel = true,
        disable = { move = true, car = true, combat = true },
        anim = { dict = 'amb@world_human_gardener_plant@male@base', clip = 'base' },
    })
    occupe = false
    if reussi then
        TriggerServerEvent('vox_plantations:planter', nomType, vector3(devant.x, devant.y, z))
    end
end)

---------------------------------------------------------------------------
-- Menu d'un plant
---------------------------------------------------------------------------

local ACTIONS = {
    arroser = { duree = 3000, label = 'Arrosage…', anim = { dict = 'weapon@w_sp_jerrycan', clip = 'fire' } },
    engrais = { duree = 3000, label = 'Ajout d\'engrais…', anim = { dict = 'amb@world_human_gardener_plant@male@base', clip = 'base' } },
    recolter = { duree = 8000, label = 'Récolte…', anim = { dict = 'amb@world_human_gardener_plant@male@base', clip = 'base' } },
    detruire = { duree = 5000, label = 'Arrachage du plant…', anim = { dict = 'amb@world_human_gardener_plant@male@base', clip = 'base' } },
}

local function agir(id, nomAction)
    local action = ACTIONS[nomAction]
    occupe = true
    local reussi = lib.progressBar({
        duration = action.duree,
        label = action.label,
        canCancel = true,
        disable = { move = true, car = true, combat = true },
        anim = action.anim,
    })
    occupe = false
    if reussi then
        TriggerServerEvent('vox_plantations:action', id, nomAction)
    end
end

local function ouvrirMenu(id)
    local infos = lib.callback.await('vox_plantations:infos', false, id)
    if not infos then
        return
    end
    local etat
    if infos.mort then
        etat = 'Ce plant est mort (trop longtemps sans eau).'
    elseif infos.prete then
        etat = 'Prêt à être récolté !'
    else
        etat = ('Pousse à %d %% (encore environ %d min avec de l\'eau)'):format(infos.croissance, infos.reste)
    end
    lib.registerContext({
        id = 'vox_plantation',
        title = infos.label,
        options = {
            { title = etat, icon = 'seedling', readOnly = true },
            { title = ('Eau : %d %%'):format(infos.eau), icon = 'droplet', progress = infos.eau, colorScheme = 'blue', readOnly = true },
            { title = ('Qualité : %d %%'):format(infos.qualite), icon = 'star', progress = infos.qualite, colorScheme = 'green', readOnly = true },
            {
                title = 'Arroser',
                description = 'Utilise une bouteille d\'eau',
                icon = 'bottle-water',
                disabled = infos.mort,
                onSelect = function() agir(id, 'arroser') end,
            },
            {
                title = 'Mettre de l\'engrais',
                description = 'Améliore la qualité, donc la récolte',
                icon = 'flask',
                disabled = infos.mort,
                onSelect = function() agir(id, 'engrais') end,
            },
            {
                title = 'Récolter',
                icon = 'hand-holding',
                disabled = not infos.prete,
                onSelect = function() agir(id, 'recolter') end,
            },
            {
                title = (infos.police and not infos.proprietaire) and 'Saisir et détruire' or 'Arracher le plant',
                icon = 'trash',
                onSelect = function() agir(id, 'detruire') end,
            },
        },
    })
    lib.showContext('vox_plantation')
end

CreateThread(function()
    while true do
        local attente = 1000
        local ped = PlayerPedId()
        local position = GetEntityCoords(ped)
        local plusProche, distanceMin = nil, 1.6
        for id, plante in pairs(plantes) do
            local d = plante.donnees
            local distance = #(position - vector3(d.x, d.y, d.z))
            if distance < 15.0 then
                attente = 0
            end
            if distance < distanceMin then
                plusProche, distanceMin = id, distance
            end
        end
        if plusProche and not occupe and not IsPedInAnyVehicle(ped, false) then
            if not aideAffichee then
                aideAffichee = true
                lib.showTextUI('[E] Voir le plant', { position = 'right-center' })
            end
            if IsControlJustReleased(0, 38) then
                ouvrirMenu(plusProche)
            end
        elseif aideAffichee then
            aideAffichee = false
            lib.hideTextUI()
        end
        Wait(attente)
    end
end)

AddEventHandler('onResourceStop', function(ressource)
    if ressource ~= GetCurrentResourceName() then
        return
    end
    for _, plante in pairs(plantes) do
        supprimerObjet(plante)
    end
    if aideAffichee then
        lib.hideTextUI()
    end
end)

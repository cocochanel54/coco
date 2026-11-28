-- Convoi de fonds : un fourgon blindé conduit par des PNJ traverse la ville.
-- Les braqueurs doivent l'immobiliser, faire sauter les portes arrière et prendre les sacs.
-- La police est prévenue dès que le fourgon est attaqué.

local cfg = Config.Evenements.convoi

local function donneesClient(instance)
    return {
        id = instance.id,
        depart = instance.depart.nom,
        arrivee = instance.arrivee.nom,
        approx = instance.depart.coords,
        destination = instance.arrivee.coords,
        netVehicule = instance.netVehicule,
        netPeds = instance.netPeds,
        etat = instance.etat,
        sacs = instance.sacs,
        position = instance.position,
    }
end

local function diffuserEtat(instance, extra)
    TriggerClientEvent('vox_events:convoi:etat', -1, instance.id, instance.etat, instance.sacs, extra)
end

-- Point derrière le fourgon (là où sont les portes)
local function arriere(instance)
    local position = GetEntityCoords(instance.vehicule)
    local devant = Outils.directions(GetEntityHeading(instance.vehicule))
    return position - devant * 3.8
end

local function choisirTrajet()
    local paires = {}
    for _, a in ipairs(Config.Lieux) do
        for _, b in ipairs(Config.Lieux) do
            local d = Outils.distance2D(a.coords, b.coords)
            if a ~= b and d >= cfg.DistanceMin and d <= cfg.DistanceMax then
                paires[#paires + 1] = { a, b }
            end
        end
    end
    if #paires == 0 then
        return nil
    end
    local paire = paires[math.random(#paires)]
    return paire[1], paire[2]
end

local function alerterPolice(instance, texte)
    if instance.alerte then
        return
    end
    instance.alerte = true
    Bridge.alerter({ 'police' }, {
        id = 'convoi' .. instance.id,
        coords = instance.position or instance.depart.coords,
        titre = 'Convoi de fonds attaqué',
        texte = texte,
        icone = 'truck',
        couleur = 1,
        duree = cfg.Duree * 60,
    })
end

local function attendreEntite(entite)
    for _ = 1, 50 do
        if entite and entite ~= 0 and DoesEntityExist(entite) then
            return true
        end
        Wait(100)
    end
    return false
end

-- Surveille le fourgon toutes les 2 secondes : arrivée, attaque, immobilisation
local function surveiller(instance)
    CreateThread(function()
        local dernierePosition, immobile, tour = nil, 0, 0
        while Directeur.actif('convoi') == instance do
            local vehicule = instance.vehicule
            if not DoesEntityExist(vehicule) then
                Directeur.terminer('convoi', 'disparu')
                return
            end
            local position = GetEntityCoords(vehicule)
            instance.position = position
            tour = tour + 1
            if tour % 2 == 0 then
                TriggerClientEvent('vox_events:convoi:position', -1, instance.id, position)
            end

            if instance.etat == 'route' then
                if Outils.distance2D(position, instance.arrivee.coords) < cfg.DistanceArrivee then
                    Bridge.notifier(-1, {
                        title = 'Convoi de fonds',
                        description = ('Le fourgon est arrivé à bon port à %s.'):format(instance.arrivee.nom),
                        type = 'inform',
                        icon = 'truck',
                    })
                    Directeur.terminer('convoi', 'arrive')
                    return
                end

                -- Attaqué : carrosserie abîmée ou garde blessé
                local attaque = GetVehicleBodyHealth(vehicule) < instance.santeDepart - 40
                for i, ped in ipairs(instance.peds) do
                    if not DoesEntityExist(ped) or GetEntityHealth(ped) < instance.santePeds[i] - 20 then
                        attaque = true
                    end
                end
                if attaque then
                    alerterPolice(instance, ('Le fourgon blindé parti de %s est attaqué ! Suis le point sur ta carte.'):format(instance.depart.nom))
                end

                if dernierePosition and #(position - dernierePosition) < 1.5 then
                    immobile = immobile + 1
                else
                    immobile = 0
                end
                -- Immobilisé : arrêté 10 s après une attaque, ou bloqué 30 s
                if (instance.alerte and immobile >= 5) or immobile >= 15 then
                    instance.etat = 'immobilise'
                    diffuserEtat(instance)
                end
            end
            dernierePosition = position
            Wait(2000)
        end
    end)
end

local function creerConvoi(instance, position, cap)
    local vehicule = CreateVehicleServerSetter(GetHashKey(cfg.Vehicule), 'automobile', position.x, position.y, position.z + 1.0, cap)
    if not attendreEntite(vehicule) then
        print('[vox_events] Impossible de créer le fourgon du convoi.')
        Directeur.terminer('convoi', 'erreur')
        return
    end
    SetEntityOrphanMode(vehicule, 2) -- le fourgon reste même si son propriétaire se déconnecte
    SetVehicleDoorsLocked(vehicule, 2)

    local peds, netPeds = {}, {}
    for _, siege in ipairs({ -1, 0 }) do
        local ped = CreatePedInsideVehicle(vehicule, 4, GetHashKey(cfg.ModeleGarde), siege, true, true)
        if attendreEntite(ped) then
            SetEntityOrphanMode(ped, 2)
            GiveWeaponToPed(ped, GetHashKey(cfg.Arme), 250, false, true)
            SetPedArmour(ped, 100)
            peds[#peds + 1] = ped
            netPeds[#netPeds + 1] = NetworkGetNetworkIdFromEntity(ped)
        end
    end

    instance.vehicule, instance.peds = vehicule, peds
    instance.netVehicule = NetworkGetNetworkIdFromEntity(vehicule)
    instance.netPeds = netPeds
    instance.position = GetEntityCoords(vehicule)
    instance.santeDepart = GetVehicleBodyHealth(vehicule)
    instance.santePeds = {}
    for i, ped in ipairs(peds) do
        instance.santePeds[i] = GetEntityHealth(ped)
    end
    instance.etat = 'route'
    TriggerClientEvent('vox_events:convoi:vehicule', -1, instance.id, instance.netVehicule, netPeds)
    diffuserEtat(instance)
    surveiller(instance)
end

local function supprimerEntites(instance)
    for _, ped in ipairs(instance.peds or {}) do
        if DoesEntityExist(ped) then
            DeleteEntity(ped)
        end
    end
    if instance.vehicule and DoesEntityExist(instance.vehicule) then
        DeleteEntity(instance.vehicule)
    end
end

Directeur.enregistrer({
    nom = 'convoi',
    titre = 'Convoi de fonds',
    config = cfg,

    demarrer = function(instance)
        local depart, arrivee = choisirTrajet()
        if not depart then
            instance.annule = 'aucun trajet assez long dans Config.Lieux'
            return
        end
        instance.depart, instance.arrivee = depart, arrivee
        instance.etat = 'attente'
        instance.sacs = cfg.Sacs
        instance.peds = {}

        TriggerClientEvent('vox_events:convoi:debut', -1, donneesClient(instance))
        local police = {}
        for _, id in ipairs(Bridge.joueursEnService('police')) do
            police[id] = true
            Bridge.notifier(id, {
                title = 'Convoi de fonds',
                description = ('Un fourgon blindé va rouler de %s à %s. Restez vigilants.'):format(depart.nom, arrivee.nom),
                type = 'inform',
                icon = 'truck',
                duration = 12000,
            })
        end
        for _, idTexte in ipairs(GetPlayers()) do
            local id = tonumber(idTexte)
            if not police[id] then
                Bridge.notifier(id, {
                    title = 'Rumeur',
                    description = ('On dit qu\'un fourgon blindé va bientôt quitter %s pour %s…'):format(depart.nom, arrivee.nom),
                    type = 'inform',
                    icon = 'user-secret',
                    duration = 12000,
                })
            end
        end
    end,

    arreter = function(instance, raison)
        TriggerClientEvent('vox_events:convoi:fin', -1, instance.id)
        for _, id in ipairs(Bridge.joueursEnService('police')) do
            TriggerClientEvent('vox_events:alerteFin', id, 'convoi' .. instance.id)
        end
        if raison == 'fini' then
            -- On laisse l'épave quelques instants avant de la faire disparaître
            SetTimeout(30000, function()
                supprimerEntites(instance)
            end)
        else
            supprimerEntites(instance)
        end
    end,

    envoyerEtat = function(instance, src)
        TriggerClientEvent('vox_events:convoi:debut', src, donneesClient(instance))
    end,
})

-- Le premier joueur proche du dépôt indique la route : le fourgon apparaît là
RegisterNetEvent('vox_events:convoi:placer', function(id, position, cap)
    local src = source
    local instance = Directeur.actif('convoi')
    position = Outils.vecteur(position)
    if not instance or instance.id ~= id or instance.etat ~= 'attente' or not position or type(cap) ~= 'number' then
        return
    end
    local joueur = Outils.positionJoueur(src)
    if Outils.distance2D(position, instance.depart.coords) > 250.0 or not joueur or Outils.distance2D(joueur, position) > 350.0 then
        return
    end
    instance.etat = 'creation'
    creerConvoi(instance, position, cap % 360)
end)

RegisterNetEvent('vox_events:convoi:charge', function(id)
    local src = source
    local instance = Directeur.actif('convoi')
    if not instance or instance.id ~= id or instance.etat ~= 'immobilise' then
        return
    end
    local joueur = Outils.positionJoueur(src)
    if not joueur or Outils.distance2D(joueur, arriere(instance)) > 4.5 then
        return
    end
    if cfg.ObjetCharge and not Bridge.retirerObjet(src, cfg.ObjetCharge, 1) then
        Bridge.notifier(src, {
            title = 'Convoi de fonds',
            description = 'Il te faut de quoi faire sauter les portes.',
            type = 'error',
        })
        return
    end

    instance.etat = 'charge'
    alerterPolice(instance, ('Explosif posé sur le fourgon blindé parti de %s ! Suis le point sur ta carte.'):format(instance.depart.nom))
    diffuserEtat(instance, cfg.DelaiExplosion)

    SetTimeout(cfg.DelaiExplosion * 1000, function()
        if Directeur.actif('convoi') ~= instance then
            return
        end
        instance.etat = 'ouvert'
        -- L'explosion est créée par un seul joueur : celui qui « possède » le fourgon, sinon le poseur
        local proprietaire = NetworkGetEntityOwner(instance.vehicule)
        local cible = (proprietaire and proprietaire > 0) and proprietaire or src
        TriggerClientEvent('vox_events:convoi:explosion', cible, instance.id, arriere(instance))
        diffuserEtat(instance)
    end)
end)

RegisterNetEvent('vox_events:convoi:sac', function(id)
    local src = source
    local instance = Directeur.actif('convoi')
    if not instance or instance.id ~= id or instance.etat ~= 'ouvert' or instance.sacs <= 0 then
        return
    end
    local joueur = Outils.positionJoueur(src)
    if not joueur or Outils.distance2D(joueur, arriere(instance)) > 4.5 then
        return
    end

    instance.sacs = instance.sacs - 1
    local montant = math.random(cfg.ArgentParSac[1], cfg.ArgentParSac[2])
    if not (cfg.ObjetArgent and Bridge.donnerObjet(src, cfg.ObjetArgent, montant)) then
        Bridge.donnerArgent(src, montant, 'vox_events:convoi')
    end
    Bridge.notifier(src, {
        title = 'Convoi de fonds',
        description = ('Sac récupéré : %s'):format(Outils.argent(montant)),
        type = 'success',
        icon = 'sack-dollar',
    })
    diffuserEtat(instance)

    if instance.sacs <= 0 then
        Bridge.notifier(-1, {
            title = 'Info',
            description = 'Un fourgon blindé a été entièrement pillé !',
            type = 'warning',
            icon = 'truck',
        })
        Directeur.terminer('convoi', 'fini')
    end
end)

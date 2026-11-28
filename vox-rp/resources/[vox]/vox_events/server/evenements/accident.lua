-- Accident de la route : des blessés à soigner (ambulanciers), des épaves à dégager
-- (mécanos) et une zone à sécuriser (police). Chacun est payé pour son travail.

local cfg = Config.Evenements.accident

local function donneesClient(instance)
    return {
        id = instance.id,
        lieu = instance.lieu,
        approx = instance.approx,
        site = instance.site,
        vehicules = instance.vehicules,
        victimes = instance.victimes,
    }
end

-- Positions des épaves et des blessés autour du point de la route
local function construireSite(position, cap, nbVictimes)
    local devant, droite = Outils.directions(cap)
    local site = { position = position, cap = cap, vehicules = {}, victimes = {} }
    site.vehicules[1] = { position = position, cap = (cap + 25.0) % 360 }
    site.vehicules[2] = { position = position + devant * 5.5 + droite * 1.5, cap = (cap + 160.0) % 360 }
    site.vehicules[3] = { position = position - devant * 6.0 - droite * 1.0, cap = (cap + 300.0) % 360 }
    for i = 1, nbVictimes do
        site.victimes[i] = {
            position = position + droite * (4.5 + i * 0.8) + devant * ((i - 2) * 2.0),
            cap = (cap + 90.0 * i) % 360,
        }
    end
    return site
end

local function finir(instance)
    for _, victime in ipairs(instance.victimes) do
        if victime.etat ~= 'soignee' then
            return
        end
    end
    for _, vehicule in ipairs(instance.vehicules) do
        if vehicule.etat ~= 'degage' then
            return
        end
    end
    Bridge.notifier(-1, {
        title = 'Info trafic',
        description = ('La route est dégagée près de %s. Merci aux secours !'):format(instance.lieu),
        type = 'success',
        icon = 'road',
    })
    Directeur.terminer('accident', 'fini')
end

local function prevenirChangement(instance, genre, index, etat)
    TriggerClientEvent('vox_events:accident:maj', -1, instance.id, genre, index, etat)
end

-- Les policiers en service qui viennent sécuriser la zone sont payés une fois
local function surveillerPolice(instance)
    CreateThread(function()
        while Directeur.actif('accident') == instance do
            if instance.site then
                for _, id in ipairs(Bridge.joueursEnService('police')) do
                    local position = Outils.positionJoueur(id)
                    if not instance.policiersPayes[id] and position and Outils.distance2D(position, instance.site.position) < 40.0 then
                        instance.policiersPayes[id] = true
                        Bridge.donnerArgent(id, cfg.Primes.police, 'vox_events:accident')
                        Bridge.notifier(id, {
                            title = 'Accident',
                            description = ('Zone sécurisée, merci ! Prime : %s'):format(Outils.argent(cfg.Primes.police)),
                            type = 'success',
                            icon = 'shield-halved',
                        })
                    end
                end
            end
            Wait(5000)
        end
    end)
end

Directeur.enregistrer({
    nom = 'accident',
    titre = 'Accident de la route',
    config = cfg,

    -- Pendant une tempête, les accidents tombent plus souvent
    poids = function()
        local poids = cfg.Poids
        if Directeur.actif('tempete') then
            poids = poids * Config.Evenements.tempete.BonusAccident
        end
        return poids
    end,

    peutDemarrer = function(ctx)
        if ctx.police + ctx.ems + ctx.mecanos < cfg.MinSecours then
            return false, 'il faut au moins un policier, ambulancier ou mécano en service'
        end
        return true
    end,

    demarrer = function(instance)
        local lieu = Config.Lieux[math.random(#Config.Lieux)]
        instance.lieu = lieu.nom
        instance.approx = lieu.coords
        instance.site = nil
        instance.policiersPayes = {}

        instance.vehicules = {}
        for i = 1, math.random(2, 3) do
            instance.vehicules[i] = { modele = cfg.Vehicules[math.random(#cfg.Vehicules)], etat = 'present' }
        end
        instance.victimes = {}
        for i = 1, math.random(cfg.Victimes[1], cfg.Victimes[2]) do
            instance.victimes[i] = { modele = cfg.ModelesVictimes[math.random(#cfg.ModelesVictimes)], etat = 'blessee' }
        end

        TriggerClientEvent('vox_events:accident:debut', -1, donneesClient(instance))
        Bridge.alerter({ 'police', 'ems', 'mecano' }, {
            id = 'accident' .. instance.id,
            coords = lieu.coords,
            titre = 'Accident de la route',
            texte = ('Accident signalé près de %s : %d blessé(s) et des véhicules sur la chaussée.'):format(lieu.nom, #instance.victimes),
            icone = 'car-burst',
            couleur = 1,
            duree = cfg.Duree * 60,
        })
        Bridge.notifier(-1, {
            title = 'Info trafic',
            description = ('Un accident vient d\'être signalé près de %s. Ralentissez !'):format(lieu.nom),
            type = 'warning',
            icon = 'car-burst',
        })
        surveillerPolice(instance)
    end,

    arreter = function(instance, raison)
        TriggerClientEvent('vox_events:accident:fin', -1, instance.id)
        for _, role in ipairs({ 'police', 'ems', 'mecano' }) do
            for _, id in ipairs(Bridge.joueursEnService(role)) do
                TriggerClientEvent('vox_events:alerteFin', id, 'accident' .. instance.id)
            end
        end
        if raison ~= 'fini' then
            Bridge.notifier(-1, {
                title = 'Info trafic',
                description = ('Les secours de l\'État ont pris le relais près de %s.'):format(instance.lieu),
                type = 'inform',
                icon = 'road',
            })
        end
    end,

    envoyerEtat = function(instance, src)
        TriggerClientEvent('vox_events:accident:debut', src, donneesClient(instance))
    end,
})

-- Le premier joueur proche indique la route la plus proche (le jeu connaît les routes autour de lui)
RegisterNetEvent('vox_events:accident:placer', function(id, position, cap)
    local src = source
    local instance = Directeur.actif('accident')
    position = Outils.vecteur(position)
    if not instance or instance.id ~= id or instance.site or not position or type(cap) ~= 'number' then
        return
    end
    local joueur = Outils.positionJoueur(src)
    if Outils.distance2D(position, instance.approx) > 250.0 or not joueur or Outils.distance2D(joueur, position) > 300.0 then
        return
    end
    instance.site = construireSite(position, cap % 360, #instance.victimes)
    TriggerClientEvent('vox_events:accident:site', -1, id, instance.site)
end)

RegisterNetEvent('vox_events:accident:soigner', function(id, index)
    local src = source
    local instance = Directeur.actif('accident')
    if not instance or instance.id ~= id or not instance.site or type(index) ~= 'number' then
        return
    end
    local victime = instance.victimes[index]
    local place = instance.site.victimes[index]
    local joueur = Outils.positionJoueur(src)
    if not victime or victime.etat ~= 'blessee' or not joueur or Outils.distance2D(joueur, place.position) > 3.5 then
        return
    end

    local prime
    if Bridge.estEnService(src, 'ems') then
        prime = cfg.Primes.soin
    elseif cfg.SoinsCivilsSansEms and Bridge.compterEnService('ems') == 0 then
        prime = cfg.PrimeCivile
    else
        Bridge.notifier(src, {
            title = 'Accident',
            description = 'Laisse faire les ambulanciers, ils sont en route !',
            type = 'error',
        })
        return
    end

    victime.etat = 'soignee'
    prevenirChangement(instance, 'victime', index, victime.etat)
    Bridge.donnerArgent(src, prime, 'vox_events:accident')
    Bridge.notifier(src, {
        title = 'Accident',
        description = ('Blessé pris en charge. Prime : %s'):format(Outils.argent(prime)),
        type = 'success',
        icon = 'kit-medical',
    })
    finir(instance)
end)

RegisterNetEvent('vox_events:accident:degager', function(id, index)
    local src = source
    local instance = Directeur.actif('accident')
    if not instance or instance.id ~= id or not instance.site or type(index) ~= 'number' then
        return
    end
    local vehicule = instance.vehicules[index]
    local place = instance.site.vehicules[index]
    local joueur = Outils.positionJoueur(src)
    if not vehicule or vehicule.etat ~= 'present' or not joueur or Outils.distance2D(joueur, place.position) > 6.0 then
        return
    end

    -- Les mécanos en priorité ; la police peut dégager s'il n'y a aucun mécano en service
    local autorise = Bridge.estEnService(src, 'mecano')
        or (Bridge.compterEnService('mecano') == 0 and Bridge.estEnService(src, 'police'))
    if not autorise then
        Bridge.notifier(src, {
            title = 'Accident',
            description = 'Il faut une dépanneuse : appelle un mécano !',
            type = 'error',
        })
        return
    end

    vehicule.etat = 'degage'
    prevenirChangement(instance, 'vehicule', index, vehicule.etat)
    Bridge.donnerArgent(src, cfg.Primes.degagement, 'vox_events:accident')
    Bridge.notifier(src, {
        title = 'Accident',
        description = ('Épave dégagée. Prime : %s'):format(Outils.argent(cfg.Primes.degagement)),
        type = 'success',
        icon = 'truck-pickup',
    })
    finir(instance)
end)

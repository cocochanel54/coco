-- Colis mystère : un colis est caché quelque part dans une zone, le premier qui l'ouvre gagne

local cfg = Config.Evenements.colis

local function donneesClient(instance)
    return {
        id = instance.id,
        lieu = instance.lieu,
        zone = instance.zone,
        rayon = cfg.RayonZone,
        approx = instance.approx,
        position = instance.position,
    }
end

local function donnerRecompense(src)
    local recompense = Outils.tirage(cfg.Recompenses)
    if recompense.objet and Bridge.donnerObjet(src, recompense.objet, recompense.quantite or 1) then
        return ('%dx %s'):format(recompense.quantite or 1, recompense.objet)
    end
    -- Argent (ou objet impossible à donner : inventaire plein, objet inconnu...)
    local fourchette = recompense.argent or { 1500, 2500 }
    local montant = math.random(fourchette[1], fourchette[2])
    Bridge.donnerArgent(src, montant, 'vox_events:colis')
    return Outils.argent(montant)
end

Directeur.enregistrer({
    nom = 'colis',
    titre = 'Colis mystère',
    config = cfg,

    demarrer = function(instance)
        local lieu = Config.Lieux[math.random(#Config.Lieux)]
        local angle = math.random() * 2 * math.pi
        local ecart = math.random() * cfg.Ecart
        instance.lieu = lieu.nom
        instance.zone = lieu.coords
        -- Point approximatif : le premier joueur qui passe à côté trouve le trottoir le plus proche
        instance.approx = vector3(lieu.coords.x + math.cos(angle) * ecart, lieu.coords.y + math.sin(angle) * ecart, lieu.coords.z)
        instance.position = nil
        instance.ouvert = false

        TriggerClientEvent('vox_events:colis:debut', -1, donneesClient(instance))
        Bridge.notifier(-1, {
            title = 'Colis mystère',
            description = ('Un colis a été perdu du côté de %s. Le cercle jaune est sur ta carte : premier arrivé, premier servi !'):format(lieu.nom),
            type = 'inform',
            icon = 'box',
            duration = 12000,
        })
    end,

    arreter = function(instance, raison)
        TriggerClientEvent('vox_events:colis:fin', -1, instance.id)
        if raison ~= 'fini' then
            Bridge.notifier(-1, {
                title = 'Colis mystère',
                description = 'Personne n\'a trouvé le colis… il a disparu.',
                type = 'inform',
                icon = 'box',
            })
        end
    end,

    envoyerEtat = function(instance, src)
        TriggerClientEvent('vox_events:colis:debut', src, donneesClient(instance))
    end,
})

-- Le premier joueur proche indique où poser le colis (sur le trottoir le plus proche)
RegisterNetEvent('vox_events:colis:placer', function(id, position)
    local src = source
    local instance = Directeur.actif('colis')
    position = Outils.vecteur(position)
    if not instance or instance.id ~= id or instance.position or not position then
        return
    end
    local joueur = Outils.positionJoueur(src)
    if Outils.distance2D(position, instance.approx) > 80.0 or not joueur or Outils.distance2D(joueur, position) > 300.0 then
        return
    end
    instance.position = position
    TriggerClientEvent('vox_events:colis:position', -1, id, position)
end)

RegisterNetEvent('vox_events:colis:ouvrir', function(id)
    local src = source
    local instance = Directeur.actif('colis')
    if not instance or instance.id ~= id or not instance.position then
        return
    end
    if instance.ouvert then
        Bridge.notifier(src, { title = 'Colis mystère', description = 'Trop tard, quelqu\'un l\'a déjà ouvert !', type = 'error' })
        return
    end
    local joueur = Outils.positionJoueur(src)
    if not joueur or Outils.distance2D(joueur, instance.position) > 3.0 or math.abs(joueur.z - instance.position.z) > 3.0 then
        return
    end

    instance.ouvert = true
    local gain = donnerRecompense(src)
    Bridge.notifier(src, {
        title = 'Colis mystère',
        description = ('Bien joué ! Dans le colis : %s'):format(gain),
        type = 'success',
        icon = 'box-open',
        duration = 10000,
    })
    Bridge.notifier(-1, {
        title = 'Colis mystère',
        description = 'Quelqu\'un a trouvé le colis. À la prochaine !',
        type = 'inform',
        icon = 'box',
    })
    Directeur.terminer('colis', 'fini')
end)

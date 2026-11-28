-- Tempête : orage et vent violent. Plus d'accidents, et la foudre peut couper le courant.

local cfg = Config.Evenements.tempete

Directeur.enregistrer({
    nom = 'tempete',
    titre = 'Tempête',
    config = cfg,

    demarrer = function(instance)
        Bridge.meteo.tempete(true, cfg.Meteo)
        Bridge.notifier(-1, {
            title = 'Alerte météo',
            description = 'Une grosse tempête arrive sur la ville : vent violent et orages. Roulez prudemment !',
            type = 'warning',
            icon = 'cloud-bolt',
            duration = 12000,
        })

        -- La foudre peut provoquer une panne pendant la tempête
        if math.random(100) <= cfg.ChancePanne then
            local delai = math.random(2, math.max(2, cfg.Duree - 5)) * 60000
            SetTimeout(delai, function()
                if Directeur.actif('tempete') == instance and not Directeur.actif('panne') then
                    Directeur.lancer('panne', { force = true, cause = 'tempete' })
                end
            end)
        end
    end,

    arreter = function()
        Bridge.meteo.tempete(false, cfg.MeteoApres)
        Bridge.notifier(-1, {
            title = 'Alerte météo',
            description = 'La tempête s\'éloigne de la ville.',
            type = 'inform',
            icon = 'cloud-sun',
        })
    end,
})

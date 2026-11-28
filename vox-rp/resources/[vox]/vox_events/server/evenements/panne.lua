-- Panne de courant : toute la ville s'éteint quelques minutes

local cfg = Config.Evenements.panne

Directeur.enregistrer({
    nom = 'panne',
    titre = 'Panne de courant',
    config = cfg,

    demarrer = function(instance)
        Bridge.meteo.panne(true)
        local texte = instance.cause == 'tempete'
            and 'La foudre a frappé une centrale : la ville est plongée dans le noir !'
            or 'Une panne géante plonge la ville dans le noir. Les caméras de surveillance sont coupées !'
        Bridge.notifier(-1, {
            title = 'Panne de courant',
            description = texte,
            type = 'warning',
            icon = 'bolt',
            duration = 12000,
        })
    end,

    arreter = function()
        Bridge.meteo.panne(false)
        Bridge.notifier(-1, {
            title = 'Courant rétabli',
            description = 'Les techniciens ont réparé le réseau électrique.',
            type = 'success',
            icon = 'lightbulb',
        })
    end,
})

-- Pour les futurs braquages : exports.vox_events:EnPanne()
exports('EnPanne', function()
    return Directeur.actif('panne') ~= nil
end)

-- Commandes du staff (réservées : permission « command.event », donnée au groupe admin)
--   /event               liste des événements et leur état
--   /event [nom]         lance un événement tout de suite
--   /eventstop [nom]     arrête un événement en cours
--   /eventauto           active / coupe le lancement automatique

local function etatDe(nom, def, ctx)
    if Directeur.enCours[nom] then
        return 'EN COURS'
    end
    local reste = (Directeur.rechargeJusqua[nom] or 0) - os.time()
    if reste > 0 then
        return ('recharge encore %d min'):format(math.ceil(reste / 60))
    end
    local ok, raison = Directeur.conditions(def, ctx)
    if not ok then
        return 'en attente : ' .. (raison or '?')
    end
    return 'prêt'
end

RegisterCommand('event', function(src, args)
    local nom = args[1] and args[1]:lower()
    if not nom then
        local ctx = Directeur.contexte()
        local noms = {}
        for n in pairs(Directeur.evenements) do
            noms[#noms + 1] = n
        end
        table.sort(noms)
        local lignes = {
            ('Lancement automatique : %s'):format(Directeur.auto and 'oui' or 'non'),
            ('En service : %d policiers, %d ambulanciers, %d mécanos'):format(ctx.police, ctx.ems, ctx.mecanos),
        }
        for _, n in ipairs(noms) do
            local def = Directeur.evenements[n]
            lignes[#lignes + 1] = ('- %s (%s) : %s'):format(n, def.titre, etatDe(n, def, ctx))
        end
        lignes[#lignes + 1] = 'Pour en lancer un : /event [nom]'
        Bridge.repondre(src, table.concat(lignes, '\n'))
        return
    end
    local ok, raison = Directeur.lancer(nom, { force = true })
    Bridge.repondre(src, ok and ('« %s » est lancé.'):format(nom) or ('Impossible : %s.'):format(raison))
end, true)

RegisterCommand('eventstop', function(src, args)
    local nom = args[1] and args[1]:lower()
    if not nom then
        Bridge.repondre(src, 'Utilise /eventstop [nom]')
        return
    end
    local ok = Directeur.terminer(nom, 'staff')
    Bridge.repondre(src, ok and ('« %s » est arrêté.'):format(nom) or ('« %s » n\'est pas en cours.'):format(nom))
end, true)

RegisterCommand('eventauto', function(src)
    Directeur.auto = not Directeur.auto
    Bridge.repondre(src, Directeur.auto and 'Lancement automatique activé.' or 'Lancement automatique coupé.')
end, true)

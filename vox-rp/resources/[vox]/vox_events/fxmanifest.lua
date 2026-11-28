fx_version 'cerulean'
game 'gta5'
lua54 'yes'

name 'vox_events'
author 'Vox RP'
description 'Directeur d\'événements : la ville lance toute seule colis mystères, accidents, pannes et tempêtes'
version '0.1.0'

shared_scripts {
    '@ox_lib/init.lua',
    'config.lua',
}

client_scripts {
    'client/outils.lua',
    'client/evenements/*.lua',
}

server_scripts {
    'server/bridge.lua',
    'server/directeur.lua',
    'server/evenements/*.lua',
    'server/commandes.lua',
}

dependencies {
    'ox_lib',
    'qbx_core',
}

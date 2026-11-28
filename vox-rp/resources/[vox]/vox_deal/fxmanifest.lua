fx_version 'cerulean'
game 'gta5'
lua54 'yes'

name 'vox_deal'
author 'Vox RP'
description 'Vente de drogue aux PNJ dans la rue (avec témoins et appels à la police)'
version '0.1.0'

shared_scripts {
    '@ox_lib/init.lua',
    'config.lua',
}

client_script 'client.lua'
server_script 'server.lua'

dependencies {
    'ox_lib',
    'qbx_core',
    'ox_inventory',
    'vox_temoins',
}

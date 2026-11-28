fx_version 'cerulean'
game 'gta5'
lua54 'yes'

name 'vox_temoins'
author 'Vox RP'
description 'Les PNJ témoins appellent la police : coups de feu, deals, enlèvements, braquages'
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
}

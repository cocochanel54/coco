fx_version 'cerulean'
game 'gta5'
lua54 'yes'

name 'vox_plantations'
author 'Vox RP'
description 'Plantations : planter, arroser, engrais, récolte. Les plants poussent même hors ligne.'
version '0.1.0'

shared_scripts {
    '@ox_lib/init.lua',
    'config.lua',
}

client_script 'client.lua'

server_scripts {
    '@oxmysql/lib/MySQL.lua',
    'server.lua',
}

dependencies {
    'ox_lib',
    'oxmysql',
    'qbx_core',
    'ox_inventory',
}

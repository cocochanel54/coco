# Génère index.html (la page à ouvrir directement dans un navigateur) à partir de jeu.html,
# qui est aussi la version publiée sur claude.ai (sans en-tête de document : claude.ai l'ajoute).
import pathlib
d = pathlib.Path(__file__).parent
src = (d / 'jeu.html').read_text()
cut = src.index('<div id="app">')
(d / 'index.html').write_text(
    '<!doctype html>\n<html lang="fr">\n<head>\n<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
    '<style>:root{padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}[hidden]{display:none!important}</style>\n'
    + src[:cut] + '</head>\n<body>\n' + src[cut:] + '</body>\n</html>\n')

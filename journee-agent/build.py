# Génère les fichiers à partir de app.src.html, apps-script/Code.src.gs et data.json (extrait de Journee_agent.xlsm) :
#   index.html               : la page (artifact claude.ai ou ouverture directe)
#   apps-script/Index.html   : la même page, à coller dans Google Apps Script
#   apps-script/Code.gs      : le script Google Apps Script
import pathlib
d = pathlib.Path(__file__).parent
data = (d / 'data.json').read_text()
html = (d / 'app.src.html').read_text().replace('/*DATA*/null', data)
(d / 'index.html').write_text(html)
(d / 'apps-script' / 'Index.html').write_text(html)
(d / 'apps-script' / 'Code.gs').write_text((d / 'apps-script' / 'Code.src.gs').read_text().replace('/*DATA*/null', data))

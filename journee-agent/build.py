# Génère les fichiers à partir de app.src.html, apps-script/Code.src.gs et data.json (extrait de Journee_agent.xlsm) :
#   index.html               : la page, pour un essai dans le navigateur (saisies gardées sur l'appareil)
#   apps-script/Index.html   : la même page, à coller dans Google Apps Script
#   apps-script/Code.gs      : le script Google Apps Script (seul fichier qui contient les prix)
import json, pathlib
d = pathlib.Path(__file__).parent
data = json.loads((d / 'data.json').read_text())
# la page ne reçoit jamais les prix : [code, libellé, désignation, unité] sans le prix
page = dict(data, bpu={p: [x[:4] for x in l] for p, l in data['bpu'].items()})
html = (d / 'app.src.html').read_text().replace('/*DATA*/null', json.dumps(page, ensure_ascii=False, separators=(',', ':')))
(d / 'index.html').write_text(html)
(d / 'apps-script' / 'Index.html').write_text(html)
gs = (d / 'apps-script' / 'Code.src.gs').read_text().replace('/*DATA*/null', json.dumps(data, ensure_ascii=False, separators=(',', ':')))
(d / 'apps-script' / 'Code.gs').write_text(gs)

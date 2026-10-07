# Génère index.html à partir de app.src.html et de data.json (extrait du classeur Journee_agent.xlsm)
import pathlib
d = pathlib.Path(__file__).parent
html = (d / 'app.src.html').read_text().replace('/*DATA*/null', (d / 'data.json').read_text())
(d / 'index.html').write_text(html)

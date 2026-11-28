#!/usr/bin/env python3
"""Fabrique les photos des annonces Etsy : JPEG de 2700 × 2025 px (format 4:3).

1. LibreOffice ouvre chaque fichier « exemple » en français et l'exporte en PDF.
2. Les pages utiles sont converties en images et recadrées.
3. Chaque photo est mise en page en HTML, puis photographiée avec Playwright (shoot.js).

Usage : python3 photos/make_photos.py
Besoin : LibreOffice, poppler (pdftotext, pdftoppm), Pillow, Node + Playwright.
Si LibreOffice a besoin d'un utilitaire spécial pour démarrer (bac à sable), indiquer
son dossier dans la variable SOFFICE_HELPER_DIR.
"""
import html
import json
import os
import subprocess
import sys
from pathlib import Path

from PIL import Image

HERE = Path(__file__).resolve().parent
SHOP = HERE.parent
BUILD = HERE / "_build"
W, H, SCALE = 1200, 900, 2.25      # 1200 × 900 à l'échelle 2,25 = 2700 × 2025 px

XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>fr-FR</value></prop></item>
<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooLocale" oor:op="fuse"><value>fr-FR</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""

# Captures : produit -> (fichier exemple, {capture: (texte qui repère la page, hauteur max / largeur)})
CAPTURES = {
    "budget": ("budget/Budget_2027_exemple.xlsx", {
        "mois": ("Mon budget de janvier", 1.12),
        "annee": ("Bilan de l", None),
        "transactions": ("Une ligne par dépense ou revenu", 0.62),
        "parametres": ("Résumé de mon budget prévu", 0.9),
    }),
    "menus": ("menus/Menus_et_courses_exemple.xlsx", {
        "semaine": ("REPAS PRÉVUS", None),
        "liste": ("Autres courses", 1.05),
        "recettes": ("Mes recettes", 0.95),
    }),
    "epargne": ("epargne/Epargne_defi_52_semaines_exemple.xlsx", {
        "objectifs": ("Mes objectifs d", 0.62),
        "defi": ("semaines réussies sur 52", None),
    }),
    "comptes": ("comptes/Comptes_a_deux_exemple.xlsx", {
        "bilan": ("Qui doit combien", None),
        "depenses": ("Une ligne par dépense commune", 0.55),
    }),
}

CHIPS = ["Excel", "Google Sheets", "En français", "Téléchargement immédiat"]

# Photos de chaque annonce. Positions en pixels CSS (image de 1200 × 900).
PHOTOS = {
    "budget": [
        dict(kind="cover", name="01-couverture", pill="Modèle Excel & Google Sheets",
             title="Planificateur de budget 2027",
             sub="Suivez vos dépenses en quelques minutes par semaine.",
             bullets=["Bilan de chaque mois", "Bilan de l'année et graphiques", "Catégories modifiables",
                      "Mode d'emploi inclus"],
             windows=[("annee", 500, 70, 640, 400), ("mois", 780, 360, 360, 470)]),
        dict(kind="wide", name="02-bilan-annee", title="Le bilan de votre année en un coup d'œil",
             sub="Revenus, dépenses, épargne et solde du compte, calculés pour vous.", shot="annee",
             callouts=[("Graphiques et totaux automatiques", 640, 790)]),
        dict(kind="split", name="03-chaque-mois", title="Chaque mois, voyez où part votre argent",
             sub="Choisissez un mois : tout se met à jour.", shot="mois",
             bullets=["Prévu et réel par catégorie", "Dépassements en rouge", "Ce qu'il vous reste, en grand"]),
        dict(kind="wide", name="04-saisie", title="Notez une dépense en 10 secondes",
             sub="Date, description, catégorie dans une liste, montant. C'est tout.", shot="transactions",
             callouts=[("Catégories en liste déroulante", 700, 790)]),
        dict(kind="howto", name="05-comment-ca-marche", product="Planificateur de budget 2027"),
    ],
    "menus": [
        dict(kind="cover", name="01-couverture", pill="Modèle Excel & Google Sheets",
             title="Menus de la semaine & liste de courses",
             sub="Choisissez vos repas : la liste de courses se fait toute seule.",
             bullets=["Liste rangée par rayon", "Quantités selon le nombre de personnes", "16 recettes incluses",
                      "Mode d'emploi inclus"],
             windows=[("semaine", 500, 70, 640, 400), ("liste", 760, 330, 380, 500)]),
        dict(kind="split", name="02-liste-automatique", title="Votre liste de courses se fait toute seule",
             sub="Rangée par rayon, avec les bonnes quantités.", shot="liste",
             bullets=["Rien à recopier", "Vos achats réguliers ajoutés", "Cochez ✓ dans le magasin"]),
        dict(kind="wide", name="03-planning", title="Planifiez midi et soir en 2 minutes",
             sub="Choisissez vos plats dans la liste de vos recettes.", shot="semaine",
             callouts=[("Le décompte se fait tout seul", 640, 170), ("Choix dans une liste", 850, 560)]),
        dict(kind="split", name="04-recettes", title="16 recettes déjà prêtes",
             sub="Lasagnes, quiche, curry, soupe… et vous ajoutez les vôtres.", shot="recettes",
             bullets=["Une ligne par ingrédient", "Quantités pour 4, adaptées tout seul", "Ingrédients en liste"]),
        dict(kind="howto", name="05-comment-ca-marche", product="Menus de la semaine & liste de courses"),
    ],
    "epargne": [
        dict(kind="cover", name="01-couverture", pill="Modèle Excel & Google Sheets",
             title="Défi 52 semaines & objectifs d'épargne",
             sub="1 378 € mis de côté en un an, presque sans y penser.",
             bullets=["Le défi à cocher chaque semaine", "Jusqu'à 12 objectifs", "La somme à épargner par mois",
                      "Mode d'emploi inclus"],
             windows=[("objectifs", 500, 70, 640, 400), ("defi", 800, 300, 340, 540)]),
        dict(kind="split", name="02-defi", title="Le défi 52 semaines",
             sub="1 € la première semaine, 2 € la deuxième… 1 378 € au bout d'un an.", shot="defi",
             bullets=["Cochez chaque semaine réussie", "Choisissez votre mise de départ", "Sens croissant ou décroissant"]),
        dict(kind="wide", name="03-objectifs", title="Vos projets, et combien épargner par mois",
             sub="Vacances, voiture, fonds d'urgence : suivez chaque objectif.", shot="objectifs",
             callouts=[("Progression automatique", 520, 650), ("« Atteint ! » quand c'est gagné", 700, 735)]),
        dict(kind="howto", name="04-comment-ca-marche", product="Défi 52 semaines & objectifs d'épargne"),
    ],
    "comptes": [
        dict(kind="cover", name="01-couverture", pill="Modèle Excel & Google Sheets",
             title="Comptes à deux",
             sub="Couple ou colocation : le tableur vous dit qui doit combien à qui.",
             bullets=["Calcul automatique", "Moitié-moitié ou selon les revenus", "Remboursements suivis",
                      "Mode d'emploi inclus"],
             windows=[("bilan", 500, 60, 640, 470), ("depenses", 620, 520, 520, 330)]),
        dict(kind="wide", name="02-bilan", title="Fini les disputes : le bilan est automatique",
             sub="Le montant à rembourser s'affiche en grand, mis à jour à chaque dépense.", shot="bilan",
             callouts=[("Calculé tout seul", 800, 285), ("Détail mois par mois", 810, 380)]),
        dict(kind="wide", name="03-partage", title="Partagez comme vous voulez",
             sub="Moitié-moitié, selon les revenus ou un pourcentage libre : la part de chacun est calculée.",
             shot="depenses", callouts=[("Un partage spécial sur une ligne", 600, 640),
                                        ("Remboursements suivis", 150, 725)]),
        dict(kind="howto", name="04-comment-ca-marche", product="Comptes à deux"),
    ],
    "pack": [
        dict(kind="pack", name="01-couverture"),
        dict(kind="howto", name="02-comment-ca-marche", product="Pack des 4 modèles"),
    ],
}

CSS = """
:root{--dark:#1F4E5F;--light:#E3EFF1;--green:#2E8B6E;--greenl:#E4F3EE;--coral:#D9644A;--corall:#FBEAE5;
--cream:#FBF6EC;--text:#1F2937;--muted:#5B6472}
*{box-sizing:border-box;margin:0;padding:0}
html,body{width:1200px;height:900px}
body{font-family:Inter,sans-serif;color:var(--text);overflow:hidden;position:relative;
background:radial-gradient(circle at 88% 8%,#DCEBEE 0,transparent 42%),
radial-gradient(circle at 6% 96%,#F8E3DC 0,transparent 38%),var(--cream)}
h1{font-family:'Inter Display',Inter,sans-serif;font-weight:800;color:var(--dark);letter-spacing:-.02em;line-height:1.04}
.sub{color:var(--muted);line-height:1.35}
.pill{display:inline-block;font-weight:700;font-size:14px;letter-spacing:.1em;text-transform:uppercase;color:var(--dark);
border:2px solid var(--dark);border-radius:999px;padding:7px 15px;background:#fff}
ul.checks{list-style:none}
ul.checks li{font-size:21px;font-weight:600;display:flex;gap:14px;align-items:center;margin:13px 0;color:var(--text)}
ul.checks li::before{content:"✓";flex:none;width:30px;height:30px;border-radius:50%;background:var(--green);color:#fff;
font-size:17px;font-weight:800;display:flex;align-items:center;justify-content:center}
.chips{display:flex;flex-wrap:wrap;gap:9px}
.chips span{background:var(--dark);color:#fff;border-radius:999px;padding:8px 15px;font-weight:600;font-size:15px}
.window{position:absolute;background:#fff;border-radius:14px;overflow:hidden;
box-shadow:0 30px 60px -22px rgba(31,78,95,.45),0 0 0 1px rgba(31,78,95,.10)}
.window .bar{height:26px;background:#EEF1F3;display:flex;align-items:center;gap:7px;padding:0 12px;
border-bottom:1px solid #E2E6EA}
.window .bar i{width:10px;height:10px;border-radius:50%;background:#CDD3D9;display:block}
.window .shot{overflow:hidden;position:relative}
.window .shot img{display:block;width:100%}
.window .shot.cut::after{content:"";position:absolute;left:0;right:0;bottom:0;height:46px;
background:linear-gradient(rgba(255,255,255,0),#fff)}
.callout{position:absolute;background:var(--coral);color:#fff;font-weight:800;font-size:22px;padding:12px 20px;
border-radius:14px;box-shadow:0 14px 26px -10px rgba(217,100,74,.75);white-space:nowrap;z-index:9}
.left{position:absolute;left:70px;top:0;bottom:0;width:400px;display:flex;flex-direction:column;justify-content:center}
.brand{position:absolute;left:70px;bottom:34px;font-size:14px;color:var(--muted);font-weight:600;letter-spacing:.04em}
"""


def soffice(args):
    helper = os.environ.get("SOFFICE_HELPER_DIR")
    if helper:
        sys.path.insert(0, helper)
        from office.soffice import run_soffice
        return run_soffice(args, capture_output=True)
    return subprocess.run(["soffice"] + args, capture_output=True)


def render_captures():
    profile = BUILD / "lo_profile"
    (profile / "user").mkdir(parents=True, exist_ok=True)
    (profile / "user" / "registrymodifications.xcu").write_text(XCU, encoding="utf-8")
    shots = {}
    for product, (xlsx, wanted) in CAPTURES.items():
        src = SHOP / xlsx
        pdf = BUILD / (src.stem + ".pdf")
        soffice([f"-env:UserInstallation={profile.as_uri()}", "--headless", "--convert-to", "pdf",
                 "--outdir", str(BUILD), str(src)])
        if not pdf.exists():
            sys.exit(f"Échec de l'export PDF de {src}")
        pages = int(next(line.split()[-1] for line in subprocess.run(
            ["pdfinfo", str(pdf)], capture_output=True, text=True).stdout.splitlines() if line.startswith("Pages")))
        texts = [subprocess.run(["pdftotext", "-f", str(p), "-l", str(p), str(pdf), "-"],
                                capture_output=True, text=True).stdout for p in range(1, pages + 1)]
        for name, (marker, max_ratio) in wanted.items():
            page = next((i + 1 for i, t in enumerate(texts) if marker in t), None)
            if page is None:
                sys.exit(f"{product}/{name} : page introuvable (« {marker} »)")
            base = BUILD / f"{product}-{name}"
            subprocess.run(["pdftoppm", "-r", "220", "-png", "-singlefile", "-f", str(page), "-l", str(page),
                            str(pdf), str(base)], check=True)
            shots[f"{product}/{name}"] = crop(base.with_suffix(".png"), max_ratio)
            print(f"capture {product}/{name} : page {page}")
    return shots


def crop(path, max_ratio):
    img = Image.open(path).convert("RGB")
    gray = img.convert("L").point(lambda v: 255 if v < 246 else 0)
    left, top, right, bottom = gray.getbbox()
    pad = 26
    left, top = max(0, left - pad), max(0, top - pad)
    right, bottom = min(img.width, right + pad), min(img.height, bottom + pad)
    cut = False
    if max_ratio and (bottom - top) > (right - left) * max_ratio:
        bottom = top + int((right - left) * max_ratio)
        cut = True
    img.crop((left, top, right, bottom)).save(path)
    return {"src": path.as_uri(), "ratio": (bottom - top) / (right - left), "cut": cut}


def window(shot, left, top, width, max_h, z=1):
    shot_h = width * shot["ratio"]
    cut = shot["cut"] or shot_h > max_h
    return (f'<div class="window" style="left:{left}px;top:{top}px;width:{width}px;z-index:{z}">'
            f'<div class="bar"><i></i><i></i><i></i></div>'
            f'<div class="shot{" cut" if cut else ""}" style="max-height:{max_h}px"><img src="{shot["src"]}"></div></div>')


def bullets(items):
    return '<ul class="checks">' + "".join(f"<li>{html.escape(b)}</li>" for b in items) + "</ul>"


def chips():
    return '<div class="chips">' + "".join(f"<span>{c}</span>" for c in CHIPS) + "</div>"


def page(body):
    return (f'<!doctype html><html lang="fr"><head><meta charset="utf-8"><style>{CSS}</style></head>'
            f'<body>{body}</body></html>')


def cover(p, shots, product):
    wins = "".join(window(shots[f"{product}/{s}"], x, y, w, h, z=i + 1)
                   for i, (s, x, y, w, h) in enumerate(p["windows"]))
    left = (f'<div class="left"><div><span class="pill">{p["pill"]}</span></div>'
            f'<h1 style="font-size:52px;margin:22px 0 16px">{html.escape(p["title"])}</h1>'
            f'<p class="sub" style="font-size:21px;margin-bottom:18px">{html.escape(p["sub"])}</p>'
            f'{bullets(p["bullets"])}<div style="margin-top:24px">{chips()}</div></div>')
    return page(left + wins)


def split(p, shots, product):
    shot = shots[f"{product}/{p['shot']}"]
    width = 600 if shot["ratio"] > 1 else 640
    top = max(60, int((H - min(width * shot["ratio"] + 26, 780)) / 2))
    left = (f'<div class="left"><h1 style="font-size:50px;margin-bottom:18px">{html.escape(p["title"])}</h1>'
            f'<p class="sub" style="font-size:21px;margin-bottom:20px">{html.escape(p["sub"])}</p>'
            f'{bullets(p["bullets"])}</div>')
    return page(left + window(shot, 1200 - 70 - width, top, width, 780))


def wide(p, shots, product):
    shot = shots[f"{product}/{p['shot']}"]
    width = 1060
    head = (f'<div style="position:absolute;left:70px;right:70px;top:46px;text-align:center">'
            f'<h1 style="font-size:46px;margin-bottom:12px">{html.escape(p["title"])}</h1>'
            f'<p class="sub" style="font-size:21px">{html.escape(p["sub"])}</p></div>')
    calls = "".join(f'<div class="callout" style="left:{x}px;top:{y}px">{html.escape(t)}</div>'
                    for t, x, y in p.get("callouts", []))
    return page(head + window(shot, 70, 190, width, 660) + calls)


def howto(p, *_):
    steps = [("1", "Achetez", "Paiement sécurisé sur Etsy."),
             ("2", "Téléchargez", "Le fichier vous attend tout de suite dans votre compte Etsy."),
             ("3", "Ouvrez", "Dans Excel, LibreOffice ou Google Sheets, sur ordinateur ou téléphone.")]
    cards = "".join(
        f'<div style="flex:1;background:#fff;border-radius:18px;padding:34px 30px;'
        f'box-shadow:0 24px 50px -26px rgba(31,78,95,.45)">'
        f'<div style="width:62px;height:62px;border-radius:50%;background:var(--dark);color:#fff;font-size:30px;'
        f'font-weight:800;display:flex;align-items:center;justify-content:center;margin-bottom:22px">{n}</div>'
        f'<div style="font-size:30px;font-weight:800;color:var(--dark);margin-bottom:10px">{t}</div>'
        f'<div style="font-size:20px;color:var(--muted);line-height:1.4">{d}</div></div>' for n, t, d in steps)
    extras = ["Mode d'emploi en première page", "Version exemple offerte", "100 % en français",
              "Rien n'est envoyé par la poste"]
    body = (f'<div style="position:absolute;left:70px;right:70px;top:70px;text-align:center">'
            f'<span class="pill">{html.escape(p["product"])}</span>'
            f'<h1 style="font-size:58px;margin:24px 0 0">Comment ça marche ?</h1></div>'
            f'<div style="position:absolute;left:70px;right:70px;top:290px;display:flex;gap:28px">{cards}</div>'
            f'<div style="position:absolute;left:70px;right:70px;top:660px;background:var(--greenl);border-radius:18px;'
            f'padding:26px 30px;display:grid;grid-template-columns:1fr 1fr;gap:14px 30px">'
            + "".join(f'<div style="font-size:21px;font-weight:600;display:flex;gap:12px;align-items:center">'
                      f'<span style="color:var(--green);font-weight:800;font-size:24px">✓</span>{e}</div>'
                      for e in extras)
            + "</div>")
    return page(body)


def pack(p, shots, product):
    tiles = [("Budget 2027", "budget/annee"), ("Menus & liste de courses", "menus/semaine"),
             ("Épargne & défi 52 semaines", "epargne/objectifs"), ("Comptes à deux", "comptes/bilan")]
    body = (f'<div style="position:absolute;left:70px;right:70px;top:44px;text-align:center">'
            f'<span class="pill">4 modèles Excel & Google Sheets</span>'
            f'<h1 style="font-size:54px;margin:18px 0 0">Le pack organisation</h1></div>')
    positions = [(70, 230), (620, 230), (70, 555), (620, 555)]
    for (label, key), (x, y) in zip(tiles, positions):
        body += window(shots[key], x, y + 36, 510, 250)
        body += (f'<div style="position:absolute;left:{x}px;top:{y}px;font-size:22px;font-weight:800;'
                 f'color:var(--dark)">{label}</div>')
    return page(body)


LAYOUTS = {"cover": cover, "split": split, "wide": wide, "howto": howto, "pack": pack}


def main():
    BUILD.mkdir(exist_ok=True)
    shots = render_captures()
    jobs = []
    for product, photos in PHOTOS.items():
        (HERE / product).mkdir(exist_ok=True)
        for p in photos:
            page_html = LAYOUTS[p["kind"]](p, shots, product)
            src = BUILD / f"{product}-{p['name']}.html"
            src.write_text(page_html, encoding="utf-8")
            jobs.append({"html": src.as_uri(), "out": str(HERE / product / f"{p['name']}.jpg")})
    jobs_file = BUILD / "jobs.json"
    jobs_file.write_text(json.dumps(jobs), encoding="utf-8")
    subprocess.run(["node", str(HERE / "shoot.js"), str(jobs_file), str(W), str(H), str(SCALE)], check=True)


if __name__ == "__main__":
    main()

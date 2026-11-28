#!/usr/bin/env python3
"""Regroupe les 4 modèles (versions vierges + exemples) dans un ZIP pour l'annonce « pack ».
Etsy limite une annonce à 5 fichiers : le pack en a 8, d'où le ZIP.

Usage : python3 pack/build_pack.py
"""
import zipfile
from pathlib import Path

SHOP = Path(__file__).resolve().parent.parent
FILES = ["budget/Budget_2027.xlsx", "budget/Budget_2027_exemple.xlsx",
         "menus/Menus_et_courses.xlsx", "menus/Menus_et_courses_exemple.xlsx",
         "epargne/Epargne_defi_52_semaines.xlsx", "epargne/Epargne_defi_52_semaines_exemple.xlsx",
         "comptes/Comptes_a_deux.xlsx", "comptes/Comptes_a_deux_exemple.xlsx"]

out = SHOP / "pack" / "Pack_4_modeles.zip"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for f in FILES:
        z.write(SHOP / f, arcname=Path(f).name)
print("écrit :", out)

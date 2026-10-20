"""Compose une illustration pleine carte : décor peint flouté + portrait Wesnoth + halo et ombre."""
import json, os, sys
import numpy as np
from PIL import Image, ImageFilter, ImageEnhance
SP = '/tmp/claude-0/-home-user-coco/20628a72-d0d0-5df4-a703-2899f6536c1d/scratchpad'
W, H = 500, 700
cards = {c['id']: c for c in json.load(open(SP + '/artgen/cards.json'))}
m = json.load(open(SP + '/wesnoth/mapping.json'))
out = sys.argv[1]
ids = sys.argv[2:] or list(m)
def hx(c): c = c.lstrip('#'); return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))
for cid in ids:
    c = cards[cid]
    bg = Image.open(f'{SP}/wesnoth/bgall/{cid}.webp').convert('RGB').filter(ImageFilter.GaussianBlur(3.2))
    bg = ImageEnhance.Brightness(bg).enhance(0.86)
    f = m[cid]
    p = Image.open(f'{SP}/wesnoth/raw/{f}.webp' if os.path.exists(f'{SP}/wesnoth/raw/{f}.webp') else f'{SP}/wesnoth/raw/{f}.png').convert('RGBA')
    p = p.crop(p.getchannel('A').getbbox())
    hero = cid.startswith('hero-')
    al = np.asarray(p.getchannel('A'))
    bust = al[-3:].mean() > 40            # portrait coupé en bas : on le pose sur le bord
    s = max(480 / p.width, 560 / p.height)
    s = min(s, 640 / p.width, 660 / p.height)
    p = p.resize((round(p.width * s), round(p.height * s)), Image.LANCZOS)
    x = (W - p.width) // 2
    y = H - p.height + 8 if bust else int(H * 0.47 - p.height / 2)
    canvas = bg.convert('RGBA')
    a = Image.new('L', (W, H), 0); a.paste(p.getchannel('A'), (x, y))
    glow_col = hx(c['sceneColors'][2])
    halo = a.filter(ImageFilter.MaxFilter(9)).filter(ImageFilter.GaussianBlur(18)).point(lambda v: int(v * 0.8))
    canvas = Image.composite(Image.new('RGBA', (W, H), glow_col + (255,)), canvas, halo)
    shadow = Image.new('L', (W, H), 0); shadow.paste(p.getchannel('A'), (x + 14, y + 10))
    shadow = shadow.filter(ImageFilter.GaussianBlur(12)).point(lambda v: int(v * 0.45))
    canvas = Image.composite(Image.new('RGBA', (W, H), (10, 6, 20, 255)), canvas, shadow)
    # léger reflet de la couleur du décor sur le bord du portrait
    rim = a.filter(ImageFilter.GaussianBlur(3))
    rim = Image.fromarray(np.clip((np.asarray(a, np.float32) - np.asarray(rim, np.float32)) * 2.2, 0, 255).astype(np.uint8))
    canvas.alpha_composite(p, (x, y))
    canvas = Image.composite(Image.new('RGBA', (W, H), glow_col + (255,)), canvas, rim.point(lambda v: int(v * 0.5)))
    # vignette
    yy, xx = np.mgrid[0:H, 0:W]
    d = np.hypot((xx - W / 2) / W, (yy - H * 0.45) / H)
    vig = np.clip((d - 0.38) * 1.6, 0, 0.55)
    arr = np.asarray(canvas.convert('RGB'), np.float32) * (1 - vig[..., None])
    Image.fromarray(arr.astype(np.uint8)).save(f'{out}/{cid}.webp', 'WEBP', quality=86, method=6)

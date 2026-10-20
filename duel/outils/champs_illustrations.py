"""Illustrations du mode Champs de bataille : portrait Wesnoth sur un décor peint (éventuellement retourné)."""
import json, os, sys
import numpy as np
from PIL import Image, ImageFilter, ImageEnhance, ImageOps
SP = '/tmp/claude-0/-home-user-coco/20628a72-d0d0-5df4-a703-2899f6536c1d/scratchpad'
W, H = 500, 700
cards = {c['id']: c for c in json.load(open(SP + '/artgen/cards.json'))}
spec = json.load(open(SP + '/wesnoth/bg_art.json'))
out = sys.argv[1]
ids = sys.argv[2:] or list(spec)
def hx(c): c = c.lstrip('#'); return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))
yy, xx = np.mgrid[0:H, 0:W]
vig = np.clip((np.hypot((xx - W / 2) / W, (yy - H * 0.45) / H) - 0.38) * 1.6, 0, 0.55)
for cid in ids:
    f, bgid, flip = spec[cid]
    bg = Image.open(f'{SP}/wesnoth/bgall/{bgid}.webp').convert('RGB')
    if flip: bg = ImageOps.mirror(bg)
    bg = ImageEnhance.Brightness(bg.filter(ImageFilter.GaussianBlur(3.2))).enhance(0.86)
    src = f'{SP}/wesnoth/raw/{f}.webp'
    p = Image.open(src if os.path.exists(src) else src[:-5] + '.png').convert('RGBA')
    p = p.crop(p.getchannel('A').getbbox())
    if flip: p = ImageOps.mirror(p)
    al = np.asarray(p.getchannel('A'))
    bust = al[-3:].mean() > 40
    s = min(max(480 / p.width, 560 / p.height), 640 / p.width, 660 / p.height)
    p = p.resize((round(p.width * s), round(p.height * s)), Image.LANCZOS)
    x = (W - p.width) // 2
    y = H - p.height + 8 if bust else int(H * 0.47 - p.height / 2)
    canvas = bg.convert('RGBA')
    a = Image.new('L', (W, H), 0); a.paste(p.getchannel('A'), (x, y))
    glow = hx(cards[bgid]['sceneColors'][2])
    halo = a.filter(ImageFilter.MaxFilter(9)).filter(ImageFilter.GaussianBlur(18)).point(lambda v: int(v * 0.8))
    canvas = Image.composite(Image.new('RGBA', (W, H), glow + (255,)), canvas, halo)
    sh = Image.new('L', (W, H), 0); sh.paste(p.getchannel('A'), (x + 14, y + 10))
    canvas = Image.composite(Image.new('RGBA', (W, H), (10, 6, 20, 255)), canvas, sh.filter(ImageFilter.GaussianBlur(12)).point(lambda v: int(v * 0.45)))
    rim = a.filter(ImageFilter.GaussianBlur(3))
    rim = Image.fromarray(np.clip((np.asarray(a, np.float32) - np.asarray(rim, np.float32)) * 2.2, 0, 255).astype(np.uint8))
    canvas.alpha_composite(p, (x, y))
    canvas = Image.composite(Image.new('RGBA', (W, H), glow + (255,)), canvas, rim.point(lambda v: int(v * 0.5)))
    arr = np.asarray(canvas.convert('RGB'), np.float32) * (1 - vig[..., None])
    Image.fromarray(arr.astype(np.uint8)).save(f'{out}/{cid}.webp', 'WEBP', quality=84, method=6)
print('ok', len(ids))

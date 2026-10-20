#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Duel des Royaumes -- full-art card illustration compositor.
Art direction B: "Pokemon Special Art Rare" -- luminous pastel fantasy.

usage:  python3 gen.py <cards.json> <outdir> [id ...]

Produces one 500x700 WebP per card: a painted, full-bleed procedural scene
(12 environments) with the card's 3D subject integrated into it (glow, rim
light, ambient grading, ground contact, haze, foreground overlap).
No frame / text / gems: the game UI draws those on top.
Dependencies: numpy + Pillow only.
"""
import sys, os, json, math, hashlib, re, time
import numpy as np
from PIL import Image, ImageDraw

W, H = 500, 700
SS = 2                                   # supersampling for vector shapes
YY, XX = np.mgrid[0:H, 0:W].astype(np.float32)
RARITY = {'Commune': 0, 'Rare': 1, 'Épique': 2, 'Epique': 2, 'Légendaire': 3, 'Legendaire': 3}

# subjects whose PNG is a bust / object cut flat at the bottom edge
CUT_SUBJECTS = {'guard', 'princess', 'troll', 'unicorn', 'woman_elf', 'woman_mage',
                'woman_vampire', 'water_wave', 'moai', 'castle', 'mountain', 'lion',
                'dragon_face', 'trophy', 'scroll'}
# subjects that naturally hover (no hard contact shadow)
FLOAT_SUBJECTS = {'baby_angel', 'fairy', 'eagle', 'dolphin', 'fire', 'tornado', 'skull_and_crossbones',
                  'comet', 'snowflake', 'glowing_star', 'high_voltage', 'collision', 'sparkles',
                  'cyclone', 'shooting_star', 'direct_hit', 'paw_prints', 'footprints', 'folded_hands',
                  'fleur-de-lis', 'herb', 'postal_horn', 'dagger', 'axe', 'shield', 'coin', 'octopus',
                  'crocodile', 'bow_and_arrow', 'military_helmet', 'candle', 'skull', 'crystal_ball'}


# =============================================================== colour / math utils
def C(h):
    if not isinstance(h, str):
        return np.asarray(h, np.float32)
    h = h.lstrip('#')
    return np.array([int(h[i:i + 2], 16) / 255.0 for i in (0, 2, 4)], np.float32)


def lerp(a, b, t):
    return a + (b - a) * t


def sstep(e0, e1, x):
    t = np.clip((x - e0) / (e1 - e0), 0.0, 1.0)
    return t * t * (3 - 2 * t)


def rgba8(col, a=1.0):
    col = np.clip(np.asarray(col, np.float32), 0, 1)
    return (int(col[0] * 255 + .5), int(col[1] * 255 + .5), int(col[2] * 255 + .5),
            int(float(np.clip(a, 0, 1)) * 255 + .5))


def _box(a, r, axis):
    n = a.shape[axis]
    pad = [(0, 0)] * a.ndim
    pad[axis] = (r + 1, r)
    c = np.cumsum(np.pad(a, pad, mode='edge'), axis=axis, dtype=np.float32)
    hi = [slice(None)] * a.ndim
    lo = [slice(None)] * a.ndim
    hi[axis] = slice(2 * r + 1, 2 * r + 1 + n)
    lo[axis] = slice(0, n)
    return (c[tuple(hi)] - c[tuple(lo)]) / (2 * r + 1)


def _resize_f(a, w, h, method):
    if a.ndim == 2:
        return np.asarray(Image.fromarray(np.ascontiguousarray(a, dtype=np.float32), 'F').resize((w, h), method))
    return np.stack([_resize_f(a[..., k], w, h, method) for k in range(a.shape[2])], -1)


def blur(a, s):
    """Gaussian-ish blur of a float array (2D or HxWxC)."""
    if s < 0.35:
        return a
    a = a.astype(np.float32, copy=False)
    if s > 6:
        f = max(2, int(s / 3))
        h, w = a.shape[:2]
        small = _resize_f(a, max(4, w // f), max(4, h // f), Image.BOX)
        small = blur(small, s / f)
        return _resize_f(small, w, h, Image.BICUBIC)
    r = max(1, int(round((math.sqrt(4 * s * s + 1) - 1) / 2)))
    for _ in range(3):
        a = _box(_box(a, r, 0), r, 1)
    return a


def shift(a, dx, dy):
    """out[y, x] = a[y - dy, x - dx] (integer shift, zero fill)."""
    dx, dy = int(round(dx)), int(round(dy))
    out = np.zeros_like(a)
    h, w = a.shape[:2]
    ys0, ys1 = max(0, dy), min(h, h + dy)
    xs0, xs1 = max(0, dx), min(w, w + dx)
    if ys1 <= ys0 or xs1 <= xs0:
        return out
    out[ys0:ys1, xs0:xs1] = a[ys0 - dy:ys1 - dy, xs0 - dx:xs1 - dx]
    return out


# =============================================================== noise
def vnoise(rng, w, h, cx, cy):
    cx = max(1, int(round(cx)))
    cy = max(1, int(round(cy)))
    g = rng.random((cy + 1, cx + 1)).astype(np.float32)
    return _resize_f(g, w, h, Image.BICUBIC)


def fbm(rng, base=3, octaves=5, pers=0.5, sx=1.0, sy=1.0, w=W, h=H):
    acc = np.zeros((h, w), np.float32)
    amp, tot = 1.0, 0.0
    for i in range(octaves):
        f = base * (2 ** i)
        acc += amp * vnoise(rng, w, h, f * sx, f * h / w * sy)
        tot += amp
        amp *= pers
    lo, hi = np.percentile(acc[::3, ::3], [1, 99])
    return np.clip((acc - lo) / (hi - lo + 1e-6), 0, 1)


def fbm1(rng, n, base=4, octaves=5, pers=0.5):
    s = np.zeros(n, np.float32)
    tot, amp = 0.0, 1.0
    for i in range(octaves):
        f = int(base * 2 ** i) + 1
        g = rng.random(f + 2)
        xs = np.linspace(0, f, n)
        i0 = np.floor(xs).astype(int).clip(0, f)
        fr = xs - i0
        fr = fr * fr * (3 - 2 * fr)
        s += amp * (g[i0] * (1 - fr) + g[i0 + 1] * fr)
        tot += amp
        amp *= pers
    return s / tot


# =============================================================== compositing
def over(img, col, a):
    a = np.clip(a, 0, 1)[..., None]
    img *= (1 - a)
    img += np.asarray(col, np.float32) * a


def add(img, col, a):
    img += np.asarray(col, np.float32) * np.clip(a, 0, None)[..., None]


def screen(img, col, a):
    v = np.clip(np.asarray(col, np.float32) * np.clip(a, 0, None)[..., None], 0, 1)
    img += v * (1 - np.clip(img, 0, 1))


def multiply(img, col, a):
    a = np.clip(a, 0, 1)[..., None]
    img *= (1 - a) + np.asarray(col, np.float32) * a


class Layer:
    """RGBA vector layer drawn at SSx; returns premultiplied float colour + alpha."""

    def __init__(self):
        self.im = Image.new('RGBA', (W * SS, H * SS), (0, 0, 0, 0))
        self.d = ImageDraw.Draw(self.im)

    @staticmethod
    def _p(pts):
        return [(float(x) * SS, float(y) * SS) for x, y in pts]

    def poly(self, pts, col, a=1.0):
        self.d.polygon(self._p(pts), fill=rgba8(col, a))

    def ellipse(self, cx, cy, rx, ry, col, a=1.0, width=0):
        box = [(cx - rx) * SS, (cy - ry) * SS, (cx + rx) * SS, (cy + ry) * SS]
        if box[2] - box[0] < 1 or box[3] - box[1] < 1:
            return
        if width:
            self.d.ellipse(box, outline=rgba8(col, a), width=max(1, int(round(width * SS))))
        else:
            self.d.ellipse(box, fill=rgba8(col, a))

    def line(self, pts, col, width, a=1.0):
        self.d.line(self._p(pts), fill=rgba8(col, a), width=max(1, int(round(width * SS))), joint='curve')

    def get(self):
        arr = np.asarray(self.im.convert('RGBa').reduce(SS), np.float32) / 255.0
        return arr[..., :3], arr[..., 3]


class Mask:
    def __init__(self):
        self.im = Image.new('L', (W * SS, H * SS), 0)
        self.d = ImageDraw.Draw(self.im)

    def poly(self, pts, v=1.0):
        self.d.polygon(Layer._p(pts), fill=int(v * 255))

    def ellipse(self, cx, cy, rx, ry, v=1.0, width=0):
        box = [(cx - rx) * SS, (cy - ry) * SS, (cx + rx) * SS, (cy + ry) * SS]
        if box[2] - box[0] < 1 or box[3] - box[1] < 1:
            return
        if width:
            self.d.ellipse(box, outline=int(v * 255), width=max(1, int(round(width * SS))))
        else:
            self.d.ellipse(box, fill=int(v * 255))

    def line(self, pts, width, v=1.0):
        self.d.line(Layer._p(pts), fill=int(v * 255), width=max(1, int(round(width * SS))), joint='curve')

    def get(self, s=0):
        a = np.asarray(self.im.reduce(SS), np.float32) / 255.0
        return blur(a, s) if s else a


def comp_layer(img, L, opacity=1.0, s=0):
    pm, a = L.get()
    if s:
        pm, a = blur(pm, s), blur(a, s)
    img *= (1 - a * opacity)[..., None]
    img += pm * opacity
    return a


# =============================================================== shape helpers
def ellipse_pts(cx, cy, rx, ry, ang=0.0, n=28):
    t = np.linspace(0, 2 * np.pi, n, endpoint=False)
    x, y = rx * np.cos(t), ry * np.sin(t)
    ca, sa = math.cos(ang), math.sin(ang)
    return list(zip(cx + x * ca - y * sa, cy + x * sa + y * ca))


def petal_pts(x, y, L, Wd, ang, n=14, tip=0.6):
    """Teardrop leaf/petal from base (x,y) pointing to angle ang (radians, screen space)."""
    t = np.linspace(0, np.pi, n)
    u = L * (1 - np.cos(t)) / 2
    hw = Wd * np.sin(t) ** tip * (0.55 + 0.45 * (1 - np.cos(t)) / 2)
    ca, sa = math.cos(ang), math.sin(ang)
    left = [(x + uu * ca - hh * sa, y + uu * sa + hh * ca) for uu, hh in zip(u, hw)]
    right = [(x + uu * ca + hh * sa, y + uu * sa - hh * ca) for uu, hh in zip(u[::-1], hw[::-1])]
    return left + right


def star_pts(x, y, r, thin=0.14, rot=0.0, n=4):
    pts = []
    for k in range(2 * n):
        a = rot + k * math.pi / n
        rr = r if k % 2 == 0 else r * thin
        pts.append((x + rr * math.cos(a), y + rr * math.sin(a)))
    return pts


# =============================================================== context
class Ctx:
    def __init__(self, card):
        self.card = card
        seed = int(hashlib.md5(card['id'].encode('utf-8')).hexdigest()[:8], 16)
        self.rng = np.random.default_rng(seed)
        self.img = np.zeros((H, W, 3), np.float32)
        self.tier = RARITY.get(card.get('rarity', 'Commune'), 0)
        self.type = card.get('type', 'serviteur')
        self.el = [C(x) for x in card.get('elementColors', ['#ffffff', '#ffe0b0', '#ff9a5a', '#ff4d3a'])]
        self.scene = card.get('scene', 'plaine')
        self.P = PAL[self.scene]
        self.light = (W * 0.5, H * 0.1)
        self.S = None

    def pts(self, n, x0, y0, x1, y1, avoid=0.75, scale=0.42):
        out, tries = [], 0
        while len(out) < n and tries < n * 12:
            tries += 1
            x, y = self.rng.uniform(x0, x1), self.rng.uniform(y0, y1)
            if avoid and self.S is not None:
                dx = (x - self.S.cx) / (self.S.sw * scale)
                dy = (y - self.S.cy) / (self.S.sh * scale)
                if dx * dx + dy * dy < 1 and self.rng.random() < avoid:
                    continue
            out.append((x, y))
        return out


# =============================================================== generic painted elements
def sky(c, stops, wobble=0.05):
    ys = np.array([s[0] for s in stops], np.float32)
    cols = np.array([C(s[1]) for s in stops], np.float32)
    n = fbm(c.rng, 2, 3)
    t = np.clip(YY / H + (n - 0.5) * wobble, 0, 1)
    c.img[:] = np.stack([np.interp(t, ys, cols[:, k]) for k in range(3)], -1)


def glow(c, x, y, r, col, s=1.0, mode='screen', ry=None, power=1.0):
    ry = r if ry is None else ry
    d2 = ((XX - x) / r) ** 2 + ((YY - y) / ry) ** 2
    g = np.exp(-d2 * 2.0) ** power * s
    (screen if mode == 'screen' else add)(c.img, col, g)


def stars(c, n, y0, y1, col='#ffffff', big=0.06):
    m = Mask()
    rng = c.rng
    for _ in range(n):
        x, y = rng.uniform(0, W), rng.uniform(y0, y1) ** 1.0
        r = rng.uniform(0.35, 1.1) ** 2 * 1.4 + 0.3
        m.ellipse(x, y, r, r, rng.uniform(0.35, 1.0))
    a = m.get()
    screen(c.img, C(col), a)
    screen(c.img, C(col), blur(a, 2.0) * 0.8)
    pts = [(rng.uniform(0, W), rng.uniform(y0, y1), rng.uniform(4, 9)) for _ in range(int(n * big))]
    sparkles(c, pts, C(col), C(col), 0.8)


def sparkles(c, pts, col, gcol, s=1.0):
    if not pts:
        return
    m = Mask()
    for (x, y, r) in pts:
        m.poly(star_pts(x, y, r, 0.11, rot=-math.pi / 2), 1)
        m.poly(star_pts(x, y, r * 0.42, 0.2, rot=-math.pi / 4), 1)
        m.ellipse(x, y, r * 0.13, r * 0.13, 1)
    a = m.get()
    add(c.img, gcol, (blur(a, 2.2) * 1.3 + blur(a, 7) * 0.9) * s)
    screen(c.img, col, a * s)


def bokeh(c, n, region, rmin, rmax, cols, amin=0.12, amax=0.4, soft=1.0, avoid=0.6, ring=True):
    rng = c.rng
    cols = [C(x) if isinstance(x, str) else x for x in cols]
    masks = [Mask() for _ in cols]
    for (x, y) in c.pts(n, *region, avoid=avoid):
        k = rng.integers(len(cols))
        r = rng.uniform(rmin, rmax)
        a = rng.uniform(amin, amax)
        masks[k].ellipse(x, y, r, r, a * 0.65)
        if ring:
            masks[k].ellipse(x, y, r, r, a, width=max(1.0, r * 0.12))
    for m, col in zip(masks, cols):
        screen(c.img, col, m.get(soft))


def god_rays(c, sx, sy, col, s=0.5, cells=40, power=2.4, reach=800, start=0.0, src_glow=True):
    ang = np.arctan2(YY - sy, XX - sx)
    n = fbm1(c.rng, 4096, cells, 4)
    n = (n - n.min()) / (n.max() - n.min() + 1e-6)
    idx = ((ang + np.pi) / (2 * np.pi) * 4095).astype(np.int32).clip(0, 4095)
    r = n[idx] ** power
    d = np.sqrt((XX - sx) ** 2 + (YY - sy) ** 2)
    fall = np.exp(-d / reach) * sstep(start, start + 80, d)
    screen(c.img, C(col) if isinstance(col, str) else col, r * fall * s)


def fill_ridge(c, yl, top, bot, depth=220, a=1.0, rim=None, rimw=4.0, rims=0.8, s=0.0, tex=None):
    m = np.clip(YY - yl[None, :] + 0.5, 0, 1)
    t = np.clip((YY - yl[None, :]) / depth, 0, 1)
    col = lerp(C(top), C(bot), t[..., None])
    if tex is not None:
        col = col * (1 + (tex[..., None] - 0.5) * 0.25)
    if s:
        m = blur(m, s)
    over(c.img, col, m * a)
    if rim is not None:
        r = m * np.exp(-np.clip(YY - yl[None, :], 0, None) / rimw)
        screen(c.img, C(rim), r * rims * a)
    return m


def mountain_ridge(c, n_peaks, base, hmin, hmax, slope=(0.7, 1.3), detail=14.0, xr=(-60, W + 60)):
    rng = c.rng
    xs = np.arange(W, dtype=np.float32)
    y = np.full(W, base, np.float32)
    for _ in range(n_peaks):
        px = rng.uniform(*xr)
        ph = rng.uniform(hmin, hmax)
        sl = rng.uniform(*slope)
        prof = base - ph + np.where(xs < px, (px - xs) * sl * rng.uniform(0.75, 1.25),
                                    (xs - px) * sl * rng.uniform(0.75, 1.25))
        y = np.minimum(y, prof)
    det = (fbm1(rng, W, 5, 5) - 0.5) * detail + (fbm1(rng, W, 28, 3) - 0.5) * detail * 0.35
    return y + det


def mountain_layer(c, yl, top, bot, lit, shade, snow=None, snow_frac=0.35, lightx=1.0, a=1.0,
                   rim=None, haze=None, haze_a=0.0, depth=260, s=0.0):
    rng = c.rng
    m = np.clip(YY - yl[None, :] + 0.5, 0, 1)
    if s:
        m = blur(m, s)
    t = np.clip((YY - yl[None, :]) / depth, 0, 1)[..., None]
    col = lerp(C(top), C(bot), t)
    dy = np.gradient(blur(yl[None, :].repeat(3, 0), 2.0)[1])
    face = sstep(-0.35, 0.35, dy * lightx)[None, :]           # 1 = lit face
    streak = vnoise(rng, W, H, 60, 4)
    face = np.clip(face + (streak - 0.5) * 0.35, 0, 1)
    fade = np.exp(-np.clip(YY - yl[None, :], 0, None) / (depth * 0.35))
    shading = lerp(C(shade), C(lit), face[..., None])
    col = lerp(col, col * 0.4 + shading * 0.75, (fade * 0.85)[..., None])
    if snow is not None:
        top_y = yl.min()
        hgt = np.clip(yl.max() - yl, 0, None)[None, :]
        sd = hgt * snow_frac * (0.45 + 0.9 * vnoise(rng, W, H, 50, 5)) + 3
        sm = sstep(0.0, -3.0, (YY - yl[None, :]) - sd) * m
        scol = lerp(C(snow[1]), C(snow[0]), face[..., None])
        col = lerp(col, scol, sm[..., None])
    if haze is not None:
        col = lerp(col, C(haze), haze_a)
    over(c.img, col, m * a)
    if rim is not None:
        r = m * np.exp(-np.clip(YY - yl[None, :], 0, None) / 3.0)
        screen(c.img, C(rim), r * 0.9 * a)
    return m


def fog_band(c, y, thick, col, a=0.6, base=4, sx=1.0):
    n = fbm(c.rng, base, 4, sx=sx, sy=0.5)
    band = np.exp(-((YY - y) / thick) ** 2)
    over(c.img, C(col), band * (0.35 + 0.65 * n) * a)


def puffs(c, blobs, lit, mid, shade, light=(-0.45, -0.9), soft=5.0, a=1.0, s=0.0, rim=None, rims=0.7, flat=None,
          style='cloud', rough=None, zdepth=0.35):
    """Volumetric puffs: height field = max of (noisy) spheres, lit with real normals.
    style: 'cloud' (soft wrap light, silver lining), 'rock' (toon facets), 'bush' (leafy noise)."""
    rng = c.rng
    lit, mid, shade = C(lit), C(mid), C(shade)
    if rough is None:
        rough = {'cloud': 0.16, 'rock': 0.10, 'bush': 0.32}[style]
    nscale = {'cloud': 9, 'rock': 6, 'bush': 22}[style]
    n = fbm(rng, nscale, 3)
    hmap = np.zeros((H, W), np.float32)
    amap = np.zeros((H, W), np.float32)
    if not blobs:
        return amap
    ymin = min(b[1] for b in blobs)
    for (x, y, r) in blobs:
        R = r * (1 + rough * 0.6)
        x0, x1 = int(max(0, x - R - 2)), int(min(W, x + R + 3))
        y0, y1 = int(max(0, y - R - 2)), int(min(H, y + R + 3))
        if x1 <= x0 or y1 <= y0:
            continue
        dx = XX[y0:y1, x0:x1] - x
        dy = YY[y0:y1, x0:x1] - y
        re = r * (1 + rough * (n[y0:y1, x0:x1] - 0.5) * 2)
        h2 = re * re - dx * dx - dy * dy
        hs = np.sqrt(np.clip(h2, 0, None))
        z = hs + (y - ymin) * zdepth
        inside = h2 > 0
        hmap[y0:y1, x0:x1] = np.where(inside, np.maximum(hmap[y0:y1, x0:x1], z), hmap[y0:y1, x0:x1])
        amap[y0:y1, x0:x1] = np.maximum(amap[y0:y1, x0:x1], np.clip(hs / 1.6, 0, 1))
    hb = blur(hmap, 1.2 if style != 'cloud' else max(1.5, soft * 0.35))
    gx = np.gradient(hb, axis=1)
    gy = np.gradient(hb, axis=0)
    inv = 1.0 / np.sqrt(gx * gx + gy * gy + 1.0)
    nx, ny, nz = -gx * inv, -gy * inv, inv
    lx, ly = light
    lz = 0.75
    ln = math.sqrt(lx * lx + ly * ly + lz * lz)
    lx, ly, lz = lx / ln, ly / ln, lz / ln
    d = nx * lx + ny * ly + nz * lz
    if style == 'rock':
        t = np.clip(d, 0, 1)
        t = np.round(t * 4) / 4
        t = blur(t, 0.6)
    else:
        t = np.clip((d + 0.35) / 1.35, 0, 1)
    ao = np.clip((blur(hmap, 7) - hmap) / 14.0, 0, 1) * (amap > 0)
    t = np.clip(t - ao * 0.55, 0, 1)
    t3 = t[..., None]
    col = np.where(t3 < 0.5, lerp(shade, mid, t3 * 2), lerp(mid, lit, (t3 - 0.5) * 2))
    if style == 'bush':
        col = col * (0.85 + 0.3 * fbm(rng, 40, 2))[..., None]
    if flat is not None:
        amap = amap * sstep(flat + 3, flat - 3, YY)
        col = lerp(col, lerp(col, shade, 0.5), sstep(flat - 26, flat, YY)[..., None])
    al = blur(amap, s) if s else amap
    over(c.img, col, al * a)
    if rim is not None:
        edge = np.clip(1 - nz, 0, 1) ** 1.5 * np.clip(nx * lx + ny * ly, 0, 1) * amap
        screen(c.img, C(rim), blur(edge, 0.8) * rims * a * 1.4)
    return al


def cumulus(c, cx, cy, w, h, n=None):
    """Blob list for a cumulus cloud sitting on y=cy (flat-ish base), dome of height h."""
    rng = c.rng
    n = n or int(10 + w / 16)
    blobs = []
    for i in range(n):
        u = rng.uniform(-1, 1)
        dome = (1 - u * u) ** 0.7
        r = h * (0.25 + 0.55 * dome) * rng.uniform(0.6, 1.2)
        y = cy - r * 0.35 - dome * h * rng.uniform(0.2, 0.75)
        blobs.append((cx + u * w * 0.5, y, r))
    blobs.sort(key=lambda b: b[1])
    return blobs


def cloud_bank(c, y, x0, x1, rmin, rmax, n, jitter=14):
    rng = c.rng
    blobs = []
    for i in range(n):
        x = lerp(x0, x1, (i + rng.uniform(-0.3, 0.3)) / max(1, n - 1))
        r = rng.uniform(rmin, rmax)
        blobs.append((x, y + rng.uniform(-jitter, jitter) - r * 0.25, r))
    blobs.sort(key=lambda b: b[2])
    return blobs


def crystal(L, x, y, h, w, ang, dark, mid, light, tipc, a=0.92, edge=True):
    """Hexagonal-prism crystal growing from (x,y) towards angle ang (0 = up)."""
    dx, dy = math.sin(ang), -math.cos(ang)
    px, py = math.cos(ang), math.sin(ang)

    def P(u, v):
        return (x + px * u * w + dx * v, y + py * u * w + dy * v)

    s = [h - w * 0.55, h - w * 0.95, h - w * 0.85, h - w * 0.5]
    T = P(0.04, h)
    us = [-0.5, -0.18, 0.24, 0.5]
    L.poly([P(us[0], 0), P(us[1], 0), P(us[1], s[1]), P(us[0], s[0])], dark, a)
    L.poly([P(us[1], 0), P(us[2], 0), P(us[2], s[2]), P(us[1], s[1])], mid, a)
    L.poly([P(us[2], 0), P(us[3], 0), P(us[3], s[3]), P(us[2], s[2])], light, a)
    L.poly([P(us[0], s[0]), P(us[1], s[1]), T], lerp(dark, tipc, 0.45), a)
    L.poly([P(us[1], s[1]), P(us[2], s[2]), T], lerp(mid, tipc, 0.6), a)
    L.poly([P(us[2], s[2]), P(us[3], s[3]), T], tipc, a)
    if edge:
        wl = max(0.6, w * 0.035)
        L.line([P(us[1], h * 0.05), P(us[1], s[1]), T], lerp(light, np.ones(3), 0.6), wl, 0.7)
        L.line([P(us[2], h * 0.15), P(us[2], s[2])], np.ones(3), wl * 0.8, 0.45)


def shard(L, x, y, ln, ang, dark, light, tipc, a=0.95):
    """Floating bipyramid crystal shard."""
    dx, dy = math.sin(ang), -math.cos(ang)
    px, py = math.cos(ang), math.sin(ang)
    w = ln * 0.24
    top = (x + dx * ln * 0.5, y + dy * ln * 0.5)
    bot = (x - dx * ln * 0.5, y - dy * ln * 0.5)
    lft = (x - px * w + dx * ln * 0.08, y - py * w + dy * ln * 0.08)
    rgt = (x + px * w + dx * ln * 0.08, y + py * w + dy * ln * 0.08)
    mid = (x + px * w * 0.15 + dx * ln * 0.1, y + py * w * 0.15 + dy * ln * 0.1)
    L.poly([top, lft, mid], lerp(dark, tipc, 0.35), a)
    L.poly([top, mid, rgt], tipc, a)
    L.poly([bot, lft, mid], dark, a)
    L.poly([bot, mid, rgt], light, a)
    L.line([top, mid, bot], np.ones(3), max(0.6, ln * 0.03), 0.6)


def crystal_cluster(c, L, x, y, scale, base_ang, cols, n=5, spread=0.55):
    rng = c.rng
    dark, mid, light, tip = [C(k) for k in cols]
    items = []
    for i in range(n):
        ang = base_ang + rng.uniform(-spread, spread) * (0.3 if i == 0 else 1.0)
        h = scale * (1.0 if i == 0 else rng.uniform(0.35, 0.8))
        w = h * rng.uniform(0.2, 0.3)
        ox = rng.uniform(-scale * 0.25, scale * 0.25) if i else 0
        items.append((h, x + ox, y + rng.uniform(0, scale * 0.08), w, ang))
    items.sort(key=lambda t: -t[0])
    for h, xx, yy, w, ang in items:
        f = rng.uniform(0.85, 1.1)
        crystal(L, xx, yy, h, w, ang, dark * f, mid * f, np.minimum(light * f, 1), tip)
    return items


def magic_circle(c, cx, cy, R, col, core='#ffffff', tilt=1.0, rot=0.0, s=1.0, glyphs=True):
    rng = c.rng
    S = int(2 * R * SS) + 8
    im = Image.new('L', (S, S), 0)
    d = ImageDraw.Draw(im)
    o = S / 2

    def circ(r, w, v=255):
        d.ellipse([o - r * SS, o - r * SS, o + r * SS, o + r * SS], outline=v, width=max(1, int(w * SS)))

    def seg(r1, a1, r2, a2, w, v=255):
        d.line([(o + r1 * SS * math.cos(a1), o + r1 * SS * math.sin(a1)),
                (o + r2 * SS * math.cos(a2), o + r2 * SS * math.sin(a2))], fill=v, width=max(1, int(w * SS)))

    circ(R * 0.98, 2.4)
    circ(R * 0.92, 1.0)
    circ(R * 0.66, 1.8)
    circ(R * 0.60, 1.0, 200)
    circ(R * 0.24, 1.4)
    for k in range(96):
        a = 2 * math.pi * k / 96
        seg(R * 0.92, a, R * (0.95 if k % 3 else 0.98), a, 1.0, 220)
    if glyphs:
        for k in range(18):
            a0 = 2 * math.pi * (k + 0.5) / 18
            for j in range(rng.integers(2, 4)):
                r1 = R * rng.uniform(0.70, 0.86)
                r2 = R * rng.uniform(0.70, 0.86)
                seg(r1, a0 + rng.uniform(-0.07, 0.07), r2, a0 + rng.uniform(-0.07, 0.07), 1.1, 230)
            d.ellipse([o + R * 0.78 * SS * math.cos(a0 + 0.17) - 2 * SS, o + R * 0.78 * SS * math.sin(a0 + 0.17) - 2 * SS,
                       o + R * 0.78 * SS * math.cos(a0 + 0.17) + 2 * SS, o + R * 0.78 * SS * math.sin(a0 + 0.17) + 2 * SS],
                      outline=230, width=SS)
    nst = int(rng.choice([5, 6, 7, 8]))
    for k in range(nst):
        a1 = 2 * math.pi * k / nst - math.pi / 2
        a2 = 2 * math.pi * (k + (2 if nst > 5 else 2)) / nst - math.pi / 2
        seg(R * 0.6, a1, R * 0.6, a2, 1.3)
        d.ellipse([o + R * 0.6 * SS * math.cos(a1) - 4 * SS, o + R * 0.6 * SS * math.sin(a1) - 4 * SS,
                   o + R * 0.6 * SS * math.cos(a1) + 4 * SS, o + R * 0.6 * SS * math.sin(a1) + 4 * SS],
                  outline=255, width=SS)
    for k in range(nst):
        a1 = 2 * math.pi * k / nst
        seg(R * 0.24, a1, R * 0.6, a1 + math.pi / nst, 0.8, 160)
    if rot:
        im = im.rotate(math.degrees(rot), resample=Image.BICUBIC)
    sz = max(4, int(2 * R))
    szh = max(4, int(2 * R * tilt))
    im = im.resize((sz, szh), Image.LANCZOS)
    m = np.zeros((H, W), np.float32)
    arr = np.asarray(im, np.float32) / 255
    x0, y0 = int(cx - sz / 2), int(cy - szh / 2)
    xa, ya = max(0, x0), max(0, y0)
    xb, yb = min(W, x0 + sz), min(H, y0 + szh)
    if xb > xa and yb > ya:
        m[ya:yb, xa:xb] = arr[ya - y0:yb - y0, xa - x0:xb - x0]
    col = C(col) if isinstance(col, str) else col
    add(c.img, col, (blur(m, 3) * 1.4 + blur(m, 10) * 0.8) * s)
    screen(c.img, lerp(C(core), col, 0.3), m * s)
    return m


def ribbon_masks(c, n, width=7.0, turns=(1.2, 1.7)):
    """Spiral ribbons of light around the subject -> (back mask, front mask, back core, front core)."""
    rng = c.rng
    S = c.S
    mb, mf = Mask(), Mask()
    for i in range(n):
        T = rng.uniform(*turns) * 2 * math.pi
        phi = rng.uniform(0, 2 * math.pi)
        rx = S.sw * rng.uniform(0.52, 0.66)
        ry = S.sw * rng.uniform(0.10, 0.17)
        y0 = S.cy + S.sh * rng.uniform(0.32, 0.45)
        y1 = S.cy - S.sh * rng.uniform(0.15, 0.4)
        tilt = rng.uniform(-0.22, 0.22)
        N = 260
        prev = None
        for k in range(N):
            t = k / (N - 1)
            a = phi + T * t
            x = S.cx + rx * math.cos(a) * (1.05 - 0.25 * t)
            y = lerp(y0, y1, t) + ry * math.sin(a) + tilt * (x - S.cx)
            z = math.sin(a)
            wd = width * (math.sin(math.pi * t) ** 0.8) * (0.55 + 0.45 * (z * 0.5 + 0.5))
            if prev is not None and wd > 0.3:
                m = mf if (z > 0 and y > S.cy + S.sh * 0.3) else mb
                m.line([prev, (x, y)], wd)
                m.ellipse(x, y, wd / 2, wd / 2)
            prev = (x, y)
    return mb.get(0.5), mf.get(0.5)


def draw_ribbon(c, m, s=1.0):
    if m.max() <= 0:
        return
    ph = (XX * 0.012 + YY * 0.018)
    field = lerp(c.el[1], c.el[3], (0.5 + 0.5 * np.sin(ph))[..., None])
    field = lerp(field, C(c.P['glow']), 0.25)
    add(c.img, field, (blur(m, 3) * 0.7 + blur(m, 10) * 0.35) * s)
    screen(c.img, lerp(np.ones(3, np.float32), c.el[1], 0.3), m * 0.9 * s)


def petals(c, n, cols, region, size=(5, 11), blur_far=0.0, avoid=0.5, a=0.95):
    rng = c.rng
    L = Layer()
    cols = [C(x) for x in cols]
    for (x, y) in c.pts(n, *region, avoid=avoid):
        col = cols[rng.integers(len(cols))]
        Lp = rng.uniform(*size)
        ang = rng.uniform(0, 2 * math.pi)
        L.poly(petal_pts(x, y, Lp, Lp * rng.uniform(0.32, 0.5), ang), col, a)
        L.poly(petal_pts(x, y, Lp * 0.75, Lp * 0.14, ang), lerp(col, np.ones(3), 0.5), a * 0.6)
    comp_layer(c.img, L, 1.0, blur_far)


def bubbles(c, n, region, rmin=2, rmax=9, col='#e8ffff', avoid=0.4):
    rng = c.rng
    m = Mask()
    hl = Mask()
    for (x, y) in c.pts(n, *region, avoid=avoid):
        r = rng.uniform(rmin, rmax) * (0.6 + 0.6 * y / H)
        m.ellipse(x, y, r, r, 0.85, width=max(0.7, r * 0.14))
        m.ellipse(x, y, r * 0.9, r * 0.9, 0.12)
        hl.ellipse(x - r * 0.35, y - r * 0.38, r * 0.22, r * 0.16, 1.0)
    a = m.get(0.4)
    screen(c.img, C(col), a * 0.8 + blur(a, 3) * 0.3)
    screen(c.img, np.ones(3, np.float32), hl.get(0.3))


def grass(c, L, x0, x1, ybase, n, hmin, hmax, cols, lean=0.0, wmin=1.5, wmax=4.0):
    rng = c.rng
    cols = [C(k) for k in cols]
    for _ in range(n):
        x = rng.uniform(x0, x1)
        y = ybase + rng.uniform(-6, 10)
        h = rng.uniform(hmin, hmax)
        w = rng.uniform(wmin, wmax)
        bend = (lean + rng.uniform(-0.5, 0.5)) * h * 0.5
        col = cols[rng.integers(len(cols))] * rng.uniform(0.85, 1.12)
        pts = [(x - w / 2, y), (x + w * 0.1 + bend * 0.5, y - h * 0.55), (x + bend, y - h),
               (x + w * 0.35 + bend * 0.5, y - h * 0.5), (x + w / 2, y)]
        L.poly(pts, col)


def flower(L, x, y, r, col, center, n=5, rot=0.0, squash=1.0):
    for k in range(n):
        a = rot + 2 * math.pi * k / n
        pts = petal_pts(x, y, r, r * 0.5, a, n=10)
        pts = [(px, y + (py - y) * squash) for px, py in pts]
        L.poly(pts, col)
    L.ellipse(x, y, r * 0.3, r * 0.3 * squash, center)


def light_leaks(c, cols, s=0.35):
    rng = c.rng
    corners = [(0, 0), (W, 0), (0, H), (W, H)]
    idx = rng.permutation(4)[:2]
    for k, i in enumerate(idx):
        x, y = corners[i]
        glow(c, x + rng.uniform(-40, 40), y + rng.uniform(-60, 60), rng.uniform(220, 320),
             C(cols[k % len(cols)]), s)


# =============================================================== CSS filter (subject tint)
def css_filter(im, spec):
    arr = np.asarray(im.convert('RGBA'), np.float32) / 255.0
    rgb, a = arr[..., :3], arr[..., 3:]
    for name, val, unit in re.findall(r'([a-z-]+)\(\s*([-\d.]+)\s*(deg|%)?\s*\)', spec):
        v = float(val)
        if unit == '%':
            v /= 100.0
        if name == 'hue-rotate':
            th = math.radians(v)
            cs, sn = math.cos(th), math.sin(th)
            M = np.array([
                [0.213 + cs * 0.787 - sn * 0.213, 0.715 - cs * 0.715 - sn * 0.715, 0.072 - cs * 0.072 + sn * 0.928],
                [0.213 - cs * 0.213 + sn * 0.143, 0.715 + cs * 0.285 + sn * 0.140, 0.072 - cs * 0.072 - sn * 0.283],
                [0.213 - cs * 0.213 - sn * 0.787, 0.715 - cs * 0.715 + sn * 0.715, 0.072 + cs * 0.928 + sn * 0.072]],
                np.float32)
        elif name == 'saturate':
            s = v
            M = np.array([
                [0.213 + 0.787 * s, 0.715 - 0.715 * s, 0.072 - 0.072 * s],
                [0.213 - 0.213 * s, 0.715 + 0.285 * s, 0.072 - 0.072 * s],
                [0.213 - 0.213 * s, 0.715 - 0.715 * s, 0.072 + 0.928 * s]], np.float32)
        elif name == 'sepia':
            k = 1 - min(1.0, v)
            M = np.array([
                [0.393 + 0.607 * k, 0.769 - 0.769 * k, 0.189 - 0.189 * k],
                [0.349 - 0.349 * k, 0.686 + 0.314 * k, 0.168 - 0.168 * k],
                [0.272 - 0.272 * k, 0.534 - 0.534 * k, 0.131 + 0.869 * k]], np.float32)
        elif name == 'grayscale':
            k = 1 - min(1.0, v)
            M = np.array([
                [0.2126 + 0.7874 * k, 0.7152 - 0.7152 * k, 0.0722 - 0.0722 * k],
                [0.2126 - 0.2126 * k, 0.7152 + 0.2848 * k, 0.0722 - 0.0722 * k],
                [0.2126 - 0.2126 * k, 0.7152 - 0.7152 * k, 0.0722 + 0.9278 * k]], np.float32)
        elif name == 'brightness':
            M = np.eye(3, dtype=np.float32) * v
        elif name == 'contrast':
            rgb = np.clip((rgb - 0.5) * v + 0.5, 0, 1)
            continue
        else:
            continue
        rgb = np.clip(rgb @ M.T, 0, 1)
    out = np.concatenate([rgb, a], -1)
    return Image.fromarray((out * 255 + 0.5).astype(np.uint8), 'RGBA')


# =============================================================== subject
class Subject:
    pass


def prep_subject(c):
    card = c.card
    name = os.path.splitext(os.path.basename(card['subject']))[0]
    im = Image.open(card['subject']).convert('RGBA')
    if name == 'shooting_star':           # emoji is a dark tile: key out the tile, keep star + trail
        arr = np.asarray(im, np.float32) / 255.0
        lum = arr[..., :3].max(-1)
        arr[..., 3] *= sstep(0.32, 0.6, lum)
        im = Image.fromarray((arr * 255).astype(np.uint8), 'RGBA')
    if card.get('tint'):
        im = css_filter(im, card['tint'])
    A = np.asarray(im.getchannel('A'))
    ys, xs = np.where(A > 10)
    im = im.crop((xs.min(), ys.min(), xs.max() + 1, ys.max() + 1))
    A = np.asarray(im.getchannel('A'))
    bw, bh = im.size
    cut = name in CUT_SUBJECTS or (A[-2:] > 200).mean(axis=1).max() > 0.9
    S = Subject()
    S.name, S.cut = name, cut
    S.float = (name in FLOAT_SUBJECTS or c.type == 'sort' or c.scene == 'ocean') and not cut
    if c.type == 'heros':
        wf = 0.80
    else:
        wf = [0.67, 0.70, 0.72, 0.75][c.tier] - (0.04 if c.type == 'sort' else 0.0)
    hmax = (0.66 if cut else 0.58) * H
    sc = min(wf * W / bw, hmax / bh)
    sw, sh = int(round(bw * sc)), int(round(bh * sc))
    big = im.resize((sw, sh), Image.LANCZOS)
    arr = np.asarray(big, np.float32) / 255.0
    rgb, a = arr[..., :3], arr[..., 3]
    # bleed colours into transparent area, then unsharp mask (hides upscale softness)
    pm = rgb * a[..., None]
    bl = blur(pm, 2.5) / np.maximum(blur(a, 2.5), 1e-3)[..., None]
    rgb = np.where(a[..., None] > 0.02, rgb, np.clip(bl, 0, 1))
    rgb = np.clip(rgb + (rgb - blur(rgb, 1.6)) * 0.75, 0, 1)
    S.sw, S.sh = sw, sh
    S.cx = W * 0.5
    S.cy = H * (0.425 if cut else 0.405)
    if c.type == 'heros':
        S.cy = H * 0.43
    ox, oy = int(round(S.cx - sw / 2)), int(round(S.cy - sh / 2))
    if cut:                               # make sure the flat cut sits well inside the fade zone
        oy = int(min(oy + 6, H * 0.72 - sh))
        oy = max(oy, int(H * 0.085))
    S.ox, S.oy = ox, oy
    S.cy = oy + sh / 2
    S.top, S.bottom = oy, oy + sh
    S.rgb = np.zeros((H, W, 3), np.float32)
    S.a = np.zeros((H, W), np.float32)
    xa, ya, xb, yb = max(0, ox), max(0, oy), min(W, ox + sw), min(H, oy + sh)
    S.rgb[ya:yb, xa:xb] = rgb[ya - oy:yb - oy, xa - ox:xb - ox]
    S.a[ya:yb, xa:xb] = a[ya - oy:yb - oy, xa - ox:xb - ox]
    # actual opaque bottom (for contact shadow)
    rows = np.where(S.a.max(axis=1) > 0.5)[0]
    S.foot = rows.max() if len(rows) else S.bottom
    c.S = S
    return S


def subject_backdrop(c):
    """Aura, glow and magic behind the subject."""
    S, P = c.S, c.P
    gcol = C(P['glow'])
    ecol = c.el[1]
    # big soft aura
    glow(c, S.cx, S.cy, max(S.sw, S.sh) * 0.75, lerp(gcol, ecol, 0.35), 0.12 + 0.04 * c.tier)
    if c.tier >= 3 or c.type == 'heros':
        # radiant burst
        ang = np.arctan2(YY - S.cy, XX - S.cx)
        k = int(c.rng.integers(14, 22))
        burst = (0.5 + 0.5 * np.cos(ang * k)) ** 8 + 0.5 * (0.5 + 0.5 * np.cos(ang * k * 2 + 1.3)) ** 12
        d = np.sqrt((XX - S.cx) ** 2 + (YY - S.cy) ** 2)
        fall = sstep(S.sw * 0.2, S.sw * 0.4, d) * np.exp(-d / (S.sw * 0.6))
        screen(c.img, lerp(gcol, np.ones(3), 0.4), burst * fall * 0.28)
    if c.type == 'sort' or c.type == 'heros' or c.tier >= 2 and c.P.get('circle', False):
        R = max(S.sw, S.sh) * (0.62 if c.type != 'heros' else 0.6)
        magic_circle(c, S.cx, S.cy - (S.sh * 0.06 if c.type == 'heros' else 0), R, lerp(ecol, gcol, 0.3),
                     s=0.55 if c.type != 'sort' else 0.75, rot=c.rng.uniform(0, 1))


def subject_ground(c):
    S, P = c.S, c.P
    if S.cut:
        return
    if S.float:
        gy = min(H * 0.70, S.foot + S.sh * 0.12)
        m = Mask()
        m.ellipse(S.cx, gy, S.sw * 0.32, S.sw * 0.06, 1)
        over(c.img, C(P['shadow']), m.get(7) * 0.35)
        glow(c, S.cx, gy, S.sw * 0.45, lerp(C(P['glow']), c.el[1], 0.4), 0.45, ry=S.sw * 0.08)
    else:
        m = Mask()
        m.ellipse(S.cx, S.foot - 3, S.sw * 0.40, S.sw * 0.055, 1)
        over(c.img, C(P['shadow']), m.get(5) * 0.55)
        m2 = Mask()
        m2.ellipse(S.cx, S.foot - 2, S.sw * 0.26, S.sw * 0.025, 1)
        over(c.img, C(P['shadow']) * 0.6, m2.get(2) * 0.55)


def composite_subject(c):
    S, P = c.S, c.P
    rgb, a = S.rgb.copy(), S.a.copy()
    lx, ly = c.light[0] - S.cx, c.light[1] - S.cy
    nrm = math.hypot(lx, ly) + 1e-6
    lx, ly = lx / nrm, ly / nrm
    key, amb, gcol = C(P['key']), C(P['amb']), C(P['glow'])
    # outer glow
    og = blur(a, 5) * 0.5 + blur(a, 20) * 0.45
    screen(c.img, lerp(gcol, c.el[1], 0.4), og * (0.3 + 0.06 * c.tier) * (1 - a))
    # directional ambient grading
    g = ((XX - S.cx) * lx + (YY - S.cy) * ly) / (max(S.sw, S.sh) * 0.5)
    g = sstep(-1.0, 0.9, g)[..., None]
    shadowF = lerp(np.ones(3, np.float32), amb / max(amb.max(), 1e-3), 0.55) * 0.86
    lightF = lerp(np.ones(3, np.float32), key / max(key.max(), 1e-3), 0.3) * 1.06
    rgb = rgb * lerp(shadowF, lightF, g)
    # scene colour harmonisation (slight)
    lum = rgb.mean(-1, keepdims=True)
    rgb = lerp(rgb, lum * lerp(np.ones(3), C(P['haze']), 0.6) * 1.05, 0.08)
    # rim light from key
    k = max(3.0, S.sw * 0.016)
    rim = np.clip(a - shift(a, -lx * k, -ly * k), 0, 1)
    rim = blur(rim, 1.2) * a
    screen(rgb, lerp(key, np.ones(3), 0.35), rim * 1.0)
    # soft edge back-light all around
    inner = np.clip(a - blur(a, 4.5), 0, 1) * 2.0 * a
    screen(rgb, lerp(gcol, c.el[1], 0.4), inner * 0.45)
    # bounce light from below
    k2 = max(2.0, S.sw * 0.01)
    rimb = np.clip(a - shift(a, 0, -k2), 0, 1)
    screen(rgb, C(P['bounce']), blur(rimb, 1.5) * a * 0.55)
    # haze + base fade
    rel = np.clip((YY - S.top) / max(S.sh, 1), 0, 1)
    hz = 0.04 + 0.22 * sstep(0.55, 1.0, rel) * (0.4 if S.cut else 1.0)
    rgb = lerp(rgb, C(P['haze']), hz[..., None] * a[..., None] * 0.6)
    if S.cut:
        a = a * sstep(S.foot + 1, S.foot - S.sh * 0.24, YY)
    over(c.img, rgb, a)
    c.S.a_final = a


# =============================================================== palettes
PAL = {
    'arcane': dict(key='#c8f4ff', amb='#6a4ad0', glow='#d9b8ff', haze='#c8a8f0', shadow='#1e1050', bounce='#ff9ae0',
                   leak=['#ff8fd8', '#7ae8ff'], circle=True),
    'feu': dict(key='#ffe2a0', amb='#b0306a', glow='#ffcf7a', haze='#ffb08a', shadow='#3a0a1e', bounce='#ff7a3a',
                leak=['#ffb050', '#ff5a8a']),
    'glace': dict(key='#e8fbff', amb='#5a7ad8', glow='#bff2ff', haze='#c8e8ff', shadow='#2a3a7a', bounce='#9af0ff',
                  leak=['#9af0ff', '#d0a8ff'], circle=True),
    'nuit': dict(key='#e4d8ff', amb='#4a3aa0', glow='#c9a8ff', haze='#9a88d8', shadow='#140c38', bounce='#a8ffd0',
                 leak=['#b08aff', '#7ab8ff'], circle=True),
    'sacre': dict(key='#fffbe8', amb='#d8a080', glow='#ffe9a0', haze='#fff0d0', shadow='#a87a50', bounce='#ffd890',
                  leak=['#fff0b0', '#ffc8d8'], circle=True),
    'montagne': dict(key='#fff0d8', amb='#6a70c0', glow='#ffe8d0', haze='#e0d0f0', shadow='#2a2e60', bounce='#ffc0a8',
                     leak=['#ffd0b0', '#c8b8ff']),
    'plaine': dict(key='#fffbe0', amb='#6a9ad0', glow='#fff6c0', haze='#e8f6ff', shadow='#2a5a30', bounce='#c8f0a0',
                   leak=['#fff0b0', '#ffc8e0']),
    'ocean': dict(key='#e0fffa', amb='#2a6ab0', glow='#a8fff4', haze='#6ad0e0', shadow='#0a3060', bounce='#7af0ff',
                  leak=['#a8fff0', '#7ab0ff']),
    'sang': dict(key='#ffd8b8', amb='#8a2050', glow='#ffb0a0', haze='#ff9a8a', shadow='#2a0614', bounce='#ff6a6a',
                 leak=['#ff7a8a', '#ffb070']),
    'or': dict(key='#fff4c8', amb='#b07020', glow='#ffe390', haze='#ffe0a0', shadow='#4a2a08', bounce='#ffd060',
               leak=['#fff0a0', '#ffb860'], circle=True),
    'pierre': dict(key='#e8f0ff', amb='#4a4878', glow='#c8e8ff', haze='#8a88b0', shadow='#14121e', bounce='#9af0ff',
                   leak=['#9af0ff', '#d0a0ff'], circle=True),
    'foret': dict(key='#f8ffd0', amb='#2a7a6a', glow='#d8ffb0', haze='#b8e8c8', shadow='#0e2a20', bounce='#c8ff9a',
                  leak=['#f0ffb0', '#a8ffe0']),
}


# =============================================================== scenes
# ---------------------------------------------------------------- ARCANE (nebula)
def bg_arcane(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#120a3a'), (0.25, '#2c1878'), (0.5, '#6a3cc0'), (0.7, '#c27ad8'), (0.85, '#f0a8e0'), (1, '#ffd0ee')])
    n1, n2, n3 = fbm(rng, 2, 6), fbm(rng, 3, 6), fbm(rng, 4, 5)
    top = sstep(0.85, 0.15, YY / H)
    screen(c.img, C('#ff78d0'), sstep(0.42, 0.95, n1) * 0.6 * top)
    screen(c.img, C('#5ae4ff'), sstep(0.55, 0.98, n2) * 0.5 * top)
    over(c.img, C('#160a3c'), sstep(0.55, 0.9, n3) * 0.4 * top)
    stars(c, 300, 0, H * 0.6)
    c.light = (W * rng.uniform(0.25, 0.75), H * 0.05)
    # ringed planet in an upper corner
    side = rng.choice([-1, 1])
    px, py, R = W * (0.5 + side * rng.uniform(0.28, 0.36)), H * rng.uniform(0.12, 0.2), rng.uniform(38, 58)
    pc1, pc2 = (C('#ffb8e8'), C('#8a5ae0')) if rng.random() < 0.5 else (C('#b8f4ff'), C('#5a6ae0'))
    ring_tilt = rng.uniform(-0.35, 0.35)
    for front in (False, True):
        m = Mask()
        pts_o = ellipse_pts(px, py, R * 2.0, R * 0.42, ring_tilt, 60)
        m.line(pts_o + [pts_o[0]], 3.0, 0.9)
        pts_i = ellipse_pts(px, py, R * 1.65, R * 0.33, ring_tilt, 60)
        m.line(pts_i + [pts_i[0]], 1.4, 0.6)
        rm = m.get(0.4)
        rel = (YY - py) * math.cos(ring_tilt) - (XX - px) * math.sin(ring_tilt)
        rm = rm * (rel > 0 if front else rel <= 0)
        if front:
            screen(c.img, lerp(pc1, np.ones(3), 0.4), rm * 0.9)
        else:
            screen(c.img, lerp(pc1, np.ones(3), 0.2), rm * 0.6)
            nx, ny = (XX - px) / R, (YY - py) / R
            d2 = nx ** 2 + ny ** 2
            inside = np.clip((1 - np.sqrt(d2)) * R + 0.5, 0, 1)
            nz = np.sqrt(np.clip(1 - d2, 0, 1))
            lit = sstep(-0.15, 0.55, -nx * 0.6 * side - ny * 0.45 + nz * 0.5)
            bands = 0.5 + 0.5 * np.sin(ny * 11 + rng.uniform(0, 6) + np.sin(nx * 4) * 0.8)
            col = lerp(C('#2a1a6a'), lerp(pc2, pc1, bands[..., None] * 0.6 + 0.3), lit[..., None])
            over(c.img, col, inside)
            screen(c.img, pc1, inside * np.clip(1 - nz, 0, 1) ** 4 * 0.7 * lit)
    glow(c, px, py, R * 2.2, pc1, 0.12)
    # far crystal spires on the horizon
    hor = H * 0.64
    L = Layer()
    for i in range(14):
        x = rng.uniform(-20, W + 20)
        hh = rng.uniform(40, 150) * (1.3 if abs(x - W / 2) > 140 else 0.7)
        crystal(L, x, hor + 6, hh, hh * rng.uniform(0.16, 0.24), rng.uniform(-0.25, 0.25),
                C('#6a50c0'), C('#8a70d8'), C('#b0a0f0'), C('#d8c8ff'), 1.0, edge=False)
    a = comp_layer(c.img, L, 0.8, 1.2)
    over(c.img, C('#c8a0f0'), a * 0.35)
    # floating islands
    for i in range(2):
        ix = W * (0.5 + (i * 2 - 1) * rng.uniform(0.3, 0.4))
        iy = H * rng.uniform(0.3, 0.45)
        iw = rng.uniform(26, 44)
        L = Layer()
        pts = [(ix - iw, iy), (ix + iw, iy)]
        for k in range(7):
            t = k / 6
            pts.append((ix + iw * (1 - 2 * t) * (1 - 0.3 * abs(0.5 - t)), iy + iw * (0.4 + 0.9 * math.sin(math.pi * t)) * rng.uniform(0.7, 1.0)))
        L.poly(pts, C('#3a2480'))
        L.ellipse(ix, iy, iw, iw * 0.16, C('#8a6ad8'))
        for k in range(3):
            hh = rng.uniform(16, 32)
            crystal(L, ix + rng.uniform(-iw * 0.6, iw * 0.6), iy + 1, hh, hh * 0.25, rng.uniform(-0.4, 0.4),
                    C('#5a8ae0'), C('#7ac8f0'), C('#c8f8ff'), C('#ffffff'), 0.95, edge=False)
        a = comp_layer(c.img, L, 0.9, 0.9)
        over(c.img, C('#b090e0'), a * 0.3)
        glow(c, ix, iy, iw * 1.6, C('#9af0ff'), 0.25)
    # ground
    yl = hor + (fbm1(rng, W, 3, 4) - 0.5) * 14
    fill_ridge(c, yl, '#7a52c8', '#1e1050', depth=260, rim='#ffc0f0', rimw=3, rims=0.7)
    over(c.img, C('#2a1468'), sstep(H * 0.7, H, YY) * 0.5)
    # glowing floor runes
    magic_circle(c, S.cx, max(S.foot, H * 0.66) + 4, S.sw * 0.62, lerp(C('#7af0ff'), c.el[2], 0.3), tilt=0.26, s=0.7)
    glow(c, S.cx, S.cy, max(S.sw, S.sh) * 0.8, C('#ffb0f0'), 0.12)


def fg_arcane(c):
    rng, S = c.rng, c.S
    palettes = [('#5a2aa8', '#9a5ae8', '#d8b0ff', '#ffffff'), ('#2a5ab0', '#4ab8e8', '#b8f4ff', '#ffffff'),
                ('#a02a80', '#e85ab8', '#ffb8e8', '#ffffff')]
    L = Layer()
    for side in (-1, 1):
        x = W / 2 + side * rng.uniform(W * 0.36, W * 0.46)
        crystal_cluster(c, L, x, H * rng.uniform(0.84, 0.92), rng.uniform(150, 210), side * 0.25,
                        palettes[rng.integers(3)], n=6, spread=0.6)
    for side in (-1, 1):
        x = S.cx + side * S.sw * rng.uniform(0.25, 0.4)
        crystal_cluster(c, L, x, S.foot + rng.uniform(10, 26), rng.uniform(60, 90), side * 0.35,
                        palettes[rng.integers(3)], n=4, spread=0.6)
    a = comp_layer(c.img, L)
    glow_m = blur(a, 10)
    add(c.img, C('#c890ff'), glow_m * 0.25 * (1 - a))
    # floating shards
    L = Layer()
    for (x, y) in c.pts(3 + c.tier, 30, 80, W - 30, H * 0.62, avoid=0.97, scale=0.6):
        shard(L, x, y, rng.uniform(22, 44), rng.uniform(-0.6, 0.6), C('#5a3ac0'), C('#a8c8ff'), C('#f0f8ff'))
    a = comp_layer(c.img, L)
    add(c.img, C('#9ae8ff'), blur(a, 6) * 0.5 * (1 - a))
    bokeh(c, 14 + 6 * c.tier, (0, 0, W, H), 3, 16, ['#ff9ae8', '#9af0ff', '#d8b0ff'])
    pts = [(x, y, rng.uniform(5, 13)) for x, y in c.pts(10 + 5 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, np.ones(3, np.float32), C('#c8a0ff'))


# ---------------------------------------------------------------- FEU (volcano)
def bg_feu(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#3a0a34'), (0.2, '#8a1c4c'), (0.4, '#e0484e'), (0.55, '#ff8a4e'), (0.66, '#ffc070'), (1, '#ffe0a0')])
    vx = W * (0.5 + rng.choice([-1, 1]) * rng.uniform(0.2, 0.3))
    c.light = (vx, H * 0.3)
    glow(c, vx, H * 0.45, 300, C('#ffd080'), 0.6)
    # smoke clouds
    n = fbm(rng, 3, 6, sx=0.7)
    cl = sstep(0.45, 0.85, n) * sstep(H * 0.62, H * 0.1, YY)
    lit = blur(n, 3) - shift(blur(n, 3), 0, -6)
    smoke = lerp(C('#a03a6a'), C('#ffb090'), np.clip(0.4 + lit[..., None] * 6, 0, 1))
    over(c.img, smoke, cl * 0.65)
    stars(c, 60, 0, H * 0.2, big=0.0)
    # far ridge
    yl = mountain_ridge(c, 5, H * 0.6, 30, 90, slope=(0.5, 0.9), detail=10)
    mountain_layer(c, yl, '#c4486a', '#a03058', '#ff9a7a', '#a02a5a', haze='#e07070', haze_a=0.35, rim='#ffc080', s=0.8)
    # volcano
    base = H * 0.66
    peak = H * rng.uniform(0.26, 0.32)
    xs = np.arange(W, dtype=np.float32)
    hw = 30
    sl = rng.uniform(1.15, 1.5)
    prof = np.where(np.abs(xs - vx) < hw, peak + (np.abs(xs - vx) / hw) ** 2 * 6, peak + 6 + (np.abs(xs - vx) - hw) * sl)
    prof = np.minimum(prof + (fbm1(rng, W, 8, 4) - 0.5) * 16, base + 40)
    vm = mountain_layer(c, prof, '#6a1a3e', '#3a0c26', '#c8506a', '#4a1030', rim='#ffb070', depth=300)
    # lava streaks
    m = Mask()
    for i in range(int(rng.integers(4, 7))):
        x = vx + rng.uniform(-hw * 0.8, hw * 0.8)
        y = peak + 6
        pts = [(x, y)]
        dirn = np.sign(x - vx) or 1
        for k in range(30):
            x += dirn * rng.uniform(0.5, 2.6) + rng.uniform(-1, 1)
            y += rng.uniform(4, 8)
            pts.append((x, y))
            if y > base:
                break
        m.line(pts, rng.uniform(1.6, 3.2))
    lava = m.get(0.5) * vm
    add(c.img, C('#ff7a2a'), blur(lava, 4) * 1.0)
    screen(c.img, C('#ffe080'), lava)
    glow(c, vx, peak, 70, C('#ffd060'), 0.9)
    # eruption plume
    pl = fbm(rng, 4, 5)
    col_m = np.exp(-((XX - vx - (peak - YY) * 0.25) / (25 + (peak - YY) * 0.5)) ** 2) * sstep(peak + 5, peak - 140, YY) * sstep(peak - 330, peak - 120, YY)
    over(c.img, lerp(C('#ff9a6a'), C('#ffe0b0'), pl[..., None]), col_m * sstep(0.3, 0.8, pl) * 0.85)
    # lava bombs
    m = Mask()
    for i in range(14):
        t = rng.uniform(0.1, 1)
        ang = rng.uniform(-1.2, 1.2)
        x, y = vx + math.sin(ang) * 120 * t, peak - 90 * math.sin(math.pi * t * 0.8) * rng.uniform(0.4, 1)
        m.ellipse(x, y, 2.2, 2.2)
    bm = m.get()
    add(c.img, C('#ff8a3a'), blur(bm, 3) * 2)
    screen(c.img, C('#fff0a0'), bm)
    # mid rocky ridge
    yl2 = mountain_ridge(c, 4, H * 0.7, 20, 60, slope=(0.6, 1.2), detail=18)
    mountain_layer(c, yl2, '#4a1030', '#26061a', '#a83a50', '#30081e', rim='#ff9a50', depth=200)
    # lava ground with glowing cracks
    g = fbm(rng, 5, 5, sy=2.8)
    cracks = (1 - np.abs(g - 0.5) * 2) ** 16
    gm = sstep(H * 0.69, H * 0.74, YY)
    bump = fbm(rng, 7, 4, sy=2.5)
    litf = np.clip((bump - shift(bump, int(np.sign(c.light[0] - W / 2) * 3), 3)) * 8 + 0.5, 0, 1)
    gcol = lerp(C('#5a1a40'), C('#2a0a22'), sstep(H * 0.7, H, YY)[..., None])
    gcol = lerp(gcol, C('#b04a5a'), (litf * 0.45 * sstep(H, H * 0.7, YY))[..., None])
    over(c.img, gcol, gm)
    # lava river winding to the volcano
    xs = np.arange(H, dtype=np.float32)
    rx = lerp(vx, W * 0.5 + rng.uniform(-80, 80), sstep(H * 0.68, H, xs)) + np.sin(xs / 37 + rng.uniform(0, 6)) * 22
    rw = lerp(3, 46, sstep(H * 0.68, H, xs))
    river = np.clip(rw[:, None] - np.abs(XX - rx[:, None]) * (1 + (fbm(rng, 6, 3) - 0.5)), 0, 1) * gm
    river = blur(river, 1.0)
    add(c.img, C('#ff5a2a'), blur(river, 8) * 0.8)
    over(c.img, lerp(C('#ff8a2a'), C('#ffe080'), fbm(rng, 8, 3)[..., None]), river)
    add(c.img, C('#ff6a2a'), blur(cracks, 2) * gm * 0.9)
    screen(c.img, C('#ffd070'), cracks * gm * 0.9)
    glow(c, S.cx, S.foot + 20, S.sw * 0.6, C('#ff8a3a'), 0.5, ry=40)


def fg_feu(c):
    rng, S = c.rng, c.S
    # foreground rocks
    items = []
    for side in (-1, 1):
        bx = W / 2 + side * rng.uniform(190, 240)
        for k in range(7):
            items.append((bx + side * rng.uniform(-40, 80) - side * k * 6, H * rng.uniform(0.8, 0.99) + k * 4, rng.uniform(34, 66)))
    items.sort(key=lambda b: b[1])
    lxx = 1 if c.light[0] > W / 2 else -1
    a = puffs(c, items, '#c0506a', '#5a1a3e', '#22061a', light=(lxx * 0.8, -0.7), style='rock', rim='#ffb070', rims=1.0)
    g = fbm(rng, 8, 4)
    cr = (1 - np.abs(g - 0.5) * 2) ** 18 * a
    add(c.img, C('#ff5a2a'), blur(cr, 2) * 1.3)
    screen(c.img, C('#ffe080'), cr)
    # amber crystals
    L = Layer()
    for side in (-1, 1):
        x = S.cx + side * S.sw * rng.uniform(0.38, 0.5)
        crystal_cluster(c, L, x, H * rng.uniform(0.78, 0.86), rng.uniform(70, 110), side * 0.35,
                        ('#a8301a', '#ff7a2a', '#ffc860', '#fff4c0'), n=4)
    a = comp_layer(c.img, L)
    add(c.img, C('#ff8a3a'), blur(a, 10) * 0.4 * (1 - a))
    # embers
    m = Mask()
    for (x, y) in c.pts(70 + 20 * c.tier, 0, 0, W, H, avoid=0.5):
        l = rng.uniform(2, 10)
        ang = -math.pi / 2 + rng.uniform(-0.6, 0.6)
        m.line([(x, y), (x + math.cos(ang) * l, y + math.sin(ang) * l)], rng.uniform(1.0, 2.4), rng.uniform(0.5, 1))
    em = m.get()
    add(c.img, C('#ff6a2a'), blur(em, 2.5) * 1.6)
    screen(c.img, C('#fff0b0'), em)
    bokeh(c, 12 + 5 * c.tier, (0, 0, W, H), 4, 18, ['#ffb060', '#ff7a8a', '#ffe0a0'])
    pts = [(x, y, rng.uniform(5, 12)) for x, y in c.pts(8 + 4 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, C('#fff8e0'), C('#ffa040'))


# ---------------------------------------------------------------- GLACE (ice + aurora)
def aurora(c, y_base, amp, height, c1, c2, s=0.8):
    rng = c.rng
    yb = y_base + (fbm1(rng, W, 2.5, 4) - 0.5) * amp
    streak = fbm1(rng, W, 70, 3)
    streak = 0.3 + 0.7 * ((streak - streak.min()) / (np.ptp(streak) + 1e-6)) ** 1.6
    env = 0.4 + 0.6 * fbm1(rng, W, 4, 3)
    d = yb[None, :] - YY
    m = np.where(d > 0, np.exp(-d / height), np.exp(d / 5)) * (streak * env)[None, :]
    t = np.clip(d / (height * 1.6), 0, 1)[..., None]
    col = lerp(C(c1), C(c2), t)
    add(c.img, col, blur(m, 1.0) * s)


def bg_glace(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#0a1640'), (0.25, '#1c3a88'), (0.5, '#4a80cc'), (0.68, '#a0d4f2'), (0.85, '#e0f6ff'), (1, '#f4fcff')])
    stars(c, 240, 0, H * 0.55)
    c.light = (W * rng.uniform(0.3, 0.7), H * 0.15)
    aurora(c, H * rng.uniform(0.28, 0.36), 90, 110, '#6affc0', '#b080ff', 0.75)
    aurora(c, H * rng.uniform(0.18, 0.26), 70, 80, '#5ae0ff', '#ff90e0', 0.45)
    glow(c, W * 0.5, H * 0.6, 300, C('#bfe8ff'), 0.2)
    yl = mountain_ridge(c, 5, H * 0.62, 60, 150, slope=(0.8, 1.4), detail=10)
    mountain_layer(c, yl, '#a8c8f0', '#8aaee8', '#f0faff', '#8aa0e0', snow=('#ffffff', '#c8d8ff'), snow_frac=0.45,
                   lightx=np.sign(c.light[0] - W / 2) or 1, haze='#d0ecff', haze_a=0.45, s=0.8)
    yl2 = mountain_ridge(c, 4, H * 0.67, 40, 110, slope=(0.9, 1.6), detail=12)
    mountain_layer(c, yl2, '#6a90d8', '#3a5aa8', '#e8f6ff', '#5a6ec8', snow=('#ffffff', '#a8c0f0'), snow_frac=0.35,
                   lightx=np.sign(c.light[0] - W / 2) or 1, rim='#ffffff')
    fog_band(c, H * 0.66, 30, '#e8f8ff', 0.6, sx=0.6)
    # frozen lake / snow field
    gm = sstep(H * 0.68, H * 0.71, YY)
    over(c.img, lerp(C('#d8f0ff'), C('#6a8ad8'), sstep(H * 0.7, H, YY)[..., None]), gm)
    refl = fbm(rng, 10, 3, sy=0.15)
    screen(c.img, C('#ffffff'), gm * sstep(0.7, 1.0, refl) * 0.35)
    glow(c, S.cx, S.foot + 14, S.sw * 0.6, C('#bff6ff'), 0.45, ry=36)


def fg_glace(c):
    rng, S = c.rng, c.S
    pal = [('#3a6ac0', '#7ac0f0', '#e0f8ff', '#ffffff'), ('#5a5ac8', '#a8a0f8', '#f0e8ff', '#ffffff'),
           ('#2a8ac0', '#6ae0f0', '#d8ffff', '#ffffff')]
    # snow mounds
    xs = np.arange(W, dtype=np.float32)
    yl = H * 0.86 - 70 * sstep(W * 0.3, 0, xs) - 60 * sstep(W * 0.7, W, xs) + (fbm1(rng, W, 4, 4) - 0.5) * 24
    m = fill_ridge(c, yl, '#f4fbff', '#8aa8e0', depth=160, rim='#ffffff', rimw=3, rims=0.8)
    gl = (rng.random((H, W)) > 0.996).astype(np.float32) * m
    screen(c.img, np.ones(3, np.float32), gl + blur(gl, 1.5) * 1.5)
    L = Layer()
    for side in (-1, 1):
        x = W / 2 + side * rng.uniform(W * 0.33, W * 0.45)
        crystal_cluster(c, L, x, H * rng.uniform(0.82, 0.9), rng.uniform(140, 200), side * 0.28, pal[rng.integers(3)], n=6)
    for side in (-1, 1):
        x = S.cx + side * S.sw * rng.uniform(0.28, 0.42)
        crystal_cluster(c, L, x, S.foot + rng.uniform(8, 24), rng.uniform(50, 80), side * 0.35, pal[rng.integers(3)], n=4)
    a = comp_layer(c.img, L)
    add(c.img, C('#9af0ff'), blur(a, 10) * 0.3 * (1 - a))
    # snowflakes (bokeh + crisp)
    bokeh(c, 26 + 8 * c.tier, (0, 0, W, H), 2, 10, ['#ffffff', '#d8f4ff'], amin=0.3, amax=0.7, soft=1.4, ring=False)
    bokeh(c, 6, (0, 0, W, H), 14, 26, ['#e8faff'], amin=0.12, amax=0.25, soft=3)
    pts = [(x, y, rng.uniform(5, 13)) for x, y in c.pts(10 + 5 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, np.ones(3, np.float32), C('#9ae8ff'))


# ---------------------------------------------------------------- SACRE (holy sky temple)
def bg_sacre(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#5a4ab8'), (0.16, '#9a6ac8'), (0.32, '#e894b4'), (0.48, '#ffc890'), (0.62, '#fff0c8'), (0.75, '#f0c8d8'), (1, '#b8a0d8')])
    c.light = (W * 0.5 + rng.uniform(-60, 60), H * 0.08)
    glow(c, c.light[0], H * 0.1, 180, C('#fff4d0'), 0.55)
    god_rays(c, c.light[0], -60, '#fff4d8', 0.32, cells=30, power=2.4, reach=650)
    for i in range(2):
        cx = W * (0.5 + (2 * i - 1) * rng.uniform(0.25, 0.42))
        cy = H * rng.uniform(0.18, 0.3)
        puffs(c, cumulus(c, cx, cy, rng.uniform(140, 220), rng.uniform(26, 40)), '#fff4ec', '#ffc8c0', '#a888c8',
              light=((c.light[0] - cx) / W, -0.6), soft=4, a=0.92, rim='#fffbe8', flat=cy)
    hor = H * 0.62
    L = Layer()
    for side in (-1, 1):
        n = int(rng.integers(3, 5))
        for k in range(n):
            x = W / 2 + side * (120 + k * 42)
            hh = 210 - k * 4
            L.poly([(x - 9, hor), (x + 9, hor), (x + 8, hor - hh), (x - 8, hor - hh)], C('#f6d6a0'))
            if side < 0:
                L.poly([(x + 3, hor), (x + 9, hor), (x + 8, hor - hh), (x + 3, hor - hh)], C('#d0a0b8'))
            else:
                L.poly([(x - 9, hor), (x - 3, hor), (x - 3, hor - hh), (x - 8, hor - hh)], C('#d0a0b8'))
            L.poly([(x - 14, hor - hh), (x + 14, hor - hh), (x + 12, hor - hh - 8), (x - 12, hor - hh - 8)], C('#fbe4b8'))
    a = comp_layer(c.img, L, 0.85, 1.0)
    over(c.img, C('#ffe8d8'), a * 0.3)
    hc = (S.cx, S.cy - S.sh * 0.12)
    m = Mask()
    for k in range(3 + c.tier):
        r = S.sw * (0.48 + 0.12 * k)
        m.ellipse(hc[0], hc[1], r, r, 0.75 - 0.12 * k, width=1.3 + (1.2 if k == 0 else 0))
    hm = m.get(0.5)
    add(c.img, C('#ffc860'), blur(hm, 4) * 0.6)
    screen(c.img, C('#fffbe8'), hm * 0.9)
    for y, hh, a in ((H * 0.62, 30, 0.9), (H * 0.72, 42, 1.0), (H * 0.86, 56, 1.0)):
        blobs = []
        for cx in np.linspace(-60, W + 60, 5):
            blobs += cumulus(c, cx + rng.uniform(-25, 25), y + rng.uniform(-8, 8), rng.uniform(140, 200), hh)
        blobs.sort(key=lambda b: b[1])
        puffs(c, blobs, '#fffaf4', '#f8c8d0', '#9a86c8', light=((c.light[0] - W / 2) / W, -1), soft=5, a=a,
              rim='#ffffff', rims=0.8)
        over(c.img, C('#c8a8e0'), sstep(y - 10, y + 120, YY) * 0.18)
    glow(c, S.cx, S.cy, max(S.sw, S.sh), C('#fff0b0'), 0.12)


def fg_sacre(c):
    rng, S = c.rng, c.S
    blobs = []
    for side in (-1, 1):
        blobs += cumulus(c, W / 2 + side * rng.uniform(170, 230), H * 0.98, rng.uniform(200, 260), rng.uniform(55, 75))
    blobs.sort(key=lambda b: b[1])
    puffs(c, blobs, '#fffcf6', '#f8ccd4', '#9a88cc', light=(0, -1), soft=6, rim='#ffffff', rims=0.9)
    L = Layer()
    for (x, y) in c.pts(4 + 2 * c.tier, 10, 70, W - 10, H * 0.75, avoid=0.95, scale=0.55):
        ln = rng.uniform(20, 34)
        ang = rng.uniform(0, 2 * math.pi)
        L.poly(petal_pts(x, y, ln, ln * 0.24, ang, tip=0.8), C('#fffaf2'), 0.97)
        L.poly(petal_pts(x, y, ln, ln * 0.12, ang + 0.05, tip=0.8), C('#f0d8e8'), 0.6)
        L.line([(x, y), (x + math.cos(ang) * ln * 0.95, y + math.sin(ang) * ln * 0.95)], C('#e0b880'), 0.7, 0.9)
    a = comp_layer(c.img, L)
    add(c.img, C('#ffd8a0'), blur(a, 5) * 0.3 * (1 - a))
    petals(c, 8 + 4 * c.tier, ['#ffc8d8', '#fff0c0', '#ffffff'], (0, 0, W, H), size=(5, 9))
    bokeh(c, 14 + 5 * c.tier, (0, 0, W, H), 3, 16, ['#ffe9a0', '#fff8e0', '#ffc8d8'], amin=0.12, amax=0.35)
    pts = [(x, y, rng.uniform(6, 14)) for x, y in c.pts(10 + 5 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, np.ones(3, np.float32), C('#ffb830'))


# ---------------------------------------------------------------- MONTAGNE
def bg_montagne(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#3a52a0'), (0.22, '#6a88d8'), (0.42, '#b8a8e8'), (0.56, '#ffc0c8'), (0.68, '#ffd8b8'), (1, '#fff0d8')])
    sx = W * rng.choice([0.25, 0.75]) + rng.uniform(-30, 30)
    c.light = (sx, H * 0.5)
    glow(c, sx, H * 0.52, 260, C('#fff0c8'), 0.75)
    glow(c, sx, H * 0.52, 40, C('#ffffff'), 0.9)
    god_rays(c, sx, H * 0.52, '#fff4e0', 0.25, cells=24, reach=500, start=20)
    # wispy cirrus
    n = fbm(rng, 3, 5, sx=0.35, sy=1.6)
    over(c.img, C('#fff0f0'), sstep(0.6, 0.95, n) * sstep(H * 0.45, H * 0.05, YY) * 0.45)
    lx = 1 if sx > W / 2 else -1
    layers = [
        (H * 0.58, 120, 220, '#d8c4ec', '#c8b4e4', '#fff0f0', '#b0a0d8', 0.55, '#f8e0ff'),
        (H * 0.63, 90, 190, '#a8a0dc', '#8a88cc', '#ffe8e8', '#7a78c0', 0.3, '#e8d8f8'),
        (H * 0.69, 70, 150, '#6e70c0', '#4a4e9a', '#f8d8d8', '#4a4c98', 0.12, None),
    ]
    for i, (base, hmin, hmax, top, bot, lit, shade, hz, fog) in enumerate(layers):
        yl = mountain_ridge(c, 4 + i, base, hmin, hmax, slope=(0.8, 1.6), detail=12 - i * 2)
        mountain_layer(c, yl, top, bot, lit, shade, snow=('#ffffff', '#c8c8f0'), snow_frac=0.4,
                       lightx=lx, haze='#f0d8f0', haze_a=hz, rim='#fff0e0' if i == 2 else None, s=max(0, 1.2 - i * 0.5))
        if fog:
            fog_band(c, base + 10, 26, fog, 0.65, sx=0.5)
    # birds
    m = Mask()
    for _ in range(int(rng.integers(3, 7))):
        x, y, s = rng.uniform(40, W - 40), rng.uniform(H * 0.15, H * 0.35), rng.uniform(3, 6)
        m.line([(x - s, y - s * 0.5), (x, y), (x + s, y - s * 0.6)], 1.0)
    over(c.img, C('#4a4a8a'), m.get() * 0.7)


def fg_montagne(c):
    rng, S = c.rng, c.S
    lx = 1 if c.light[0] > W / 2 else -1
    fog_band(c, H * 0.7, 40, '#fff0f4', 0.5, sx=0.6)
    # far cloud sea below the peaks
    blobs = []
    for cx in np.linspace(-40, W + 40, 5):
        blobs += cumulus(c, cx + rng.uniform(-20, 20), H * 0.8, rng.uniform(130, 180), 34)
    puffs(c, blobs, '#fff4f0', '#f0c8d8', '#8a84c0', light=(lx * 0.6, -1), soft=5, rim='#fff8f0', a=0.95)
    # boulders framing the bottom corners
    L_items = []
    for side in (-1, 1):
        bx = W / 2 + side * rng.uniform(170, 230)
        for k in range(7):
            L_items.append((bx + side * rng.uniform(-40, 90) - side * k * 8, H * rng.uniform(0.76, 0.98) + k * 6,
                            rng.uniform(34, 70)))
    if not S.cut and not S.float:
        for k in range(6):
            L_items.append((S.cx + rng.uniform(-S.sw * 0.45, S.sw * 0.45), S.foot + rng.uniform(25, 60), rng.uniform(30, 50)))
    L_items.sort(key=lambda b: b[1])
    al = puffs(c, L_items, '#b8b0e8', '#6a66b0', '#2a2a60', light=(lx * 0.8, -0.8), style='rock', rim='#ffd8c0', rims=0.9)
    # tufts and alpine flowers on the boulders
    L = Layer()
    tops = [(x, y - r * 0.8) for x, y, r in L_items]
    for (x, y) in tops:
        grass(c, L, x - 20, x + 20, y + 6, 12, 8, 22, ['#6a9a7a', '#8ab888', '#4a7a6a'], lean=lx * 0.3, wmin=1.5, wmax=3)
        for k in range(int(rng.integers(1, 4))):
            col = C(['#ffffff', '#ffd0e0', '#c8d0ff', '#fff0a0'][rng.integers(4)])
            flower(L, x + rng.uniform(-18, 18), y + rng.uniform(-4, 6), rng.uniform(4, 7), col, C('#ffd040'),
                   n=int(rng.choice([5, 6, 8])), rot=rng.uniform(0, 1), squash=0.8)
    comp_layer(c.img, L)
    petals(c, 8 + 4 * c.tier, ['#ffd0e0', '#ffffff', '#ffe8b0'], (0, 0, W, H))
    bokeh(c, 12 + 5 * c.tier, (0, 0, W, H), 3, 16, ['#fff0d0', '#ffd0e0', '#d0d8ff'])
    pts = [(x, y, rng.uniform(5, 12)) for x, y in c.pts(8 + 4 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, np.ones(3, np.float32), C('#ffd8a0'))


# ---------------------------------------------------------------- PLAINE (meadow)
def bg_plaine(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#3a8ae0'), (0.3, '#7ac0f4'), (0.55, '#c8eafc'), (0.68, '#f4fbff'), (1, '#fff8e0')])
    sx = W * rng.uniform(0.15, 0.85)
    c.light = (sx, H * 0.12)
    glow(c, sx, H * 0.12, 200, C('#fffbe0'), 0.85)
    god_rays(c, sx, H * 0.12, '#fffbe8', 0.3, cells=30, reach=600, start=10)
    for i in range(3):
        cx, cy = rng.uniform(-40, W + 40), H * rng.uniform(0.12, 0.42)
        blobs = cloud_bank(c, cy, cx - rng.uniform(60, 120), cx + rng.uniform(60, 120), 18, 44, 7, 10)
        puffs(c, blobs, '#ffffff', '#f0f6ff', '#a8c0e8', light=((sx - cx) / W, -1), soft=4, rim='#ffffff')
    lx = 1 if sx > W / 2 else -1
    hills = [(H * 0.6, '#b8dca8', '#a0d090', 0.45), (H * 0.65, '#8ccb7a', '#6ab060', 0.2), (H * 0.71, '#5aa858', '#3a8048', 0.0)]
    for i, (base, top, bot, hz) in enumerate(hills):
        yl = base + (fbm1(rng, W, 1.5, 4) - 0.5) * 70
        tex = fbm(rng, 8, 3, sy=0.3)
        m = fill_ridge(c, yl, top, bot, depth=150, rim='#f8ffd0', rimw=4, rims=0.5, tex=tex)
        over(c.img, C('#e0f0ff'), m * hz * 0.7)
        if i == 0:
            # distant castle
            x = W * rng.choice([0.2, 0.8]) + rng.uniform(-30, 30)
            yb = yl[int(np.clip(x, 0, W - 1))] + 4
            L = Layer()
            cc = C('#9ab8c8')
            for dx, hh, ww in [(-16, 34, 8), (0, 48, 10), (16, 30, 8), (-8, 24, 22)]:
                L.poly([(x + dx - ww / 2, yb), (x + dx + ww / 2, yb), (x + dx + ww / 2, yb - hh), (x + dx - ww / 2, yb - hh)], cc)
                L.poly([(x + dx - ww / 2 - 2, yb - hh), (x + dx + ww / 2 + 2, yb - hh), (x + dx, yb - hh - ww * 1.2)], C('#c890b0'))
            comp_layer(c.img, L, 0.9, 0.6)
        # trees on hills
        L = Layer()
        for _ in range(int(rng.integers(4, 10))):
            x = rng.uniform(0, W)
            y = yl[int(np.clip(x, 0, W - 1))] + rng.uniform(2, 20)
            r = rng.uniform(6, 12) * (1 + i * 0.6)
            col = C(['#4a9a5a', '#3a8a5a', '#5aa86a'][i]) * (1 + 0.15 * (2 - i))
            L.ellipse(x, y - r, r, r * 1.15, col)
            L.ellipse(x + lx * r * 0.3, y - r * 1.25, r * 0.55, r * 0.55, lerp(col, C('#f0ffb0'), 0.4))
        comp_layer(c.img, L, 1.0 - hz, 0.5 * (2 - i))
    # flower field gradient
    for k in range(3):
        L = Layer()
        n = [120, 70, 30][k]
        for _ in range(n):
            y = H * rng.uniform(0.66 + k * 0.08, 0.8 + k * 0.1)
            x = rng.uniform(0, W)
            r = (1.2 + (y / H - 0.6) * 14) * rng.uniform(0.7, 1.2)
            col = C(['#ff9ec0', '#fff07a', '#ffffff', '#c8a0ff', '#ff8a7a', '#9ad8ff'][rng.integers(6)])
            if r < 4:
                L.ellipse(x, y, r, r * 0.7, col)
            else:
                flower(L, x, y, r, col, C('#ffd040'), n=5, rot=rng.uniform(0, 1), squash=0.7)
        comp_layer(c.img, L)
    glow(c, S.cx, S.foot, S.sw * 0.6, C('#fff8c0'), 0.3, ry=40)


def fg_plaine(c):
    rng, S = c.rng, c.S
    lx = 1 if c.light[0] > W / 2 else -1
    L = Layer()
    grass(c, L, -10, W + 10, H * 0.98, 160, 30, 110, ['#4a9a50', '#6ab858', '#3a7a48', '#8acc6a'], lean=lx * 0.3,
          wmin=3, wmax=7)
    for _ in range(14 + 3 * c.tier):
        x = rng.choice([rng.uniform(-10, W * 0.3), rng.uniform(W * 0.7, W + 10)])
        y = H * rng.uniform(0.8, 0.98)
        r = rng.uniform(10, 20)
        col = C(['#ff9ec0', '#ffffff', '#fff07a', '#c8a0ff'][rng.integers(4)])
        L.line([(x, y), (x + rng.uniform(-6, 6), y + 90)], C('#4a8a48'), 2.5)
        flower(L, x, y, r, col, C('#ffc030'), n=int(rng.choice([5, 6])), rot=rng.uniform(0, 1), squash=0.8)
    comp_layer(c.img, L)
    petals(c, 14 + 5 * c.tier, ['#ffc8dc', '#ffffff', '#ffe8a0'], (0, 0, W, H))
    petals(c, 4, ['#ffc8dc', '#ffffff'], (0, 0, W, H), size=(16, 26), blur_far=3)
    # butterflies
    L = Layer()
    for (x, y) in c.pts(2 + c.tier // 2, 30, 80, W - 30, H * 0.6, avoid=0.95, scale=0.55):
        col = C(['#ffb0e0', '#a8e0ff', '#fff0a0'][rng.integers(3)])
        s = rng.uniform(6, 10)
        for sgn in (-1, 1):
            L.poly(petal_pts(x, y, s * 1.2, s * 0.6, -math.pi / 2 + sgn * 0.9), col)
            L.poly(petal_pts(x, y, s * 0.8, s * 0.4, math.pi / 2 + sgn * 0.7), lerp(col, np.ones(3), 0.3))
    comp_layer(c.img, L)
    bokeh(c, 12 + 5 * c.tier, (0, 0, W, H), 3, 16, ['#fff8d0', '#ffd8e8', '#e0ffd0'])
    pts = [(x, y, rng.uniform(5, 11)) for x, y in c.pts(8 + 4 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, np.ones(3, np.float32), C('#ffe080'))


# ---------------------------------------------------------------- OCEAN (underwater)
def coral(c, L, x, y, length, ang, width, col, tipc, depth=0):
    rng = c.rng
    if depth > 4 or length < 6:
        L.ellipse(x, y, width * 0.9, width * 0.9, tipc)
        return
    x2 = x + math.sin(ang) * length
    y2 = y - math.cos(ang) * length
    L.line([(x, y), (x2, y2)], col, width)
    L.ellipse(x2, y2, width / 2, width / 2, col)
    nb = 2 if rng.random() < 0.8 else 3
    for k in range(nb):
        coral(c, L, x2, y2, length * rng.uniform(0.6, 0.8), ang + rng.uniform(-0.7, 0.7), width * 0.78,
              lerp(col, tipc, 0.2), tipc, depth + 1)


def caustics(rng, base=7, sharp=10):
    g = fbm(rng, base, 3)
    g2 = fbm(rng, base * 1.4, 3)
    return ((1 - np.abs(g - 0.5) * 2) ** sharp + (1 - np.abs(g2 - 0.5) * 2) ** sharp) * 0.6


def bg_ocean(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#c8fff6'), (0.1, '#6ae0ea'), (0.35, '#2aa8cc'), (0.65, '#1a6aa8'), (1, '#0c3a78')])
    sx = W * rng.uniform(0.3, 0.7)
    c.light = (sx, -100)
    surf = sstep(H * 0.14, 0, YY)
    screen(c.img, C('#ffffff'), caustics(rng, 6, 8) * surf * 0.8)
    glow(c, sx, -20, 200, C('#ffffff'), 0.8)
    god_rays(c, sx, -140, '#e8fffa', 0.6, cells=36, power=2.6, reach=650)
    # far reef silhouettes
    for i, (base, col, hz) in enumerate([(H * 0.6, '#3a8ab8', 0.5), (H * 0.66, '#2a6aa0', 0.3)]):
        yl = base + (fbm1(rng, W, 3, 5) - 0.5) * 70 - (fbm1(rng, W, 20, 2) - 0.5) * 20
        fill_ridge(c, yl, col, '#14406e', depth=200, a=0.85, s=1.5 - i * 0.6, rim='#9af0ff', rims=0.3)
        L = Layer()
        for _ in range(8):
            x = rng.uniform(0, W)
            coral(c, L, x, yl[int(x)] + 4, rng.uniform(16, 30), rng.uniform(-0.4, 0.4), rng.uniform(3, 5),
                  C(col) * 1.1, C(col) * 1.4)
        comp_layer(c.img, L, 0.7, 1.2 - i * 0.5)
    # kelp
    L = Layer()
    for _ in range(int(rng.integers(5, 9))):
        x = rng.choice([rng.uniform(0, W * 0.3), rng.uniform(W * 0.7, W)])
        pts = []
        ph = rng.uniform(0, 6)
        top = H * rng.uniform(0.15, 0.45)
        for k in range(30):
            t = k / 29
            y = lerp(H * 0.75, top, t)
            pts.append((x + math.sin(ph + t * 5) * 12 * t, y))
        L.line(pts, C('#1a6a80'), rng.uniform(4, 7))
        for k in range(3, 29, 3):
            px, py = pts[k]
            L.poly(petal_pts(px, py, 18, 5, -math.pi / 2 + (0.8 if k % 2 else -0.8)), C('#2a8a90'))
    a = comp_layer(c.img, L, 0.8, 1.2)
    over(c.img, C('#3aa0c0'), a * 0.35)
    # fish school
    m = Mask()
    fx, fy = rng.uniform(60, W - 60), H * rng.uniform(0.2, 0.4)
    dirn = rng.choice([-1, 1])
    for _ in range(int(rng.integers(10, 22))):
        x, y = fx + rng.normal(0, 40), fy + rng.normal(0, 18)
        s = rng.uniform(3, 6)
        m.poly(ellipse_pts(x, y, s, s * 0.4, 0, 10))
        m.poly([(x - dirn * s * 0.9, y), (x - dirn * s * 1.7, y - s * 0.5), (x - dirn * s * 1.7, y + s * 0.5)])
    over(c.img, C('#1a5a8a'), m.get(0.6) * 0.6)
    # sand floor with caustics
    gm = sstep(H * 0.7, H * 0.74, YY)
    over(c.img, lerp(C('#3ab0c0'), C('#14507a'), sstep(H * 0.7, H, YY)[..., None]), gm)
    screen(c.img, C('#b8fff4'), caustics(rng, 9, 9) * gm * 0.4)
    glow(c, S.cx, S.cy, max(S.sw, S.sh) * 0.85, C('#a8fff4'), 0.15)


def fg_ocean(c):
    rng, S = c.rng, c.S
    L = Layer()
    for side in (-1, 1):
        for k in range(3):
            x = W / 2 + side * rng.uniform(W * 0.28, W * 0.5)
            col = C(['#ff8ab0', '#ffa878', '#c8a0ff', '#ff9ad8'][rng.integers(4)])
            coral(c, L, x, H * rng.uniform(0.86, 0.98), rng.uniform(34, 56), side * rng.uniform(0.0, 0.4), rng.uniform(8, 12),
                  col * 0.8, lerp(col, np.ones(3), 0.55))
        # sea fan
        x = W / 2 + side * rng.uniform(W * 0.3, W * 0.42)
        y = H * rng.uniform(0.8, 0.88)
        fc = C(['#ff7ab0', '#b888ff'][rng.integers(2)])
        for k in range(26):
            a = side * 0.3 + (k / 25 - 0.5) * 1.8
            ln = rng.uniform(50, 80)
            L.line([(x, y), (x + math.sin(a) * ln, y - math.cos(a) * ln)], fc, 1.6, 0.9)
    a = comp_layer(c.img, L)
    # anemone light & glow
    add(c.img, C('#ffb0d0'), blur(a, 8) * 0.2 * (1 - a))
    # seaweed in front
    L = Layer()
    for side in (-1, 1):
        for _ in range(3):
            x = W / 2 + side * rng.uniform(W * 0.42, W * 0.55)
            pts = []
            ph = rng.uniform(0, 6)
            for k in range(25):
                t = k / 24
                pts.append((x + math.sin(ph + t * 4) * 16 * t, lerp(H + 10, H * rng.uniform(0.45, 0.6), t)))
            L.line(pts, C('#1e8a7a'), rng.uniform(9, 14))
    comp_layer(c.img, L, 1.0, 1.5)
    bubbles(c, 30 + 10 * c.tier, (0, 0, W, H * 0.95))
    bokeh(c, 14 + 5 * c.tier, (0, 0, W, H), 3, 14, ['#b8fff4', '#ffd0f0', '#a8d8ff'])
    pts = [(x, y, rng.uniform(5, 11)) for x, y in c.pts(8 + 4 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, np.ones(3, np.float32), C('#7af0ff'))


# ---------------------------------------------------------------- FORET (enchanted forest)
def trunk(L, x, y_bot, y_top, w, lean, col, rng, branches=2, lit=None, lx=1):
    pts_l, pts_r = [], []
    ph = rng.uniform(0, 6)
    for k in range(13):
        t = k / 12
        y = lerp(y_bot, y_top, t)
        xx = x + lean * t * 60 + math.sin(ph + t * 3) * w * 0.25
        ww = w * (1.6 - 0.7 * t) if t < 0.15 else w * (1.0 - 0.35 * t)
        if t < 0.1:
            ww = w * (2.2 - 8 * t)
        pts_l.append((xx - ww / 2, y))
        pts_r.append((xx + ww / 2, y))
    L.poly(pts_l + pts_r[::-1], col)
    if lit is not None:
        side = pts_r if lx > 0 else pts_l
        L.line(side, lit, max(1.0, w * 0.12), 0.8)
    for b in range(branches):
        t = rng.uniform(0.45, 0.85)
        k = int(t * 12)
        bx, by = pts_l[k][0] + w / 2, pts_l[k][1]
        d = rng.choice([-1, 1])
        L.line([(bx, by), (bx + d * w * 1.4, by - w * 1.2), (bx + d * w * 2.4, by - w * 1.4)], col, max(1.5, w * 0.25))


def bg_foret(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#0c3a40'), (0.25, '#1e6a62'), (0.48, '#6ac0a0'), (0.62, '#d8f8c8'), (0.75, '#f0ffe0'), (1, '#8ac898')])
    sx = W * rng.uniform(0.35, 0.65)
    c.light = (sx, H * 0.35)
    glow(c, sx, H * 0.42, 260, C('#fbffe0'), 0.85)
    god_rays(c, sx, H * 0.05, '#f8ffd8', 0.5, cells=26, power=2.4, reach=900)
    layers = [(14, '#a8dcc0', 0.6, 2.5, 6, 0.6), (9, '#5aa890', 0.35, 1.5, 11, 0.85), (6, '#2a6a5a', 0.0, 0.6, 18, 1.0)]
    for i, (n, col, hz, bl, wbase, op) in enumerate(layers):
        L = Layer()
        for k in range(n):
            x = rng.uniform(-20, W + 20)
            if i == 2 and abs(x - W / 2) < 120:
                x = W / 2 + np.sign(x - W / 2 + 1e-3) * rng.uniform(140, 260)
            w = wbase * rng.uniform(0.8, 1.4)
            trunk(L, x, H * (0.68 + i * 0.05), -20, w, rng.uniform(-0.3, 0.3), C(col), rng,
                  branches=int(rng.integers(1, 3)), lit=lerp(C(col), C('#f8ffd0'), 0.6), lx=1 if sx > x else -1)
        comp_layer(c.img, L, op, bl)
        # canopy masses at top
        blobs = [(rng.uniform(-30, W + 30), rng.uniform(-30, H * 0.12 + i * 18), rng.uniform(30, 60)) for _ in range(10)]
        puffs(c, blobs, lerp(C(col), C('#e8ffb0'), 0.45), col, C(col) * 0.75, light=((sx - W / 2) / W, 1), soft=4, a=op, s=bl)
        fog_band(c, H * (0.62 + 0.04 * i), 40, '#e8ffe0', 0.35 - 0.1 * i, sx=0.6)
    # ground
    yl = H * 0.69 + (fbm1(rng, W, 3, 4) - 0.5) * 20
    tex = fbm(rng, 10, 3, sy=0.3)
    fill_ridge(c, yl, '#5aa070', '#1a4a38', depth=200, rim='#e0ffb0', rimw=3, rims=0.5, tex=tex)
    # light pool on ground
    glow(c, S.cx, S.foot, S.sw * 0.7, C('#f0ffc0'), 0.5, ry=45)
    # glowing mushrooms mid
    L = Layer()
    for _ in range(int(rng.integers(5, 9))):
        x = rng.uniform(20, W - 20)
        y = yl[int(x)] + rng.uniform(4, 30)
        s = rng.uniform(4, 9)
        col = C(['#ff9ad0', '#9ae8ff', '#e0ff9a'][rng.integers(3)])
        L.line([(x, y), (x, y - s * 1.4)], C('#f0f0e0'), s * 0.35)
        L.poly(ellipse_pts(x, y - s * 1.4, s, s * 0.6, 0, 16)[8:] + [(x + s, y - s * 1.4)], col)
    a = comp_layer(c.img, L)
    add(c.img, C('#c8ffd0'), blur(a, 6) * 0.6)


def fg_foret(c):
    rng, S = c.rng, c.S
    lx = 1 if c.light[0] > W / 2 else -1
    # ferns
    L = Layer()
    for side in (-1, 1):
        for _ in range(3):
            bx = W / 2 + side * rng.uniform(W * 0.32, W * 0.55)
            by = H * rng.uniform(0.86, 1.0)
            for k in range(7):
                a = side * rng.uniform(0.1, 1.2) + (-0.4 if side > 0 else 0.4)
                ln = rng.uniform(70, 130)
                col = C(['#2a7a50', '#3a9a5a', '#4aa868'][rng.integers(3)])
                pts = []
                for j in range(10):
                    t = j / 9
                    pts.append((bx + math.sin(a + t * 0.8 * side) * ln * t, by - math.cos(a + t * 0.3) * ln * t))
                L.line(pts, col * 0.8, 2.0)
                for j in range(1, 9):
                    px, py = pts[j]
                    sz = 14 * (1 - j / 10)
                    L.poly(petal_pts(px, py, sz, sz * 0.32, a - math.pi / 2 + 0.9 * side), col)
                    L.poly(petal_pts(px, py, sz, sz * 0.32, a - math.pi / 2 - 0.9 * side + math.pi * 0), col * 1.1)
    comp_layer(c.img, L)
    L = Layer()
    grass(c, L, -10, W + 10, H * 0.99, 120, 20, 70, ['#3a8a50', '#5aa860', '#2a6a48'], lean=lx * 0.2, wmin=2, wmax=5)
    for _ in range(10 + 3 * c.tier):
        x, y = rng.uniform(0, W), H * rng.uniform(0.72, 0.98)
        r = rng.uniform(5, 11)
        flower(L, x, y, r, C(['#ffffff', '#ffd0f0', '#d0e8ff', '#fff4a0'][rng.integers(4)]), C('#ffd040'),
               n=5, rot=rng.uniform(0, 1), squash=0.75)
    comp_layer(c.img, L)
    # big glowing mushrooms
    L = Layer()
    for side in (-1, 1):
        x = S.cx + side * S.sw * rng.uniform(0.38, 0.55)
        y = H * rng.uniform(0.78, 0.86)
        for k in range(int(rng.integers(2, 4))):
            s = rng.uniform(12, 24)
            xx = x + rng.uniform(-30, 30)
            yy = y + rng.uniform(-6, 16)
            col = C(['#ff8ac8', '#7ae0ff', '#c8a0ff'][rng.integers(3)])
            L.line([(xx, yy), (xx + side * 3, yy - s * 1.4)], C('#f6f0e0'), s * 0.32)
            cap = ellipse_pts(xx + side * 3, yy - s * 1.4, s, s * 0.7, 0, 24)[12:] + [(xx + side * 3 + s, yy - s * 1.4)]
            L.poly(cap, col)
            for d in range(3):
                L.ellipse(xx + side * 3 + rng.uniform(-s * 0.6, s * 0.6), yy - s * 1.4 - rng.uniform(s * 0.2, s * 0.5),
                          s * 0.1, s * 0.1, np.ones(3))
    a = comp_layer(c.img, L)
    add(c.img, C('#c8f0ff'), blur(a, 10) * 0.4 * (1 - a))
    # fireflies
    m = Mask()
    for (x, y) in c.pts(30 + 8 * c.tier, 0, 0, W, H, avoid=0.6):
        r = rng.uniform(1.2, 2.6)
        m.ellipse(x, y, r, r, rng.uniform(0.5, 1))
    fm = m.get()
    add(c.img, C('#d8ff8a'), blur(fm, 3) * 1.5 + blur(fm, 8) * 0.8)
    screen(c.img, C('#ffffe0'), fm)
    petals(c, 8 + 3 * c.tier, ['#c8f0a0', '#ffe0f0', '#f0ffd0'], (0, 0, W, H))
    bokeh(c, 12 + 5 * c.tier, (0, 0, W, H), 3, 16, ['#f0ffc0', '#c0ffe8', '#ffe0f0'])
    pts = [(x, y, rng.uniform(5, 11)) for x, y in c.pts(8 + 4 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, np.ones(3, np.float32), C('#c8ff8a'))


# ---------------------------------------------------------------- NUIT (moonlit night)
def bg_nuit(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#0e0a34'), (0.3, '#2c2278'), (0.52, '#6a50b8'), (0.66, '#c090d8'), (0.78, '#f0b8e0'), (1, '#6a50a0')])
    stars(c, 380, 0, H * 0.62, big=0.08)
    mx = W * rng.choice([0.25, 0.75]) + rng.uniform(-25, 25)
    my = H * rng.uniform(0.14, 0.2)
    R = rng.uniform(46, 62)
    c.light = (mx, my)
    glow(c, mx, my, R * 5, C('#c8b0ff'), 0.6)
    glow(c, mx, my, R * 2.2, C('#fff0f8'), 0.55)
    d = np.sqrt((XX - mx) ** 2 + (YY - my) ** 2)
    moon = np.clip(R - d + 0.5, 0, 1)
    tex = fbm(rng, 10, 4)
    mc = lerp(C('#fff6e0'), C('#e0d0f0'), sstep(0.5, 0.9, tex)[..., None] * 0.8)
    over(c.img, mc, moon)
    m = Mask()
    for k in range(6):
        cr = rng.uniform(4, 10)
        a = rng.uniform(0, 6.28)
        rr = rng.uniform(0, R * 0.7)
        m.ellipse(mx + math.cos(a) * rr, my + math.sin(a) * rr, cr, cr, 0.5)
    over(c.img, C('#d8c8ec'), m.get(1.2) * moon * 0.5)
    # moonlit clouds
    for i in range(3):
        cy = H * rng.uniform(0.25, 0.48)
        cx = rng.uniform(0, W)
        blobs = cloud_bank(c, cy, cx - 120, cx + 120, 14, 32, 8, 8)
        puffs(c, blobs, '#e8d0f8', '#9a80d0', '#4a3a90', light=(mx - cx, my - cy), soft=4, a=0.75, rim='#fff0ff')
    # castle on a hill
    side = -1 if mx > W / 2 else 1
    hx = W / 2 + side * rng.uniform(90, 160)
    xs = np.arange(W, dtype=np.float32)
    yl = H * 0.6 - 60 * np.exp(-((xs - hx) / 90) ** 2) + (fbm1(rng, W, 4, 4) - 0.5) * 16
    fill_ridge(c, yl, '#5a46a0', '#3a2a78', depth=200, a=1.0, s=0.8, rim='#d8c0ff', rims=0.5)
    L = Layer()
    cc = C('#4a3a90')
    yb = yl[int(np.clip(hx, 0, W - 1))] + 6
    for dx, hh, ww in [(-30, 70, 14), (-12, 110, 16), (8, 86, 14), (26, 60, 12), (-20, 40, 60)]:
        x = hx + dx
        L.poly([(x - ww / 2, yb), (x + ww / 2, yb), (x + ww / 2, yb - hh), (x - ww / 2, yb - hh)], cc)
        if ww < 30:
            L.poly([(x - ww / 2 - 3, yb - hh), (x + ww / 2 + 3, yb - hh), (x, yb - hh - ww * 1.6)], cc)
    comp_layer(c.img, L, 0.95, 0.6)
    m = Mask()
    for k in range(7):
        m.ellipse(hx + rng.uniform(-30, 30), yb - rng.uniform(15, 90), 1.6, 2.2, 1)
    wm = m.get()
    add(c.img, C('#ffd080'), wm * 1.2 + blur(wm, 3) * 1.5)
    # second hill + mist
    yl2 = H * 0.67 + (fbm1(rng, W, 2.5, 4) - 0.5) * 40
    fill_ridge(c, yl2, '#3a2c80', '#1a1240', depth=220, rim='#c8b0ff', rims=0.5)
    fog_band(c, H * 0.66, 30, '#b8a0f0', 0.5, sx=0.6)
    # twisted trees framing
    L = Layer()
    for sd in (-1, 1):
        x = W / 2 + sd * rng.uniform(W * 0.4, W * 0.5)
        trunk(L, x, H * 0.85, H * 0.15, 14, -sd * 0.6, C('#140c34'), rng, branches=4, lit=C('#8a70d0'), lx=-sd)
    comp_layer(c.img, L)
    glow(c, S.cx, S.cy, max(S.sw, S.sh) * 0.85, C('#b8a0ff'), 0.15)


def fg_nuit(c):
    rng, S = c.rng, c.S
    fog_band(c, S.foot - 4, 40, '#c8b8ff', 0.5, sx=0.6)
    L = Layer()
    grass(c, L, -10, W + 10, H * 0.99, 140, 20, 80, ['#2a2060', '#3a2a7a', '#1e1650'], wmin=2, wmax=5)
    comp_layer(c.img, L)
    # glowing bell flowers
    L = Layer()
    for _ in range(14 + 4 * c.tier):
        x = rng.choice([rng.uniform(0, W * 0.35), rng.uniform(W * 0.65, W)])
        y = H * rng.uniform(0.74, 0.98)
        r = rng.uniform(5, 11) * (0.6 + y / H * 0.6)
        col = C(['#9ab8ff', '#e0a8ff', '#a8fff0', '#ffb8e8'][rng.integers(4)])
        L.line([(x, y + 30), (x + rng.uniform(-6, 6), y - r)], C('#3a3a8a'), 1.6)
        L.poly(petal_pts(x, y - r, r * 1.4, r * 0.7, math.pi / 2, tip=0.4), col)
    a = comp_layer(c.img, L)
    add(c.img, C('#b8a8ff'), blur(a, 6) * 0.9 * (1 - a) + blur(a, 16) * 0.4)
    m = Mask()
    for (x, y) in c.pts(30 + 8 * c.tier, 0, 0, W, H, avoid=0.6):
        r = rng.uniform(1.2, 2.6)
        m.ellipse(x, y, r, r, rng.uniform(0.5, 1))
    fm = m.get()
    add(c.img, C('#b0ffb8'), blur(fm, 3) * 1.4 + blur(fm, 8) * 0.7)
    screen(c.img, C('#f0ffe8'), fm)
    bokeh(c, 14 + 5 * c.tier, (0, 0, W, H), 3, 16, ['#c8a8ff', '#a8c8ff', '#ffc8f0'])
    pts = [(x, y, rng.uniform(5, 12)) for x, y in c.pts(8 + 4 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, np.ones(3, np.float32), C('#b090ff'))


# ---------------------------------------------------------------- SANG (crimson battlefield)
def bg_sang(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#2a0620'), (0.22, '#6a1438'), (0.42, '#c83a58'), (0.56, '#ff7a6a'), (0.66, '#ffb48a'), (1, '#ffd8b0')])
    sx = W * rng.uniform(0.3, 0.7)
    sy = H * 0.5
    c.light = (sx, sy)
    glow(c, sx, sy, 300, C('#ff8a6a'), 0.6)
    d = np.sqrt((XX - sx) ** 2 + (YY - sy) ** 2)
    over(c.img, C('#fff0d0'), np.clip(70 - d + 0.5, 0, 1) * 0.9)
    glow(c, sx, sy, 110, C('#ffe0c0'), 0.6)
    n = fbm(rng, 3, 5, sx=0.3, sy=2.2)
    cl = sstep(0.5, 0.9, n) * sstep(H * 0.6, H * 0.1, YY)
    over(c.img, lerp(C('#7a1a3a'), C('#ff9a8a'), np.clip(1 - d / 300, 0, 1)[..., None]), cl * 0.7)
    god_rays(c, sx, sy, '#ffd0b0', 0.3, cells=20, reach=400, start=60)
    # far fortress + hills
    yl = H * 0.6 + (fbm1(rng, W, 2, 4) - 0.5) * 40
    fill_ridge(c, yl, '#b8405a', '#8a2a48', depth=200, s=1.0, rim='#ffc0a0', rims=0.5)
    L = Layer()
    fx = rng.choice([W * 0.18, W * 0.82])
    yb = yl[int(fx)] + 4
    for dx, hh, ww in [(-34, 50, 12), (-16, 74, 14), (4, 60, 12), (22, 80, 14), (40, 46, 12), (0, 32, 80)]:
        x = fx + dx
        L.poly([(x - ww / 2, yb), (x + ww / 2, yb), (x + ww / 2, yb - hh), (x - ww / 2, yb - hh)], C('#8a2a48'))
        if ww < 30:
            for b in range(3):
                bxx = x - ww / 2 + b * ww / 2.5
                L.poly([(bxx, yb - hh), (bxx + ww / 5, yb - hh), (bxx + ww / 5, yb - hh - 4), (bxx, yb - hh - 4)], C('#8a2a48'))
    comp_layer(c.img, L, 0.9, 0.8)
    # battlefield banners & spears (mid)
    yl2 = H * 0.66 + (fbm1(rng, W, 3, 4) - 0.5) * 30
    L = Layer()
    for _ in range(int(rng.integers(10, 16))):
        x = rng.uniform(0, W)
        y = yl2[int(x)] + rng.uniform(4, 16)
        hh = rng.uniform(60, 130)
        lean = rng.uniform(-0.25, 0.25)
        top = (x + lean * hh, y - hh)
        L.line([(x, y), top], C('#3a0a20'), 2.2)
        if rng.random() < 0.55:
            dirn = rng.choice([-1, 1])
            fw, fh = rng.uniform(22, 38), rng.uniform(14, 22)
            pts = [top, (top[0] + dirn * fw * 0.5, top[1] + 3), (top[0] + dirn * fw, top[1] + 1),
                   (top[0] + dirn * fw * 0.8, top[1] + fh * 0.6), (top[0] + dirn * fw, top[1] + fh), (top[0], top[1] + fh)]
            L.poly(pts, C(['#5a0a2a', '#7a1a3a', '#3a0a2a'][rng.integers(3)]))
        else:
            L.poly([(top[0] - 3, top[1]), (top[0] + 3, top[1]), (top[0] + lean * 12, top[1] - 12)], C('#3a0a20'))
    comp_layer(c.img, L, 0.85, 0.8)
    fill_ridge(c, yl2, '#7a1a38', '#2a0614', depth=220, rim='#ffb090', rims=0.6)
    fog_band(c, H * 0.66, 34, '#ff9a8a', 0.4, sx=0.5)
    glow(c, S.cx, S.foot, S.sw * 0.6, C('#ff8a6a'), 0.4, ry=40)


def fg_sang(c):
    rng, S = c.rng, c.S
    L = Layer()
    for side in (-1, 1):
        for k in range(2):
            x = W / 2 + side * rng.uniform(W * 0.3, W * 0.46)
            y = H * rng.uniform(0.86, 0.96)
            hh = rng.uniform(110, 170)
            a = side * rng.uniform(0.05, 0.3)
            dx, dy = math.sin(a), -math.cos(a)
            px, py = math.cos(a), math.sin(a)
            bw = 9
            tip = (x + dx * hh, y + dy * hh)
            L.poly([(x - px * bw, y - py * bw), (x + px * bw, y + py * bw),
                    (x + px * bw + dx * hh * 0.85, y + py * bw + dy * hh * 0.85), tip,
                    (x - px * bw + dx * hh * 0.85, y - py * bw + dy * hh * 0.85)], C('#d8c8d8'))
            L.poly([(x, y), (x + px * bw, y + py * bw), (x + px * bw + dx * hh * 0.85, y + py * bw + dy * hh * 0.85), tip],
                   C('#fff0f0'))
            gy = 0.05
            L.poly([(x - px * 26 - dx * 4, y - py * 26 - dy * 4), (x + px * 26 - dx * 4, y + py * 26 - dy * 4),
                    (x + px * 26 - dx * 12, y + py * 26 - dy * 12), (x - px * 26 - dx * 12, y - py * 26 - dy * 12)],
                   C('#c89040'))
            L.line([(x - dx * 12, y - dy * 12), (x - dx * 50, y - dy * 50)], C('#4a1a20'), 7)
    a = comp_layer(c.img, L)
    add(c.img, C('#ff9a8a'), blur(a, 6) * 0.35 * (1 - a))
    # ground
    yl = H * 0.82 + (fbm1(rng, W, 3, 4) - 0.5) * 30 + ((np.arange(W) - W / 2) / (W / 2)) ** 2 * -30
    fill_ridge(c, yl, '#4a0c22', '#1a0410', depth=150, rim='#ff9a7a', rims=0.6)
    # red roses / petals swirl
    L = Layer()
    for _ in range(8 + 2 * c.tier):
        x = rng.choice([rng.uniform(0, W * 0.3), rng.uniform(W * 0.7, W)])
        y = H * rng.uniform(0.84, 0.98)
        r = rng.uniform(8, 14)
        for k in range(3):
            L.ellipse(x, y, r * (1 - k * 0.28), r * (1 - k * 0.28) * 0.85, C(['#a01a3a', '#d83a5a', '#ff7a8a'][k]))
    comp_layer(c.img, L)
    petals(c, 22 + 6 * c.tier, ['#ff4a6a', '#ff8aa0', '#d01a3a', '#ffb0b8'], (0, 0, W, H))
    petals(c, 4, ['#ff4a6a', '#ff8aa0'], (0, 0, W, H), size=(18, 28), blur_far=3)
    m = Mask()
    for (x, y) in c.pts(40 + 10 * c.tier, 0, 0, W, H, avoid=0.5):
        l = rng.uniform(2, 8)
        ang = -math.pi / 2 + rng.uniform(-0.8, 0.8)
        m.line([(x, y), (x + math.cos(ang) * l, y + math.sin(ang) * l)], rng.uniform(1.0, 2.0))
    em = m.get()
    add(c.img, C('#ff7a4a'), blur(em, 2.5) * 1.4)
    screen(c.img, C('#fff0c0'), em)
    bokeh(c, 12 + 5 * c.tier, (0, 0, W, H), 3, 16, ['#ffb0a0', '#ff8aa0', '#ffd8b0'])
    pts = [(x, y, rng.uniform(5, 11)) for x, y in c.pts(8 + 4 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, C('#fff4e8'), C('#ff7a6a'))


# ---------------------------------------------------------------- OR (golden treasure vault)
def coin_pile(c, L, cx, cy, w, h, n):
    rng = c.rng
    items = []
    for _ in range(n):
        u = rng.uniform(-1, 1)
        v = rng.uniform(0, 1)
        top = cy - h * (1 - u * u)
        y = lerp(top, cy + 10, v ** 1.5)
        x = cx + u * w
        items.append((y, x))
    items.sort()
    for y, x in items:
        r = rng.uniform(7, 12) * (0.7 + 0.5 * y / H)
        tilt = rng.uniform(0.35, 0.9)
        L.ellipse(x, y + r * 0.16, r, r * tilt, C('#a86a10'))
        L.ellipse(x, y, r, r * tilt, C('#ffc83a') * rng.uniform(0.9, 1.08))
        L.ellipse(x, y, r * 0.72, r * 0.72 * tilt, C('#ffdc6a'))
        L.ellipse(x - r * 0.2, y - r * 0.2 * tilt, r * 0.3, r * 0.18 * tilt, C('#fff6c8'))


def gem(L, x, y, r, col, rot=0.0):
    col = C(col)
    pts = [(x + r * math.cos(rot + k * math.pi / 3), y + r * 0.8 * math.sin(rot + k * math.pi / 3)) for k in range(6)]
    L.poly(pts, col * 0.7)
    L.poly([(x, y)] + pts[3:6], col)
    L.poly([(x, y), pts[0], pts[1], pts[2]], lerp(col, np.ones(3), 0.25))
    L.poly([(x - r * 0.3, y - r * 0.3), (x, y - r * 0.6), (x + r * 0.2, y - r * 0.25)], lerp(col, np.ones(3), 0.7))


def bg_or(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#5a3008'), (0.25, '#a8661a'), (0.5, '#f0b448'), (0.7, '#ffe090'), (1, '#c88a30')])
    c.light = (W * 0.5, -40)
    glow(c, W / 2, H * 0.1, 260, C('#fff8d0'), 0.7)
    god_rays(c, W / 2 + rng.uniform(-50, 50), -60, '#fff0b8', 0.45, cells=30, reach=800)
    # vault arches
    L = Layer()
    for k, side in enumerate((-1, 1)):
        for j in range(2):
            x = W / 2 + side * (150 + j * 80)
            L.poly([(x - 16, H * 0.68), (x + 16, H * 0.68), (x + 14, H * 0.05), (x - 14, H * 0.05)], C('#c8862a'))
            L.poly([(x - 20, H * 0.05), (x + 20, H * 0.05), (x + 20, H * 0.03), (x - 20, H * 0.03)], C('#c8862a'))
    for r in (210, 235):
        L.ellipse(W / 2, H * 0.32, r, r * 1.1, C('#b87a24'), 1.0, width=10)
    a = comp_layer(c.img, L, 0.6, 1.5)
    over(c.img, C('#ffe0a0'), a * 0.25)
    # halo / sunburst
    ang = np.arctan2(YY - S.cy, XX - S.cx)
    burst = (0.5 + 0.5 * np.cos(ang * 16)) ** 6
    d = np.sqrt((XX - S.cx) ** 2 + (YY - S.cy) ** 2)
    screen(c.img, C('#fff0b0'), burst * np.exp(-d / 260) * sstep(40, 140, d) * 0.4)
    # coin mountains
    L = Layer()
    coin_pile(c, L, W * 0.12, H * 0.66, 110, 120, 140)
    coin_pile(c, L, W * 0.88, H * 0.66, 110, 110, 140)
    coin_pile(c, L, W * 0.5, H * 0.7, 260, 50, 160)
    for _ in range(14):
        x, y = rng.uniform(0, W), H * rng.uniform(0.58, 0.75)
        gem(L, x, y, rng.uniform(6, 11), ['#ff4a8a', '#4ad8ff', '#7aff8a', '#b07aff', '#ff6a4a'][rng.integers(5)], rng.uniform(0, 1))
    a = comp_layer(c.img, L, 1.0, 0.8)
    over(c.img, C('#ffe8b0'), a * 0.15)
    glow(c, S.cx, S.cy, max(S.sw, S.sh) * 0.9, C('#fff0b0'), 0.15)


def fg_or(c):
    rng, S = c.rng, c.S
    L = Layer()
    coin_pile(c, L, -10, H * 0.9, 150, 110, 120)
    coin_pile(c, L, W + 10, H * 0.9, 150, 120, 120)
    coin_pile(c, L, W / 2, H * 0.94, 260, 40, 110)
    for _ in range(8 + 2 * c.tier):
        x = rng.choice([rng.uniform(0, W * 0.32), rng.uniform(W * 0.68, W)])
        y = H * rng.uniform(0.78, 0.98)
        gem(L, x, y, rng.uniform(10, 18), ['#ff4a8a', '#4ad8ff', '#7aff8a', '#b07aff', '#ff6a4a'][rng.integers(5)], rng.uniform(0, 1))
    a = comp_layer(c.img, L)
    add(c.img, C('#ffc850'), blur(a, 10) * 0.25 * (1 - a))
    bokeh(c, 18 + 6 * c.tier, (0, 0, W, H), 3, 18, ['#ffe090', '#fff4c8', '#ffc060'], amin=0.15, amax=0.45)
    pts = [(x, y, rng.uniform(5, 14)) for x, y in c.pts(18 + 6 * c.tier, 10, 40, W - 10, H * 0.95, avoid=0.5)]
    sparkles(c, pts, np.ones(3, np.float32), C('#ffc030'))


# ---------------------------------------------------------------- PIERRE (crystal cave)
def bg_pierre(c):
    rng, S = c.rng, c.S
    sky(c, [(0, '#1c1a34'), (0.4, '#3a3660'), (0.7, '#5a4e78'), (1, '#2a2440')])
    tex = fbm(rng, 4, 6)
    c.img *= (0.8 + 0.4 * tex)[..., None]
    # cave mouth behind subject
    mx, my = S.cx + rng.uniform(-40, 40), H * 0.3
    n = fbm(rng, 3, 4)
    d = np.sqrt(((XX - mx) / 150) ** 2 + ((YY - my) / 200) ** 2) + (n - 0.5) * 0.35
    mouth = sstep(1.0, 0.92, d)
    sky_in = lerp(C('#fff8f0'), C('#a8d0ff'), sstep(0.0, 1.0, d)[..., None])
    over(c.img, sky_in, mouth)
    c.light = (mx, my - 60)
    god_rays(c, mx, my - 100, '#fff4e8', 0.5, cells=26, reach=600, start=40)
    glow(c, mx, my, 220, C('#e0f0ff'), 0.5)
    # rock rim of mouth
    rim = sstep(1.08, 1.0, d) * (1 - mouth)
    screen(c.img, C('#c8b8f0'), rim * 0.5)
    # stalactites
    L = Layer()
    for _ in range(18):
        x = rng.uniform(-10, W + 10)
        ln = rng.uniform(40, 150) * (1.3 if abs(x - W / 2) > 150 else 0.7)
        w = rng.uniform(12, 30)
        L.poly([(x - w / 2, -5), (x + w / 2, -5), (x + rng.uniform(-4, 4), ln)], C('#221e3a') * rng.uniform(0.9, 1.3))
    comp_layer(c.img, L)
    # walls
    for side in (-1, 1):
        xs = np.arange(H, dtype=np.float32)
        wl = (W / 2 + side * (150 + (fbm1(rng, H, 3, 4) - 0.5) * 80 + 60 * (1 - xs / H)))
        mm = (XX - wl[:, None]) * side
        wm = np.clip(mm + 0.5, 0, 1)
        over(c.img, lerp(C('#3a3460'), C('#1a1830'), np.clip(mm / 120, 0, 1)[..., None]), wm)
        screen(c.img, C('#8a7ac8'), wm * np.exp(-np.clip(mm, 0, None) / 4) * 0.6)
    # glowing wall crystals
    L = Layer()
    pals = [('#2a6ab0', '#4ad0f0', '#b8f8ff', '#ffffff'), ('#6a2ab0', '#b06af0', '#e8c8ff', '#ffffff'),
            ('#b02a80', '#f06ac0', '#ffc8e8', '#ffffff')]
    for _ in range(5):
        side = rng.choice([-1, 1])
        x = W / 2 + side * rng.uniform(170, 240)
        y = rng.uniform(H * 0.2, H * 0.6)
        crystal_cluster(c, L, x, y, rng.uniform(30, 60), -side * rng.uniform(0.6, 1.4), pals[rng.integers(3)], n=4)
    a = comp_layer(c.img, L)
    add(c.img, C('#9ae0ff'), blur(a, 12) * 0.5 * (1 - a))
    # floor
    yl = H * 0.7 + (fbm1(rng, W, 3, 4) - 0.5) * 24
    fill_ridge(c, yl, '#5a5080', '#1e1a30', depth=200, rim='#c8c0ff', rims=0.5, tex=fbm(rng, 10, 3))
    glow(c, S.cx, S.foot, S.sw * 0.6, C('#c8e8ff'), 0.35, ry=36)


def fg_pierre(c):
    rng, S = c.rng, c.S
    L = Layer()
    for side in (-1, 1):
        bx = W / 2 + side * rng.uniform(W * 0.4, W * 0.52)
        by = H * rng.uniform(0.88, 0.96)
        for k in range(3):
            r = rng.uniform(50, 90)
            cx = bx - side * k * 40
            pts = []
            for j in range(9):
                a = math.pi + j * math.pi / 8
                rr = r * rng.uniform(0.75, 1.05)
                pts.append((cx + math.cos(a) * rr * 1.2, by + 25 * k + math.sin(a) * rr))
            L.poly(pts, C('#26223e') * rng.uniform(0.9, 1.3))
    a = comp_layer(c.img, L)
    edge = a * (1 - shift(a, 0, 4))
    screen(c.img, C('#b8a8ff'), blur(edge, 1.2) * 0.8)
    pals = [('#2a6ab0', '#4ad0f0', '#b8f8ff', '#ffffff'), ('#6a2ab0', '#b06af0', '#e8c8ff', '#ffffff'),
            ('#b02a80', '#f06ac0', '#ffc8e8', '#ffffff')]
    L = Layer()
    for side in (-1, 1):
        x = W / 2 + side * rng.uniform(W * 0.32, W * 0.44)
        crystal_cluster(c, L, x, H * rng.uniform(0.84, 0.9), rng.uniform(130, 190), side * 0.28, pals[rng.integers(3)], n=6)
        x = S.cx + side * S.sw * rng.uniform(0.3, 0.42)
        crystal_cluster(c, L, x, S.foot + rng.uniform(8, 20), rng.uniform(50, 80), side * 0.35, pals[rng.integers(3)], n=4)
    a = comp_layer(c.img, L)
    add(c.img, C('#a8e0ff'), blur(a, 12) * 0.35 * (1 - a))
    m = Mask()
    for (x, y) in c.pts(40, S.cx - 160, 0, S.cx + 160, H * 0.7, avoid=0.4):
        r = rng.uniform(0.6, 1.6)
        m.ellipse(x, y, r, r, rng.uniform(0.4, 1))
    dm = m.get()
    screen(c.img, C('#fff8e8'), dm + blur(dm, 2) * 0.6)
    bokeh(c, 12 + 5 * c.tier, (0, 0, W, H), 3, 16, ['#9af0ff', '#d0a0ff', '#ffb0e0'])
    pts = [(x, y, rng.uniform(5, 12)) for x, y in c.pts(8 + 4 * c.tier, 10, 40, W - 10, H * 0.9)]
    sparkles(c, pts, np.ones(3, np.float32), C('#9ae0ff'))


COVER = {
    'sacre': ('cloud', '#fffaf4', '#ffd8d4', '#b898d0'),
    'montagne': ('cloud', '#fff6f2', '#f0d4e4', '#9890c8'),
    'glace': ('cloud', '#ffffff', '#d4e8ff', '#7e9ad8'),
    'nuit': ('cloud', '#e0d0ff', '#9480d8', '#3a2c80'),
    'arcane': ('cloud', '#ffd4f6', '#b688e8', '#4a2a96'),
    'feu': ('cloud', '#ffd8a8', '#e8806a', '#6a1a3a'),
    'sang': ('cloud', '#ffd0b8', '#d86a78', '#5a1030'),
    'ocean': ('cloud', '#c8fff6', '#5ac8dc', '#1a6a9a'),
    'foret': ('bush', '#c8f49a', '#4aa060', '#1a5040'),
    'plaine': ('bush', '#d0f4a0', '#5ab060', '#2a6a40'),
    'pierre': ('rock', '#b0a8d8', '#5a5488', '#26223e'),
    'or': ('coins', '', '', ''),
}


def cover_cut(c):
    """Foreground mass hiding the flat bottom edge of bust-type subjects."""
    S = c.S
    if not S.cut:
        return
    kind, lit, mid, shade = COVER.get(c.scene, COVER['sacre'])
    rng = c.rng
    lxs = 1 if c.light[0] > S.cx else -1
    if kind == 'coins':
        L = Layer()
        coin_pile(c, L, S.cx, S.foot + 40, S.sw * 0.75, S.sh * 0.32, 170)
        comp_layer(c.img, L)
        return
    w = S.sw * 1.4
    base = S.foot + S.sh * 0.12
    blobs = cumulus(c, S.cx, base, w, S.sh * 0.19, n=int(14 + w / 14))
    if kind == 'bush':
        blobs = [(x, y, r * 0.75) for x, y, r in blobs] + [(x + rng.uniform(-8, 8), y - r * 0.4, r * 0.5) for x, y, r in blobs]
        blobs.sort(key=lambda b: b[1])
        puffs(c, blobs, lit, mid, shade, light=(lxs * 0.5, -1), soft=2.5, rim=lit, rims=0.5, style='bush')
        L = Layer()
        for x, y, r in blobs[::2]:
            for k in range(3):
                a = rng.uniform(-2.6, -0.5)
                L.poly(petal_pts(x + rng.uniform(-r, r) * 0.6, y - r * 0.5, r * 0.6, r * 0.22, a), C(lit) * rng.uniform(0.8, 1.05), 0.9)
        comp_layer(c.img, L)
    elif kind == 'rock':
        puffs(c, blobs, lit, mid, shade, light=(lxs * 0.6, -1), soft=1.5, rim='#d8d0ff', rims=0.6, style='rock')
    else:
        puffs(c, blobs, lit, mid, shade, light=(lxs * 0.4, -1), soft=5, rim='#ffffff', rims=0.8)


SCENES = {
    'arcane': (bg_arcane, fg_arcane), 'feu': (bg_feu, fg_feu), 'glace': (bg_glace, fg_glace),
    'sacre': (bg_sacre, fg_sacre), 'montagne': (bg_montagne, fg_montagne), 'plaine': (bg_plaine, fg_plaine),
    'ocean': (bg_ocean, fg_ocean), 'foret': (bg_foret, fg_foret), 'nuit': (bg_nuit, fg_nuit),
    'sang': (bg_sang, fg_sang), 'or': (bg_or, fg_or), 'pierre': (bg_pierre, fg_pierre),
}


# =============================================================== spell energy
def spell_fx(c, front=False):
    S = c.S
    rng = c.rng
    if not front:
        # energy vortex behind
        ang = np.arctan2(YY - S.cy, XX - S.cx)
        d = np.sqrt((XX - S.cx) ** 2 + (YY - S.cy) ** 2) + 1e-3
        sw = 0.5 + 0.5 * np.sin(ang * 3 + np.log(d) * 5)
        env = np.exp(-((d - S.sw * 0.45) / (S.sw * 0.3)) ** 2)
        add(c.img, lerp(c.el[2], c.el[1], 0.5), sw ** 3 * env * 0.45)
    else:
        m = Mask()
        for k in range(int(10 + 4 * c.tier)):
            a0 = rng.uniform(0, 2 * math.pi)
            r0 = S.sw * rng.uniform(0.38, 0.6)
            pts = []
            for j in range(16):
                t = j / 15
                a = a0 + t * rng.uniform(0.5, 0.9)
                r = r0 * (1 + t * 0.25)
                pts.append((S.cx + math.cos(a) * r, S.cy + math.sin(a) * r * 0.9))
            m.line(pts, rng.uniform(1.0, 2.2))
        em = m.get()
        add(c.img, c.el[2], blur(em, 3) * 1.2)
        screen(c.img, lerp(c.el[0], c.el[1], 0.3), em)


# =============================================================== final grade
def grade(c):
    img = c.img
    P = c.P
    np.clip(img, 0, 4, out=img)
    br = np.clip(img.max(-1, keepdims=True) - 0.8, 0, None) * img
    bloom = blur(br, 4) * 0.35 + blur(br, 16) * 0.4
    img += bloom
    # soft shoulder highlight compression
    k = 0.82
    hi = img > k
    img[hi] = k + (1 - k) * (1 - np.exp(-(img[hi] - k) / (1 - k)))
    # vibrance (pastel-safe)
    lum = img.mean(-1, keepdims=True)
    sat = img.max(-1, keepdims=True) - img.min(-1, keepdims=True)
    img[:] = lum + (img - lum) * (1.0 + 0.18 * (1 - sat))
    # gentle S-curve for a crisp, polished look
    img[:] = np.clip(img, 0, 1)
    img[:] = img + 0.22 * img * (1 - img) * (img - 0.45) * 4 * 0.5
    # tinted vignette (never black)
    v = ((XX - W / 2) / (W * 0.62)) ** 2 + ((YY - H * 0.45) / (H * 0.72)) ** 2
    over(img, C(P['shadow']), sstep(0.55, 1.6, v) * 0.45)
    # light leaks
    light_leaks(c, P['leak'], 0.14)
    # holographic shimmer for legendaries
    if c.tier >= 3:
        hue = (XX * 0.9 + YY * 0.5) / 140.0 + fbm(c.rng, 2, 2) * 1.5
        rainbow = np.stack([0.5 + 0.5 * np.cos(2 * np.pi * (hue + o)) for o in (0, 1 / 3, 2 / 3)], -1)
        lum = img.mean(-1)
        screen(img, rainbow, sstep(0.45, 0.95, lum) * 0.07)
    np.clip(img, 0, 1, out=img)


# =============================================================== render
def render(card):
    c = Ctx(card)
    prep_subject(c)
    bg, fg = SCENES.get(c.scene, SCENES['plaine'])
    bg(c)
    if c.type == 'sort':
        spell_fx(c, front=False)
    subject_backdrop(c)
    nrib = [0, 1, 1, 2][c.tier] + (1 if c.type in ('sort', 'heros') else 0)
    mb = mf = None
    if nrib:
        mb, mf = ribbon_masks(c, min(nrib, 2), width=3.2 + 0.5 * c.tier)
        draw_ribbon(c, mb, 0.85)
    subject_ground(c)
    composite_subject(c)
    cover_cut(c)
    if mf is not None:
        draw_ribbon(c, mf, 0.7)
    if c.type == 'sort':
        spell_fx(c, front=True)
    fg(c)
    grade(c)
    out = np.clip(c.img * 255 + c.rng.uniform(-0.6, 0.6, c.img.shape), 0, 255).astype(np.uint8)
    return Image.fromarray(out, 'RGB')


def save_webp(im, path):
    q = 82
    for _ in range(4):
        im.save(path, 'WEBP', quality=q, method=6)
        sz = os.path.getsize(path)
        if sz > 72000 and q > 70:
            q -= 4
        elif sz < 28000 and q < 92:
            q += 4
        else:
            break
    return os.path.getsize(path), q


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    cards = json.load(open(sys.argv[1], encoding='utf-8'))
    outdir = sys.argv[2]
    ids = set(sys.argv[3:])
    os.makedirs(outdir, exist_ok=True)
    for card in cards:
        if ids and card['id'] not in ids:
            continue
        t = time.time()
        im = render(card)
        sz, q = save_webp(im, os.path.join(outdir, card['id'] + '.webp'))
        print('%-14s %-9s %-11s %5.1f KB q%d %.2fs' % (card['id'], card['scene'], card['rarity'], sz / 1024, q, time.time() - t))


if __name__ == '__main__':
    main()

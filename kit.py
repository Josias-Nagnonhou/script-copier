"""Kopi — kit de vente d'un ebook : mockup 3D, visuel 16:9 pour la boutique, visuel carré."""
import os, re, random
import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageEnhance, ImageOps

HERE = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(HERE, "fonts")

DISPLAY = {"moderne": "Poppins-ExtraBold", "impact": "BebasNeue-Regular", "sacre": "Marcellus-Regular",
           "luxe": "DMSerifDisplay-Regular", "doux": "DMSerifDisplay-Regular", "editorial": "AbrilFatface-Regular"}
BODY = {"moderne": ("Poppins-Medium", "Poppins-SemiBold"), "impact": ("Poppins-Medium", "Poppins-Bold"),
        "sacre": ("Spectral-Regular", "Spectral-SemiBold"), "luxe": ("Lato-Regular", "Lato-Bold"),
        "doux": ("Lato-Regular", "Lato-Bold"), "editorial": ("Spectral-Regular", "Spectral-SemiBold")}


def font(name, size):
    return ImageFont.truetype(os.path.join(FONT_DIR, name + ".ttf"), int(size))


def rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


# ---------------------------------------------------------------- couleurs d'une couverture importée
def palette_from_image(path):
    im = Image.open(path).convert("RGB").resize((120, 170))
    q = im.quantize(colors=8, method=Image.Quantize.MEDIANCUT).convert("RGB")
    cols = sorted(q.getcolors(120 * 170), reverse=True)
    def sat(c):
        mx, mn = max(c), min(c)
        return (mx - mn) / (mx + 1)
    def lum(c):
        return .2126 * c[0] + .7152 * c[1] + .0722 * c[2]
    cs = [c for _, c in cols]
    main = max(cs, key=lambda c: sat(c) * (1 if 50 < lum(c) < 210 else .3))
    dark = min(cs, key=lum)
    dark = tuple(int(v * .45) for v in dark) if lum(dark) > 40 else dark
    accent = max((c for c in cs if c != main), key=lambda c: sat(c) + lum(c) / 400, default=(240, 180, 40))
    light = tuple(int(v + (255 - v) * .9) for v in main)
    hx = lambda c: "#%02X%02X%02X" % c
    return {"main": hx(main), "dark": hx(dark), "light": hx(light), "accent": hx(accent)}


# ---------------------------------------------------------------- mockup 3D
def _coeffs(dst, src):
    """Coefficients PIL PERSPECTIVE : point de sortie (dst) → point source (src)."""
    A, B = [], []
    for (x, y), (u, v) in zip(dst, src):
        A.append([x, y, 1, 0, 0, 0, -u * x, -u * y]); B.append(u)
        A.append([0, 0, 0, x, y, 1, -v * x, -v * y]); B.append(v)
    return np.linalg.solve(np.array(A, float), np.array(B, float)).tolist()


def _warp(img, quad, size):
    w, h = img.size
    src = [(0, 0), (w, 0), (w, h), (0, h)]
    return img.transform(size, Image.PERSPECTIVE, _coeffs(quad, src), Image.BICUBIC)


def _rmask(size, r, ss=4):
    """Masque arrondi anti-crénelé."""
    w, h = size
    m = Image.new("L", (w * ss, h * ss), 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, w * ss - 1, h * ss - 1], radius=int(r * ss), fill=255)
    return m.resize((w, h), Image.LANCZOS)


def _fit(img, size):
    """Recadre l'image pour remplir exactement « size »."""
    return ImageOps.fit(img.convert("RGB"), size, Image.LANCZOS, centering=(.5, .3))


def _gloss(size, strength=46, start=.18, end=.55):
    """Reflet diagonal doux (vitre / pelliculage)."""
    w, h = size
    g = Image.new("L", (w, h), 0)
    ImageDraw.Draw(g).polygon([(w * start, 0), (w * end, 0), (w * (end - .42), h), (w * (start - .42), h)], fill=strength)
    g = g.filter(ImageFilter.GaussianBlur(max(2, w * .05)))
    out = Image.new("RGBA", (w, h), (255, 255, 255, 0))
    out.putalpha(g)
    return out


def book_mockup(cover_path, height):
    """Livre relié vu de face, légèrement tourné : couverture pelliculée, charnière, tranche de pages."""
    cover = Image.open(cover_path).convert("RGB")
    ratio = cover.width / cover.height if cover.height else .707
    H = int(height)
    W = int(H * ratio)
    face = cover.resize((W, H), Image.LANCZOS).convert("RGBA")
    fx = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    g = ImageDraw.Draw(fx)
    # dos arrondi : liseré clair puis ombre de charnière
    sp = int(W * .055)
    for i in range(sp):
        t = i / sp
        a = int(70 * (1 - t) ** 2)
        g.line([(i, 0), (i, H)], fill=(255, 255, 255, a) if i < sp * .25 else (0, 0, 0, int(60 * (1 - t))))
    hx = int(W * .062)
    for k in range(6):
        g.line([(hx + k, 0), (hx + k, H)], fill=(0, 0, 0, 70 - k * 11))
    g.line([(hx + 7, 0), (hx + 7, H)], fill=(255, 255, 255, 42), width=2)
    # léger assombrissement vers le bord qui s'éloigne
    for i in range(int(W * .35)):
        x = W - 1 - i
        g.line([(x, 0), (x, H)], fill=(0, 0, 0, int(38 * (1 - i / (W * .35)) ** 1.5)))
    face.alpha_composite(fx)
    face.alpha_composite(_gloss((W, H), 40))
    face.putalpha(_rmask((W, H), max(3, W * .012)))
    # géométrie : la couverture tourne un peu, on voit la tranche de pages à droite
    th = int(W * .075)           # épaisseur visible
    dy = int(H * .018)           # fuite de perspective
    cw, ch = W + th + 6, H + 6
    out = Image.new("RGBA", (cw, ch), (0, 0, 0, 0))
    face_q = [(0, 0), (W, dy), (W, H - dy), (0, H)]
    # tranche des pages (avec les cartons de la couverture qui débordent)
    pg = Image.new("RGBA", (th * 3, H), (250, 247, 240, 255))
    d = ImageDraw.Draw(pg)
    for y in range(0, H, 2):
        c = 236 - (y // 2 % 3) * 5
        d.line([(0, y), (pg.width, y)], fill=(c, c - 3, c - 9, 255))
    sh = Image.new("RGBA", pg.size, (0, 0, 0, 0))
    ImageDraw.Draw(sh).rectangle([0, 0, int(pg.width * .5), H], fill=(0, 0, 0, 30))
    pg.alpha_composite(sh.filter(ImageFilter.GaussianBlur(pg.width * .2)))
    edge = tuple(int(v * .62) for v in cover.getpixel((cover.width - 2, cover.height // 2)))
    board = int(H * .012)
    pages_q = [(W - 1, dy + board), (W + th, dy * 2 + board), (W + th, H - dy * 2 - board), (W - 1, H - dy - board)]
    back_q = [(W - 1, dy), (W + th, dy * 2), (W + th, H - dy * 2), (W - 1, H - dy)]
    bk = Image.new("RGBA", (8, 8), edge + (255,))
    out.alpha_composite(_warp(bk, back_q, (cw, ch)))
    out.alpha_composite(_warp(pg, pages_q, (cw, ch)))
    out.alpha_composite(_warp(face, face_q, (cw, ch)))
    return out


def readable(c, min_l=120):
    """Éclaircit une couleur trop sombre pour rester lisible sur fond sombre."""
    l = .2126 * c[0] + .7152 * c[1] + .0722 * c[2]
    if l >= min_l:
        return c
    t = min(.85, (min_l - l) / (255 - l + 1) * 1.4)
    return tuple(int(v + (255 - v) * t) for v in c)


def _rot(im, angle):
    return im.rotate(angle, resample=Image.BICUBIC, expand=True) if angle else im


def place(canvas, im, cx, cy, angle=0, shadow=150, lift=1.0):
    """Pose un élément (centré sur cx, cy) avec rotation et ombre portée douce en deux couches."""
    im = _rot(im, angle)
    w, h = im.size
    a = im.split()[3]
    pad = int(max(w, h) * .25)
    for blur, op, off in ((.05, .55, .045), (.014, .45, .012)):
        s = Image.new("L", (w + 2 * pad, h + 2 * pad), 0)
        s.paste(a.point(lambda v, op=op: int(v * shadow * op / 255)), (pad, pad))
        s = s.filter(ImageFilter.GaussianBlur(max(w, h) * blur * lift))
        layer = Image.new("RGBA", s.size, (8, 6, 14, 0))
        layer.putalpha(s)
        canvas.alpha_composite(layer, (int(cx - w / 2 - pad + h * off * .3 * lift), int(cy - h / 2 - pad + h * off * lift)))
    canvas.alpha_composite(im, (int(cx - w / 2), int(cy - h / 2)))


def page_card(page, w):
    """Page intérieure du PDF en carte blanche, coins légèrement arrondis."""
    h = int(w * page.height / page.width)
    im = page.convert("RGB").resize((w, h), Image.LANCZOS).convert("RGBA")
    im.alpha_composite(_gloss((w, h), 18, .3, .6))
    im.putalpha(_rmask((w, h), w * .018))
    return im


def tablet(screen, w):
    """Tablette moderne (bords fins, cadre alu) affichant « screen »."""
    h = int(w * 1.38)
    out = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    R = w * .075
    rim = Image.new("RGBA", (w, h), (196, 198, 204, 255))
    rg = Image.new("L", (1, h))
    for y in range(h):
        rg.putpixel((0, y), int(150 + 90 * abs(.5 - y / h) * 2))
    rim = Image.merge("RGBA", (rg.resize((w, h)),) * 3 + (Image.new("L", (w, h), 255),))
    out.paste(rim, (0, 0), _rmask((w, h), R))
    b = max(3, int(w * .008))
    out.paste((18, 18, 22, 255), (b, b), _rmask((w - 2 * b, h - 2 * b), R - b))
    bz = int(w * .045)
    sw, sh_ = w - 2 * bz, h - 2 * bz
    scr = _fit(screen, (sw, sh_)).convert("RGBA")
    scr.alpha_composite(_gloss((sw, sh_), 34))
    out.paste(scr, (bz, bz), _rmask((sw, sh_), R * .55))
    d = ImageDraw.Draw(out)
    cam = max(3, int(w * .008))
    d.ellipse([w / 2 - cam, bz / 2 - cam, w / 2 + cam, bz / 2 + cam], fill=(40, 42, 50, 255))
    return out


def phone(cover, w, accent=(255, 255, 255)):
    """Téléphone : fond flouté de la couverture, la couverture au centre, barre de lecture."""
    h = int(w * 2.06)
    out = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    R = w * .17
    out.paste((28, 28, 32, 255), (0, 0), _rmask((w, h), R))
    b = max(2, int(w * .012))
    out.paste((8, 8, 10, 255), (b, b), _rmask((w - 2 * b, h - 2 * b), R - b))
    bz = int(w * .04)
    sw, sh_ = w - 2 * bz, h - 2 * bz
    cv = cover.convert("RGB")
    bg = _fit(cv, (sw, sh_)).filter(ImageFilter.GaussianBlur(sw * .08))
    bg = ImageEnhance.Brightness(bg).enhance(.55).convert("RGBA")
    cwid = int(sw * .78)
    ch = int(cwid * cv.height / cv.width)
    cimg = cv.resize((cwid, ch), Image.LANCZOS).convert("RGBA")
    cimg.putalpha(_rmask((cwid, ch), cwid * .02))
    place(bg, cimg, sw / 2, sh_ * .44, 0, shadow=200)
    d = ImageDraw.Draw(bg)
    y = int(sh_ * .44 + ch / 2 + sh_ * .07)
    x0, x1 = int(sw * .11), int(sw * .89)
    d.rounded_rectangle([x0, y, x1, y + max(4, sw // 70)], radius=4, fill=(255, 255, 255, 70))
    d.rounded_rectangle([x0, y, x0 + int((x1 - x0) * .38), y + max(4, sw // 70)], radius=4, fill=accent + (255,))
    bg.alpha_composite(_gloss((sw, sh_), 30))
    out.paste(bg, (bz, bz), _rmask((sw, sh_), R * .8))
    d = ImageDraw.Draw(out)
    iw, ih = int(w * .3), int(w * .085)
    d.rounded_rectangle([w / 2 - iw / 2, bz + w * .035, w / 2 + iw / 2, bz + w * .035 + ih], radius=ih / 2, fill=(0, 0, 0, 255))
    return out


def badge(size, top, bottom, col, fname, angle=-12):
    """Pastille ronde « PDF · 120 pages »."""
    S = int(size)
    ss = 3
    im = Image.new("RGBA", (S * ss, S * ss), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.ellipse([0, 0, S * ss - 1, S * ss - 1], fill=(255, 255, 255, 255))
    m = int(S * ss * .045)
    d.ellipse([m, m, S * ss - m, S * ss - m], fill=col + (255,))
    tc = (20, 20, 20) if sum(col) > 400 else (255, 255, 255)
    d.text((S * ss / 2, S * ss * .45), top, font=font(fname, S * ss * .3), fill=tc, anchor="mm")
    d.text((S * ss / 2, S * ss * .7), bottom, font=font("Poppins-SemiBold", S * ss * .11), fill=tc, anchor="mm")
    im = im.resize((S, S), Image.LANCZOS)
    return _rot(im, angle)


def mesh_bg(size, pal, dark=True, seed=7):
    """Fond « mesh gradient » aux couleurs du livre, avec grain et vignette."""
    W, H = size
    rnd = random.Random(seed)
    base = rgb(pal["dark"]) if dark else tuple(int(v + (255 - v) * .55) for v in rgb(pal["light"]))
    if dark:
        base = tuple(max(8, int(v * .7)) for v in base)
    sc = 8
    w, h = W // sc, H // sc
    im = Image.new("RGB", (w, h), base)
    lay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(lay)
    main, acc, light = rgb(pal["main"]), rgb(pal["accent"]), rgb(pal["light"])
    blobs = [(main, .62, .45, .55, 170 if dark else 120), (acc, .9, .15, .32, 120 if dark else 90),
             (main, .1, .95, .4, 110 if dark else 80), (light if not dark else acc, .35, .05, .3, 60 if dark else 120)]
    for col, x, y, r, a in blobs:
        x += rnd.uniform(-.05, .05); y += rnd.uniform(-.05, .05)
        R = r * max(w, h)
        d.ellipse([x * w - R, y * h - R, x * w + R, y * h + R], fill=col + (a,))
    lay = lay.filter(ImageFilter.GaussianBlur(max(w, h) * .12))
    im = im.convert("RGBA"); im.alpha_composite(lay)
    im = im.convert("RGB").resize((W, H), Image.BICUBIC)
    arr = np.asarray(im).astype(np.float32)
    yy, xx = np.mgrid[0:H, 0:W]
    v = ((xx - W * .6) / W) ** 2 + ((yy - H * .5) / H) ** 2
    arr *= (1 - np.clip(v * (.9 if dark else .12), 0, .5))[..., None]
    arr += np.random.default_rng(seed).normal(0, 5.5, (H, W, 1)).astype(np.float32)
    return Image.fromarray(np.clip(arr, 0, 255).astype(np.uint8)).convert("RGBA")


def pdf_pages(pdf_path, want=3):
    """Choisit les pages intérieures les plus « visuelles » (titres, images, encadrés)."""
    try:
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(pdf_path)
        n = len(pdf)
        scored = []
        for i in range(1, min(n, 12)):
            th = pdf[i].render(scale=.25).to_pil().convert("RGB")
            a = np.asarray(th).astype(np.float32)
            ink = 1 - a.mean() / 255
            color = a.std(axis=2).mean() / 60
            if ink < .03:
                continue
            scored.append((ink + color, i))
        pick = sorted(i for _, i in sorted(scored, reverse=True)[:want])
        pages = [pdf[i].render(scale=1.6).to_pil().convert("RGB") for i in pick]
        pdf.close()
        best = max(range(len(pick)), key=lambda k: dict((i, s) for s, i in scored)[pick[k]]) if pick else 0
        return pages, n, best
    except Exception:
        return [], 0, 0


# ---------------------------------------------------------------- texte
def wrap(draw, text, fnt, maxw):
    try:
        from ebook import clean
        text = clean(text)
    except Exception:
        pass
    words, lines, cur = text.split(), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if draw.textlength(t, font=fnt) <= maxw or not cur:
            cur = t
        else:
            lines.append(cur); cur = w
    if cur:
        lines.append(cur)
    return lines


def fit_title(draw, text, fname, maxw, maxh, start, upper=False):
    text = text.upper() if upper else text
    size = start
    while size > 30:
        f = font(fname, size)
        lines = wrap(draw, text, f, maxw)
        lh = size * (.95 if upper else 1.08)
        if len(lines) * lh <= maxh and len(lines) <= 4:
            return f, lines, lh
        size -= 4
    f = font(fname, 30)
    return f, wrap(draw, text, f, maxw)[:4], 32


def _unused_bg_from_cover(cover, size, pal, dark=True):
    W, H = size
    im = Image.open(cover).convert("RGB")
    sc = max(W / im.width, H / im.height) * 1.15
    im = im.resize((int(im.width * sc), int(im.height * sc)), Image.LANCZOS)
    im = im.crop(((im.width - W) // 2, (im.height - H) // 2, (im.width - W) // 2 + W, (im.height - H) // 2 + H))
    im = im.filter(ImageFilter.GaussianBlur(W * .03))
    over = Image.new("RGB", size, rgb(pal["dark"] if dark else pal["light"]))
    return Image.blend(im, over, .78 if dark else .82).convert("RGBA")


def _scene(canvas, cover, pages, best, cx, cy, scale, pal, uni, n_pages, label="PDF", compact=False):
    """Compose la scène produit : pages en éventail, tablette, livre relié, téléphone, pastille."""
    cov = Image.open(cover).convert("RGB")
    acc = readable(rgb(pal["accent"]), 95)
    S = scale
    others = [p for k, p in enumerate(pages) if k != best]
    # pages intérieures en éventail derrière
    for k, (pg, ang, dx, dy) in enumerate(zip(others[:2], (-14, -5), (.34, .25), (-.09, -.12))):
        place(canvas, page_card(pg, int(S * .40)), cx + S * dx, cy + S * dy, ang, shadow=120)
    # tablette
    scr = pages[best] if pages else cov
    place(canvas, tablet(scr, int(S * .46)), cx - S * .33, cy - S * .05, 6, shadow=170)
    # livre relié
    book = book_mockup(cover, S * .78)
    place(canvas, book, cx + S * .03, cy + S * .04, 0, shadow=210, lift=1.2)
    # téléphone
    ph = phone(cov, int(S * .22), acc)
    place(canvas, ph, cx + S * .33, cy + S * .2, -7, shadow=200)
    # pastille
    if n_pages:
        bd = badge(S * .17, label, f"{n_pages} pages", acc, DISPLAY.get("impact"), -12)
        place(canvas, bd, cx - S * .39, cy + S * .33, 0, shadow=110)


def wide_visual(cover, out, info, pal, uni="moderne", pdf=None):
    """Visuel 16:9 pour la boutique : scène produit à droite, titre + promesse + points forts + prix à gauche."""
    W, H = 1920, 1080
    canvas = mesh_bg((W, H), pal, dark=True)
    pages, n, best = pdf_pages(pdf) if pdf else ([], 0, 0)
    n = info.get("pages") or n
    _scene(canvas, cover, pages, best, W * .755, H * .52, H * .84, pal, uni, n)
    d = ImageDraw.Draw(canvas)
    acc, main = readable(rgb(pal["accent"])), rgb(pal["main"])
    x, maxw = 110, int(W * .385)
    body, bold = BODY.get(uni, BODY["moderne"])
    upper = uni == "impact"
    y = 150
    if info.get("author"):
        d.text((x, 86), info["author"].upper(), font=font(bold, 22), fill=(255, 255, 255, 190))
    label = (info.get("label") or "Ebook").upper()
    lf = font(bold, 24)
    lw = d.textlength(label, font=lf) + 46
    d.rounded_rectangle([x, y, x + lw, y + 50], radius=25, fill=acc)
    tc = (20, 20, 20) if sum(acc) > 380 else (255, 255, 255)
    d.text((x + 23, y + 25), label, font=lf, fill=tc, anchor="lm")
    y += 86
    tf, lines, lh = fit_title(d, info["title"], DISPLAY.get(uni, DISPLAY["moderne"]), maxw, 320, 100 if upper else 80, upper)
    for ln in lines:
        d.text((x, y), ln, font=tf, fill=(255, 255, 255)); y += lh
    y += 14
    if info.get("subtitle"):
        sf = font(body, 31)
        for ln in wrap(d, info["subtitle"], sf, maxw)[:3]:
            d.text((x, y), ln, font=sf, fill=(226, 224, 230)); y += 43
    y += 28
    bf = font(bold, 29)
    for b in (info.get("points") or [])[:3]:
        cy = y + 20
        d.ellipse([x, cy - 16, x + 32, cy + 16], fill=acc)
        d.line([(x + 9, cy), (x + 14, cy + 6), (x + 24, cy - 6)], fill=tc, width=4, joint="curve")
        ls = wrap(d, b, bf, maxw - 54)[:2]
        for i, ln in enumerate(ls):
            d.text((x + 50, y + i * 38), ln, font=bf, fill=(255, 255, 255))
        y += 38 * len(ls) + 20
    y = max(y + 14, H - 200)
    cta = "Disponible maintenant"
    cf = font(bold, 32)
    cwid = d.textlength(cta, font=cf) + 120
    d.rounded_rectangle([x, y, x + cwid, y + 82], radius=41, fill=(255, 255, 255))
    dark_txt = main if sum(main) < 520 else (20, 20, 20)
    d.text((x + 40, y + 41), cta, font=cf, fill=dark_txt, anchor="lm")
    ax = x + cwid - 46
    d.ellipse([ax - 22, y + 19, ax + 22, y + 63], fill=acc)
    d.line([(ax - 9, y + 41), (ax + 9, y + 41)], fill=tc, width=4)
    d.line([(ax + 2, y + 33), (ax + 10, y + 41), (ax + 2, y + 49)], fill=tc, width=4, joint="curve")
    if info.get("price"):
        pf = font(DISPLAY.get(uni, "Poppins-ExtraBold") if uni != "impact" else "BebasNeue-Regular", 60 if uni == "impact" else 46)
        d.text((x + cwid + 36, y + 41), info["price"], font=pf, fill=(255, 255, 255), anchor="lm")
    canvas.convert("RGB").save(out, "JPEG", quality=92, optimize=True)


def square_visual(cover, out, pal, uni="moderne", pdf=None, n_pages=0):
    """Visuel carré : la scène produit seule, centrée, sur un fond dégradé moderne."""
    S = 1600
    dark = uni in ("luxe", "sacre", "impact", "moderne")
    canvas = mesh_bg((S, S), pal, dark=dark, seed=11)
    pages, n, best = pdf_pages(pdf) if pdf else ([], 0, 0)
    _scene(canvas, cover, pages, best, S * .52, S * .52, S * .76, pal, uni, n_pages or n)
    canvas.convert("RGB").save(out, "JPEG", quality=93, optimize=True)


def story_visual(cover, out, info, pal, uni="moderne", pdf=None, n_pages=0):
    """Visuel 9:16 (story TikTok / Instagram / Facebook) : accroche, scène produit, bénéfices, « lien dans la bio »."""
    W, H = 1080, 1920
    canvas = mesh_bg((W, H), pal, dark=True, seed=23)
    pages, n, best = pdf_pages(pdf) if pdf else ([], 0, 0)
    n = n_pages or n
    d = ImageDraw.Draw(canvas)
    acc, main = readable(rgb(pal["accent"])), rgb(pal["main"])
    tc = (20, 20, 20) if sum(acc) > 380 else (255, 255, 255)
    body, bold = BODY.get(uni, BODY["moderne"])
    upper = uni == "impact"
    # zone haute (sous la barre de la story) : étiquette + accroche
    y = 230
    label = (info.get("label") or "Nouvel ebook").upper()
    lf = font(bold, 30)
    lw = d.textlength(label, font=lf) + 56
    d.rounded_rectangle([(W - lw) / 2, y, (W + lw) / 2, y + 60], radius=30, fill=acc)
    d.text((W / 2, y + 30), label, font=lf, fill=tc, anchor="mm")
    y += 96
    hook = (info.get("hook") or info.get("title") or "").strip()
    tf, lines, lh = fit_title(d, hook, DISPLAY.get(uni, DISPLAY["moderne"]), W - 150, 300, 96 if upper else 74, upper)
    for ln in lines[:4]:
        d.text((W / 2, y), ln, font=tf, fill=(255, 255, 255), anchor="ma"); y += lh
    top_end = y
    # scène produit
    scene_cy = top_end + 395
    _scene(canvas, cover, pages, best, W * .51, scene_cy, W * .76, pal, uni, n)
    d = ImageDraw.Draw(canvas)
    # bénéfices
    y = scene_cy + W * .76 * .46
    bf = font(bold, 36)
    pts = (info.get("points") or [])[:3]
    if pts:
        widths = [min(W - 220, d.textlength(p, font=bf)) for p in pts]
        x = int((W - (max(widths) + 56)) / 2)
        for p in pts:
            cy = y + 24
            d.ellipse([x, cy - 19, x + 38, cy + 19], fill=acc)
            d.line([(x + 10, cy), (x + 16, cy + 7), (x + 28, cy - 7)], fill=tc, width=5, joint="curve")
            ln = wrap(d, p, bf, W - 220)[0]
            d.text((x + 56, y + 2), ln, font=bf, fill=(255, 255, 255))
            y += 62
    # prix
    if info.get("price"):
        y += 18
        pf = font("BebasNeue-Regular" if upper else DISPLAY.get(uni, "Poppins-ExtraBold"), 84 if upper else 64)
        d.text((W / 2, y), info["price"], font=pf, fill=(255, 255, 255), anchor="ma")
        y += 84 if upper else 78
    # appel à l'action « lien dans la bio » (au-dessus de la zone des boutons de l'appli)
    cy0 = max(y + 30, H - 500)
    cta = "LIEN DANS LA BIO"
    cf = font(bold, 46)
    cw = d.textlength(cta, font=cf) + 200
    x0, x1 = (W - cw) / 2, (W + cw) / 2
    glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    ImageDraw.Draw(glow).rounded_rectangle([x0 - 10, cy0 - 10, x1 + 10, cy0 + 120], radius=70, fill=acc + (150,))
    canvas.alpha_composite(glow.filter(ImageFilter.GaussianBlur(26)))
    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle([x0, cy0, x1, cy0 + 110], radius=55, fill=(255, 255, 255))
    ink = main if sum(main) < 520 else (20, 20, 20)
    ax, ay = x0 + 66, cy0 + 55
    dot = acc if sum(acc) < 600 else main
    tc = (20, 20, 20) if sum(dot) > 380 else (255, 255, 255)
    d.ellipse([ax - 34, ay - 34, ax + 34, ay + 34], fill=dot)
    d.line([(ax, ay + 17), (ax, ay - 17)], fill=tc, width=6)
    d.line([(ax - 13, ay - 4), (ax, ay - 18), (ax + 13, ay - 4)], fill=tc, width=6, joint="curve")
    d.text((ax + 52, ay), cta, font=cf, fill=ink, anchor="lm")
    sub = info.get("cta_sub") or "Clique et télécharge-le tout de suite"
    d.text((W / 2, cy0 + 146), sub, font=font(body, 32), fill=(236, 234, 240), anchor="ma")
    canvas.convert("RGB").save(out, "JPEG", quality=92, optimize=True)


def cover_from_upload(path, out):
    """Image ou PDF importé → image de couverture JPEG."""
    if path.lower().endswith(".pdf"):
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(path)
        pdf[0].render(scale=2).to_pil().convert("RGB").save(out, "JPEG", quality=92)
        pdf.close()
    else:
        im = Image.open(path)
        im = ImageOps.exif_transpose(im).convert("RGB")
        im.save(out, "JPEG", quality=92)
    return out


def pdf_text(path, max_chars=5000):
    try:
        from pypdf import PdfReader
        r = PdfReader(path)
        txt = " ".join((p.extract_text() or "") for p in r.pages[:12])
        return re.sub(r"\s+", " ", txt)[:max_chars]
    except Exception:
        return ""

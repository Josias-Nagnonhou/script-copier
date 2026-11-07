"""Kopi — kit de vente d'un ebook : mockup 3D, visuel 16:9 pour la boutique, visuel carré."""
import os, re, random
import numpy as np
from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont, ImageEnhance, ImageOps

HERE = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(HERE, "fonts")

DISPLAY = {"moderne": "Poppins-ExtraBold", "impact": "BebasNeue-Regular", "sacre": "Marcellus-Regular",
           "luxe": "DMSerifDisplay-Regular", "doux": "DMSerifDisplay-Regular", "editorial": "AbrilFatface-Regular"}
BODY = {"moderne": ("Poppins-Medium", "Poppins-SemiBold"), "impact": ("Poppins-Medium", "Poppins-Bold"),
        "sacre": ("Spectral-Regular", "Spectral-SemiBold"), "luxe": ("Lato-Regular", "Lato-Bold"),
        "doux": ("Lato-Regular", "Lato-Bold"), "editorial": ("Spectral-Regular", "Spectral-SemiBold")}


# Duos de polices (titre, texte, texte gras, MAJUSCULES, échelle du titre) par univers : à chaque kit, Kopi en tire un.
PAIRS = {
    "moderne": [("Studio", "Unbounded-Black", "Sora-Regular", "Sora-SemiBold", False, .8),
                ("Grotesk", "BricolageGrotesque-ExtraBold", "DMSans-Regular", "DMSans-Bold", False, 1.0),
                ("Créatif", "Syne-ExtraBold", "SpaceGrotesk-Regular", "SpaceGrotesk-Bold", False, .78),
                ("Net", "PlusJakartaSans-ExtraBold", "PlusJakartaSans-Regular", "PlusJakartaSans-Bold", False, 1.0),
                ("Rond", "Outfit-ExtraBold", "Lexend-Regular", "Lexend-Bold", False, 1.0)],
    "impact": [("Affiche", "Anton-Regular", "Montserrat-Regular", "Montserrat-Bold", True, 1.12),
               ("Stade", "BebasNeue-Regular", "Montserrat-Regular", "Montserrat-Bold", True, 1.25),
               ("Bloc", "DelaGothicOne-Regular", "Outfit-Regular", "Outfit-SemiBold", False, .82),
               ("Large", "ArchivoExpanded-Black", "Manrope-Medium", "Manrope-Bold", True, .74),
               ("Choc", "Montserrat-Black", "Montserrat-Regular", "Montserrat-Bold", True, .88)],
    "editorial": [("Magazine", "PlayfairDisplay-ExtraBold", "DMSans-Regular", "DMSans-Bold", False, 1.0),
                  ("Une", "AbrilFatface-Regular", "Lato-Regular", "Lato-Bold", False, 1.0),
                  ("Revue", "YoungSerif-Regular", "PlusJakartaSans-Regular", "PlusJakartaSans-Bold", False, .95),
                  ("Galerie", "InstrumentSerif-Regular", "Manrope-Medium", "Manrope-Bold", False, 1.22)],
    "luxe": [("Prestige", "Gloock-Regular", "Sora-Regular", "Sora-SemiBold", False, .95),
             ("Maison", "CormorantGaramond-Bold", "Manrope-Medium", "Manrope-Bold", False, 1.12),
             ("Velours", "DMSerifDisplay-Regular", "Lato-Regular", "Lato-Bold", False, 1.0),
             ("Signature", "PlayfairDisplay-BoldItalic", "Montserrat-Regular", "Montserrat-Bold", False, 1.0)],
    "doux": [("Tendre", "Fraunces-Bold", "Outfit-Regular", "Outfit-SemiBold", False, 1.0),
             ("Cocon", "YoungSerif-Regular", "Lexend-Regular", "Lexend-Bold", False, .95),
             ("Pétale", "DMSerifDisplay-Regular", "Lato-Regular", "Lato-Bold", False, 1.0),
             ("Plume", "InstrumentSerif-Italic", "PlusJakartaSans-Regular", "PlusJakartaSans-Bold", False, 1.22)],
    "sacre": [("Temple", "Cinzel-Bold", "Lato-Regular", "Lato-Bold", False, .9),
              ("Psaume", "Marcellus-Regular", "Spectral-Regular", "Spectral-SemiBold", False, 1.0),
              ("Lumière", "CormorantGaramond-Bold", "Lato-Regular", "Lato-Bold", False, 1.12),
              ("Encens", "Gloock-Regular", "Manrope-Medium", "Manrope-Bold", False, .95)],
}


def pick_pair(uni="moderne", seed=None, any_style=False):
    """Choisit un duo de polices adapté au sujet ; change à chaque kit."""
    pool = (PAIRS["moderne"] + PAIRS["editorial"] + PAIRS["luxe"] + PAIRS["impact"]) if any_style \
        else PAIRS.get(uni, PAIRS["moderne"])
    name, disp, body, bold, upper, sc = random.Random(seed).choice(pool)
    return {"name": name, "display": disp, "body": body, "bold": bold, "upper": upper, "scale": sc}


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
    small = np.asarray(im.convert("RGB")).astype(np.float32)
    yy, xx = np.mgrid[0:h, 0:w]
    v = ((xx - w * .6) / w) ** 2 + ((yy - h * .5) / h) ** 2
    small *= (1 - np.clip(v * (.9 if dark else .12), 0, .5))[..., None]
    im = Image.fromarray(np.clip(small, 0, 255).astype(np.uint8)).resize((W, H), Image.BICUBIC)
    arr = np.array(im, dtype=np.uint8)
    del im
    rng = np.random.default_rng(seed)
    for y0 in range(0, H, 256):  # grain ajouté par bandes pour limiter la mémoire
        band = arr[y0:y0 + 256].astype(np.int16)
        band += rng.normal(0, 2.6, (band.shape[0], W, 1)).astype(np.int16)
        np.clip(band, 0, 255, out=band)
        arr[y0:y0 + 256] = band
    return Image.fromarray(arr).convert("RGBA")


PAGE_SCALE = 2.6  # pages intérieures nettes en 4K


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
        pages = [pdf[i].render(scale=PAGE_SCALE).to_pil().convert("RGB") for i in pick]
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


def fit_title(draw, text, fname, maxw, maxh, start, upper=False, k=1):
    text = text.upper() if upper else text
    size = start
    while size > 30 * k:
        f = font(fname, size)
        lines = wrap(draw, text, f, maxw)
        bb = f.getbbox("ÉÀHgjpq" if not upper else "ÉÀH")
        lh = max(size * (.95 if upper else 1.08), (bb[3] - bb[1]) * 1.08)
        if len(lines) * lh <= maxh and len(lines) <= 4:
            return f, lines, lh
        size -= 4 * k
    f = font(fname, 30 * k)
    return f, wrap(draw, text, f, maxw)[:4], 32 * k


MAX_BYTES = int(1.8 * 1024 * 1024)


def save_hq(canvas, out):
    import gc
    """JPEG 4K de haute qualité : couleurs non sous-échantillonnées (texte net) + aperçu léger pour l'écran."""
    rgb_img = canvas.convert("RGB")
    # 4K mais léger : on baisse la compression petit à petit jusqu'à passer sous ~1,8 Mo
    for q in (93, 90, 88, 86, 84, 82):
        rgb_img.save(out, "JPEG", quality=q, subsampling=0, optimize=True, progressive=True, dpi=(300, 300))
        if os.path.getsize(out) <= MAX_BYTES:
            break
    else:
        rgb_img.save(out, "JPEG", quality=85, subsampling=2, optimize=True, progressive=True, dpi=(300, 300))
    prev = rgb_img.copy()
    prev.thumbnail((1400, 1400), Image.LANCZOS)
    prev.save(out.rsplit(".", 1)[0] + "_apercu.jpg", "JPEG", quality=86, optimize=True)
    del rgb_img, prev
    gc.collect()


def _unused_bg_from_cover(cover, size, pal, dark=True):
    W, H = size
    im = Image.open(cover).convert("RGB")
    sc = max(W / im.width, H / im.height) * 1.15
    im = im.resize((int(im.width * sc), int(im.height * sc)), Image.LANCZOS)
    im = im.crop(((im.width - W) // 2, (im.height - H) // 2, (im.width - W) // 2 + W, (im.height - H) // 2 + H))
    im = im.filter(ImageFilter.GaussianBlur(W * .03))
    over = Image.new("RGB", size, rgb(pal["dark"] if dark else pal["light"]))
    return Image.blend(im, over, .78 if dark else .82).convert("RGBA")


def _reflect(canvas, im, cx, bottom, strength=.22, length=.16):
    """Reflet au sol (façon vitrine) sous un objet posé."""
    h = int(im.height * length)
    if h < 4:
        return
    ref = im.transpose(Image.FLIP_TOP_BOTTOM).crop((0, 0, im.width, h))
    grad = Image.linear_gradient("L").resize((1, h)).point(lambda v: int((255 - v) * strength)).resize((im.width, h))
    a = ImageChops.multiply(ref.split()[3], grad)
    ref.putalpha(a)
    ref = ref.filter(ImageFilter.GaussianBlur(max(1, im.width * .004)))
    canvas.alpha_composite(ref, (int(cx - im.width / 2), int(bottom + im.height * .004)))


def _backlight(canvas, cx, cy, r, col, alpha=120):
    """Halo lumineux doux derrière le produit."""
    W, H = canvas.size
    sc = 8
    lay = Image.new("RGBA", (W // sc, H // sc), (0, 0, 0, 0))
    ImageDraw.Draw(lay).ellipse([(cx - r) / sc, (cy - r * .8) / sc, (cx + r) / sc, (cy + r * .8) / sc], fill=col + (alpha,))
    lay = lay.filter(ImageFilter.GaussianBlur(r / sc * .45)).resize((W, H), Image.BICUBIC)
    canvas.alpha_composite(lay)


def _scene(canvas, cover, pages, best, cx, cy, scale, pal, uni, n_pages, label="PDF", compact=False, badge_font=None):
    """Scène produit réaliste : livre relié 3D + téléphone, halo de lumière, ombres de contact et reflet au sol."""
    cov = Image.open(cover).convert("RGB")
    acc = readable(rgb(pal["accent"]), 95)
    S = scale
    ground = cy + S * .40                       # ligne du « sol » où les objets sont posés
    _backlight(canvas, cx - S * .02, cy - S * .02, S * .62, readable(rgb(pal["main"]), 70), 110)
    # livre relié
    book = book_mockup(cover, S * .80)
    bx = cx - S * .09
    _reflect(canvas, book, bx, ground)
    place(canvas, book, bx, ground - book.height / 2, 0, shadow=215, lift=1.15)
    # ombre de contact sous le livre (là où il touche le sol)
    _contact(canvas, bx, ground, book.width * .95)
    # téléphone, posé devant et un peu de côté
    ph = phone(cov, int(S * .27), acc)
    px = cx + S * .30
    _reflect(canvas, ph, px, ground + S * .035, .18, .12)
    place(canvas, ph, px, ground + S * .035 - ph.height / 2, 0, shadow=200, lift=1.0)
    _contact(canvas, px, ground + S * .035, ph.width * .9)
    # pastille
    if n_pages:
        bd = badge(S * .16, label, f"{n_pages} pages", acc, badge_font or DISPLAY.get("impact"), -12)
        place(canvas, bd, bx - book.width * .53, ground - S * .045, 0, shadow=110)


def _contact(canvas, x, y, w):
    """Ombre de contact très douce et serrée au pied d'un objet."""
    h = max(6, int(w * .07))
    lay = Image.new("RGBA", (int(w * 1.3), h * 4), (0, 0, 0, 0))
    ImageDraw.Draw(lay).ellipse([w * .15, h * 1.5, w * 1.15, h * 2.5], fill=(0, 0, 0, 150))
    lay = lay.filter(ImageFilter.GaussianBlur(h * .45))
    canvas.alpha_composite(lay, (int(x - w * .65), int(y - h * 2)))


def wide_visual(cover, out, info, pal, uni="moderne", pdf=None):
    """Visuel 16:9 4K (3840×2160) pour la boutique : scène produit à droite, titre + promesse + points forts + prix."""
    K = 2
    W, H = 1920 * K, 1080 * K
    P = lambda v: int(round(v * K))
    fp = info.get("pair") or pick_pair(uni, info.get("title"))
    canvas = mesh_bg((W, H), pal, dark=True)
    pages, n, best = pdf_pages(pdf) if pdf else ([], 0, 0)
    n = info.get("pages") or n
    _scene(canvas, cover, pages, best, W * .755, H * .52, H * .84, pal, uni, n, badge_font=fp["display"])
    d = ImageDraw.Draw(canvas)
    acc, main = readable(rgb(pal["accent"])), rgb(pal["main"])
    x, maxw = P(110), int(W * .385)
    body, bold, upper, disp = fp["body"], fp["bold"], fp["upper"], fp["display"]
    y = P(150)
    if info.get("author"):
        d.text((x, P(86)), info["author"].upper(), font=font(bold, P(22)), fill=(255, 255, 255, 190))
    label = (info.get("label") or "Ebook").upper()
    lf = font(bold, P(24))
    lw = d.textlength(label, font=lf) + P(46)
    d.rounded_rectangle([x, y, x + lw, y + P(50)], radius=P(25), fill=acc)
    tc = (20, 20, 20) if sum(acc) > 380 else (255, 255, 255)
    d.text((x + P(23), y + P(25)), label, font=lf, fill=tc, anchor="lm")
    y += P(86)
    tf, lines, lh = fit_title(d, info["title"], disp, maxw, P(320), P(84 * fp["scale"]), upper, K)
    for ln in lines:
        d.text((x, y), ln, font=tf, fill=(255, 255, 255)); y += lh
    y += P(26)
    if info.get("subtitle"):
        sf = font(body, P(31))
        for ln in wrap(d, info["subtitle"], sf, maxw)[:3]:
            d.text((x, y), ln, font=sf, fill=(226, 224, 230)); y += P(43)
    y += P(28)
    bf = font(bold, P(29))
    for b in (info.get("points") or [])[:3]:
        cy = y + P(20)
        d.ellipse([x, cy - P(16), x + P(32), cy + P(16)], fill=acc)
        d.line([(x + P(9), cy), (x + P(14), cy + P(6)), (x + P(24), cy - P(6))], fill=tc, width=P(4), joint="curve")
        ls = wrap(d, b, bf, maxw - P(54))[:2]
        for i, ln in enumerate(ls):
            d.text((x + P(50), y + i * P(38)), ln, font=bf, fill=(255, 255, 255))
        y += P(38) * len(ls) + P(20)
    y = max(y + P(14), H - P(200))
    cta = "Disponible maintenant"
    cf = font(bold, P(32))
    cwid = d.textlength(cta, font=cf) + P(120)
    d.rounded_rectangle([x, y, x + cwid, y + P(82)], radius=P(41), fill=(255, 255, 255))
    dark_txt = main if sum(main) < 520 else (20, 20, 20)
    d.text((x + P(40), y + P(41)), cta, font=cf, fill=dark_txt, anchor="lm")
    ax = x + cwid - P(46)
    d.ellipse([ax - P(22), y + P(19), ax + P(22), y + P(63)], fill=acc)
    d.line([(ax - P(9), y + P(41)), (ax + P(9), y + P(41))], fill=tc, width=P(4))
    d.line([(ax + P(2), y + P(33)), (ax + P(10), y + P(41)), (ax + P(2), y + P(49))], fill=tc, width=P(4), joint="curve")
    if info.get("price"):
        pf = font(disp, P(48 * min(1.2, fp["scale"])))
        d.text((x + cwid + P(36), y + P(41)), info["price"], font=pf, fill=(255, 255, 255), anchor="lm")
    save_hq(canvas, out)


def square_visual(cover, out, pal, uni="moderne", pdf=None, n_pages=0, pair=None):
    """Visuel carré 3000×3000 : la scène produit seule, centrée, sur un fond dégradé moderne."""
    S = 3000
    dark = uni in ("luxe", "sacre", "impact", "moderne")
    canvas = mesh_bg((S, S), pal, dark=dark, seed=11)
    pages, n, best = pdf_pages(pdf) if pdf else ([], 0, 0)
    _scene(canvas, cover, pages, best, S * .5, S * .47, S * .9, pal, uni, n_pages or n,
           badge_font=(pair or {}).get("display"))
    save_hq(canvas, out)


def story_visual(cover, out, info, pal, uni="moderne", pdf=None, n_pages=0):
    """Visuel story 4K (2160×3840) pour TikTok / Instagram / Facebook : accroche, scène produit, bénéfices, « lien dans la bio »."""
    K = 2
    W, H = 1080 * K, 1920 * K
    P = lambda v: int(round(v * K))
    fp = info.get("pair") or pick_pair(uni, info.get("title"))
    canvas = mesh_bg((W, H), pal, dark=True, seed=23)
    pages, n, best = pdf_pages(pdf) if pdf else ([], 0, 0)
    n = n_pages or n
    d = ImageDraw.Draw(canvas)
    acc, main = readable(rgb(pal["accent"])), rgb(pal["main"])
    tc = (20, 20, 20) if sum(acc) > 380 else (255, 255, 255)
    body, bold, upper, disp = fp["body"], fp["bold"], fp["upper"], fp["display"]
    y = P(230)
    label = (info.get("label") or "Nouvel ebook").upper()
    lf = font(bold, P(30))
    lw = d.textlength(label, font=lf) + P(56)
    d.rounded_rectangle([(W - lw) / 2, y, (W + lw) / 2, y + P(60)], radius=P(30), fill=acc)
    d.text((W / 2, y + P(30)), label, font=lf, fill=tc, anchor="mm")
    y += P(96)
    hook = (info.get("hook") or info.get("title") or "").strip()
    tf, lines, lh = fit_title(d, hook, disp, W - P(150), P(300), P(78 * fp["scale"]), upper, K)
    for ln in lines[:4]:
        d.text((W / 2, y), ln, font=tf, fill=(255, 255, 255), anchor="ma"); y += lh
    scene_cy = y + P(395)
    _scene(canvas, cover, pages, best, W * .51, scene_cy, W * .76, pal, uni, n, badge_font=disp)
    d = ImageDraw.Draw(canvas)
    y = scene_cy + W * .76 * .46
    bf = font(bold, P(36))
    pts = (info.get("points") or [])[:3]
    if pts:
        widths = [min(W - P(220), d.textlength(p, font=bf)) for p in pts]
        x = int((W - (max(widths) + P(56))) / 2)
        for p in pts:
            cy = y + P(24)
            d.ellipse([x, cy - P(19), x + P(38), cy + P(19)], fill=acc)
            d.line([(x + P(10), cy), (x + P(16), cy + P(7)), (x + P(28), cy - P(7))], fill=tc, width=P(5), joint="curve")
            d.text((x + P(56), y + P(2)), wrap(d, p, bf, W - P(220))[0], font=bf, fill=(255, 255, 255))
            y += P(62)
    if info.get("price"):
        y += P(18)
        pf = font(disp, P(66 * min(1.2, fp["scale"])))
        d.text((W / 2, y), info["price"], font=pf, fill=(255, 255, 255), anchor="ma")
        y += P(66 * min(1.2, fp["scale"]) + 14)
    cy0 = max(y + P(30), H - P(500))
    cta = "LIEN DANS LA BIO"
    cf = font(bold, P(46))
    cw = d.textlength(cta, font=cf) + P(200)
    x0, x1 = (W - cw) / 2, (W + cw) / 2
    # halo lumineux calculé en petit puis agrandi (rapide et léger en mémoire)
    gw, gh = W // 8, H // 8
    glow = Image.new("RGBA", (gw, gh), (0, 0, 0, 0))
    ImageDraw.Draw(glow).rounded_rectangle([(x0 - P(10)) / 8, (cy0 - P(10)) / 8, (x1 + P(10)) / 8, (cy0 + P(120)) / 8],
                                           radius=P(70) / 8, fill=acc + (150,))
    canvas.alpha_composite(glow.filter(ImageFilter.GaussianBlur(P(26) / 8)).resize((W, H), Image.BICUBIC))
    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle([x0, cy0, x1, cy0 + P(110)], radius=P(55), fill=(255, 255, 255))
    ink = main if sum(main) < 520 else (20, 20, 20)
    ax, ay = x0 + P(66), cy0 + P(55)
    dot = acc if sum(acc) < 600 else main
    tc = (20, 20, 20) if sum(dot) > 380 else (255, 255, 255)
    d.ellipse([ax - P(34), ay - P(34), ax + P(34), ay + P(34)], fill=dot)
    d.line([(ax, ay + P(17)), (ax, ay - P(17))], fill=tc, width=P(6))
    d.line([(ax - P(13), ay - P(4)), (ax, ay - P(18)), (ax + P(13), ay - P(4))], fill=tc, width=P(6), joint="curve")
    d.text((ax + P(52), ay), cta, font=cf, fill=ink, anchor="lm")
    sub = info.get("cta_sub") or "Clique et télécharge-le tout de suite"
    d.text((W / 2, cy0 + P(146)), sub, font=font(body, P(32)), fill=(236, 234, 240), anchor="ma")
    save_hq(canvas, out)


def bio_square_visual(cover, out, info, pal, uni="moderne", pdf=None, n_pages=0):
    """Carré 3000×3000 « lien en bio » pour la fin des vidéos TikTok : promesse, produit, prix, réassurance, gros bouton."""
    K = 3000 / 1080
    S = 3000
    P = lambda v: int(round(v * K))
    fp = info.get("pair") or pick_pair(uni, info.get("title"))
    canvas = mesh_bg((S, S), pal, dark=True, seed=31)
    pages, n, best = pdf_pages(pdf) if pdf else ([], 0, 0)
    n = n_pages or n
    acc, main = readable(rgb(pal["accent"])), rgb(pal["main"])
    tc = (20, 20, 20) if sum(acc) > 380 else (255, 255, 255)
    body, bold, upper, disp = fp["body"], fp["bold"], fp["upper"], fp["display"]
    d = ImageDraw.Draw(canvas)
    # promesse en haut
    y = P(70)
    hook = (info.get("hook") or info.get("title") or "").strip()
    tf, lines, lh = fit_title(d, hook, disp, S - P(160), P(150), P(58 * fp["scale"]), upper, K)
    for ln in lines[:2]:
        d.text((S / 2, y), ln, font=tf, fill=(255, 255, 255), anchor="ma"); y += lh
    # le produit
    by0, by1 = S - P(210), S - P(70)
    ry = by0 - P(78)                       # ligne de réassurance
    sc = min(S * .55, (ry - P(24) - (y + P(44))) / .93)
    scene_cy = y + P(44) + .45 * sc
    _scene(canvas, cover, pages, best, S * .5, scene_cy, sc, pal, uni, n, badge_font=disp)
    d = ImageDraw.Draw(canvas)
    # pastille prix (en haut à droite du produit)
    if info.get("price"):
        R = P(118)
        st = Image.new("RGBA", (R * 2 + 8, R * 2 + 8), (0, 0, 0, 0))
        g = ImageDraw.Draw(st)
        import math
        pts = []
        for k in range(48):  # étoile festonnée façon étiquette promo
            a = math.pi * 2 * k / 48
            r = R if k % 2 == 0 else R * .9
            pts.append((R + 4 + r * math.cos(a), R + 4 + r * math.sin(a)))
        g.polygon(pts, fill=acc + (255,))
        price = info["price"]
        num, cur = (price.rsplit(" ", 1) + [""])[:2] if " " in price else (price, "")
        nf = font(disp, P(40) if len(num) <= 6 else P(32))
        g.text((R + 4, R + 4 - (P(14) if cur else 0)), num, font=nf, fill=tc, anchor="mm")
        if cur:
            g.text((R + 4, R + 4 + P(30)), cur.upper(), font=font(bold, P(19)), fill=tc, anchor="mm")
        place(canvas, _rot(st, 12), S * .5 + sc * .47, scene_cy - sc * .10, 0, shadow=150)
        d = ImageDraw.Draw(canvas)
    # réassurance
    y = ry
    items = ["PDF" + (f" · {n} pages" if n else ""), "Accès immédiat", "Lisible sur ton téléphone"]
    rf = font(bold, P(25))
    widths = [d.textlength(t, font=rf) for t in items]
    gap = P(44)
    x = (S - (sum(widths) + gap * (len(items) - 1) + P(34) * len(items))) / 2
    for t, w in zip(items, widths):
        cy = y + P(16)
        d.ellipse([x, cy - P(13), x + P(26), cy + P(13)], fill=acc)
        d.line([(x + P(7), cy), (x + P(11), cy + P(5)), (x + P(19), cy - P(5))], fill=tc, width=P(3), joint="curve")
        d.text((x + P(34), y), t, font=rf, fill=(236, 234, 240))
        x += P(34) + w + gap
    # flèches vers le haut + gros bouton
    cf = font(disp if upper else bold, P(56 * (1.1 if upper else 1)))
    cta = "CLIQUE SUR LE LIEN EN BIO"
    tw = d.textlength(cta, font=cf)
    while tw > S - P(260) and cf.size > P(30):
        cf = font(disp if upper else bold, cf.size - P(2)); tw = d.textlength(cta, font=cf)
    bw = tw + P(190)
    bx0, bx1 = (S - bw) / 2, (S + bw) / 2
    gw = S // 8
    glow = Image.new("RGBA", (gw, gw), (0, 0, 0, 0))
    ImageDraw.Draw(glow).rounded_rectangle([(bx0 - P(16)) / 8, (by0 - P(16)) / 8, (bx1 + P(16)) / 8, (by1 + P(16)) / 8],
                                           radius=P(80) / 8, fill=acc + (190,))
    canvas.alpha_composite(glow.filter(ImageFilter.GaussianBlur(P(30) / 8)).resize((S, S), Image.BICUBIC))
    d = ImageDraw.Draw(canvas)
    d.rounded_rectangle([bx0, by0, bx1, by1], radius=(by1 - by0) / 2, fill=acc)
    ax, ay = bx0 + P(78), (by0 + by1) / 2
    d.ellipse([ax - P(42), ay - P(42), ax + P(42), ay + P(42)], fill=tc)
    d.line([(ax, ay + P(22)), (ax, ay - P(22))], fill=acc, width=P(8))
    d.line([(ax - P(17), ay - P(5)), (ax, ay - P(23)), (ax + P(17), ay - P(5))], fill=acc, width=P(8), joint="curve")
    d.text((ax + P(66), ay), cta, font=cf, fill=tc, anchor="lm")
    save_hq(canvas, out)


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

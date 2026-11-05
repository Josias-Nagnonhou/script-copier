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


def book_mockup(cover_path, height):
    """Livre debout vu de trois quarts : couverture, tranche, pages, reflet. Renvoie une image RGBA."""
    cover = Image.open(cover_path).convert("RGB")
    ratio = cover.width / cover.height if cover.height else .707
    H = int(height)
    W = int(H * ratio)
    face = cover.resize((W, H), Image.LANCZOS).convert("RGBA")
    spine_w, pages_w, slant = int(W * .085), int(W * .035), int(H * .022)
    cw, ch = spine_w + int(W * .93) + pages_w + 4, H + 4
    out = Image.new("RGBA", (cw, ch), (0, 0, 0, 0))
    fx0, fx1 = spine_w, spine_w + int(W * .93)
    face_q = [(fx0, slant), (fx1, 0), (fx1, H), (fx0, H - slant)]
    # tranche : bande gauche de la couverture, assombrie
    strip = cover.crop((0, 0, max(4, int(cover.width * .05)), cover.height)).resize((spine_w * 2, H))
    strip = ImageEnhance.Brightness(strip).enhance(.55).convert("RGBA")
    spine_q = [(0, slant * 2.2), (fx0, slant), (fx0, H - slant), (0, H - slant * 2.2)]
    out.alpha_composite(_warp(strip, spine_q, (cw, ch)))
    # pages (bord droit)
    pg = Image.new("RGBA", (pages_w * 4, H), (246, 242, 235, 255))
    d = ImageDraw.Draw(pg)
    for x in range(0, pg.width, 3):
        d.line([(x, 0), (x, H)], fill=(222, 216, 206, 255), width=1)
    pages_q = [(fx1, int(H * .012)), (fx1 + pages_w, int(H * .022)), (fx1 + pages_w, H - int(H * .022)), (fx1, H - int(H * .012))]
    out.alpha_composite(_warp(pg, pages_q, (cw, ch)))
    # couverture
    out.alpha_composite(_warp(face, face_q, (cw, ch)))
    # ombrage près de la reliure + reflet
    fx = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    g = ImageDraw.Draw(fx)
    for i in range(int(W * .07)):
        a = int(110 * (1 - i / (W * .07)) ** 1.6)
        g.line([(i, 0), (i, H)], fill=(0, 0, 0, a))
    g.line([(int(W * .075), 0), (int(W * .075), H)], fill=(255, 255, 255, 38), width=max(2, W // 160))
    gloss = Image.new("L", (W, H), 0)
    gd = ImageDraw.Draw(gloss)
    gd.polygon([(W * .35, 0), (W * .62, 0), (W * .25, H), (0, H), (0, H * .85)], fill=34)
    gloss = gloss.filter(ImageFilter.GaussianBlur(W * .06))
    fx.alpha_composite(Image.merge("RGBA", (Image.new("L", (W, H), 255),) * 3 + (gloss,)))
    out.alpha_composite(_warp(fx, face_q, (cw, ch)))
    return out


def place_with_shadow(canvas, book, x, y, strength=150, floor=True):
    """Pose le livre avec une ombre douce : ombre portée diffuse + contact au sol."""
    bw, bh = book.size
    alpha = book.split()[3]
    pad = int(bh * .2)
    sh = Image.new("L", (bw + 2 * pad, bh + 2 * pad), 0)
    sh.paste(alpha.point(lambda a: int(a * strength * .55 / 255)), (pad, pad))
    sh = sh.filter(ImageFilter.GaussianBlur(bh * .06))
    shadow = Image.new("RGBA", sh.size, (0, 0, 0, 0))
    shadow.putalpha(sh)
    canvas.alpha_composite(shadow, (int(x - pad + bh * .05), int(y - pad + bh * .035)))
    if floor:
        fw, fh = int(bw * .92), int(bh * .05)
        fl = Image.new("L", (fw + 200, fh + 200), 0)
        ImageDraw.Draw(fl).ellipse([100, 100, 100 + fw, 100 + fh], fill=int(strength * .75))
        fl = fl.filter(ImageFilter.GaussianBlur(bh * .018))
        f2 = Image.new("RGBA", fl.size, (0, 0, 0, 0)); f2.putalpha(fl)
        canvas.alpha_composite(f2, (int(x + bw * .06 - 100), int(y + bh - fh * .6 - 100)))
    canvas.alpha_composite(book, (int(x), int(y)))


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


def bg_from_cover(cover, size, pal, dark=True):
    W, H = size
    im = Image.open(cover).convert("RGB")
    sc = max(W / im.width, H / im.height) * 1.15
    im = im.resize((int(im.width * sc), int(im.height * sc)), Image.LANCZOS)
    im = im.crop(((im.width - W) // 2, (im.height - H) // 2, (im.width - W) // 2 + W, (im.height - H) // 2 + H))
    im = im.filter(ImageFilter.GaussianBlur(W * .03))
    over = Image.new("RGB", size, rgb(pal["dark"] if dark else pal["light"]))
    return Image.blend(im, over, .78 if dark else .82).convert("RGBA")


def wide_visual(cover, out, info, pal, uni="moderne"):
    """Visuel 16:9 pour la boutique : livre 3D à droite, titre + promesse + points forts + prix à gauche."""
    W, H = 1920, 1080
    canvas = bg_from_cover(cover, (W, H), pal, dark=True)
    lay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(lay)
    acc, main = rgb(pal["accent"]), rgb(pal["main"])
    d.ellipse([W * .52, H * .02, W * 1.08, H * 1.02], fill=main + (70,))  # halo derrière le livre
    d.ellipse([W * .6, H * .12, W * 1.0, H * .9], fill=acc + (40,))
    for k in range(3):
        r = 330 + k * 120
        d.ellipse([W * .79 - r, H * .5 - r, W * .79 + r, H * .5 + r], outline=(255, 255, 255, 22), width=2)
    lay = lay.filter(ImageFilter.GaussianBlur(40))
    canvas.alpha_composite(lay)
    book = book_mockup(cover, H * .80)
    place_with_shadow(canvas, book, W * .79 - book.width / 2, H * .5 - book.height / 2 - 10, strength=170)
    d = ImageDraw.Draw(canvas)
    x, maxw = 120, int(W * .5)
    body, bold = BODY.get(uni, BODY["moderne"])
    upper = uni == "impact"
    y = 140
    label = (info.get("label") or "Ebook").upper()
    lf = font(bold, 26)
    lw = d.textlength(label, font=lf) + 48
    d.rounded_rectangle([x, y, x + lw, y + 52], radius=26, fill=acc)
    tc = (20, 20, 20) if sum(acc) > 380 else (255, 255, 255)
    d.text((x + 24, y + 26), label, font=lf, fill=tc, anchor="lm")
    y += 92
    tf, lines, lh = fit_title(d, info["title"], DISPLAY.get(uni, DISPLAY["moderne"]), maxw, 330, 104 if upper else 84, upper)
    for ln in lines:
        d.text((x, y), ln, font=tf, fill=(255, 255, 255)); y += lh
    y += 18
    if info.get("subtitle"):
        sf = font(body, 34)
        for ln in wrap(d, info["subtitle"], sf, maxw)[:3]:
            d.text((x, y), ln, font=sf, fill=(232, 228, 222)); y += 46
    y += 26
    bf = font(bold, 31)
    for b in (info.get("points") or [])[:3]:
        cy = y + 21
        d.ellipse([x, cy - 17, x + 34, cy + 17], fill=acc)
        d.line([(x + 9, cy), (x + 15, cy + 7), (x + 26, cy - 7)], fill=tc, width=4, joint="curve")
        ls = wrap(d, b, bf, maxw - 56)[:2]
        for i, ln in enumerate(ls):
            d.text((x + 54, y + i * 40), ln, font=bf, fill=(255, 255, 255))
        y += 40 * len(ls) + 22
    y = max(y + 14, H - 210)
    cta = "Disponible maintenant"
    cf = font(bold, 34)
    cwid = d.textlength(cta, font=cf) + 80
    d.rounded_rectangle([x, y, x + cwid, y + 84], radius=42, fill=(255, 255, 255))
    d.text((x + 40, y + 42), cta, font=cf, fill=main if sum(main) < 520 else (20, 20, 20), anchor="lm")
    if info.get("price"):
        pf = font(DISPLAY.get(uni, "Poppins-ExtraBold") if uni != "impact" else "BebasNeue-Regular", 60 if uni == "impact" else 48)
        d.text((x + cwid + 40, y + 42), info["price"], font=pf, fill=acc, anchor="lm")
    if info.get("author"):
        d.text((x, 70), info["author"].upper(), font=font(bold, 24), fill=(255, 255, 255, 200))
    canvas.convert("RGB").save(out, "JPEG", quality=92, optimize=True)


def square_visual(cover, out, pal, uni="moderne"):
    """Visuel carré : seulement le livre en 3D, centré, sur un fond propre."""
    S = 1600
    dark = uni in ("luxe", "sacre", "impact")
    top, bot = (rgb(pal["dark"]), tuple(int(v * .55) for v in rgb(pal["dark"]))) if dark else \
        ((250, 248, 244), rgb(pal["light"]))
    grad = Image.new("RGB", (1, S))
    for y in range(S):
        t = y / S
        grad.putpixel((0, y), tuple(int(top[i] + (bot[i] - top[i]) * t) for i in range(3)))
    canvas = grad.resize((S, S)).convert("RGBA")
    lay = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(lay)
    glow = rgb(pal["main"]) + ((90,) if dark else (55,))
    d.ellipse([S * .14, S * .1, S * .86, S * .82], fill=glow)
    lay = lay.filter(ImageFilter.GaussianBlur(120))
    canvas.alpha_composite(lay)
    book = book_mockup(cover, S * .66)
    place_with_shadow(canvas, book, (S - book.width) / 2 + S * .01, S * .14, strength=120 if not dark else 190)
    canvas.convert("RGB").save(out, "JPEG", quality=93, optimize=True)


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

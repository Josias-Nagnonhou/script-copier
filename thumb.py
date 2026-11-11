"""Miniatures : génération de l'image (Cloudflare FLUX.2 klein, avec photo du personnage et miniature modèle)
puis pose du texte par Kopi (police nette, accents justes, contour, mot en couleur, badge, flèche…)."""
import base64
import io
import math
import os
import random
import re
import time

import requests
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter, ImageFont, ImageOps

HERE = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(HERE, "fonts")

FORMATS = {"16:9": (1280, 720), "9:16": (720, 1280), "1:1": (1080, 1080), "4:5": (1080, 1350)}
GEN_SIZE = {"16:9": (1280, 720), "9:16": (720, 1280), "1:1": (1024, 1024), "4:5": (896, 1120)}
MODELS = {"rapide": "@cf/black-forest-labs/flux-2-klein-4b", "qualite": "@cf/black-forest-labs/flux-2-klein-9b"}

FONTS = {  # nom affiché → fichier, majuscules ?
    "impact": ("Anton-Regular", True), "bebas": ("BebasNeue-Regular", True), "montserrat": ("Montserrat-Black", True),
    "archivo": ("ArchivoBlack-Regular", True), "gothic": ("DelaGothicOne-Regular", True),
    "unbounded": ("Unbounded-Black", True), "cinzel": ("Cinzel-Bold", True), "serif": ("Fraunces-Black", False),
    "poppins": ("Poppins-ExtraBold", True), "rond": ("PaytoneOne-Regular", True),
}

STYLES = {
    "viral": {"label": "Viral / choc", "font": "montserrat", "color": "#FFFFFF", "accent": "#FFE500",
              "img": "ultra vibrant saturated colors, high contrast, bright punchy lighting, glossy, MrBeast-style viral "
                     "YouTube thumbnail look, strong rim light, clean background with depth"},
    "documentaire": {"label": "Documentaire / enquête", "font": "impact", "color": "#FFFFFF", "accent": "#E8202A",
                     "img": "dark cinematic documentary look, moody teal and orange grading, dramatic side light, film "
                            "grain, investigative atmosphere, shallow depth of field"},
    "mystere": {"label": "Mystère / occulte", "font": "cinzel", "color": "#F5E6C8", "accent": "#FF2D2D",
                "img": "mysterious dark atmosphere, fog, candle light and deep shadows, glowing red and gold accents, "
                       "eerie cinematic mood, volumetric light"},
    "spirituel": {"label": "Spirituel / lumière", "font": "cinzel", "color": "#FFFFFF", "accent": "#FFC93C",
                  "img": "divine golden light rays, heavenly glow, warm sacred atmosphere, soft haze, majestic and "
                         "emotional, cinematic"},
    "foot": {"label": "Football / sport", "font": "impact", "color": "#FFFFFF", "accent": "#00E676",
             "img": "stadium at night with floodlights, sports broadcast look, dynamic energy, sweat and intensity, "
                    "motion, confetti or smoke, high contrast"},
    "argent": {"label": "Business / argent", "font": "archivo", "color": "#FFFFFF", "accent": "#2BE37A",
               "img": "luxury and success atmosphere, money, gold, modern city, sleek premium lighting, sharp and "
                      "glossy, aspirational"},
    "histoire": {"label": "Histoire / épopée", "font": "serif", "color": "#FFF4E0", "accent": "#E8A33D",
                 "img": "epic historical painting meets cinema, warm dusty light, ancient textures, grand scale, "
                        "dramatic sky"},
    "minimal": {"label": "Minimaliste / propre", "font": "poppins", "color": "#FFFFFF", "accent": "#FF4D4D",
                "img": "clean minimal composition, one strong subject, plain colored background, soft studio light, "
                       "modern and premium"},
    "avantapres": {"label": "Avant / après", "font": "montserrat", "color": "#FFFFFF", "accent": "#FFE500",
                   "img": "split-screen before and after composition, left side dull and sad, right side bright and "
                          "successful, strong contrast between the two halves"},
}

NICHES = ["Football", "Spiritualité chrétienne", "Islam", "Mystère / ésotérisme", "Business / argent", "Histoire",
          "Couple / relations", "Santé / astuces", "Motivation", "Faits divers / enquête", "Tech / IA", "Cuisine",
          "Autre"]

EMOTIONS = {"choc": "shocked face, wide eyes, open mouth", "colere": "angry intense face, frowning",
            "peur": "scared worried face", "serieux": "serious intense stare into the camera",
            "joie": "big genuine smile, excited", "mystere": "mysterious half-lit face, intense eyes",
            "triste": "sad emotional face, teary eyes", "confiant": "confident powerful pose, slight smirk"}


def hexrgb(h):
    h = (h or "#FFFFFF").lstrip("#")
    if len(h) != 6:
        return (255, 255, 255)
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def font(name, size):
    return ImageFont.truetype(os.path.join(FONT_DIR, name + ".ttf"), max(8, int(size)))


# ---------------------------------------------------------------- génération de l'image
def _ref_bytes(path, side=504):
    """Les images de référence doivent faire moins de 512×512 pour FLUX.2 klein."""
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    im.thumbnail((side, side), Image.LANCZOS)
    b = io.BytesIO()
    im.save(b, "JPEG", quality=92)
    return b.getvalue()


def generate(prompt, fmt, account, token, refs=(), quality="rapide", seed=None, log=print):
    """refs : chemins d'images (0 = modèle de style, puis photos du personnage). Retourne une image PIL."""
    w, h = GEN_SIZE.get(fmt, GEN_SIZE["16:9"])
    model = MODELS.get(quality, MODELS["rapide"])
    url = f"https://api.cloudflare.com/client/v4/accounts/{account}/ai/run/{model}"
    files = {"prompt": (None, prompt[:2000]), "width": (None, str(w)), "height": (None, str(h))}
    if seed is not None:
        files["seed"] = (None, str(int(seed) % 2_000_000_000))
    for i, p in enumerate(list(refs)[:4]):
        files[f"input_image_{i}"] = (f"ref{i}.jpg", _ref_bytes(p), "image/jpeg")
    last = ""
    for attempt in range(3):
        try:
            r = requests.post(url, headers={"Authorization": f"Bearer {token}"}, files=files, timeout=180)
        except requests.RequestException as e:
            last = f"connexion impossible ({type(e).__name__})"
            time.sleep(4 + 4 * attempt)
            continue
        ct = r.headers.get("content-type", "")
        if r.status_code == 200:
            if ct.startswith("image/"):
                return Image.open(io.BytesIO(r.content)).convert("RGB")
            try:
                res = r.json().get("result") or {}
                b64 = res.get("image") if isinstance(res, dict) else None
                if b64:
                    return Image.open(io.BytesIO(base64.b64decode(b64))).convert("RGB")
            except Exception:
                pass
            last = "réponse vide"
        else:
            txt = r.text[:300]
            log(f"[miniature] Cloudflare {r.status_code} : {txt}")
            if r.status_code in (401, 403):
                raise RuntimeError("Cloudflare refuse la clé : vérifie ton Account ID et ton jeton.")
            if r.status_code == 429 or "neuron" in txt.lower() or "limit" in txt.lower() and "rate" not in txt.lower():
                raise RuntimeError("Quota gratuit Cloudflare du jour atteint (environ 50 miniatures). "
                                   "Il revient demain à 1 h du matin (heure du Bénin).")
            last = f"erreur {r.status_code}"
        time.sleep(4 + 4 * attempt)
    raise RuntimeError(f"La génération de l'image a échoué ({last}). Réessaie dans un instant.")


def fit_cover(im, size):
    """Recadre au centre pour remplir exactement le format voulu."""
    return ImageOps.fit(im, size, Image.LANCZOS, centering=(0.5, 0.42))


# ---------------------------------------------------------------- pose du texte
def _lines(draw, words, fnt, maxw):
    lines, cur = [], []
    for w in words:
        t = " ".join(cur + [w])
        if not cur or draw.textlength(t, font=fnt) <= maxw:
            cur.append(w)
        else:
            lines.append(cur)
            cur = [w]
    if cur:
        lines.append(cur)
    return lines


def _fit(draw, text, fname, maxw, maxh, start, max_lines=3):
    words = text.split()
    size = start
    while size > 22:
        f = font(fname, size)
        lines = _lines(draw, words, f, maxw)
        lh = size * 1.02
        if len(lines) <= max_lines and len(lines) * lh <= maxh and \
                all(draw.textlength(" ".join(l), font=f) <= maxw for l in lines):
            return f, lines, lh
        size -= 3
    f = font(fname, 22)
    return f, _lines(draw, words, f, maxw)[:max_lines], 23


def _darken_side(img, side, strength=150):
    """Dégradé sombre discret côté texte : le texte reste lisible sur n'importe quelle image."""
    w, h = img.size
    grad = Image.new("L", (w, h), 0)
    g = ImageDraw.Draw(grad)
    if side in ("left", "right"):
        span = int(w * .62)
        for x in range(span):
            a = int(strength * (1 - x / span) ** 1.6)
            xx = x if side == "left" else w - 1 - x
            g.line([(xx, 0), (xx, h)], fill=a)
    else:
        span = int(h * .5)
        for y in range(span):
            a = int(strength * (1 - y / span) ** 1.6)
            yy = y if side == "top" else h - 1 - y
            g.line([(0, yy), (w, yy)], fill=a)
    black = Image.new("RGB", (w, h), (0, 0, 0))
    return Image.composite(black, img, grad)


def _vignette(img, amount=0.35):
    w, h = img.size
    m = Image.new("L", (w, h), 0)
    ImageDraw.Draw(m).ellipse([-w * .25, -h * .3, w * 1.25, h * 1.3], fill=255)
    m = m.filter(ImageFilter.GaussianBlur(min(w, h) * .18))
    dark = ImageEnhance.Brightness(img).enhance(1 - amount)
    return Image.composite(img, dark, m)


def _arrow(draw, p0, p1, color, width):
    (x0, y0), (x1, y1) = p0, p1
    ang = math.atan2(y1 - y0, x1 - x0)
    head = width * 3.2
    bx, by = x1 - head * math.cos(ang), y1 - head * math.sin(ang)
    for col, extra in (((255, 255, 255), width * .45), (color, 0)):
        wd = int(width + extra)
        draw.line([(x0, y0), (bx, by)], fill=col, width=wd)
        l = (bx + (head * .62 + extra) * math.cos(ang + math.pi / 2), by + (head * .62 + extra) * math.sin(ang + math.pi / 2))
        r = (bx + (head * .62 + extra) * math.cos(ang - math.pi / 2), by + (head * .62 + extra) * math.sin(ang - math.pi / 2))
        tip = (x1 + extra * .8 * math.cos(ang), y1 + extra * .8 * math.sin(ang))
        draw.polygon([l, tip, r], fill=col)


def compose(base, spec, fmt="16:9"):
    """base : image PIL au bon format. spec : texte, mot_fort, badge, position, couleurs, police, éléments."""
    W, H = FORMATS.get(fmt, FORMATS["16:9"])
    img = fit_cover(base, (W, H))
    img = ImageEnhance.Contrast(img).enhance(1.06)
    img = ImageEnhance.Color(img).enhance(1.12)
    img = img.filter(ImageFilter.UnsharpMask(radius=2, percent=70, threshold=3))
    text = re.sub(r"\s+", " ", (spec.get("texte") or "")).strip()
    pos = spec.get("position") or "auto"
    if pos == "auto":
        pos = "left" if fmt == "16:9" else ("top" if fmt == "9:16" else "bottom")
    if spec.get("vignette", True):
        img = _vignette(img, .28)
    if text and pos in ("left", "right", "top", "bottom"):
        img = _darken_side(img, pos, 165 if spec.get("contraste", True) else 90)
    d = ImageDraw.Draw(img)
    fkey = spec.get("police") or "montserrat"
    fname, upper = FONTS.get(fkey, FONTS["montserrat"])
    main = hexrgb(spec.get("couleur") or "#FFFFFF")
    acc = hexrgb(spec.get("accent") or "#FFE500")
    m = int(min(W, H) * .055)
    if pos in ("left", "right"):
        bw, bh = int(W * .5), int(H * .74)
        bx = m if pos == "left" else W - m - bw
    else:
        bw, bh = W - 2 * m, int(H * (.34 if fmt == "9:16" else .4))
        bx = m
    scale = float(spec.get("taille") or 1)
    tx = text.upper() if upper else text
    hl = [w.strip(".,!?:;«»\"'").upper() for w in (spec.get("mot_fort") or "").split()]
    block = []
    if tx:
        f, lines, lh = _fit(d, tx, fname, bw, bh, int(H * (.2 if fmt == "16:9" else .11) * scale),
                            max_lines=3 if fmt == "16:9" else 4)
        total_h = len(lines) * lh
        badge = (spec.get("badge") or "").strip()
        bf = font("Montserrat-Black", max(18, int(f.size * .32)))
        badge_h = int(bf.size * 1.7) if badge else 0
        if pos in ("left", "right"):
            by = (H - total_h - badge_h) / 2 + badge_h
        elif pos == "top":
            by = m + badge_h
        else:
            by = H - m - total_h
        sw = max(3, int(f.size * .075))
        maxlw = max(d.textlength(" ".join(l), font=f) for l in lines)
        box_style = spec.get("style_mot") == "boite"
        for i, line in enumerate(lines):
            lw = d.textlength(" ".join(line), font=f)
            if pos == "right":
                x = bx + bw - lw
            elif pos in ("top", "bottom"):
                x = (W - lw) / 2
            else:
                x = bx
            y = by + i * lh
            for word in line:
                ww = d.textlength(word, font=f)
                strong = word.strip(".,!?:;«»\"'").upper() in hl
                # ombre portée douce
                sh = Image.new("L", (int(ww + sw * 6), int(lh + sw * 6)), 0)
                ImageDraw.Draw(sh).text((sw * 3, sw * 3), word, font=f, fill=200, stroke_width=sw,
                                        stroke_fill=200)
                sh = sh.filter(ImageFilter.GaussianBlur(sw * 1.2))
                img.paste((0, 0, 0), (int(x - sw * 3 + sw * .8), int(y - sw * 3 + sw * 1.4)), sh)
                d = ImageDraw.Draw(img)
                if strong and box_style:
                    bb = d.textbbox((x, y), word, font=f)
                    pad = f.size * .1
                    d.rounded_rectangle([bb[0] - pad, bb[1] - pad, bb[2] + pad, bb[3] + pad], radius=pad * .9, fill=acc)
                    d.text((x, y), word, font=f, fill=(15, 15, 15))
                else:
                    d.text((x, y), word, font=f, fill=acc if strong else main, stroke_width=sw, stroke_fill=(0, 0, 0))
                x += ww + d.textlength(" ", font=f)
        if badge:
            bt = badge.upper()
            tw = d.textlength(bt, font=bf)
            pad = bf.size * .55
            bx0 = bx if pos == "left" else (bx + bw - tw - 2 * pad if pos == "right" else (W - tw - 2 * pad) / 2)
            by0 = by - badge_h - bf.size * .2
            d.rounded_rectangle([bx0, by0, bx0 + tw + 2 * pad, by0 + bf.size * 1.5], radius=bf.size * .4, fill=acc)
            lum = 0.299 * acc[0] + 0.587 * acc[1] + 0.114 * acc[2]
            d.text((bx0 + pad, by0 + bf.size * .2), bt, font=bf, fill=(12, 12, 12) if lum > 140 else (255, 255, 255))
        if pos == "right":
            block = [bx + bw - maxlw, by, bx + bw, by + total_h]
        elif pos in ("top", "bottom"):
            block = [(W - maxlw) / 2, by, (W + maxlw) / 2, by + total_h]
        else:
            block = [bx, by, bx + maxlw, by + total_h]
    # éléments graphiques
    red = (235, 30, 35)
    if spec.get("fleche") and block:
        mid = (block[1] + block[3]) / 2
        if pos == "left":
            p0, p1 = (block[2] + W * .025, block[3] - (block[3] - block[1]) * .2), (min(W * .66, block[2] + W * .2), mid + H * .12)
        elif pos == "right":
            p0, p1 = (block[0] - W * .025, block[3] - (block[3] - block[1]) * .2), (max(W * .34, block[0] - W * .2), mid + H * .12)
        elif pos == "top":
            p0, p1 = (W * .3, block[3] + H * .04), (W * .4, block[3] + H * .17)
        else:
            p0, p1 = (W * .3, block[1] - H * .04), (W * .4, block[1] - H * .17)
        _arrow(d, p0, p1, red, max(8, int(min(W, H) * .022)))
    if spec.get("cercle"):
        if pos == "left":
            cx, cy = W * .74, H * .5
        elif pos == "right":
            cx, cy = W * .26, H * .5
        else:
            cx, cy = W * .5, H * (.62 if pos == "top" else .38)
        r = min(W, H) * .2
        for col, wd in (((255, 255, 255), int(r * .13)), (red, int(r * .08))):
            d.ellipse([cx - r * 1.15, cy - r, cx + r * 1.15, cy + r], outline=col, width=wd)
    return img


def save(img, path, max_bytes=1_900_000):
    """JPEG net sous 2 Mo (limite de YouTube)."""
    for q in (95, 92, 90, 87, 84, 80):
        img.save(path, "JPEG", quality=q, optimize=True, subsampling=0 if q >= 90 else 2)
        if os.path.getsize(path) <= max_bytes:
            break
    return path


# ---------------------------------------------------------------- prompts
def build_prompt(scene, style, emotion, fmt, has_ref, n_people, text_side, african=True, no_text=True):
    st = STYLES.get(style, STYLES["viral"])["img"]
    parts = ["Professional YouTube thumbnail photograph" if fmt == "16:9" else "Professional vertical social media cover photo"]
    parts.append(scene.strip().rstrip("."))
    if n_people:
        who = "the person from image 1" if n_people == 1 else "the people from image 1 and image 2"
        idx = 1 if has_ref else 0
        if not has_ref:
            who = "the person from image 0" if n_people == 1 else "the people from image 0 and image 1"
        parts.append(f"The main character is {who}: keep exactly the same face, identity, skin tone and hairstyle, "
                     f"close-up, {EMOTIONS.get(emotion, EMOTIONS['choc'])}")
    elif african:
        parts.append(f"Main character: Black African person, close-up, {EMOTIONS.get(emotion, EMOTIONS['choc'])}")
    if has_ref:
        parts.append("Use image 0 only as a style and composition reference (framing, colors, lighting, energy); "
                     "do not copy its people, logos or text")
    side = {"left": "on the right side of the frame, leaving the left 45% of the image clean and simple for a title",
            "right": "on the left side of the frame, leaving the right 45% clean and simple for a title",
            "top": "in the lower part of the frame, leaving the top third clean for a title",
            "bottom": "in the upper part of the frame, leaving the bottom third clean for a title"}.get(text_side)
    if side:
        parts.append("Place the main subject " + side)
    parts.append(st)
    parts.append("ultra sharp focus, 4k detail, high dynamic range, eye-catching, clickable")
    if no_text:
        parts.append("absolutely no text, no letters, no words, no logo, no watermark")
    return ". ".join(parts) + "."

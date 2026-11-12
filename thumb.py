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

# clé → (fichier, majuscules, description pour l'IA)
FONTS = {
    "montserrat": ("Montserrat-Black", True, "très grasse et nette, polyvalente, style viral MrBeast"),
    "impact": ("Anton-Regular", True, "condensée haute type Impact, info, enquête, foot"),
    "bebas": ("BebasNeue-Regular", True, "condensée fine et haute, élégante, documentaire moderne"),
    "archivo": ("ArchivoBlack-Regular", True, "massive, business, argent, tech"),
    "gothic": ("DelaGothicOne-Regular", True, "ultra grasse japonisante, choc, énergie"),
    "unbounded": ("Unbounded-Black", True, "large futuriste, tech, IA, crypto"),
    "cinzel": ("Cinzel-Bold", True, "romaine gravée, histoire, Bible, sacré, empire"),
    "cinzeldeco": ("CinzelDecorative-Black", True, "romaine décorative, royal, mystique, légende"),
    "serif": ("Fraunces-Black", False, "serif élégante, luxe, littérature"),
    "poppins": ("Poppins-Black", True, "ronde géométrique très propre, minimaliste, conseils"),
    "poppinsit": ("Poppins-BlackItalic", True, "ronde italique, dynamique, sport, motivation"),
    "rond": ("PaytoneOne-Regular", True, "arrondie amicale, cuisine, famille, astuces"),
    "luckiest": ("LuckiestGuy-Regular", True, "cartoon épaisse, fun, divertissement, jeunesse, réaction"),
    "bangers": ("Bangers-Regular", True, "bande dessinée, action, humour, gaming"),
    "titan": ("TitanOne-Regular", True, "très ronde et grasse, joyeuse, enfants, lifestyle"),
    "lilita": ("LilitaOne-Regular", True, "grasse arrondie, très lisible, YouTube grand public"),
    "blackops": ("BlackOpsOne-Regular", True, "pochoir militaire, guerre, crime, danger"),
    "russo": ("RussoOne-Regular", True, "sport mécanique, foot, voitures, gaming"),
    "bungee": ("Bungee-Regular", True, "urbaine en blocs, street, rap, ville"),
    "rubikmono": ("RubikMonoOne-Regular", True, "géométrique large, tech, chiffres, classement"),
    "passion": ("PassionOne-Black", True, "condensée ronde, émotion forte, faits divers"),
    "alfaslab": ("AlfaSlabOne-Regular", True, "slab épaisse western, histoire, vieux temps, chocs"),
    "staatliches": ("Staatliches-Regular", True, "condensée affiche, documentaire, politique"),
    "marker": ("PermanentMarker-Regular", True, "écrite au marqueur, authentique, vlog, personnel"),
    "creepster": ("Creepster-Regular", True, "horreur dégoulinante, peur, surnaturel"),
    "nosifer": ("Nosifer-Regular", True, "horreur sanglante, crime, terreur (à utiliser rarement)"),
    "righteous": ("Righteous-Regular", True, "rétro moderne, musique, culture"),
    "bowlby": ("BowlbyOneSC-Regular", True, "massive à petites capitales, puissance, autorité"),
    "ultra": ("Ultra-Regular", True, "serif ultra grasse, presse, gros titres, scandale"),
    "barlow": ("BarlowCondensed-Black", True, "condensée sport/news très lisible, foot, actu"),
    "kanit": ("Kanit-Black", True, "moderne grasse, business, motivation"),
    "kanitit": ("Kanit-ExtraBoldItalic", True, "italique grasse, vitesse, foot, urgence"),
    "pirata": ("PirataOne-Regular", False, "gothique, occulte, sectes, mystère, sorcellerie"),
    "fraktur": ("UnifrakturCook-Bold", False, "gothique médiévale, ésotérisme, ancien, interdit"),
    "metal": ("MetalMania-Regular", True, "métal agressive, sombre, rebelle"),
    "changa": ("ChangaOne-Regular", True, "grasse arrondie italique, fun, sport"),
    "knewave": ("Knewave-Regular", True, "pinceau énergique, motivation, sport"),
    "shrikhand": ("Shrikhand-Regular", False, "rétro grasse italique, lifestyle, food"),
    "graffiti": ("SedgwickAveDisplay-Regular", True, "graffiti à la main, street, rap"),
    "playfairsc": ("PlayfairDisplaySC-Black", True, "serif luxe, richesse, élégance, spirituel"),
    "abril": ("AbrilFatface-Regular", False, "serif contrastée, mode, féminin, luxe"),
    "fugaz": ("FugazOne-Regular", True, "italique sportive, course, compétition"),
    "racing": ("RacingSansOne-Regular", True, "course italique, vitesse, foot"),
    "coiny": ("Coiny-Regular", True, "ronde bulle, fun, enfants"),
    "wetpaint": ("RubikWetPaint-Regular", True, "peinture qui coule, horreur légère, choc"),
    "glitch": ("RubikGlitch-Regular", True, "glitch numérique, hacking, IA, bug"),
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


# ---------------------------------------------------------------- personnage détouré (photo intacte)
def place_people(img, people, pos, fmt, accent, glow=True):
    """Pose la ou les personnes détourées en gros plan, sans retoucher le visage."""
    W, H = img.size
    if not people:
        return img
    side_x = {"left": .70, "right": .30}.get(pos, .5)
    if fmt == "16:9":
        face_h, face_cy = .40, .40
    elif fmt == "9:16":
        face_h, face_cy = .19, (.58 if pos == "top" else .36)
    else:
        face_h, face_cy = .28, (.56 if pos == "top" else .40)
    layers = []
    for n, pp in enumerate(people[:2]):
        cut = pp["img"]
        face = pp.get("face")
        if face:
            fh = max(10, face[3] - face[1])
            s = (H * face_h * (.84 if n else 1)) / fh
            fcx, fcy = (face[0] + face[2]) / 2 * s, (face[1] + face[3]) / 2 * s
        else:  # pas de visage détecté : la personne remplit 95 % de la hauteur
            s = H * .95 / cut.height
            fcx, fcy = cut.width * s / 2, cut.height * s * .25
        s = min(s, 2.6)  # au-delà, la photo deviendrait floue
        cw, ch = max(1, int(cut.width * s)), max(1, int(cut.height * s))
        sub = cut.resize((cw, ch), Image.LANCZOS)
        cx = side_x if len(people) == 1 else ((.62, .86)[n] if pos == "left" else (.38, .14)[n] if pos == "right" else (.32, .68)[n])
        x = int(W * cx - fcx)
        y = int(H * face_cy - fcy)
        if y + ch < H:  # le corps doit toucher le bas de l'image (pas de personne qui flotte)
            y = H - ch if ch >= H * .6 else y
        layers.append((sub, x, y))
    for sub, x, y in reversed(layers):  # la 1re personne passe devant
        a = sub.split()[3]
        pad = max(8, int(min(W, H) * .02))
        big = Image.new("L", (sub.width + 4 * pad, sub.height + 4 * pad), 0)
        big.paste(a, (2 * pad, 2 * pad))
        # ombre portée douce
        sh = big.filter(ImageFilter.GaussianBlur(pad * 1.4)).point(lambda v: int(v * .55))
        img.paste((0, 0, 0), (x - 2 * pad + pad // 2, y - 2 * pad + pad), sh)
        if glow:
            # halo de couleur + liseré blanc : la personne se détache du fond (style miniature pro)
            halo = big.filter(ImageFilter.MaxFilter(pad // 2 * 2 + 1)).filter(ImageFilter.GaussianBlur(pad * .9))
            img.paste(accent, (x - 2 * pad, y - 2 * pad), halo.point(lambda v: int(v * .75)))
            ring = big.filter(ImageFilter.MaxFilter(max(3, pad // 3 * 2 + 1))).filter(ImageFilter.GaussianBlur(.8))
            img.paste((255, 255, 255), (x - 2 * pad, y - 2 * pad), ring)
        sub2 = ImageEnhance.Contrast(sub).enhance(1.04)
        sub2 = sub2.filter(ImageFilter.UnsharpMask(radius=1.4, percent=60, threshold=2))
        img.paste(sub2, (x, y), a)
    return img


# ---------------------------------------------------------------- couleurs qui captent l'œil
ACCENTS = ["#FFE500", "#FF2D2D", "#2BE37A", "#00D9FF", "#FF8A00", "#FF3DCB", "#FFFFFF"]


def _hsv(rgb):
    import colorsys
    return colorsys.rgb_to_hsv(*[c / 255 for c in rgb])


def auto_accent(img, box, prefer=None):
    """Choisit la couleur du mot fort qui ressort le plus sur le fond de la zone du texte."""
    crop = img.crop(tuple(int(v) for v in box)).resize((40, 40)).convert("RGB")
    px = list(crop.getdata())
    avg = tuple(sum(p[i] for p in px) // len(px) for i in range(3))
    h0, s0, v0 = _hsv(avg)
    best, score = prefer or ACCENTS[0], -1
    for c in ([prefer] if prefer else []) + ACCENTS:
        r = hexrgb(c)
        h, s_, v = _hsv(r)
        dh = min(abs(h - h0), 1 - abs(h - h0)) * 2 if s0 > .15 else .6
        sc = dh * .6 + abs(v - v0) * .3 + s_ * .1 + (.15 if c == prefer else 0)
        if sc > score:
            best, score = c, sc
    return best


# ---------------------------------------------------------------- texte empilé façon pro
def _stack(d, text, fname, maxw, maxh, max_lines, hl):
    """Découpe en lignes et donne à chaque ligne sa propre taille pour former un bloc compact."""
    words = text.split()
    if not words:
        return []
    best = None
    for nl in range(1, min(max_lines, len(words)) + 1):
        # répartit les mots en nl lignes de longueur proche
        tot = sum(len(w) for w in words) + len(words) - 1
        lines, cur, acc = [], [], 0
        for w in words:
            if cur and acc + len(w) > tot / nl * (len(lines) + 1) + 2 and len(lines) < nl - 1:
                lines.append(cur); cur = []
            cur.append(w); acc += len(w) + 1
        lines.append(cur)
        sizes = []
        for l in lines:
            t = " ".join(l)
            f = font(fname, 100)
            wpx = d.textlength(t, font=f)
            sizes.append(100 * maxw / max(wpx, 1))
        mn = min(sizes)
        sizes = [min(sz, mn * 1.45) for sz in sizes]  # une ligne courte grossit, sans exagérer
        total = sum(sizes) * 1.0
        k = min(1, maxh / total)
        sizes = [sz * k for sz in sizes]
        score = min(sizes) * (1 if nl <= 3 else .9)
        if best is None or score > best[0]:
            best = (score, lines, sizes)
    _, lines, sizes = best
    return [(l, font(fname, max(20, int(sz)))) for l, sz in zip(lines, sizes)]


def compose(base, spec, fmt="16:9", people=None):
    """base : fond (image IA). people : personnes détourées. spec : texte, couleurs, police, flou…"""
    W, H = FORMATS.get(fmt, FORMATS["16:9"])
    img = fit_cover(base, (W, H))
    flou = float(spec.get("flou") or 0)
    if flou > 0:
        img = img.filter(ImageFilter.GaussianBlur(min(W, H) * .0045 * flou))
    img = ImageEnhance.Contrast(img).enhance(1.08)
    img = ImageEnhance.Color(img).enhance(1.15)
    if not people:
        img = img.filter(ImageFilter.UnsharpMask(radius=2, percent=70, threshold=3))
    text = re.sub(r"\s+", " ", (spec.get("texte") or "")).strip()
    pos = spec.get("position") or "auto"
    if pos == "auto":
        pos = "left" if fmt == "16:9" else ("top" if fmt == "9:16" else "bottom")
    if spec.get("vignette", True):
        img = _vignette(img, .3)
    if text and pos in ("left", "right", "top", "bottom"):
        img = _darken_side(img, pos, 175 if spec.get("contraste", True) else 90)
    m = int(min(W, H) * .05)
    if pos in ("left", "right"):
        bw, bh = int(W * (.47 if people else .5)), int(H * .72)
        bx = m if pos == "left" else W - m - bw
    else:
        bw, bh = W - 2 * m, int(H * (.3 if fmt == "9:16" else .36))
        bx = m
    acc_pref = spec.get("accent")
    acc = hexrgb(acc_pref if acc_pref and acc_pref != "auto" else auto_accent(img, (bx, m, bx + bw, H - m)))
    if people:
        img = place_people(img, people, pos, fmt, acc, glow=spec.get("halo", True))
    d = ImageDraw.Draw(img)
    fkey = spec.get("police") or "montserrat"
    fname, upper = FONTS.get(fkey, FONTS["montserrat"])[:2]
    main = hexrgb(spec.get("couleur") if spec.get("couleur") not in (None, "", "auto") else "#FFFFFF")
    scale = float(spec.get("taille") or 1)
    tx = text.upper() if upper else text
    hl = [w.strip(".,!?:;«»\"'").upper() for w in (spec.get("mot_fort") or "").split()]
    block = []
    if tx:
        badge = (spec.get("badge") or "").strip()
        lines = _stack(d, tx, fname, bw * min(1, scale), bh * scale * (.8 if badge else 1),
                       3 if fmt == "16:9" else 4, hl)
        heights = [f.size * (1.0 if upper else 1.12) for _, f in lines]
        total_h = sum(heights)
        bf = font("Montserrat-Black", max(18, int(min(f.size for _, f in lines) * .36)))
        badge_h = int(bf.size * 1.9) if badge else 0
        if pos in ("left", "right"):
            by = (H - total_h - badge_h) / 2 + badge_h
        elif pos == "top":
            by = m + badge_h
        else:
            by = H - m - total_h
        box_style = spec.get("style_mot") == "boite"
        maxlw = 0
        y = by
        for (line, f), lh in zip(lines, heights):
            sw = max(3, int(f.size * .085))
            lw = d.textlength(" ".join(line), font=f)
            maxlw = max(maxlw, lw)
            x = bx + bw - lw if pos == "right" else ((W - lw) / 2 if pos in ("top", "bottom") else bx)
            top_off = f.getbbox("ÉH")[1] * .55  # place pour les accents des majuscules
            for word in line:
                ww = d.textlength(word, font=f)
                strong = word.strip(".,!?:;«»\"'").upper() in hl
                # ombre dure décalée (relief) + ombre douce
                off = max(2, int(f.size * .055))
                sh = Image.new("L", (int(ww + sw * 8), int(lh + sw * 8)), 0)
                ImageDraw.Draw(sh).text((sw * 4, sw * 4 - top_off), word, font=f, fill=255, stroke_width=sw, stroke_fill=255)
                img.paste((0, 0, 0), (int(x - sw * 4 + off), int(y - sw * 4 + off)), sh.point(lambda v: int(v * .85)))
                img.paste((0, 0, 0), (int(x - sw * 4 + off), int(y - sw * 4 + off * 1.6)),
                          sh.filter(ImageFilter.GaussianBlur(sw * 1.5)).point(lambda v: int(v * .5)))
                d = ImageDraw.Draw(img)
                if strong and box_style:
                    bb = d.textbbox((x, y - top_off), word, font=f)
                    pad = f.size * .09
                    d.rounded_rectangle([bb[0] - pad, bb[1] - pad, bb[2] + pad, bb[3] + pad], radius=pad, fill=acc)
                    lum = .299 * acc[0] + .587 * acc[1] + .114 * acc[2]
                    d.text((x, y - top_off), word, font=f, fill=(15, 15, 15) if lum > 140 else (255, 255, 255))
                else:
                    d.text((x, y - top_off), word, font=f, fill=acc if strong else main, stroke_width=sw,
                           stroke_fill=(0, 0, 0))
                x += ww + d.textlength(" ", font=f)
            y += lh
        if badge:
            bt = badge.upper()
            tw = d.textlength(bt, font=bf)
            pad = bf.size * .55
            bx0 = bx if pos == "left" else (bx + bw - tw - 2 * pad if pos == "right" else (W - tw - 2 * pad) / 2)
            by0 = by - badge_h
            d.rounded_rectangle([bx0, by0, bx0 + tw + 2 * pad, by0 + bf.size * 1.5], radius=bf.size * .35, fill=acc)
            lum = 0.299 * acc[0] + 0.587 * acc[1] + 0.114 * acc[2]
            d.text((bx0 + pad, by0 + bf.size * .2), bt, font=bf, fill=(12, 12, 12) if lum > 140 else (255, 255, 255))
        if pos == "right":
            block = [bx + bw - maxlw, by, bx + bw, by + total_h]
        elif pos in ("top", "bottom"):
            block = [(W - maxlw) / 2, by, (W + maxlw) / 2, by + total_h]
        else:
            block = [bx, by, bx + maxlw, by + total_h]
    red = (235, 30, 35)
    if spec.get("fleche") and block:
        mid = (block[1] + block[3]) / 2
        if pos == "left":
            p0, p1 = (block[2] + W * .02, block[3] - (block[3] - block[1]) * .25), (min(W * .6, block[2] + W * .14), mid + H * .1)
        elif pos == "right":
            p0, p1 = (block[0] - W * .02, block[3] - (block[3] - block[1]) * .25), (max(W * .4, block[0] - W * .14), mid + H * .1)
        elif pos == "top":
            p0, p1 = (W * .3, block[3] + H * .04), (W * .4, block[3] + H * .15)
        else:
            p0, p1 = (W * .3, block[1] - H * .04), (W * .4, block[1] - H * .15)
        _arrow(d, p0, p1, red, max(8, int(min(W, H) * .02)))
    if spec.get("cercle"):
        cx, cy = ((W * .74, H * .5) if pos == "left" else (W * .26, H * .5) if pos == "right"
                  else (W * .5, H * (.62 if pos == "top" else .38)))
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
    """n_people > 0 : la vraie photo sera posée par-dessus, donc l'IA ne fait QUE le décor (aucune personne)."""
    st = STYLES.get(style, STYLES["viral"])["img"]
    if n_people:
        parts = ["Background plate for a professional YouTube thumbnail, photographic, cinematic depth"
                 if fmt == "16:9" else "Background plate for a vertical social media cover, photographic"]
        parts.append(scene.strip().rstrip("."))
        parts.append("IMPORTANT: no people, no person, no human, no face, no body, no hands, no crowd in the "
                     "foreground; only the setting, objects and atmosphere")
        side = {"left": "the right half has an open area where a person will be added later, the left half is darker "
                        "and simple for a big title",
                "right": "the left half has an open area where a person will be added later, the right half is darker "
                         "and simple for a big title",
                "top": "the bottom has an open area for a person, the top is simple for a title",
                "bottom": "the top has an open area for a person, the bottom is simple for a title"}.get(text_side)
        if side:
            parts.append(side)
    else:
        parts = ["Professional YouTube thumbnail photograph" if fmt == "16:9" else "Professional vertical social media cover photo"]
        parts.append(scene.strip().rstrip("."))
        if african:
            parts.append(f"Main character: Black African person, close-up, {EMOTIONS.get(emotion, EMOTIONS['choc'])}")
        side = {"left": "on the right side of the frame, leaving the left 45% of the image clean and simple for a title",
                "right": "on the left side of the frame, leaving the right 45% clean and simple for a title",
                "top": "in the lower part of the frame, leaving the top third clean for a title",
                "bottom": "in the upper part of the frame, leaving the bottom third clean for a title"}.get(text_side)
        if side:
            parts.append("Place the main subject " + side)
    if has_ref:
        parts.append("Use image 0 only as a style and composition reference (colors, lighting, energy); "
                     "do not copy its people, logos or text")
    parts.append(st)
    parts.append("ultra sharp, rich detail, high dynamic range, eye-catching color contrast")
    if no_text:
        parts.append("absolutely no text, no letters, no words, no logo, no watermark")
    return ". ".join(parts) + "."

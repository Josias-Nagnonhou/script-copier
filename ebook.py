"""Kopi — créateur d'ebooks : plan + rédaction par l'IA, illustrations, PDF pro et Word modifiable."""
import os, re, io, json, math, random, time, datetime, urllib.parse

import requests
from PIL import Image, ImageDraw, ImageFilter
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, PageBreak,
                                NextPageTemplate, Table, TableStyle, KeepTogether, Flowable, CondPageBreak)
from reportlab.platypus import Image as RLImage
_RLParagraph = Paragraph


class Paragraph(_RLParagraph):
    """Paragraphe qui ne fait jamais planter le livre : en cas de balisage invalide, on garde le texte brut."""
    def __init__(self, text, *a, **k):
        try:
            super().__init__(text, *a, **k)
        except Exception:
            import re as _re
            bare = _re.sub(r"<[^>]+>", "", str(text))
            super().__init__(bare.replace("&", "&amp;").replace("&amp;amp;", "&amp;").replace("<", "&lt;"), *a, **k)
from reportlab.platypus.tableofcontents import TableOfContents
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

HERE = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(HERE, "fonts")
_FONTS_OK = False


def register_fonts():
    global _FONTS_OK
    if _FONTS_OK:
        return
    files = {"Pop": "Poppins-Regular", "Pop-It": "Poppins-Italic", "Pop-Md": "Poppins-Medium", "Pop-Sb": "Poppins-SemiBold",
             "Pop-Bd": "Poppins-Bold", "Pop-Xb": "Poppins-ExtraBold", "DMS": "DMSerifDisplay-Regular",
             "DMS-It": "DMSerifDisplay-Italic", "Marc": "Marcellus-Regular", "Bebas": "BebasNeue-Regular",
             "Archivo": "ArchivoBlack-Regular", "Abril": "AbrilFatface-Regular", "Spec": "Spectral-Regular",
             "Spec-It": "Spectral-Italic", "Spec-Sb": "Spectral-SemiBold", "Spec-Bd": "Spectral-Bold",
             "Lato": "Lato-Regular", "Lato-It": "Lato-Italic", "Lato-Bd": "Lato-Bold", "Lato-Bk": "Lato-Black"}
    for name, f in files.items():
        pdfmetrics.registerFont(TTFont(name, os.path.join(FONT_DIR, f + ".ttf")))
    pdfmetrics.registerFontFamily("Pop", normal="Pop", bold="Pop-Sb", italic="Pop-It", boldItalic="Pop-Sb")
    pdfmetrics.registerFontFamily("Spec", normal="Spec", bold="Spec-Sb", italic="Spec-It", boldItalic="Spec-Bd")
    pdfmetrics.registerFontFamily("Lato", normal="Lato", bold="Lato-Bd", italic="Lato-It", boldItalic="Lato-Bd")
    _FONTS_OK = True


# ---------------------------------------------------------------- univers graphiques
# Chaque univers a ses polices, sa couverture, ses ouvertures de chapitre, ses encadrés et ses en-têtes.
STYLES = {
    "moderne": {"label": "Moderne", "display": "Pop-Xb", "head": "Pop-Bd", "sub": "Pop-Sb", "body": "Pop", "it": "Pop-It",
                "lab": "Pop-Sb", "size": 10.6, "lead": 17.2, "upper": False, "cover": "bleed", "opener": "band",
                "callout": "bar", "marker": "dot", "header": "topbar", "quote": "bar",
                "img": "clean modern editorial illustration, bright natural light, uncluttered background",
                "for": "argent, business, entrepreneuriat, marketing, technologie, productivité, carrière"},
    "impact": {"label": "Impact", "display": "Bebas", "head": "Archivo", "sub": "Pop-Bd", "body": "Pop", "it": "Pop-It",
               "lab": "Pop-Bd", "size": 10.6, "lead": 17.2, "upper": True, "cover": "diagonal", "opener": "giant",
               "callout": "block", "marker": "square", "header": "sidebar", "quote": "big",
               "img": "dramatic cinematic scene, strong contrast, dynamic energy, bold rim lighting",
               "for": "motivation, discipline, sport, football, mental de gagnant, fitness, défi"},
    "sacre": {"label": "Sacré", "display": "Marc", "head": "Marc", "sub": "Spec-Sb", "body": "Spec", "it": "Spec-It",
              "lab": "Spec-Sb", "size": 11.4, "lead": 18.4, "upper": False, "cover": "arch", "opener": "centered",
              "callout": "frame", "marker": "diamond", "header": "centered", "quote": "centered",
              "img": "luminous spiritual painting, soft golden light rays, serene and reverent atmosphere, symbolic",
              "for": "foi chrétienne, islam, prière, psaumes, spiritualité, délivrance, méditation"},
    "luxe": {"label": "Luxe", "display": "DMS", "head": "DMS", "sub": "Lato-Bd", "body": "Lato", "it": "Lato-It",
             "lab": "Lato-Bd", "size": 10.8, "lead": 17.6, "upper": False, "cover": "medallion", "opener": "dark",
             "callout": "frame", "marker": "diamond", "header": "gold", "quote": "centered",
             "img": "mysterious elegant dark atmosphere, deep shadows, glowing gold highlights, cinematic depth",
             "for": "mystère, ésotérisme, secrets, sociétés secrètes, occulte, luxe, haut de gamme"},
    "doux": {"label": "Doux", "display": "DMS", "head": "DMS", "sub": "Lato-Bd", "body": "Lato", "it": "Lato-It",
             "lab": "Lato-Bd", "size": 10.9, "lead": 17.8, "upper": False, "cover": "card", "opener": "rounded",
             "callout": "soft", "marker": "circle", "header": "dot", "quote": "soft",
             "img": "soft warm pastel illustration, gentle diffused light, tender and comforting mood",
             "for": "couple, famille, maternité, santé, bien-être, beauté, cuisine, développement personnel doux"},
    "editorial": {"label": "Éditorial", "display": "Abril", "head": "DMS", "sub": "Spec-Sb", "body": "Spec", "it": "Spec-It",
                  "lab": "Spec-Sb", "size": 11.2, "lead": 18.2, "upper": False, "cover": "magazine", "opener": "split",
                  "callout": "rule", "marker": "dash", "header": "running", "quote": "bar",
                  "img": "rich documentary painting, textured oil paint look, historical atmosphere, warm tones",
                  "for": "histoire, culture africaine, contes, biographies, éducation, langues, savoirs"},
}

PALETTES = {
    "moderne": {"main": "#2E4BB8", "dark": "#121A3D", "light": "#EAEEFB", "accent": "#F0B429"},
    "impact": {"main": "#E0412B", "dark": "#141414", "light": "#FBEAE6", "accent": "#F5C518"},
    "sacre": {"main": "#7A5A1E", "dark": "#1B2238", "light": "#F6F0E2", "accent": "#C9A227"},
    "luxe": {"main": "#6B3FA0", "dark": "#120D1C", "light": "#F1ECF7", "accent": "#C9A84A"},
    "doux": {"main": "#C4627A", "dark": "#3A2430", "light": "#FBEFF2", "accent": "#7FB7A4"},
    "editorial": {"main": "#A2461E", "dark": "#22201C", "light": "#F7F0E6", "accent": "#2F6F6A"},
}


def hexrgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def tohex(rgb):
    return "#%02X%02X%02X" % tuple(max(0, min(255, int(round(v)))) for v in rgb)


def lum(h):
    def ch(v):
        v /= 255
        return v / 12.92 if v <= .03928 else ((v + .055) / 1.055) ** 2.4
    r, g, b = hexrgb(h)
    return .2126 * ch(r) + .7152 * ch(g) + .0722 * ch(b)


def mix(h1, h2, t):
    a, b = hexrgb(h1), hexrgb(h2)
    return tohex([a[i] + (b[i] - a[i]) * t for i in range(3)])


def resolve_design(design):
    """Univers + couleurs proposés par l'IA, vérifiés (lisibilité) et complétés si besoin."""
    design = design if isinstance(design, dict) else {}
    uni = str(design.get("univers") or "").lower().strip()
    uni = next((k for k in STYLES if uni.startswith(k[:4]) or STYLES[k]["label"].lower() == uni), "moderne")
    base = dict(PALETTES[uni])
    cols = design.get("couleurs") if isinstance(design.get("couleurs"), dict) else {}
    ok = lambda v: isinstance(v, str) and re.fullmatch(r"#?[0-9A-Fa-f]{6}", v.strip())
    for k, alias in (("main", "principale"), ("dark", "sombre"), ("light", "claire"), ("accent", "accent")):
        v = cols.get(alias) or cols.get(k)
        if ok(v):
            base[k] = "#" + v.strip().lstrip("#").upper()
    if lum(base["dark"]) > .06:
        base["dark"] = mix(base["dark"], "#000000", .55)
    if lum(base["light"]) < .80:
        base["light"] = mix(base["main"], "#FFFFFF", .9)
    if lum(base["main"]) > .45:
        base["main"] = mix(base["main"], "#000000", .3)
    if abs(lum(base["accent"]) - lum(base["dark"])) < .15:
        base["accent"] = mix(base["accent"], "#FFFFFF", .45)
    return uni, base


# ---------------------------------------------------------------- texte
EMOJI = re.compile("[\U0001F000-\U0001FAFF️​-‍⁠]")
# caractères que les polices du livre n'ont pas → équivalents lisibles
GLYPHS = {" ": " ", " ": " ", " ": " ", " ": " ", " ": " ",
          "→": "–", "⇒": "–", "➜": "–", "←": "–", "✓": "", "✔": "", "✅": "",
          "☑": "", "✗": "×", "✘": "×", "❌": "×", "★": "•", "☆": "•", "♥": "",
          "❤": "", "■": "•", "□": "•", "●": "•", "○": "•", "◆": "•", "◇": "•",
          "▪": "•", "▫": "•", "‣": "•", "◦": "•", "‐": "-", "‑": "-", "‒": "–",
          "―": "—", "‟": "\"", "‛": "'", "№": "N°", " ": " ", "\t": " ", "\r": ""}
_SAFE = None


def safe_chars():
    """Caractères présents dans les 3 polices de texte (Poppins, Spectral, Lato)."""
    global _SAFE
    if _SAFE is None:
        try:
            from fontTools.ttLib import TTFont as FT
            sets = [set(FT(os.path.join(FONT_DIR, f + ".ttf")).getBestCmap()) for f in
                    ("Poppins-Regular", "Spectral-Regular", "Lato-Regular")]
            _SAFE = set.intersection(*sets)
        except Exception:
            _SAFE = set()
    return _SAFE


def clean(t):
    t = EMOJI.sub("", str(t or ""))
    for a, b in GLYPHS.items():
        t = t.replace(a, b)
    ok = safe_chars()
    if ok:
        t = "".join(ch for ch in t if ord(ch) < 128 or ord(ch) in ok)
    t = re.sub(r"(?<=\S)[ \u00a0]{2,}(?=\S)", " ", t)
    return t.strip()


def tidy_md(md):
    """Nettoie un texte rédigé par l'IA : phrases parasites, doublons, balises cassées."""
    md = clean(md).replace("```markdown", "").replace("```", "")
    md = re.sub(r"<br\s*/?>", "\n", md)
    md = re.sub(r"</?(?:b|i|strong|em|u|p|span|div)[^>]*>", "", md)
    out, seen, open_box = [], set(), False
    for ln in md.split("\n"):
        s = ln.strip()
        low = s.lower()
        if re.match(r"^(voici|bien sûr|certainement|d'accord|here is|here's|sure)\b.*[:：]$", low):
            continue  # phrase d'introduction de l'IA
        if re.match(r"^(#+\s*)?(chapitre|chapter)\s+\d+\s*[:\-–—.]?\s*$", low):
            continue  # « Chapitre 3 » répété seul
        if s.startswith(":::"):
            if s == ":::":
                if not open_box:
                    continue
                open_box = False
            else:
                if open_box:
                    out.append(":::")
                open_box = True
        elif s.startswith("#") and open_box:
            out.append(":::")
            open_box = False
        if s.count("**") % 2:
            ln = ln.replace("**", "")
        key = re.sub(r"\W+", "", low)
        if len(key) > 60 and key in seen:
            continue  # paragraphe répété mot pour mot
        if len(key) > 60:
            seen.add(key)
        out.append(ln.rstrip())
    if open_box:
        out.append(":::")
    md = "\n".join(out)
    md = re.sub(r"(?<=[.!?…])\s+([^.!?…\n]{40,}[.!?…])(\s+\1)+", r" \1", md)  # phrase collée deux fois
    return re.sub(r"\n{3,}", "\n\n", md).strip()


def de_loop(t, cap=1800):
    """Coupe les répétitions en boucle de l'IA (« perle noire, perle blanche, perle noire, … ») et les lignes géantes."""
    if len(t) < 600:
        return t
    parts = re.split(r"(?<=[,;])\s+", t)
    out, seen = [], set()
    for k, p in enumerate(parts):
        key = " ".join(x.strip().lower() for x in parts[k:k + 3])
        if k >= 3 and key in seen:
            out[-1] = out[-1].rstrip(",; ") + "…"
            break
        seen.add(key)
        out.append(p)
    t = " ".join(out)
    if len(t) > cap:
        cut = t[:cap]
        end = max(cut.rfind(". "), cut.rfind("! "), cut.rfind("? "))
        t = cut[:end + 1] if end > cap * .5 else cut.rsplit(" ", 1)[0] + "…"
    return t


def inline(t):
    """Markdown en ligne → balises ReportLab (après échappement)."""
    t = de_loop(clean(t)).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"<i>\1</i>", t)
    t = re.sub(r"__(.+?)__", r"<b>\1</b>", t)
    t = t.replace("`", "")
    return balance_tags(t)


def balance_tags(t):
    """Répare les balises <b>/<i> mal imbriquées (« **gras *italique*** ») qui font planter ReportLab."""
    out, stack, pos = [], [], 0
    for m in re.finditer(r"</?([bi])>", t):
        out.append(t[pos:m.start()]); pos = m.end()
        tag, closing = m.group(1), m.group(0).startswith("</")
        if not closing:
            if tag in stack:
                continue  # déjà ouvert : on ignore
            stack.append(tag); out.append(f"<{tag}>")
        elif tag in stack:
            reopen = []
            while stack[-1] != tag:
                x = stack.pop(); out.append(f"</{x}>"); reopen.append(x)
            stack.pop(); out.append(f"</{tag}>")
            for x in reversed(reopen):
                stack.append(x); out.append(f"<{x}>")
        # balise fermante orpheline : supprimée
    out.append(t[pos:])
    out.extend(f"</{x}>" for x in reversed(stack))
    return re.sub(r"<([bi])></\1>", "", "".join(out))


def plain(t):
    return re.sub(r"\*\*|__|`|(?<![\w])\*(?=\S)|(?<=\S)\*(?![\w])", "", clean(t))


CALLOUTS = {"astuce": "Astuce", "retenir": "À retenir", "attention": "Attention", "exercice": "Exercice",
            "exemple": "Exemple", "action": "Passe à l'action", "verset": "Parole", "cas": "Cas concret"}


def parse_md(md):
    """Découpe le texte rédigé par l'IA en blocs (titres, paragraphes, listes, encadrés…)."""
    blocks, para, lines = [], [], (md or "").replace("\r", "").split("\n")

    def flush():
        if para:
            blocks.append({"t": "p", "text": " ".join(para)})
            para.clear()
    i = 0
    while i < len(lines):
        raw = lines[i]
        s = raw.strip()
        if not s:
            flush(); i += 1; continue
        m = re.match(r"^:::\s*(\w+)\s*(.*)$", s)
        if m:
            flush()
            kind = m.group(1).lower()
            kind = next((k for k in CALLOUTS if kind.startswith(k[:5])), "astuce")
            title, body = m.group(2).strip(), []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith(":::"):
                body.append(lines[i]); i += 1
            i += 1
            blocks.append({"t": "callout", "kind": kind, "title": title, "body": parse_md("\n".join(body))})
            continue
        if s.startswith("#"):
            flush()
            lvl = len(s) - len(s.lstrip("#"))
            blocks.append({"t": "h2" if lvl <= 2 else "h3", "text": s.lstrip("#").strip()})
            i += 1; continue
        if re.match(r"^[-*•]\s+\[[ xX]?\]\s+", s):
            flush(); items = []
            while i < len(lines) and re.match(r"^\s*[-*•]\s+\[[ xX]?\]\s+", lines[i]):
                items.append(re.sub(r"^\s*[-*•]\s+\[[ xX]?\]\s+", "", lines[i])); i += 1
            blocks.append({"t": "check", "items": items}); continue
        if re.match(r"^[-*•]\s+", s):
            flush(); items = []
            while i < len(lines) and re.match(r"^\s*[-*•]\s+", lines[i]):
                items.append(re.sub(r"^\s*[-*•]\s+", "", lines[i])); i += 1
            blocks.append({"t": "ul", "items": items}); continue
        if re.match(r"^\d+[.)]\s+", s):
            flush(); items = []
            while i < len(lines) and re.match(r"^\s*\d+[.)]\s+", lines[i]):
                items.append(re.sub(r"^\s*\d+[.)]\s+", "", lines[i])); i += 1
            blocks.append({"t": "ol", "items": items}); continue
        if s.startswith(">"):
            flush(); q = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                q.append(lines[i].strip().lstrip(">").strip()); i += 1
            src = ""
            if q and re.match(r"^[—–-]\s*", q[-1]) and len(q) > 1:
                src = re.sub(r"^[—–-]\s*", "", q.pop())
            blocks.append({"t": "quote", "text": " ".join(q), "src": src}); continue
        if s.startswith("|"):
            flush(); rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.match(r"^:?-{2,}:?$", c) for c in cells if c):
                    rows.append(cells)
                i += 1
            if rows:
                blocks.append({"t": "table", "rows": rows})
            continue
        para.append(s); i += 1
    flush()
    return blocks


# ---------------------------------------------------------------- IA : plan et rédaction
def make_context(ctx):
    a = ctx.get("audience") or {}
    lines = [f"Niche : {ctx.get('niche') or ''}",
             f"Audience : {a.get('profil', '')} · {a.get('langue_pays', '')} · niveau {a.get('niveau', '')}",
             f"Pouvoir d'achat : {a.get('pouvoir_achat', '')}"]
    if ctx.get("douleurs"):
        lines.append("Douleurs de l'audience : " + " ; ".join(ctx["douleurs"][:6]))
    if ctx.get("desirs"):
        lines.append("Désirs de l'audience : " + " ; ".join(ctx["desirs"][:6]))
    p = ctx.get("product") or {}
    for k, lab in (("promesse", "Promesse du produit"), ("pour_qui", "Pour qui"), ("pourquoi_ca_marchera", "Pourquoi il répond à la demande")):
        if p.get(k):
            lines.append(f"{lab} : {p[k]}")
    if p.get("contenu"):
        lines.append("Contenu suggéré : " + " ; ".join(p["contenu"][:10]))
    if ctx.get("brief"):
        lines.append("Consignes de l'auteur : " + ctx["brief"])
    return "\n".join(l for l in lines if l.split(":", 1)[-1].strip())


STYLE_RULES = (
    "Règles d'écriture : français clair et chaleureux, tutoiement, phrases courtes. Chaque chapitre doit apporter "
    "une vraie valeur qu'on ne trouve pas en 2 minutes sur Internet : méthode pas à pas, exemples concrets tirés "
    "de la vie en Afrique francophone (prénoms, villes, situations, monnaie en FCFA), erreurs fréquentes et comment "
    "les éviter, exercices à faire. Interdits : remplissage, répétitions, généralités creuses, promesses miracles, "
    "chiffres ou études inventés, fausses citations, mélange de langues (tout en français). Si le sujet touche à "
    "l'actualité récente (résultats, votes, transferts, élections, événements de l'année en cours), n'invente AUCUN fait, "
    "nom, score ou déclaration : appuie-toi seulement sur ce qui est établi, explique les mécanismes et invite le lecteur "
    "à vérifier les dernières informations. Pour un verset, un hadith ou une citation, donne la référence "
    "exacte et ne cite que des textes très connus ; sinon reformule sans guillemets. Pas d'emoji."
)

FORMAT_RULES = (
    "Format du texte (Markdown simple) : « ## » pour les parties, « ### » pour les sous-parties, paragraphes séparés "
    "par une ligne vide, « - » pour les listes, « 1. » pour les étapes, « - [ ] » pour une checklist, « > » pour une "
    "citation (ligne suivante « > — Source »), tableaux Markdown si utile, **gras** pour les mots clés. "
    "Encadrés : une ligne « :::astuce Titre court », le texte, puis une ligne « ::: ». Types d'encadrés possibles : "
    "astuce, retenir, attention, exercice, exemple, cas, action. Mets 2 à 4 encadrés par chapitre, dont au moins un "
    "exercice ou une action concrète. Ne répète pas le titre du chapitre au début."
)


def gen_plan(chat, ctx, pages):
    n_ch = max(4, min(20, round(pages / 5.5)))
    n_sec = max(3, min(8, round(pages * 230 / n_ch / 900)))  # parties par chapitre : ~900 mots chacune
    system = ("Tu es un éditeur expert en ebooks qui se vendent bien auprès des audiences francophones africaines "
              "et de la diaspora. Tu conçois des guides pratiques qui tiennent leur promesse. " + STYLE_RULES)
    title_rule = (f"Le titre est IMPOSÉ par l'auteur : « {ctx['title']} ». Mets exactement ce titre dans « titre », sans le modifier, "
                  "et construis le sous-titre et le contenu autour de lui.\n") if ctx.get("keep_title") else \
        f"Titre provisoire : « {ctx['title']} » (tu peux l'améliorer pour qu'il vende mieux).\n"
    user = (f"Conçois le plan d'un ebook de {pages} pages environ.\n{title_rule}"
            f"{make_context(ctx)}\n\n"
            f"Réponds UNIQUEMENT avec un objet JSON valide :\n"
            "{\"titre\": \"titre accrocheur et clair (max 70 caractères)\", \"sous_titre\": \"promesse concrète (max 110 caractères)\", "
            "\"pour_qui\": \"1 phrase\", "
            "\"design\": {\"univers\": \"une valeur parmi : " + ", ".join(f"{k} ({v['for']})" for k, v in STYLES.items()) + "\", "
            "\"couleurs\": {\"principale\": \"#hex\", \"sombre\": \"#hex très foncé\", \"claire\": \"#hex très clair\", \"accent\": \"#hex\"}, "
            "\"etiquette\": \"2 ou 3 mots pour la couverture, ex. GUIDE PRATIQUE, PROGRAMME 21 JOURS, MANUEL\"}, "
            "\"cover_prompt\": \"description en ANGLAIS d'une illustration de couverture (scène, ambiance, lumière, sans texte)\", "
            "\"ce_que_tu_vas_obtenir\": [\"5 bénéfices concrets\"], \"avertissement\": \"une phrase si le sujet touche à la santé, l'argent, le droit ou la religion, sinon vide\", "
            f"\"chapitres\": [{{\"titre\": \"\", \"objectif\": \"ce que le lecteur saura faire après\", \"sections\": [\"" + str(n_sec) + " titres de parties\"], "
            "\"image_prompt\": \"description en ANGLAIS d'une illustration pour ce chapitre (scène concrète, sans texte)\", "
            "\"photo_query\": \"3 à 5 mots-clés en ANGLAIS pour trouver une vraie photo dans une banque d'images (ex. : african woman praying sunrise) ; ajoute african si des personnes apparaissent ; jamais le nom d'une personne célèbre : une scène, un lieu ou un objet réel\"}}], "
            "\"bonus\": {\"titre\": \"\", \"type\": \"checklist, plan sur 7 jours, fiches pratiques ou tableau de suivi\"}}\n"
            f"Il faut exactement {n_ch} chapitres, dans un ordre logique du plus simple au plus avancé. "
            "Direction artistique : choisis l'univers qui colle le mieux au sujet et à l'émotion recherchée, et invente une "
            "palette de 4 couleurs harmonieuses propre à CE livre (pas des couleurs génériques). "
            "Les illustrations : si des personnes apparaissent ce sont des personnes noires africaines ; "
            "ne représente jamais Dieu, un prophète ou une figure sacrée (préfère des scènes symboliques : lumière, mains, nature).")
    from_json = chat([{"role": "system", "content": system}, {"role": "user", "content": user}],
                     max_tokens=min(8000, 2200 + 260 * n_ch), json_mode=True)
    return from_json, n_ch


def gen_titles(chat, ctx, n=5):
    """Propose des titres d'ebook accrocheurs pour cette audience."""
    user = (f"Propose {n} titres d'ebook différents et accrocheurs (max 60 caractères chacun) pour ce produit : "
            f"« {ctx['title']} ».\n{make_context(ctx)}\n"
            "Varie les angles : bénéfice concret, chiffre ou durée, question, promesse rassurante, mot fort. "
            "Pas de promesse mensongère. Réponds UNIQUEMENT en JSON : {\"titres\": [\"\"]}")
    raw = chat([{"role": "system", "content": "Tu es un expert en titres de livres qui se vendent auprès des audiences "
                 "francophones africaines et de la diaspora."}, {"role": "user", "content": user}],
               max_tokens=800, json_mode=True)
    return raw


def gen_section(chat, plan, ctx, kind, words, idx=None, prev="", part=None):
    outline = "\n".join(f"{k + 1}. {c.get('titre', '')}" for k, c in enumerate(plan.get("chapitres") or []))
    base = (f"Ebook : « {plan.get('titre')} » — {plan.get('sous_titre', '')}\nPour qui : {plan.get('pour_qui', '')}\n"
            f"{make_context(ctx)}\n\nPlan complet :\n{outline}\n")
    if kind == "intro":
        task = (f"Écris l'INTRODUCTION (environ {words} mots) : accroche sur le problème vécu par le lecteur, ce qu'il va "
                "obtenir, comment utiliser ce guide, et une phrase d'encouragement. Pas de sous-titres « ## ».")
    elif kind == "chapter":
        ch = plan["chapitres"][idx]
        task = (f"Écris le CHAPITRE {idx + 1} « {ch.get('titre')} » (environ {words} mots).\nObjectif : {ch.get('objectif', '')}\n"
                f"Parties à couvrir, chacune avec « ## » : {' | '.join(ch.get('sections') or [])}\n"
                + (f"Le chapitre précédent a déjà couvert : {prev}. Ne le répète pas.\n" if prev else "")
                + "Termine par un encadré « :::retenir À retenir » de 3 à 5 points.")
    elif kind == "part":
        ch = plan["chapitres"][idx]
        secs = ch.get("sections") or []
        task = (f"Chapitre {idx + 1} « {ch.get('titre')} » — objectif : {ch.get('objectif', '')}\n"
                f"Parties du chapitre : {' | '.join(secs)}\n"
                f"Écris UNIQUEMENT la partie {part + 1} « {secs[part]} » (environ {words} mots), en commençant par la ligne "
                f"« ## {secs[part]} ». Va en profondeur : explications, exemples concrets, étapes, au moins un encadré. "
                + ("Ne fais pas d'introduction générale du chapitre, elle est déjà écrite.\n" if part else "")
                + ("Termine cette dernière partie par un encadré « :::retenir À retenir » de 3 à 5 points sur tout le chapitre."
                   if part == len(secs) - 1 else "Ne conclus pas le chapitre : d'autres parties suivent."))
    elif kind == "bonus":
        b = plan.get("bonus") or {}
        task = (f"Écris le BONUS « {b.get('titre', 'Plan d’action')} » ({b.get('type', 'checklist')}), environ {words} mots : "
                "outil pratique prêt à l'emploi (checklists « - [ ] », plan jour par jour ou tableau). Court texte d'intro, puis l'outil.")
    else:
        task = (f"Écris la CONCLUSION (environ {words} mots) : résumé du chemin parcouru, les 3 actions à faire dès demain, "
                "et un message motivant. Pas de sous-titres « ## ».")
    system = "Tu es l'auteur de cet ebook, expert du sujet. " + STYLE_RULES + "\n" + FORMAT_RULES
    tokens = int(min(8000, words * 2.3 + 900))
    txt = chat([{"role": "system", "content": system}, {"role": "user", "content": base + "\n" + task}], max_tokens=tokens)
    return tidy_md(txt)


# ---------------------------------------------------------------- illustrations
IMG_STYLE = (", high detail, professional book illustration, any people shown are Black African, "
             "absolutely no text, no letters, no words, no watermark, no logo")


def cover_size(uni):
    return {"bleed": (896, 1280), "diagonal": (896, 1280), "arch": (896, 1152), "medallion": (1024, 1024),
            "card": (1024, 960), "magazine": (1152, 976)}[STYLES[uni]["cover"]]


DOWN = {}  # service d'images en panne ou payant → on ne le réessaie pas avant cette heure


def _down(name):
    return time.time() < DOWN.get(name, 0)


def fetch_image(prompt, w, h, seed, path, log=print, uni="moderne"):
    if _down("pollinations"):
        return False
    style = ", " + STYLES.get(uni, STYLES["moderne"])["img"]
    url = ("https://image.pollinations.ai/prompt/" + urllib.parse.quote(clean(prompt)[:500] + style + IMG_STYLE, safe="")
           + f"?width={w}&height={h}&seed={seed}&nologo=true&model=flux&private=true")
    for attempt in range(3):
        try:
            r = requests.get(url, timeout=120, headers={"User-Agent": "Kopi-ebook/1.0"})
            if r.status_code == 200 and r.headers.get("content-type", "").startswith("image") and len(r.content) > 15000:
                im = Image.open(io.BytesIO(r.content)).convert("RGB")
                if im.width >= 300:
                    im.save(path, "JPEG", quality=86, optimize=True)
                    return True
            log(f"[ebook] image refusée ({r.status_code}, {len(r.content)} o)")
            if r.status_code in (401, 402, 403):
                DOWN["pollinations"] = time.time() + 3 * 3600
                log("[ebook] Pollinations payant : ignoré pendant 3 h")
                return False
        except Exception as e:
            log(f"[ebook] image en échec : {type(e).__name__} {str(e)[:80]}")
            if "Timeout" in type(e).__name__:
                DOWN["pollinations"] = time.time() + 3600
                log("[ebook] Pollinations trop lent : ignoré pendant 1 h")
                return False
        time.sleep(6 + attempt * 6)
    return False


GEMINI_MODELS = [m.strip() for m in os.environ.get("GEMINI_IMAGE_MODELS", "gemini-nano-banana-2.1,gemini-3.1-flash-image,gemini-2.5-flash-image").split(",") if m.strip()]
ASPECTS = {"1:1": 1, "3:2": 1.5, "2:3": .667, "3:4": .75, "4:3": 1.333, "4:5": .8, "5:4": 1.25, "9:16": .5625, "16:9": 1.778}


def nearest_aspect(w, h):
    r = w / h
    return min(ASPECTS, key=lambda k: abs(ASPECTS[k] - r))


def _find_image_b64(obj, depth=0):
    """Cherche l'image (base64) dans une réponse Gemini, quel que soit le format exact."""
    if depth > 10:
        return None
    if isinstance(obj, dict):
        inl = obj.get("inlineData") or obj.get("inline_data")
        if isinstance(inl, dict) and inl.get("data"):
            return inl["data"]
        if obj.get("data") and isinstance(obj["data"], str) and len(obj["data"]) > 1000 and \
                (obj.get("type") == "image" or str(obj.get("mime_type") or obj.get("mimeType") or "").startswith("image")):
            return obj["data"]
        vals = obj.values()
    elif isinstance(obj, list):
        vals = obj
    else:
        return None
    for v in vals:
        r = _find_image_b64(v, depth + 1)
        if r:
            return r
    return None


def fetch_gemini(prompt, w, h, path, key, log=print, uni="moderne", size="2K"):
    """Image sur mesure par l'IA de Google (Nano Banana). Renvoie True si l'image est enregistrée."""
    import base64
    style = STYLES.get(uni, STYLES["moderne"])["img"]
    full = (f"{clean(prompt)[:700]}. Style: {style}. Photorealistic or high-end editorial illustration, beautiful composition, "
            "professional book cover art quality, any people shown are Black African, no text, no letters, no words, no logo, no watermark.")
    if _down("gemini_img"):
        return False
    aspect = nearest_aspect(w, h)
    hdr = {"x-goog-api-key": key, "Content-Type": "application/json"}
    zero = 0
    for model in GEMINI_MODELS:
        attempts = (
            ("https://generativelanguage.googleapis.com/v1beta/interactions",
             {"model": model, "input": full, "response_format": {"type": "image", "mime_type": "image/jpeg",
                                                                  "aspect_ratio": aspect, "image_size": size}}),
            (f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
             {"contents": [{"parts": [{"text": full}]}],
              "generationConfig": {"responseModalities": ["IMAGE"], "imageConfig": {"aspectRatio": aspect}}}),
        )
        for url, body in attempts:
            try:
                r = requests.post(url, headers=hdr, json=body, timeout=180)
            except Exception as e:
                log(f"[ebook] Gemini {model} : {type(e).__name__}")
                continue
            if r.status_code == 200:
                b64 = _find_image_b64(r.json())
                if b64:
                    im = Image.open(io.BytesIO(base64.b64decode(b64))).convert("RGB")
                    im.save(path, "JPEG", quality=92)
                    log(f"[ebook] image Google OK ({model}, {im.width}x{im.height})")
                    return True
                log(f"[ebook] Gemini {model} : réponse sans image")
            elif r.status_code in (401, 403):
                log(f"[ebook] Gemini refusé ({r.status_code}) : clé invalide ou facturation non activée · {r.text[:160]}")
                return False
            else:
                log(f"[ebook] Gemini {model} {url.rsplit('/', 1)[-1][:20]} : {r.status_code} {r.text[:140]}")
                if r.status_code == 429 and "limit: 0" in r.text:
                    zero += 1
    if zero:
        DOWN["gemini_img"] = time.time() + 6 * 3600
        log("[ebook] Images Google indisponibles sur l'offre gratuite : ignorées pendant 6 h")
    return False


def fetch_pexels(query, path, key, used, log=print, orientation="landscape"):
    """Vraie photo Pexels pour ces mots-clés (jamais deux fois la même dans un livre). Renvoie le crédit ou None."""
    q = re.sub(r"[^\w\s-]", " ", clean(query))[:80].strip()
    tries = [q] + ([" ".join(q.split()[:3])] if len(q.split()) > 3 else [])
    for qq in tries:
        try:
            r = requests.get("https://api.pexels.com/v1/search", headers={"Authorization": key}, timeout=30,
                             params={"query": qq, "orientation": orientation, "per_page": 15, "size": "large"})
        except Exception as e:
            log(f"[ebook] Pexels : {type(e).__name__}"); return None
        if r.status_code in (401, 403):
            log("[ebook] Pexels : clé refusée"); return None
        if r.status_code != 200:
            log(f"[ebook] Pexels {r.status_code}"); continue
        for ph in r.json().get("photos") or []:
            if ph["id"] in used:
                continue
            try:
                img = requests.get(ph["src"].get("large2x") or ph["src"]["original"], timeout=60)
                im = Image.open(io.BytesIO(img.content)).convert("RGB")
                im.save(path, "JPEG", quality=90)
                used.add(ph["id"])
                return {"name": ph.get("photographer", ""), "url": ph.get("url", ""), "src": "Pexels"}
            except Exception:
                continue
    return None


def fetch_pixabay(query, path, key, used, log=print, orientation="horizontal"):
    """Vraie photo Pixabay (téléchargée sur le serveur, jamais deux fois la même). Renvoie le crédit ou None."""
    q = re.sub(r"[^\w\s-]", " ", clean(query))[:95].strip()
    tries = [q] + ([" ".join(q.split()[:3])] if len(q.split()) > 3 else []) + ([" ".join(q.split()[:2])] if len(q.split()) > 2 else [])
    for qq in tries:
        try:
            r = requests.get("https://pixabay.com/api/", timeout=30,
                             params={"key": key, "q": qq, "image_type": "photo", "orientation": orientation,
                                     "safesearch": "true", "per_page": 30, "min_width": 1200, "order": "popular"})
        except Exception as e:
            log(f"[ebook] Pixabay : {type(e).__name__}"); return None
        if r.status_code in (400, 401, 403):
            log(f"[ebook] Pixabay refusé ({r.status_code}) : {r.text[:120]}"); return None
        if r.status_code != 200:
            log(f"[ebook] Pixabay {r.status_code}"); continue
        for ph in r.json().get("hits") or []:
            if ("pb", ph["id"]) in used:
                continue
            try:
                img = requests.get(ph.get("fullHDURL") or ph["largeImageURL"], timeout=60)
                im = Image.open(io.BytesIO(img.content)).convert("RGB")
                if im.width < 900:
                    continue
                im.save(path, "JPEG", quality=90)
                used.add(("pb", ph["id"]))
                return {"name": ph.get("user", ""), "url": ph.get("pageURL", ""), "src": "Pixabay"}
            except Exception:
                continue
    return None


def fetch_openverse(query, path, used, log=print, orientation="landscape"):
    """Vraie photo d'Internet sous licence libre autorisant l'usage commercial (Wikimedia, Flickr…), via Openverse."""
    q = re.sub(r"[^\w\s-]", " ", clean(query))[:90].strip()
    tries = [q] + ([" ".join(q.split()[:3])] if len(q.split()) > 3 else []) + ([" ".join(q.split()[:2])] if len(q.split()) > 2 else [])
    for qq in tries:
        try:
            r = requests.get("https://api.openverse.org/v1/images/", timeout=40,
                             headers={"User-Agent": "Kopi-ebook/1.0"},
                             params={"q": qq, "license_type": "commercial,modification", "page_size": 20, "mature": "false",
                                     "aspect_ratio": "wide" if orientation == "landscape" else "tall", "size": "large"})
        except Exception as e:
            log(f"[ebook] Openverse : {type(e).__name__}"); return None
        if r.status_code == 429:
            log("[ebook] Openverse : trop de demandes, pause"); time.sleep(20); continue
        if r.status_code != 200:
            log(f"[ebook] Openverse {r.status_code} {r.text[:100]}"); continue
        for ph in r.json().get("results") or []:
            if ("ov", ph.get("id")) in used or (ph.get("width") or 0) < 1000:
                continue
            try:
                img = requests.get(ph["url"], timeout=60, headers={"User-Agent": "Kopi-ebook/1.0"})
                im = Image.open(io.BytesIO(img.content)).convert("RGB")
                if im.width < 900:
                    continue
                im.thumbnail((2000, 2000))
                im.save(path, "JPEG", quality=90)
                used.add(("ov", ph.get("id")))
                lic = (ph.get("license") or "").upper()
                lic = "domaine public" if lic in ("CC0", "PDM") else f"CC {lic} {ph.get('license_version') or ''}".strip()
                return {"name": ph.get("creator") or "auteur inconnu", "url": ph.get("foreign_landing_url") or "",
                        "src": "Openverse", "title": ph.get("title") or "", "license": lic,
                        "provider": (ph.get("source") or "").replace("_", " ").title()}
            except Exception:
                continue
    return None


def make_art(pal, path, w, h, seed):
    """Visuel graphique de secours aux couleurs du livre (dégradé + formes)."""
    rnd = random.Random(seed)
    a, b, c = hexrgb(pal["main"]), hexrgb(pal["dark"]), hexrgb(pal["accent"])
    im = Image.new("RGB", (w, h), b)
    px = im.load()
    for y in range(h):
        for x in range(0, w, 2):
            t = min(1, max(0, (x / w) * .55 + (y / h) * .45))
            col = tuple(int(b[k] + (a[k] - b[k]) * t) for k in range(3))
            px[x, y] = col
            if x + 1 < w:
                px[x + 1, y] = col
    layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)
    for _ in range(7):
        r = rnd.randint(int(h * .2), int(h * .7))
        x, y = rnd.randint(-r // 2, w), rnd.randint(-r // 2, h)
        col = rnd.choice([c, a, (255, 255, 255)])
        d.ellipse([x - r, y - r, x + r, y + r], fill=col + (rnd.randint(18, 60),))
    for k in range(5):
        r = int(h * (.35 + k * .12))
        cx, cy = int(w * .78), int(h * .3)
        d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=(255, 255, 255, 40), width=max(2, w // 400))
    layer = layer.filter(ImageFilter.GaussianBlur(2))
    im = Image.alpha_composite(im.convert("RGBA"), layer).convert("RGB")
    im.save(path, "JPEG", quality=88)
    return True


# ---------------------------------------------------------------- PDF
def crop_to(path, aspect, suffix):
    """Recadre une image au ratio largeur/hauteur demandé (sans déformer)."""
    out = os.path.splitext(path)[0] + f"_{suffix}.jpg"
    if os.path.exists(out):
        return out
    im = Image.open(path).convert("RGB")
    if im.width / im.height > aspect:
        nw = int(im.height * aspect)
        im = im.crop(((im.width - nw) // 2, 0, (im.width - nw) // 2 + nw, im.height))
    else:
        nh = int(im.width / aspect)
        im = im.crop((0, (im.height - nh) // 2, im.width, (im.height - nh) // 2 + nh))
    im.save(out, "JPEG", quality=90)
    return out


def veil(path, dark, start=.32, strength=.94, suffix="veil"):
    """Image A4 avec voile sombre progressif vers le bas (titre lisible)."""
    out = os.path.splitext(path)[0] + f"_{suffix}.jpg"
    if os.path.exists(out):
        return out
    im = Image.open(crop_to(path, 595 / 842, "a4")).convert("RGB").resize((1240, 1754), Image.LANCZOS)
    tw, th = im.size
    mask = Image.new("L", (1, th))
    for yy in range(th):
        t = (yy / th - start) / .5
        mask.putpixel((0, yy), int(255 * max(0, min(strength, t * 1.05))))
    im = Image.composite(Image.new("RGB", (tw, th), hexrgb(dark)), im, mask.resize((tw, th)))
    im.save(out, "JPEG", quality=90)
    return out


class Marker(Flowable):
    """Puce dessinée : point, carré, losange, cercle ou tiret selon l'univers."""
    def __init__(self, kind, color, lead):
        super().__init__()
        self.kind, self.color, self.lead = kind, color, lead
        self.width, self.height = 10, lead

    def draw(self):
        c = self.canv
        y = self.lead - self.lead * .52
        c.setFillColor(self.color); c.setStrokeColor(self.color)
        if self.kind == "square":
            c.rect(1, y - 3, 6, 6, stroke=0, fill=1)
        elif self.kind == "diamond":
            p = c.beginPath(); p.moveTo(4, y + 4); p.lineTo(8, y); p.lineTo(4, y - 4); p.lineTo(0, y); p.close()
            c.drawPath(p, stroke=0, fill=1)
        elif self.kind == "circle":
            c.setLineWidth(1.4); c.circle(4, y, 3.2, stroke=1, fill=0)
        elif self.kind == "dash":
            c.setLineWidth(1.6); c.line(0, y, 8, y)
        else:
            c.circle(3.5, y, 2.8, stroke=0, fill=1)


class NumBadge(Flowable):
    def __init__(self, n, color, font, shape="circle", lead=16, text_color=colors.white):
        super().__init__()
        self.n, self.color, self.font, self.shape, self.lead, self.tc = n, color, font, shape, lead, text_color
        self.width, self.height = 16, lead

    def draw(self):
        c, s = self.canv, 15
        y0 = self.lead - s - 1
        c.setFillColor(self.color)
        if self.shape == "square":
            c.rect(0, y0, s, s, stroke=0, fill=1)
        elif self.shape == "plain":
            c.setFont(self.font, 13); c.drawString(0, y0 + 2, f"{self.n}.")
            return
        else:
            c.circle(s / 2, y0 + s / 2, s / 2, stroke=0, fill=1)
        c.setFillColor(self.tc); c.setFont(self.font, 8.6)
        c.drawCentredString(s / 2, y0 + s / 2 - 3.1, str(self.n))


class CheckBox(Flowable):
    def __init__(self, size, color, round_=False):
        super().__init__()
        self.size, self.color, self.round = size, color, round_
        self.width = self.height = size

    def draw(self):
        self.canv.setStrokeColor(self.color); self.canv.setLineWidth(1.2)
        self.canv.roundRect(0, 0, self.size, self.size, self.size / 2 if self.round else 2.5, stroke=1, fill=0)


class Ornament(Flowable):
    """Ligne — losange — ligne, centrée."""
    def __init__(self, width, color, w=70):
        super().__init__()
        self.cw, self.color, self.w = width, color, w
        self.width, self.height = width, 14

    def draw(self):
        c, cx, y = self.canv, self.cw / 2, 7
        c.setStrokeColor(self.color); c.setFillColor(self.color); c.setLineWidth(.9)
        c.line(cx - self.w, y, cx - 9, y); c.line(cx + 9, y, cx + self.w, y)
        p = c.beginPath(); p.moveTo(cx, y + 4.5); p.lineTo(cx + 4.5, y); p.lineTo(cx, y - 4.5); p.lineTo(cx - 4.5, y); p.close()
        c.drawPath(p, stroke=0, fill=1)


class ShapedImage(Flowable):
    """Image découpée en arche, en cercle ou en rectangle arrondi, avec liseré optionnel."""
    def __init__(self, path, width, height, shape, frame_w, ring=None):
        super().__init__()
        self.path, self.w, self.h, self.shape, self.fw, self.ring = path, width, height, shape, frame_w, ring
        self.width, self.height = frame_w, height + (14 if ring else 0)

    def _path(self, c, x, y, w, h, pad=0):
        p = c.beginPath()
        x, y, w, h = x - pad, y - pad, w + 2 * pad, h + 2 * pad
        if self.shape == "arch":
            r = w / 2
            p.moveTo(x, y); p.lineTo(x + w, y); p.lineTo(x + w, y + h - r)
            p.arcTo(x, y + h - 2 * r, x + w, y + h, 0, 180); p.lineTo(x, y); p.close()
        elif self.shape == "circle":
            p.circle(x + w / 2, y + h / 2, min(w, h) / 2)
        else:
            p.roundRect(x, y, w, h, 16 + pad)
        return p

    def draw(self):
        c = self.canv
        x, y = (self.fw - self.w) / 2, (7 if self.ring else 0)
        img = crop_to(self.path, self.w / self.h, f"s{int(self.w)}x{int(self.h)}")
        if self.ring:
            c.setStrokeColor(self.ring); c.setLineWidth(1.1)
            c.drawPath(self._path(c, x, y, self.w, self.h, pad=6), stroke=1, fill=0)
        c.saveState()
        c.clipPath(self._path(c, x, y, self.w, self.h), stroke=0, fill=0)
        c.drawImage(ImageReader(img), x, y, self.w, self.h)
        c.restoreState()


class GhostNumber(Flowable):
    """Énorme numéro de chapitre en filigrane (univers Impact)."""
    def __init__(self, n, width, font, color, label_font, label_color):
        super().__init__()
        self.n, self.cw, self.font, self.color, self.lf, self.lc = n, width, font, color, label_font, label_color
        self.width, self.height = width, 118

    def draw(self):
        c = self.canv
        c.setFillColor(self.color); c.setFont(self.font, 170)
        c.drawRightString(self.cw + 4, -6, f"{self.n:02d}")
        c.setFillColor(self.lc); c.setFont(self.lf, 10)
        c.drawString(0, 6, f"CHAPITRE {self.n:02d}")


class Doc(BaseDocTemplate):
    def afterFlowable(self, f):
        if isinstance(f, Paragraph) and getattr(f, "_toc", None):
            self.notify("TOCEntry", (0, f._toc, self.page))


ROMAN = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV"]


SAFE = {"on": False}


def build_pdf(book, path):
    """Construit le PDF ; si un bloc refuse de tenir dans la page, on recommence en mode « sécurité »
    (encadrés et citations en texte simple) plutôt que de perdre le livre."""
    try:
        SAFE["on"] = False
        return _build_pdf(book, path)
    except Exception as e:
        if "too large" not in str(e) and "Flowable" not in str(e) and "LayoutError" not in type(e).__name__:
            raise
        print(f"[ebook] mise en page : un bloc ne tient pas dans la page, reconstruction en mode sécurité ({str(e)[:100]})", flush=True)
        SAFE["on"] = True
        try:
            return _build_pdf(book, path)
        finally:
            SAFE["on"] = False


def _build_pdf(book, path):
    register_fonts()
    uni = book.get("univers") if book.get("univers") in STYLES else "moderne"
    S = STYLES[uni]
    pal = book.get("colors") or PALETTES[uni]
    MAIN, DARK, LIGHT, ACC = (colors.HexColor(pal[k]) for k in ("main", "dark", "light", "accent"))
    INK, MUTED = colors.HexColor("#1D1B19"), colors.HexColor("#6E6A64")
    PAPER = colors.HexColor("#FFFFFF")
    W, H = A4
    M = 2.1 * cm
    CW = W - 2 * M
    title = plain(book["titre"])
    up = (lambda t: t.upper()) if S["upper"] else (lambda t: t)
    label_txt = plain(book.get("etiquette") or "Guide pratique")
    author = plain(book.get("auteur") or "")

    st = {
        "body": ParagraphStyle("body", fontName=S["body"], fontSize=S["size"], leading=S["lead"], textColor=INK, spaceAfter=7.5),
        "h2": ParagraphStyle("h2", fontName=S["head"], fontSize=14.5 if S["head"] == "Archivo" else 17 if S["head"] in ("DMS", "Marc") else 15.5,
                             leading=21, textColor=DARK if uni != "impact" else colors.HexColor("#111111"), spaceBefore=15, spaceAfter=7),
        "h3": ParagraphStyle("h3", fontName=S["sub"], fontSize=12.2, leading=16, textColor=MAIN, spaceBefore=9, spaceAfter=5),
        "li": ParagraphStyle("li", fontName=S["body"], fontSize=S["size"], leading=S["lead"] - .6, textColor=INK),
        "quote": ParagraphStyle("quote", fontName=S["it"], fontSize=13, leading=20, textColor=DARK),
        "qsrc": ParagraphStyle("qsrc", fontName=S["lab"], fontSize=9, leading=12, textColor=MAIN, spaceBefore=4),
        "eyebrow": ParagraphStyle("eyebrow", fontName=S["lab"], fontSize=8.6, leading=11, textColor=MAIN, spaceAfter=4),
        "ctitle": ParagraphStyle("ctitle", fontName=S["display"], fontSize=40 if uni == "impact" else 27, leading=40 if uni == "impact" else 32,
                                 textColor=DARK, spaceAfter=6),
        "obj": ParagraphStyle("obj", fontName=S["it"], fontSize=11, leading=16.5, textColor=MUTED),
        "ctl": ParagraphStyle("ctl", fontName=S["lab"], fontSize=9.6, leading=13, textColor=MAIN, spaceAfter=3),
        "cbody": ParagraphStyle("cbody", fontName=S["body"], fontSize=S["size"] - .4, leading=S["lead"] - 1.4, textColor=INK, spaceAfter=4),
        "th": ParagraphStyle("th", fontName=S["lab"], fontSize=9.4, leading=12.5, textColor=colors.white),
        "td": ParagraphStyle("td", fontName=S["body"], fontSize=9.4, leading=13, textColor=INK),
        "big": ParagraphStyle("big", fontName=S["display"], fontSize=30 if uni == "impact" else 24, leading=31 if uni == "impact" else 29,
                              textColor=DARK, spaceAfter=10),
        "toc": ParagraphStyle("toc", fontName=S["sub"], fontSize=11.5, leading=17, textColor=INK, spaceAfter=6),
    }
    center = lambda s, **kw: ParagraphStyle(s.name + "c", parent=s, alignment=TA_CENTER, **kw)

    def eyebrow(t, style=None, centered=False):
        txt = " ".join(t.upper())if uni in ("sacre", "luxe") else t.upper()
        sty = style or st["eyebrow"]
        return Paragraph(txt, center(sty) if centered else sty)

    def para(c, text, style, x, y, w, align=TA_LEFT):
        p = Paragraph(text, ParagraphStyle("x", parent=style, alignment=align))
        _, ph = p.wrap(w, H)
        p.drawOn(c, x, y - ph)
        return ph

    def fit_draw(c, img, x, y, w, h):
        c.drawImage(ImageReader(crop_to(img, w / h, f"c{int(w)}x{int(h)}")), x, y, w, h)

    cover = book.get("cover_img") if book.get("cover_img") and os.path.exists(book["cover_img"]) else None
    tsize = 1 if len(title) < 34 else .86 if len(title) < 52 else .74

    # ================================================================ couvertures
    def cover_bleed(c):
        if cover:
            c.drawImage(ImageReader(veil(cover, pal["dark"])), 0, 0, W, H)
        else:
            c.setFillColor(DARK); c.rect(0, 0, W, H, fill=1, stroke=0)
        c.setFillColor(ACC); c.rect(M, H * .43, 2.2 * cm, .17 * cm, fill=1, stroke=0)
        y = H * .43 - .6 * cm
        y -= para(c, inline(book["titre"]), ParagraphStyle("t", fontName="Pop-Xb", fontSize=36 * tsize, leading=41 * tsize,
                                                           textColor=colors.white), M, y, CW) + .45 * cm
        para(c, inline(book.get("sous_titre") or ""), ParagraphStyle("s", fontName="Pop-Md", fontSize=13.5, leading=19,
                                                                     textColor=colors.HexColor("#F1ECE6")), M, y, CW - 1 * cm)
        c.setFont("Pop-Sb", 10.5); c.setFillColor(colors.white); c.drawString(M, 1.7 * cm, author.upper())
        c.setFillColor(ACC); c.setFont("Pop-Sb", 8.5); c.drawRightString(W - M, 1.7 * cm, label_txt.upper())

    def cover_diagonal(c):
        if cover:
            fit_draw(c, cover, 0, H * .28, W, H * .72)
        else:
            c.setFillColor(DARK); c.rect(0, H * .28, W, H * .72, fill=1, stroke=0)
        p = c.beginPath(); p.moveTo(0, 0); p.lineTo(W, 0); p.lineTo(W, H * .47); p.lineTo(0, H * .33); p.close()
        c.setFillColor(MAIN); c.drawPath(p, stroke=0, fill=1)
        p = c.beginPath(); p.moveTo(0, H * .33); p.lineTo(W, H * .47); p.lineTo(W, H * .485); p.lineTo(0, H * .345); p.close()
        c.setFillColor(ACC); c.drawPath(p, stroke=0, fill=1)
        c.setFillColor(colors.HexColor("#111111")); c.rect(M, H - 2.6 * cm, 4.6 * cm, .9 * cm, fill=1, stroke=0)
        c.setFillColor(ACC); c.setFont("Bebas", 15); c.drawString(M + .3 * cm, H - 2.33 * cm, label_txt.upper())
        y = H * .31
        y -= para(c, inline(book["titre"]).upper(), ParagraphStyle("t", fontName="Bebas", fontSize=64 * tsize, leading=58 * tsize,
                                                                   textColor=colors.white), M, y, CW) + .35 * cm
        para(c, inline(book.get("sous_titre") or ""), ParagraphStyle("s", fontName="Pop-Sb", fontSize=12.5, leading=17.5,
                                                                     textColor=colors.white), M, y, CW - 1.5 * cm)
        c.setFillColor(colors.white); c.setFont("Bebas", 17); c.drawString(M, 1.5 * cm, author.upper())

    def cover_arch(c):
        c.setFillColor(DARK); c.rect(0, 0, W, H, fill=1, stroke=0)
        c.setStrokeColor(ACC); c.setLineWidth(.6)
        c.rect(.9 * cm, .9 * cm, W - 1.8 * cm, H - 1.8 * cm, stroke=1, fill=0)
        c.setStrokeColor(ACC, alpha=.35); c.rect(1.1 * cm, 1.1 * cm, W - 2.2 * cm, H - 2.2 * cm, stroke=1, fill=0)
        aw, ah = W * .58, H * .46
        ax, ay = (W - aw) / 2, H * .43
        img = ShapedImage(cover, aw, ah, "arch", W, ring=ACC) if cover else None
        if img:
            img.canv = c
            c.saveState(); c.translate(0, ay - 7); img.draw(); c.restoreState()
        y = ay - 1.0 * cm
        c.setFillColor(ACC); c.setFont("Spec-Sb", 8.5)
        c.drawCentredString(W / 2, y, " ".join(label_txt.upper())); y -= .55 * cm
        y -= para(c, inline(book["titre"]), ParagraphStyle("t", fontName="Marc", fontSize=33 * tsize, leading=39 * tsize,
                                                           textColor=colors.white), M, y, CW, TA_CENTER) + .3 * cm
        o = Ornament(W, ACC, 60); o.canv = c
        c.saveState(); c.translate(0, y - 14); o.draw(); c.restoreState(); y -= 24
        para(c, inline(book.get("sous_titre") or ""), ParagraphStyle("s", fontName="Spec-It", fontSize=13.5, leading=19,
                                                                     textColor=colors.HexColor("#E9E2D3")), M + 1 * cm, y, CW - 2 * cm, TA_CENTER)
        c.setFillColor(colors.white); c.setFont("Spec-Sb", 10.5); c.drawCentredString(W / 2, 1.9 * cm, " ".join(author.upper()))

    def cover_medallion(c):
        c.setFillColor(DARK); c.rect(0, 0, W, H, fill=1, stroke=0)
        cx, cy, r = W / 2, H * .64, W * .27
        for k in range(18, 0, -1):
            c.setFillColor(MAIN, alpha=.035); c.circle(cx, cy, r + k * 11, stroke=0, fill=1)
        c.setStrokeColor(ACC); c.setLineWidth(1.2); c.circle(cx, cy, r + 9, stroke=1, fill=0)
        c.setLineWidth(.5); c.circle(cx, cy, r + 15, stroke=1, fill=0)
        if cover:
            c.saveState(); p = c.beginPath(); p.circle(cx, cy, r); c.clipPath(p, stroke=0, fill=0)
            fit_draw(c, cover, cx - r, cy - r, 2 * r, 2 * r); c.restoreState()
        for ang in range(0, 360, 45):  # petites étoiles dorées autour du médaillon
            x, y = cx + (r + 28) * math.cos(math.radians(ang)), cy + (r + 28) * math.sin(math.radians(ang))
            c.setFillColor(ACC); c.circle(x, y, 1.4, stroke=0, fill=1)
        y = cy - r - 1.6 * cm
        c.setFillColor(ACC); c.setFont("Lato-Bd", 8.5); c.drawCentredString(W / 2, y, " ".join(label_txt.upper())); y -= .5 * cm
        y -= para(c, inline(book["titre"]), ParagraphStyle("t", fontName="DMS", fontSize=36 * tsize, leading=40 * tsize,
                                                           textColor=colors.white), M, y, CW, TA_CENTER) + .35 * cm
        c.setStrokeColor(ACC); c.setLineWidth(.8); c.line(W / 2 - 1.2 * cm, y, W / 2 + 1.2 * cm, y); y -= .4 * cm
        para(c, inline(book.get("sous_titre") or ""), ParagraphStyle("s", fontName="Lato-It", fontSize=13, leading=18.5,
                                                                     textColor=colors.HexColor("#D9D2E6")), M + 1 * cm, y, CW - 2 * cm, TA_CENTER)
        c.setFillColor(ACC); c.setFont("Lato-Bd", 10); c.drawCentredString(W / 2, 1.8 * cm, " ".join(author.upper()))

    def cover_card(c):
        c.setFillColor(LIGHT); c.rect(0, 0, W, H, fill=1, stroke=0)
        c.setFillColor(ACC, alpha=.22); c.circle(W * .92, H * .95, 5.2 * cm, stroke=0, fill=1)
        c.setFillColor(MAIN, alpha=.12); c.circle(W * .05, H * .12, 6.5 * cm, stroke=0, fill=1)
        iw, ih, ix, iy = W - 2 * M, H * .5, M, H * .41
        if cover:
            c.saveState(); p = c.beginPath(); p.roundRect(ix, iy, iw, ih, 22); c.clipPath(p, stroke=0, fill=0)
            fit_draw(c, cover, ix, iy, iw, ih); c.restoreState()
        else:
            c.setFillColor(MAIN); c.roundRect(ix, iy, iw, ih, 22, stroke=0, fill=1)
        lw = pdfmetrics.stringWidth(label_txt.upper(), "Lato-Bd", 8.5) + 1 * cm
        c.setFillColor(colors.white); c.roundRect(ix + .5 * cm, iy + ih - 1.25 * cm, lw, .75 * cm, 10, stroke=0, fill=1)
        c.setFillColor(MAIN); c.setFont("Lato-Bd", 8.5); c.drawString(ix + 1 * cm, iy + ih - .98 * cm, label_txt.upper())
        y = iy - .9 * cm
        y -= para(c, inline(book["titre"]), ParagraphStyle("t", fontName="DMS", fontSize=35 * tsize, leading=39 * tsize,
                                                           textColor=DARK), M, y, CW) + .3 * cm
        para(c, inline(book.get("sous_titre") or ""), ParagraphStyle("s", fontName="Lato", fontSize=13, leading=18.5,
                                                                     textColor=colors.HexColor("#5E5458")), M, y, CW - 1.5 * cm)
        c.setFillColor(MAIN); c.circle(M + 4, 1.85 * cm + 3, 4, stroke=0, fill=1)
        c.setFillColor(DARK); c.setFont("Lato-Bd", 10.5); c.drawString(M + 14, 1.85 * cm, author)

    def cover_magazine(c):
        c.setFillColor(colors.HexColor("#FBF8F3")); c.rect(0, 0, W, H, fill=1, stroke=0)
        top = H * .41
        if cover:
            fit_draw(c, cover, 0, top, W, H - top)
        else:
            c.setFillColor(DARK); c.rect(0, top, W, H - top, fill=1, stroke=0)
        c.setFillColor(MAIN); c.rect(0, top - .22 * cm, W, .22 * cm, fill=1, stroke=0)
        y = top - 1.15 * cm
        c.setFillColor(MAIN); c.setFont("Spec-Sb", 9); c.drawString(M, y, label_txt.upper())
        c.setFillColor(MUTED); c.drawRightString(W - M, y, str(datetime.date.today().year)); y -= .45 * cm
        y -= para(c, inline(book["titre"]), ParagraphStyle("t", fontName="Abril", fontSize=38 * tsize, leading=42 * tsize,
                                                           textColor=DARK), M, y, CW) + .3 * cm
        c.setStrokeColor(DARK); c.setLineWidth(.8); c.line(M, y, W - M, y); y -= .35 * cm
        para(c, inline(book.get("sous_titre") or ""), ParagraphStyle("s", fontName="Spec-It", fontSize=13.5, leading=19,
                                                                     textColor=colors.HexColor("#4A453F")), M, y, CW - 2 * cm)
        c.setStrokeColor(MAIN); c.setLineWidth(1.6); c.line(M, 2.35 * cm, M + 1.2 * cm, 2.35 * cm)
        c.setFillColor(DARK); c.setFont("Spec-Sb", 11); c.drawString(M, 1.75 * cm, author)

    COVERS = {"bleed": cover_bleed, "diagonal": cover_diagonal, "arch": cover_arch, "medallion": cover_medallion,
              "card": cover_card, "magazine": cover_magazine}

    def on_cover(c, d):
        c.saveState(); COVERS[S["cover"]](c); c.restoreState()

    # ================================================================ en-têtes et pieds de page
    def on_page(c, d):
        c.saveState()
        h, n = S["header"], str(d.page)
        if h == "topbar":
            c.setFillColor(MAIN); c.rect(0, H - .32 * cm, W, .32 * cm, fill=1, stroke=0)
            c.setFont("Pop-Md", 7.8); c.setFillColor(MUTED); c.drawString(M, 1.15 * cm, title[:80])
            c.setFont("Pop-Sb", 8.5); c.setFillColor(DARK); c.drawRightString(W - M, 1.15 * cm, n)
            c.setStrokeColor(colors.HexColor("#E4DFD8")); c.setLineWidth(.6); c.line(M, 1.55 * cm, W - M, 1.55 * cm)
        elif h == "sidebar":
            c.setFillColor(MAIN); c.rect(0, 0, .42 * cm, H, fill=1, stroke=0)
            c.setFillColor(ACC); c.rect(0, H * .5, .42 * cm, 2.5 * cm, fill=1, stroke=0)
            c.setFont("Bebas", 22); c.setFillColor(colors.HexColor("#111111")); c.drawRightString(W - M, 1.05 * cm, n)
            c.setFont("Pop-Bd", 7.5); c.setFillColor(MUTED); c.drawString(M, 1.15 * cm, title.upper()[:70])
        elif h == "centered":
            c.setFillColor(ACC); c.setFont("Spec-Sb", 9.5); c.drawCentredString(W / 2, 1.1 * cm, n)
            c.setStrokeColor(ACC); c.setLineWidth(.6)
            c.line(W / 2 - 1.5 * cm, 1.22 * cm, W / 2 - .45 * cm, 1.22 * cm); c.line(W / 2 + .45 * cm, 1.22 * cm, W / 2 + 1.5 * cm, 1.22 * cm)
            c.setFont("Spec-It", 8.5); c.setFillColor(MUTED); c.drawCentredString(W / 2, H - 1.25 * cm, title[:80])
        elif h == "gold":
            c.setStrokeColor(ACC); c.setLineWidth(.6)
            c.line(M, H - 1.45 * cm, W - M, H - 1.45 * cm); c.line(M, 1.5 * cm, W - M, 1.5 * cm)
            c.setFont("Lato-Bd", 7.2); c.setFillColor(MUTED); c.drawCentredString(W / 2, H - 1.22 * cm, " ".join(title.upper()[:48]))
            c.setFont("DMS", 11); c.setFillColor(MAIN); c.drawCentredString(W / 2, 1.0 * cm, n)
        elif h == "dot":
            c.setFillColor(LIGHT); c.circle(W / 2, 1.25 * cm, .42 * cm, stroke=0, fill=1)
            c.setFont("Lato-Bd", 9); c.setFillColor(MAIN); c.drawCentredString(W / 2, 1.12 * cm, n)
            c.setFillColor(ACC, alpha=.25); c.circle(W - 1 * cm, H - 1 * cm, 1.4 * cm, stroke=0, fill=1)
        else:  # running
            c.setFont("Spec-It", 8.8); c.setFillColor(MUTED); c.drawString(M, H - 1.3 * cm, title[:80])
            c.setFont("Spec-Sb", 9); c.setFillColor(DARK); c.drawRightString(W - M, H - 1.3 * cm, n)
            c.setStrokeColor(DARK); c.setLineWidth(.5); c.line(M, H - 1.45 * cm, W - M, H - 1.45 * cm)
            c.setFillColor(MAIN); c.rect(M, 1.2 * cm, .9 * cm, .09 * cm, fill=1, stroke=0)
        c.restoreState()

    def on_dark(c, d):
        c.saveState()
        c.setFillColor(DARK); c.rect(0, 0, W, H, fill=1, stroke=0)
        for k in range(14, 0, -1):
            c.setFillColor(MAIN, alpha=.03); c.circle(W / 2, H * .44, 3 * cm + k * 14, stroke=0, fill=1)
        c.setStrokeColor(ACC); c.setLineWidth(.6); c.rect(.9 * cm, .9 * cm, W - 1.8 * cm, H - 1.8 * cm, stroke=1, fill=0)
        c.restoreState()

    doc = Doc(path, pagesize=A4, leftMargin=M, rightMargin=M, topMargin=2.2 * cm, bottomMargin=2.1 * cm,
              title=title, author=author, subject=plain(book.get("sous_titre") or ""), creator="Kopi")
    frame = Frame(M + (.2 * cm if S["header"] == "sidebar" else 0), 2.1 * cm, CW, H - 4.3 * cm, id="f")
    doc.addPageTemplates([PageTemplate("cover", [Frame(0, 0, W, H, id="c")], onPage=on_cover),
                          PageTemplate("page", [frame], onPage=on_page),
                          PageTemplate("dark", [Frame(M, 2.1 * cm, CW, H - 4.3 * cm, id="d")], onPage=on_dark)])

    # ================================================================ blocs de texte
    def boxed(inner, cmds, bar_w=None, pad=(12, 10)):
        """Encadré dont chaque élément est une ligne : il peut se couper sur deux pages au lieu de bloquer le PDF."""
        if SAFE["on"]:
            return inner
        if bar_w:
            rows = [["", x] for x in inner]
            t = Table(rows, colWidths=[bar_w, CW - bar_w])
        else:
            rows = [[x] for x in inner]
            t = Table(rows, colWidths=[CW])
        c = 1 if bar_w else 0
        n = len(rows)
        t.setStyle(TableStyle(cmds + [("TOPPADDING", (c, 0), (c, -1), 1), ("BOTTOMPADDING", (c, 0), (c, -1), 1),
                                      ("TOPPADDING", (c, 0), (c, 0), pad[0]), ("BOTTOMPADDING", (c, n - 1), (c, n - 1), pad[1])]))
        return t

    def callout(kind, ttl, body_blocks):
        label = CALLOUTS.get(kind, "Astuce")
        warn = kind == "attention"
        style = S["callout"]
        if style == "block":
            lab = ParagraphStyle("cl", parent=st["ctl"], fontName="Bebas", fontSize=15, leading=16, textColor=ACC if not warn else colors.white)
            bod = ParagraphStyle("cb", parent=st["cbody"], textColor=colors.white)
            inner = [Paragraph(up(label) + (f" · {up(inline(ttl))}" if ttl else ""), lab)]
            for b in body_blocks:
                inner += render(b, bod, inner_w=CW - 1.2 * cm, on_dark=True)
            t = boxed(inner, [("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#B3261E") if warn else colors.HexColor("#141414")),
                              ("LEFTPADDING", (0, 0), (-1, -1), 16), ("RIGHTPADDING", (0, 0), (-1, -1), 16),
                              ("LINEBEFORE", (0, 0), (0, -1), 5, MAIN if not warn else ACC)])
            if SAFE["on"]:
                inner[0] = Paragraph(up(label) + (f" · {up(inline(ttl))}" if ttl else ""), st["ctl"])
                inner[1:] = [f for b in body_blocks for f in render(b, st["cbody"], inner_w=CW)]
                return [Spacer(1, 6)] + inner + [Spacer(1, 10)]
            return [Spacer(1, 6), t, Spacer(1, 10)]
        if style in ("frame", "rule"):
            centered = style == "frame"
            lab = ParagraphStyle("cl", parent=st["ctl"], alignment=TA_CENTER if centered else TA_LEFT,
                                 textColor=colors.HexColor("#B3261E") if warn else MAIN)
            head = (" ".join(label.upper()) if centered else label.upper()) + (f" · {inline(ttl)}" if ttl else "")
            inner = ([Ornament(CW - 1.6 * cm, ACC, 40)] if centered else []) + [Paragraph(head, lab)]
            for b in body_blocks:
                inner += render(b, st["cbody"], inner_w=CW - 1.6 * cm)
            sty = [("LEFTPADDING", (0, 0), (-1, -1), 18), ("RIGHTPADDING", (0, 0), (-1, -1), 18)]
            if centered:
                sty += [("BOX", (0, 0), (-1, -1), .8, ACC), ("BACKGROUND", (0, 0), (-1, -1), LIGHT)]
            else:
                sty += [("LINEABOVE", (0, 0), (-1, 0), 1.6, MAIN), ("LINEBELOW", (0, -1), (-1, -1), .6, MAIN),
                        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]
            t = boxed(inner, sty)
            return [Spacer(1, 6)] + (t if isinstance(t, list) else [t]) + [Spacer(1, 10)]
        if style == "soft":
            inner = [Paragraph(label.upper() + (f" · {inline(ttl)}" if ttl else ""),
                               ParagraphStyle("cl", parent=st["ctl"], textColor=colors.HexColor("#B3261E") if warn else MAIN))]
            for b in body_blocks:
                inner += render(b, st["cbody"], inner_w=CW - 1.4 * cm)
            t = boxed(inner, [("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FDECEA") if warn else LIGHT),
                              ("ROUNDEDCORNERS", [14, 14, 14, 14]), ("LEFTPADDING", (0, 0), (-1, -1), 16),
                              ("RIGHTPADDING", (0, 0), (-1, -1), 16)])
            return [Spacer(1, 6)] + (t if isinstance(t, list) else [t]) + [Spacer(1, 10)]
        # bar (moderne)
        inner = [Paragraph(label.upper() + (f" · {inline(ttl)}" if ttl else ""), st["ctl"])]
        for b in body_blocks:
            inner += render(b, st["cbody"], inner_w=CW - 1.4 * cm)
        bg = {"attention": colors.HexColor("#FDEDEA"), "exercice": colors.HexColor("#F3F1EC"),
              "action": colors.HexColor("#F3F1EC")}.get(kind, LIGHT)
        bar = colors.HexColor("#C8372D") if warn else ACC if kind in ("exercice", "action") else MAIN
        t = boxed(inner, [("BACKGROUND", (0, 0), (0, -1), bar), ("BACKGROUND", (1, 0), (1, -1), bg),
                          ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (1, 0), (1, -1), 14),
                          ("RIGHTPADDING", (1, 0), (1, -1), 14), ("LEFTPADDING", (0, 0), (0, -1), 0),
                          ("RIGHTPADDING", (0, 0), (0, -1), 0)], bar_w=.18 * cm, pad=(11, 9))
        return [Spacer(1, 5)] + (t if isinstance(t, list) else [t]) + [Spacer(1, 9)]

    def bullets(items, numbered=False, check=False, inner_w=None, style=None, on_dark=False):
        w = inner_w or CW
        li = style or st["li"]
        if SAFE["on"]:
            sty = ParagraphStyle("lisafe", parent=li, leftIndent=14, bulletIndent=0)
            return [Paragraph(inline(it), sty, bulletText=(f"{k}." if numbered else "•"))
                    for k, it in enumerate(items, 1)] + [Spacer(1, 6)]
        mcol = ACC if on_dark else MAIN
        rows = []
        for k, it in enumerate(items, 1):
            if check:
                mark = CheckBox(9.5, mcol, round_=S["marker"] == "circle")
            elif numbered:
                shape = {"square": "square", "dash": "plain"}.get(S["marker"], "circle")
                mark = NumBadge(k, mcol, S["display"] if shape == "plain" else S["lab"], shape, li.leading)
            else:
                mark = Marker(S["marker"], mcol, li.leading)
            rows.append([mark, Paragraph(inline(it), li)])
        t = Table(rows, colWidths=[.75 * cm, w - .75 * cm])
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("TOPPADDING", (0, 0), (-1, -1), 1)]
                              + ([("TOPPADDING", (0, 0), (0, -1), 4)] if check else [])))
        return [t, Spacer(1, 6)]

    def render(b, body=None, inner_w=None, on_dark=False):
        body = body or st["body"]
        t = b["t"]
        if t == "p":
            return [Paragraph(inline(b["text"]), body)]
        if t == "h2":
            return [CondPageBreak(3.2 * cm), Paragraph(up(inline(b["text"])), st["h2"])]
        if t == "h3":
            return [CondPageBreak(2.4 * cm), Paragraph(inline(b["text"]), st["h3"])]
        if t in ("ul", "ol", "check"):
            return bullets(b["items"], numbered=t == "ol", check=t == "check", inner_w=inner_w,
                           style=ParagraphStyle("li2", parent=st["li"], textColor=body.textColor) if on_dark else None, on_dark=on_dark)
        if t == "quote":
            q = "« " + inline(b["text"]).strip("«» \"") + " »"
            qs = S["quote"]
            if qs == "centered":
                inner = [Ornament(CW, ACC, 50), Spacer(1, 4), Paragraph(q, center(st["quote"], fontSize=14, leading=21))]
                if b.get("src"):
                    inner.append(Paragraph(inline(b["src"]), center(st["qsrc"])))
                return [Spacer(1, 8)] + inner + [Spacer(1, 12)]
            if qs == "big":
                inner = [Paragraph(q.upper(), ParagraphStyle("qb", fontName="Bebas", fontSize=26, leading=27, textColor=MAIN))]
                if b.get("src"):
                    inner.append(Paragraph(inline(b["src"]).upper(), ParagraphStyle("qbs", fontName="Pop-Bd", fontSize=8.5, leading=12,
                                                                                    textColor=colors.HexColor("#111111"), spaceBefore=4)))
                return [Spacer(1, 8)] + inner + [Spacer(1, 12)]
            inner = [Paragraph(q, st["quote"])] + ([Paragraph(inline(b["src"]), st["qsrc"])] if b.get("src") else [])
            if SAFE["on"]:
                return [Spacer(1, 6)] + inner + [Spacer(1, 10)]
            tb = Table([["", inner]] if qs == "bar" else [[inner]],
                       colWidths=[.12 * cm, (inner_w or CW) - .12 * cm] if qs == "bar" else [inner_w or CW])
            sty = ([("BACKGROUND", (0, 0), (0, 0), ACC), ("LEFTPADDING", (1, 0), (1, 0), 16), ("LEFTPADDING", (0, 0), (0, 0), 0),
                    ("RIGHTPADDING", (0, 0), (0, 0), 0), ("TOPPADDING", (1, 0), (1, 0), 4), ("BOTTOMPADDING", (1, 0), (1, 0), 4)]
                   if qs == "bar" else
                   [("BACKGROUND", (0, 0), (-1, -1), LIGHT), ("ROUNDEDCORNERS", [16, 16, 16, 16]),
                    ("LEFTPADDING", (0, 0), (-1, -1), 20), ("RIGHTPADDING", (0, 0), (-1, -1), 20),
                    ("TOPPADDING", (0, 0), (-1, -1), 14), ("BOTTOMPADDING", (0, 0), (-1, -1), 12)])
            tb.setStyle(TableStyle(sty))
            return [Spacer(1, 6), tb, Spacer(1, 10)]
        if t == "table":
            rows = b["rows"]
            ncol = max(len(r) for r in rows)
            rows = [r + [""] * (ncol - len(r)) for r in rows]
            data = [[Paragraph(inline(c), st["th"] if i == 0 else st["td"]) for c in r] for i, r in enumerate(rows)]
            tb = Table(data, colWidths=[(inner_w or CW) / ncol] * ncol, repeatRows=1)
            tsty = [("BACKGROUND", (0, 0), (-1, 0), DARK if uni in ("luxe", "editorial") else MAIN), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
                    ("LINEBELOW", (0, 0), (-1, -1), .4, colors.HexColor("#E4DFD8")),
                    ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]
            if uni == "doux":
                tsty.append(("ROUNDEDCORNERS", [10, 10, 10, 10]))
            tb.setStyle(TableStyle(tsty))
            return [Spacer(1, 4), tb, Spacer(1, 10)]
        if t == "callout":
            return callout(b["kind"], b.get("title"), b["body"])
        return []

    def heading_toc(text, style):
        p = Paragraph(up(inline(text)), style)
        p._toc = plain(text)
        return p

    # ================================================================ ouvertures de chapitre
    def opener(k, ch):
        img = ch.get("img") if ch.get("img") and os.path.exists(ch["img"]) else None
        obj = inline(ch.get("objectif") or "")
        o = S["opener"]
        if o == "giant":
            out = [GhostNumber(k, CW, "Bebas", LIGHT, "Pop-Bd", MAIN), Spacer(1, 6),
                   heading_toc(ch["titre"], st["ctitle"])]
            if obj:
                out.append(Paragraph(obj, ParagraphStyle("o", parent=st["obj"], fontName="Pop-Md", textColor=colors.HexColor("#444"))))
            if img:
                out += [Spacer(1, .45 * cm), RLImage(crop_to(img, CW / (7.6 * cm), "g"), width=CW, height=7.6 * cm)]
                bar = Table([[""]], colWidths=[CW], rowHeights=[.2 * cm])
                bar.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), MAIN)]))
                out.append(bar)
            return [PageBreak()] + out + [Spacer(1, .5 * cm)]
        if o == "centered":
            out = [Spacer(1, .6 * cm), Ornament(CW, ACC, 70), Spacer(1, 6),
                   eyebrow(f"Chapitre {ROMAN[k - 1] if k <= len(ROMAN) else k}", centered=True),
                   heading_toc(ch["titre"], center(st["ctitle"]))]
            if obj:
                out.append(Paragraph(obj, center(st["obj"])))
            if img:
                out += [Spacer(1, .5 * cm), ShapedImage(img, CW * .62, 8.4 * cm, "arch", CW, ring=ACC)]
            return [PageBreak()] + out + [Spacer(1, .6 * cm)]
        if o == "dark":
            lt = colors.HexColor("#F4EEF9")
            out = [Spacer(1, 1.2 * cm), eyebrow(f"Chapitre {k:02d}", ParagraphStyle("ed", parent=st["eyebrow"], textColor=ACC), centered=True),
                   heading_toc(ch["titre"], center(st["ctitle"], textColor=colors.white, fontSize=32, leading=37)),
                   Spacer(1, 4), Ornament(CW, ACC, 50), Spacer(1, 6)]
            if obj:
                out.append(Paragraph(obj, center(st["obj"], textColor=colors.HexColor("#CFC6DC"))))
            if img:
                out += [Spacer(1, .9 * cm), ShapedImage(img, 9 * cm, 9 * cm, "circle", CW, ring=ACC)]
            return [NextPageTemplate("dark"), PageBreak()] + out + [NextPageTemplate("page"), PageBreak()]
        if o == "rounded":
            num = Table([[Paragraph(f"{k}", ParagraphStyle("rn", fontName="DMS", fontSize=26, leading=28, textColor=MAIN, alignment=TA_CENTER))]],
                        colWidths=[1.6 * cm], rowHeights=[1.6 * cm])
            num.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), LIGHT), ("ROUNDEDCORNERS", [23, 23, 23, 23]),
                                     ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 0)]))
            head = Table([[num, [eyebrow(f"Chapitre {k}"), heading_toc(ch["titre"], ParagraphStyle("rt", parent=st["ctitle"], fontSize=25, leading=29))]]],
                         colWidths=[2.1 * cm, CW - 2.1 * cm])
            head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
            out = [head]
            if obj:
                out += [Spacer(1, 4), Paragraph(obj, st["obj"])]
            if img:
                out += [Spacer(1, .5 * cm), ShapedImage(img, CW, 8 * cm, "round", CW)]
            return [PageBreak()] + out + [Spacer(1, .55 * cm)]
        if o == "split":
            num = Paragraph(f"{k:02d}", ParagraphStyle("sn", fontName="Abril", fontSize=64, leading=64, textColor=MAIN))
            right = [eyebrow(f"Chapitre {k}"), heading_toc(ch["titre"], ParagraphStyle("st", parent=st["ctitle"], fontSize=26, leading=30))]
            if obj:
                right.append(Paragraph(obj, st["obj"]))
            head = Table([[num, right]], colWidths=[3.3 * cm, CW - 3.3 * cm])
            head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                      ("LINEBEFORE", (1, 0), (1, 0), .8, DARK), ("LEFTPADDING", (1, 0), (1, 0), 16)]))
            out = [head]
            if img:
                out += [Spacer(1, .5 * cm), RLImage(crop_to(img, CW / (8.2 * cm), "sp"), width=CW, height=8.2 * cm)]
                cap = Table([[""]], colWidths=[CW], rowHeights=[.08 * cm])
                cap.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), .8, DARK)]))
                out.append(cap)
            return [PageBreak()] + out + [Spacer(1, .55 * cm)]
        # band (moderne)
        out = [Paragraph(f"{k:02d}", ParagraphStyle("cn", fontName="Pop-Xb", fontSize=54, leading=56, textColor=MAIN)),
               eyebrow(f"Chapitre {k}"), heading_toc(ch["titre"], st["ctitle"])]
        if obj:
            out.append(Paragraph("Objectif : " + obj, st["obj"]))
        if img:
            out += [Spacer(1, .45 * cm), RLImage(crop_to(img, CW / (8.4 * cm), "b"), width=CW, height=8.4 * cm)]
        return [PageBreak()] + out + [Spacer(1, .5 * cm)]

    def section_head(label, text):
        c = uni in ("sacre", "luxe")
        out = [eyebrow(label, centered=c), heading_toc(text, center(st["ctitle"]) if c else st["ctitle"])]
        if c:
            out += [Ornament(CW, ACC, 50)]
        return out + [Spacer(1, .3 * cm)]

    # ================================================================ assemblage
    story = [NextPageTemplate("page"), PageBreak()]
    story += [Spacer(1, 1.2 * cm), eyebrow("Ce guide est pour toi si…"), Paragraph(inline(book.get("pour_qui") or ""), st["big"]), Spacer(1, .3 * cm)]
    if book.get("benefices"):
        story += [Paragraph(up("Ce que tu vas obtenir"), st["h2"])] + bullets(book["benefices"], check=False)
    legal = (f"© {datetime.date.today().year} {author}. Tous droits réservés. "
             "Ce guide est réservé à son acheteur : merci de ne pas le partager ni le revendre.")
    if book.get("avertissement"):
        legal += " " + plain(book["avertissement"])
    story += [Spacer(1, 1.4 * cm), Paragraph(inline(legal), ParagraphStyle("legal", parent=st["body"], fontSize=8.4, leading=12.5, textColor=MUTED))]
    toc = TableOfContents()
    toc.levelStyles = [st["toc"]]
    toc.dotsMinLevel = 0
    story += [PageBreak(), eyebrow("Sommaire"), Paragraph(up("Au programme"), st["big"]), Spacer(1, .2 * cm), toc]
    story += [PageBreak()] + section_head("Introduction", "Avant de commencer")
    for b in parse_md(book.get("intro", "")):
        story += render(b)
    for k, ch in enumerate(book["chapitres"], 1):
        story += opener(k, ch)
        for b in parse_md(ch.get("texte", "")):
            story += render(b)
    if book.get("bonus_texte"):
        story += [PageBreak()] + section_head("Bonus", book.get("bonus_titre") or "Ton plan d'action")
        for b in parse_md(book["bonus_texte"]):
            story += render(b)
    if book.get("conclusion"):
        story += [PageBreak()] + section_head("Conclusion", "Et maintenant ?")
        for b in parse_md(book["conclusion"]):
            story += render(b)
    credits = [c for c in book.get("credits") or [] if c.get("name")]
    if credits:
        bits = []
        for srcname in dict.fromkeys(c.get("src", "Pexels") for c in credits if c.get("src") != "Openverse"):
            names = ", ".join(dict.fromkeys(plain(c["name"]) for c in credits if c.get("src", "Pexels") == srcname))
            bits.append(f"{srcname} ({names})")
        libres = [c for c in credits if c.get("src") == "Openverse"]
        for c in libres:
            t = f"« {plain(c['title'])[:70]} »" if c.get("title") else "Photo"
            bits.append(f"{t} par {plain(c['name'])[:60]}, {c.get('license', '')}" + (f", via {c['provider']}" if c.get("provider") else "")
                        + (f" ({c['url']})" if c.get("url") else ""))
        story += [Spacer(1, .6 * cm), Paragraph(inline("Crédits photos : " + ("photos fournies par " if not libres else "") + " ; ".join(bits) + "."),
                                                ParagraphStyle("cred", parent=st["body"], fontSize=8, leading=11.5, textColor=MUTED))]
    others = [o for o in book.get("autres_produits") or [] if o.get("nom")]
    if others or book.get("boutique"):
        story += [PageBreak(), eyebrow("Pour aller plus loin"), Paragraph(up("Continue ta progression"), st["big"]),
                  Paragraph(inline(f"Merci pour ta confiance. Si ce guide t'a aidé, voici d'autres ressources de "
                                   f"{author or 'l’auteur'} pour aller plus loin :"), st["body"]), Spacer(1, .3 * cm)]
        for o in others[:5]:
            card = [Paragraph(inline(o["nom"]), ParagraphStyle("on", fontName=S["head"], fontSize=13, leading=17, textColor=DARK)),
                    Paragraph(inline(o.get("promesse") or ""), st["cbody"])]
            t = Table([[card]], colWidths=[CW])
            sty = [("BACKGROUND", (0, 0), (-1, -1), LIGHT), ("LEFTPADDING", (0, 0), (-1, -1), 16), ("RIGHTPADDING", (0, 0), (-1, -1), 16),
                   ("TOPPADDING", (0, 0), (-1, -1), 11), ("BOTTOMPADDING", (0, 0), (-1, -1), 9)]
            if uni == "doux":
                sty.append(("ROUNDEDCORNERS", [14, 14, 14, 14]))
            elif uni in ("sacre", "luxe"):
                sty.append(("BOX", (0, 0), (-1, -1), .7, ACC))
            else:
                sty.append(("LINEBEFORE", (0, 0), (0, -1), 4, MAIN))
            t.setStyle(TableStyle(sty))
            story += [KeepTogether([t]), Spacer(1, 8)]
        if book.get("boutique"):
            link = book["boutique"].replace("&", "&amp;")
            story += [Spacer(1, .4 * cm), Paragraph(f'Retrouve-les sur : <a href="{link}" color="{pal["main"]}"><b>{link}</b></a>',
                                                    ParagraphStyle("shop", parent=st["body"], fontSize=12, leading=18))]
    doc.multiBuild(story)


# ---------------------------------------------------------------- Word modifiable
def build_docx(book, path):
    pal = book.get("colors") or PALETTES["moderne"]
    main = RGBColor(*hexrgb(pal["main"]))
    d = Document()
    d.styles["Normal"].font.name, d.styles["Normal"].font.size = "Calibri", Pt(11)
    if book.get("cover_img") and os.path.exists(book["cover_img"]):
        d.add_picture(book["cover_img"], width=Cm(15))
    t = d.add_heading(plain(book["titre"]), 0)
    if book.get("sous_titre"):
        p = d.add_paragraph(); r = p.add_run(plain(book["sous_titre"])); r.italic = True; r.font.size = Pt(13)
    if book.get("auteur"):
        d.add_paragraph(plain(book["auteur"]))
    d.add_page_break()

    def add_blocks(md):
        for b in parse_md(md):
            if b["t"] == "p":
                add_par(b["text"])
            elif b["t"] in ("h2", "h3"):
                d.add_heading(plain(b["text"]), 2 if b["t"] == "h2" else 3)
            elif b["t"] in ("ul", "ol", "check"):
                for it in b["items"]:
                    add_par(("☐ " if b["t"] == "check" else "") + it, "List Number" if b["t"] == "ol" else "List Bullet")
            elif b["t"] == "quote":
                p = d.add_paragraph(); r = p.add_run("« " + plain(b["text"]) + " »"); r.italic = True
                if b.get("src"):
                    p.add_run("\n— " + plain(b["src"])).bold = True
            elif b["t"] == "table":
                rows = b["rows"]; ncol = max(len(r) for r in rows)
                tb = d.add_table(rows=len(rows), cols=ncol); tb.style = "Light Grid Accent 1"
                for i, r in enumerate(rows):
                    for j, c in enumerate(r):
                        tb.cell(i, j).text = plain(c)
            elif b["t"] == "callout":
                p = d.add_paragraph(); r = p.add_run(CALLOUTS.get(b["kind"], "Astuce").upper()
                                                     + (f" · {plain(b['title'])}" if b.get("title") else ""))
                r.bold = True; r.font.color.rgb = main
                for sb in b["body"]:
                    if sb["t"] == "p":
                        add_par(sb["text"])
                    elif sb["t"] in ("ul", "ol", "check"):
                        for it in sb["items"]:
                            add_par(it, "List Bullet")

    def add_par(text, style=None):
        p = d.add_paragraph(style=style)
        for part in re.split(r"(\*\*.+?\*\*)", clean(text)):
            if part.startswith("**") and part.endswith("**"):
                p.add_run(part[2:-2]).bold = True
            elif part:
                p.add_run(part.replace("*", ""))

    d.add_heading("Introduction", 1); add_blocks(book.get("intro", ""))
    for k, ch in enumerate(book["chapitres"], 1):
        d.add_page_break()
        d.add_heading(f"Chapitre {k} — {plain(ch['titre'])}", 1)
        if ch.get("img") and os.path.exists(ch["img"]):
            d.add_picture(ch["img"], width=Cm(16))
        if ch.get("objectif"):
            p = d.add_paragraph(); r = p.add_run("Objectif : " + plain(ch["objectif"])); r.italic = True
        add_blocks(ch.get("texte", ""))
    if book.get("bonus_texte"):
        d.add_page_break(); d.add_heading(plain(book.get("bonus_titre") or "Bonus"), 1); add_blocks(book["bonus_texte"])
    if book.get("conclusion"):
        d.add_page_break(); d.add_heading("Et maintenant ?", 1); add_blocks(book["conclusion"])
    d.save(path)

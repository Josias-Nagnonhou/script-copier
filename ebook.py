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
    for name, f in (("Pop", "Regular"), ("Pop-It", "Italic"), ("Pop-Md", "Medium"), ("Pop-Sb", "SemiBold"),
                    ("Pop-Bd", "Bold"), ("Pop-Xb", "ExtraBold")):
        pdfmetrics.registerFont(TTFont(name, os.path.join(FONT_DIR, f"Poppins-{f}.ttf")))
    pdfmetrics.registerFontFamily("Pop", normal="Pop", bold="Pop-Sb", italic="Pop-It", boldItalic="Pop-Sb")
    _FONTS_OK = True


# ---------------------------------------------------------------- palettes (une par univers)
PALETTES = {
    "ocre":     {"main": "#C2581B", "dark": "#2B1A10", "light": "#FBEFE5", "accent": "#E9A23B"},
    "emeraude": {"main": "#0F7A5A", "dark": "#0B2A22", "light": "#E6F4EE", "accent": "#D6A73A"},
    "royal":    {"main": "#2E4BB8", "dark": "#121A3D", "light": "#EAEEFB", "accent": "#F0B429"},
    "bordeaux": {"main": "#8E2340", "dark": "#2A0E17", "light": "#F8E9EE", "accent": "#D9A441"},
    "nuit":     {"main": "#5B3FC4", "dark": "#15102E", "light": "#EFEBFB", "accent": "#36C5A8"},
    "corail":   {"main": "#D9473A", "dark": "#2D1311", "light": "#FCECEA", "accent": "#2F8F83"},
    "ardoise":  {"main": "#1F6F8B", "dark": "#0E232C", "light": "#E7F2F6", "accent": "#F2A541"},
}


def hexrgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


# ---------------------------------------------------------------- texte
EMOJI = re.compile("[\U0001F000-\U0001FAFF☀-➿️‍]")


def clean(t):
    t = EMOJI.sub("", str(t or ""))
    return t.replace(" ", " ").replace("\t", " ").strip()


def inline(t):
    """Markdown en ligne → balises ReportLab (après échappement)."""
    t = clean(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"<i>\1</i>", t)
    t = re.sub(r"__(.+?)__", r"<b>\1</b>", t)
    t = t.replace("`", "")
    return t


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
    "chiffres ou études inventés, fausses citations. Pour un verset, un hadith ou une citation, donne la référence "
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
    n_ch = max(4, min(12, round(pages / 5.5)))
    system = ("Tu es un éditeur expert en ebooks qui se vendent bien auprès des audiences francophones africaines "
              "et de la diaspora. Tu conçois des guides pratiques qui tiennent leur promesse. " + STYLE_RULES)
    user = (f"Conçois le plan d'un ebook de {pages} pages environ, intitulé provisoirement « {ctx['title']} ».\n"
            f"{make_context(ctx)}\n\n"
            f"Réponds UNIQUEMENT avec un objet JSON valide :\n"
            "{\"titre\": \"titre accrocheur et clair (max 70 caractères)\", \"sous_titre\": \"promesse concrète (max 110 caractères)\", "
            "\"pour_qui\": \"1 phrase\", \"palette\": \"une valeur parmi : " + ", ".join(PALETTES) + "\", "
            "\"cover_prompt\": \"description en ANGLAIS d'une illustration de couverture (scène, ambiance, lumière, sans texte)\", "
            "\"ce_que_tu_vas_obtenir\": [\"5 bénéfices concrets\"], \"avertissement\": \"une phrase si le sujet touche à la santé, l'argent, le droit ou la religion, sinon vide\", "
            f"\"chapitres\": [{{\"titre\": \"\", \"objectif\": \"ce que le lecteur saura faire après\", \"sections\": [\"3 à 5 titres de parties\"], "
            "\"image_prompt\": \"description en ANGLAIS d'une illustration pour ce chapitre (scène concrète, sans texte)\"}}], "
            "\"bonus\": {\"titre\": \"\", \"type\": \"checklist, plan sur 7 jours, fiches pratiques ou tableau de suivi\"}}\n"
            f"Il faut exactement {n_ch} chapitres, dans un ordre logique du plus simple au plus avancé. "
            "Les illustrations : style éditorial moderne, si des personnes apparaissent ce sont des personnes noires africaines ; "
            "ne représente jamais Dieu, un prophète ou une figure sacrée (préfère des scènes symboliques : lumière, mains, nature).")
    from_json = chat([{"role": "system", "content": system}, {"role": "user", "content": user}], max_tokens=3500, json_mode=True)
    return from_json, n_ch


def gen_section(chat, plan, ctx, kind, words, idx=None, prev=""):
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
    elif kind == "bonus":
        b = plan.get("bonus") or {}
        task = (f"Écris le BONUS « {b.get('titre', 'Plan d’action')} » ({b.get('type', 'checklist')}), environ {words} mots : "
                "outil pratique prêt à l'emploi (checklists « - [ ] », plan jour par jour ou tableau). Court texte d'intro, puis l'outil.")
    else:
        task = (f"Écris la CONCLUSION (environ {words} mots) : résumé du chemin parcouru, les 3 actions à faire dès demain, "
                "et un message motivant. Pas de sous-titres « ## ».")
    system = "Tu es l'auteur de cet ebook, expert du sujet. " + STYLE_RULES + "\n" + FORMAT_RULES
    tokens = int(min(8000, words * 2.3 + 900))
    return chat([{"role": "system", "content": system}, {"role": "user", "content": base + "\n" + task}], max_tokens=tokens)


# ---------------------------------------------------------------- illustrations
IMG_STYLE = (", modern editorial illustration, rich warm colors, soft natural light, high detail, professional book art, "
             "any people shown are Black African, absolutely no text, no letters, no words, no watermark, no logo")


def fetch_image(prompt, w, h, seed, path, log=print):
    url = ("https://image.pollinations.ai/prompt/" + urllib.parse.quote(clean(prompt)[:600] + IMG_STYLE, safe="")
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
        except Exception as e:
            log(f"[ebook] image en échec : {type(e).__name__} {str(e)[:80]}")
        time.sleep(6 + attempt * 6)
    return False


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
def cover_with_veil(path, pal):
    """Recadre l'image au format A4 et ajoute un voile sombre progressif en bas (titre lisible)."""
    out = os.path.splitext(path)[0] + "_cover.jpg"
    if os.path.exists(out):
        return out
    im = Image.open(path).convert("RGB")
    tw, th = 1240, 1754
    sc = max(tw / im.width, th / im.height)
    im = im.resize((int(im.width * sc) + 1, int(im.height * sc) + 1), Image.LANCZOS)
    x, y = (im.width - tw) // 2, (im.height - th) // 2
    im = im.crop((x, y, x + tw, y + th))
    dark = Image.new("RGB", (tw, th), hexrgb(pal["dark"]))
    mask = Image.new("L", (1, th))
    for yy in range(th):
        t = (yy / th - .32) / .5
        mask.putpixel((0, yy), int(255 * max(0, min(.94, t * 1.05))))
    im = Image.composite(dark, im, mask.resize((tw, th)))
    im.save(out, "JPEG", quality=90)
    return out


class NumBadge(Flowable):
    def __init__(self, n, color, size=15):
        super().__init__()
        self.n, self.color, self.size = n, color, size
        self.width, self.height = size, size

    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.circle(self.size / 2, self.size / 2, self.size / 2, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Pop-Bd", 8.6)
        c.drawCentredString(self.size / 2, self.size / 2 - 3.1, str(self.n))


class CheckBox(Flowable):
    def __init__(self, size, color):
        super().__init__()
        self.size, self.color = size, color
        self.width = self.height = size

    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(1.2)
        self.canv.roundRect(0, 0, self.size, self.size, 2.5, stroke=1, fill=0)


class Doc(BaseDocTemplate):
    def afterFlowable(self, f):
        if isinstance(f, Paragraph) and getattr(f, "_toc", None):
            self.notify("TOCEntry", (0, f._toc, self.page))


def build_pdf(book, path):
    register_fonts()
    pal = PALETTES.get(book.get("palette"), PALETTES["ocre"])
    MAIN, DARK, LIGHT, ACC = (colors.HexColor(pal[k]) for k in ("main", "dark", "light", "accent"))
    INK, MUTED = colors.HexColor("#1D1B19"), colors.HexColor("#6E6A64")
    W, H = A4
    M = 2.1 * cm
    title = plain(book["titre"])

    st = {
        "body": ParagraphStyle("body", fontName="Pop", fontSize=10.6, leading=17.2, textColor=INK, spaceAfter=7.5),
        "h2": ParagraphStyle("h2", fontName="Pop-Bd", fontSize=15.5, leading=20, textColor=DARK, spaceBefore=14, spaceAfter=7),
        "h3": ParagraphStyle("h3", fontName="Pop-Sb", fontSize=12.2, leading=16, textColor=MAIN, spaceBefore=9, spaceAfter=5),
        "li": ParagraphStyle("li", fontName="Pop", fontSize=10.6, leading=16.4, textColor=INK),
        "quote": ParagraphStyle("quote", fontName="Pop-It", fontSize=12.5, leading=19, textColor=DARK),
        "qsrc": ParagraphStyle("qsrc", fontName="Pop-Sb", fontSize=9, leading=12, textColor=MAIN, spaceBefore=4),
        "cnum": ParagraphStyle("cnum", fontName="Pop-Xb", fontSize=54, leading=56, textColor=MAIN),
        "ctitle": ParagraphStyle("ctitle", fontName="Pop-Xb", fontSize=25, leading=30, textColor=DARK, spaceAfter=6),
        "obj": ParagraphStyle("obj", fontName="Pop-Md", fontSize=11, leading=16.5, textColor=MUTED),
        "eyebrow": ParagraphStyle("eyebrow", fontName="Pop-Sb", fontSize=8.5, leading=11, textColor=MAIN),
        "ctl": ParagraphStyle("ctl", fontName="Pop-Bd", fontSize=10.2, leading=13, textColor=MAIN, spaceAfter=3),
        "cbody": ParagraphStyle("cbody", fontName="Pop", fontSize=10.2, leading=15.6, textColor=INK, spaceAfter=4),
        "th": ParagraphStyle("th", fontName="Pop-Sb", fontSize=9.4, leading=12.5, textColor=colors.white),
        "td": ParagraphStyle("td", fontName="Pop", fontSize=9.4, leading=13, textColor=INK),
        "big": ParagraphStyle("big", fontName="Pop-Xb", fontSize=24, leading=29, textColor=DARK, spaceAfter=10),
        "toc": ParagraphStyle("toc", fontName="Pop-Md", fontSize=11.5, leading=17, textColor=INK, leftIndent=0, spaceAfter=5),
    }

    def eyebrow(t):
        return Paragraph(t.upper().replace(" ", "&nbsp;") if len(t) < 18 else t.upper(),
                         ParagraphStyle("e2", parent=st["eyebrow"], fontName="Pop-Sb", spaceAfter=4))

    # -------- fonds de page
    def on_cover(c, d):
        c.saveState()
        cover = book.get("cover_img")
        comp = cover_with_veil(cover, pal) if cover and os.path.exists(cover) else None
        if comp:
            c.drawImage(ImageReader(comp), 0, 0, W, H)
        else:
            c.setFillColor(DARK); c.rect(0, 0, W, H, fill=1, stroke=0)
        c.setFillColor(ACC); c.rect(M, H * .43, 2.2 * cm, .17 * cm, fill=1, stroke=0)
        y = H * .43 - .6 * cm
        p = Paragraph(inline(book["titre"]), ParagraphStyle("ct", fontName="Pop-Xb", fontSize=34 if len(title) < 42 else 28,
                                                             leading=40 if len(title) < 42 else 33, textColor=colors.white))
        _, ph = p.wrap(W - 2 * M, H)
        p.drawOn(c, M, y - ph); y -= ph + .45 * cm
        if book.get("sous_titre"):
            s = Paragraph(inline(book["sous_titre"]), ParagraphStyle("cs", fontName="Pop-Md", fontSize=13.5, leading=19,
                                                                     textColor=colors.HexColor("#F1ECE6")))
            _, sh = s.wrap(W - 2 * M - 1 * cm, H); s.drawOn(c, M, y - sh)
        c.setFont("Pop-Sb", 10.5); c.setFillColor(colors.white)
        c.drawString(M, 1.7 * cm, plain(book.get("auteur") or "").upper())
        c.setFillColor(ACC); c.setFont("Pop-Sb", 8.5)
        c.drawRightString(W - M, 1.7 * cm, (book.get("etiquette") or "GUIDE PRATIQUE").upper())
        c.restoreState()

    def on_page(c, d):
        c.saveState()
        c.setFillColor(MAIN); c.rect(0, H - .32 * cm, W, .32 * cm, fill=1, stroke=0)
        c.setFont("Pop-Md", 7.8); c.setFillColor(MUTED)
        c.drawString(M, 1.15 * cm, title[:80])
        c.setFont("Pop-Sb", 8.5); c.setFillColor(DARK)
        c.drawRightString(W - M, 1.15 * cm, str(d.page))
        c.setStrokeColor(colors.HexColor("#E4DFD8")); c.setLineWidth(.6)
        c.line(M, 1.55 * cm, W - M, 1.55 * cm)
        c.restoreState()

    doc = Doc(path, pagesize=A4, leftMargin=M, rightMargin=M, topMargin=2.2 * cm, bottomMargin=2.1 * cm,
              title=title, author=plain(book.get("auteur") or ""), subject=plain(book.get("sous_titre") or ""),
              creator="Kopi")
    frame = Frame(M, 2.1 * cm, W - 2 * M, H - 4.3 * cm, id="f")
    doc.addPageTemplates([PageTemplate("cover", [Frame(0, 0, W, H, id="c")], onPage=on_cover),
                          PageTemplate("page", [frame], onPage=on_page)])
    CW = W - 2 * M

    def callout(kind, ttl, body_blocks):
        label = CALLOUTS.get(kind, "Astuce")
        inner = [Paragraph(f"{label.upper()}" + (f" · {inline(ttl)}" if ttl else ""), st["ctl"])]
        for b in body_blocks:
            inner += render(b, st["cbody"], inner_w=CW - 1.4 * cm)
        bg = {"attention": colors.HexColor("#FDEDEA"), "exercice": colors.HexColor("#F3F1EC"),
              "action": colors.HexColor("#F3F1EC")}.get(kind, LIGHT)
        bar = {"attention": colors.HexColor("#C8372D")}.get(kind, ACC if kind in ("exercice", "action") else MAIN)
        t = Table([["", inner]], colWidths=[.18 * cm, CW - .18 * cm])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), bar), ("BACKGROUND", (1, 0), (1, 0), bg),
                               ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (1, 0), (1, 0), 14),
                               ("RIGHTPADDING", (1, 0), (1, 0), 14), ("TOPPADDING", (1, 0), (1, 0), 11),
                               ("BOTTOMPADDING", (1, 0), (1, 0), 9), ("LEFTPADDING", (0, 0), (0, 0), 0),
                               ("RIGHTPADDING", (0, 0), (0, 0), 0)]))
        return [Spacer(1, 5), t, Spacer(1, 9)]

    def bullets(items, numbered=False, check=False, inner_w=None):
        w = inner_w or CW
        rows = []
        for k, it in enumerate(items, 1):
            if check:
                mark = CheckBox(9.5, MAIN)
            elif numbered:
                mark = NumBadge(k, MAIN)
            else:
                mark = Paragraph("•", ParagraphStyle("b", fontName="Pop-Bd", fontSize=12, leading=16.4, textColor=MAIN))
            rows.append([mark, Paragraph(inline(it), st["li"])])
        t = Table(rows, colWidths=[.75 * cm, w - .75 * cm])
        sty = [("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
               ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("TOPPADDING", (0, 0), (-1, -1), 1)]
        if check:
            sty += [("TOPPADDING", (0, 0), (0, -1), 4)]
        if numbered:
            sty += [("TOPPADDING", (0, 0), (0, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]
        t.setStyle(TableStyle(sty))
        return [t, Spacer(1, 6)]

    def render(b, body=None, inner_w=None):
        body = body or st["body"]
        t = b["t"]
        if t == "p":
            return [Paragraph(inline(b["text"]), body)]
        if t == "h2":
            return [CondPageBreak(3.2 * cm), Paragraph(inline(b["text"]), st["h2"])]
        if t == "h3":
            return [CondPageBreak(2.4 * cm), Paragraph(inline(b["text"]), st["h3"])]
        if t in ("ul", "ol", "check"):
            return bullets(b["items"], numbered=t == "ol", check=t == "check", inner_w=inner_w)
        if t == "quote":
            inner = [Paragraph("« " + inline(b["text"]).strip("«» \"") + " »", st["quote"])]
            if b.get("src"):
                inner.append(Paragraph(inline(b["src"]), st["qsrc"]))
            q = Table([["", inner]], colWidths=[.12 * cm, (inner_w or CW) - .12 * cm])
            q.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), ACC), ("LEFTPADDING", (1, 0), (1, 0), 16),
                                   ("TOPPADDING", (1, 0), (1, 0), 4), ("BOTTOMPADDING", (1, 0), (1, 0), 4),
                                   ("LEFTPADDING", (0, 0), (0, 0), 0), ("RIGHTPADDING", (0, 0), (0, 0), 0)]))
            return [Spacer(1, 6), q, Spacer(1, 10)]
        if t == "table":
            rows = b["rows"]
            ncol = max(len(r) for r in rows)
            rows = [r + [""] * (ncol - len(r)) for r in rows]
            data = [[Paragraph(inline(c), st["th"] if i == 0 else st["td"]) for c in r] for i, r in enumerate(rows)]
            tb = Table(data, colWidths=[(inner_w or CW) / ncol] * ncol, repeatRows=1)
            tb.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), MAIN), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
                                    ("GRID", (0, 0), (-1, -1), .4, colors.HexColor("#E4DFD8")),
                                    ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
            return [Spacer(1, 4), tb, Spacer(1, 10)]
        if t == "callout":
            return callout(b["kind"], b.get("title"), b["body"])
        return []

    def img(pathimg, width, max_h):
        if not pathimg or not os.path.exists(pathimg):
            return []
        im = Image.open(pathimg)
        h = min(max_h, width * im.height / im.width)
        w = h * im.width / im.height
        return [RLImage(pathimg, width=w, height=h)]

    story = [NextPageTemplate("page"), PageBreak()]
    # -------- page 2 : présentation
    story += [Spacer(1, 1.2 * cm), eyebrow("Ce guide est pour toi si…"),
              Paragraph(inline(book.get("pour_qui") or ""), st["big"]), Spacer(1, .3 * cm)]
    if book.get("benefices"):
        story += [Paragraph("Ce que tu vas obtenir", st["h2"])] + bullets(book["benefices"], check=False)
    story += [Spacer(1, 1.4 * cm)]
    legal = (f"© {datetime.date.today().year} {plain(book.get('auteur') or '')}. Tous droits réservés. "
             "Ce guide est réservé à son acheteur : merci de ne pas le partager ni le revendre.")
    if book.get("avertissement"):
        legal += " " + plain(book["avertissement"])
    story += [Paragraph(inline(legal), ParagraphStyle("legal", parent=st["body"], fontSize=8.4, leading=12.5, textColor=MUTED))]
    # -------- sommaire
    toc = TableOfContents()
    toc.levelStyles = [ParagraphStyle("t0", parent=st["toc"])]
    toc.dotsMinLevel = 0
    story += [PageBreak(), eyebrow("Sommaire"), Paragraph("Au programme", st["big"]), toc]

    def heading_toc(text, style):
        p = Paragraph(inline(text), style)
        p._toc = plain(text)
        return p

    # -------- introduction
    story += [PageBreak(), eyebrow("Introduction"), heading_toc(book.get("intro_titre") or "Avant de commencer", st["ctitle"]),
              Spacer(1, .2 * cm)]
    for b in parse_md(book.get("intro", "")):
        story += render(b)
    # -------- chapitres
    for k, ch in enumerate(book["chapitres"], 1):
        story += [PageBreak(), Paragraph(f"{k:02d}", st["cnum"]), eyebrow(f"Chapitre {k}"),
                  heading_toc(ch["titre"], st["ctitle"])]
        if ch.get("objectif"):
            story += [Paragraph(inline("Objectif : " + ch["objectif"]), st["obj"])]
        story += [Spacer(1, .45 * cm)] + img(ch.get("img"), CW, 8.6 * cm) + [Spacer(1, .5 * cm)]
        for b in parse_md(ch.get("texte", "")):
            story += render(b)
    # -------- bonus
    if book.get("bonus_texte"):
        story += [PageBreak(), eyebrow("Bonus"), heading_toc(book.get("bonus_titre") or "Ton plan d'action", st["ctitle"]),
                  Spacer(1, .3 * cm)]
        for b in parse_md(book["bonus_texte"]):
            story += render(b)
    # -------- conclusion
    if book.get("conclusion"):
        story += [PageBreak(), eyebrow("Conclusion"), heading_toc("Et maintenant ?", st["ctitle"]), Spacer(1, .2 * cm)]
        for b in parse_md(book["conclusion"]):
            story += render(b)
    # -------- pour aller plus loin (fidélisation)
    others = [o for o in book.get("autres_produits") or [] if o.get("nom")]
    if others or book.get("boutique"):
        story += [PageBreak(), eyebrow("Pour aller plus loin"),
                  Paragraph("Continue ta progression", st["big"]),
                  Paragraph(inline(f"Merci pour ta confiance. Si ce guide t'a aidé, voici d'autres ressources de "
                                   f"{plain(book.get('auteur') or 'l’auteur')} pour aller plus loin :"), st["body"]),
                  Spacer(1, .3 * cm)]
        for o in others[:5]:
            card = [Paragraph(inline(o["nom"]), ParagraphStyle("on", fontName="Pop-Bd", fontSize=12, leading=16, textColor=DARK)),
                    Paragraph(inline(o.get("promesse") or ""), st["cbody"])]
            t = Table([[card]], colWidths=[CW])
            t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), LIGHT), ("LEFTPADDING", (0, 0), (-1, -1), 14),
                                   ("RIGHTPADDING", (0, 0), (-1, -1), 14), ("TOPPADDING", (0, 0), (-1, -1), 10),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
            story += [KeepTogether([t]), Spacer(1, 8)]
        if book.get("boutique"):
            link = book["boutique"].replace("&", "&amp;")
            story += [Spacer(1, .4 * cm), Paragraph(
                f'Retrouve-les sur : <a href="{link}" color="{pal["main"]}"><b>{link}</b></a>',
                ParagraphStyle("shop", parent=st["body"], fontSize=12, leading=18))]
    doc.multiBuild(story)


# ---------------------------------------------------------------- Word modifiable
def build_docx(book, path):
    pal = PALETTES.get(book.get("palette"), PALETTES["ocre"])
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

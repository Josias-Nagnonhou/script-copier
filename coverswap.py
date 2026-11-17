"""Remplacer la couverture d'un ebook déjà fait (PDF ou Word) par une nouvelle couverture."""
import io
import os

from PIL import Image, ImageFilter, ImageOps


def load_cover(path):
    """La nouvelle couverture : image (JPG, PNG, WEBP, HEIC) ou 1re page d'un PDF."""
    if path.lower().endswith(".pdf"):
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(path)
        page = doc[0]
        w, _ = page.get_size()
        return page.render(scale=2400 / max(w, 1)).to_pil().convert("RGB")
    try:
        import pillow_heif
        pillow_heif.register_heif_opener()
    except Exception:
        pass
    return ImageOps.exif_transpose(Image.open(path)).convert("RGB")


def _page_image(cover, w_pt, h_pt, fit):
    """Image à la taille exacte de la page. fit = "remplir" (recadre) ou "entiere" (fond flouté derrière)."""
    W = 2000
    H = max(1, int(W * h_pt / w_pt))
    if fit == "entiere":
        bg = ImageOps.fit(cover, (W, H), Image.LANCZOS).filter(ImageFilter.GaussianBlur(40))
        bg = Image.blend(bg, Image.new("RGB", bg.size, (0, 0, 0)), .35)
        fg = ImageOps.contain(cover, (W, H), Image.LANCZOS)
        bg.paste(fg, ((W - fg.width) // 2, (H - fg.height) // 2))
        return bg
    return ImageOps.fit(cover, (W, H), Image.LANCZOS, centering=(0.5, 0.5))


def cover_pdf_page(cover, w_pt, h_pt, fit):
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    img = _page_image(cover, w_pt, h_pt, fit)
    b = io.BytesIO()
    img.save(b, "JPEG", quality=92, subsampling=0)
    b.seek(0)
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=(w_pt, h_pt))
    c.drawImage(ImageReader(b), 0, 0, width=w_pt, height=h_pt)
    c.showPage()
    c.save()
    out.seek(0)
    return out


def swap_pdf(src, cover_path, dst, mode="remplacer", fit="remplir", remove=1):
    """mode : "remplacer" (enlève les `remove` premières pages) ou "ajouter" (la nouvelle couverture passe devant)."""
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(src)
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            raise ValueError("Ce PDF est protégé par un mot de passe : Kopi ne peut pas le modifier.")
    n = len(reader.pages)
    if n == 0:
        raise ValueError("Ce PDF ne contient aucune page.")
    first = reader.pages[0]
    w_pt, h_pt = float(first.mediabox.width), float(first.mediabox.height)
    if (first.get("/Rotate") or 0) % 180:
        w_pt, h_pt = h_pt, w_pt
    cover = load_cover(cover_path)
    cpage = PdfReader(cover_pdf_page(cover, w_pt, h_pt, fit)).pages[0]
    writer = PdfWriter(clone_from=reader)   # garde le sommaire, les liens et les infos du livre
    k = max(0, min(int(remove), n - 1)) if mode == "remplacer" else 0
    for _ in range(k):
        writer.remove_page(0)
    writer.insert_page(cpage, 0)
    if k:
        _fix_bookmarks(writer)
    writer.add_metadata({"/Producer": "Kopi"})
    with open(dst, "wb") as f:
        writer.write(f)
    return {"pages_avant": n, "pages_apres": n - k + 1, "retirees": k}


def _fix_bookmarks(writer):
    """Les signets qui visaient une page supprimée (l'ancienne couverture) visent maintenant la nouvelle."""
    from pypdf.generic import ArrayObject, NameObject
    try:
        pages = {id(p.indirect_reference.get_object()) for p in writer.pages}
        first = writer.pages[0].indirect_reference
        root = writer._root_object.get("/Outlines")
        stack = [root.get_object().get("/First")] if root else []
        seen = 0
        while stack and seen < 5000:
            node = stack.pop()
            if node is None:
                continue
            o = node.get_object()
            seen += 1
            dest = o.get("/Dest")
            if dest is None and o.get("/A") is not None:
                dest = o["/A"].get_object().get("/D")
            dest = dest.get_object() if dest is not None else None
            if isinstance(dest, ArrayObject) and dest:
                try:
                    ok = id(dest[0].get_object()) in pages
                except Exception:
                    ok = False
                if not ok:
                    dest[0] = first
                    if len(dest) < 2:
                        dest.append(NameObject("/Fit"))
            for key in ("/Next", "/First"):
                if o.get(key) is not None:
                    stack.append(o[key])
    except Exception:
        pass  # un signet cassé n'empêche jamais de livrer le livre


def swap_docx(src, cover_path, dst, mode="remplacer"):
    """Word : remplace l'image de la 1re page (la couverture) en gardant sa taille, ou ajoute une page devant."""
    from docx import Document
    from docx.enum.text import WD_BREAK
    from docx.shared import Emu
    doc = Document(src)
    cover = load_cover(cover_path)
    b = io.BytesIO()
    cover.save(b, "JPEG", quality=92)
    data = b.getvalue()
    body = doc.element.body
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}
    if mode == "remplacer":
        # la couverture = 1re image du document, si elle arrive avant le premier vrai texte long
        blips = body.findall(".//a:blip", ns)
        if blips:
            rid = blips[0].get("{%s}embed" % ns["r"])
            part = doc.part.related_parts.get(rid)
            if part is not None:
                # on garde le cadre de l'image d'origine : la nouvelle couverture prend ses proportions
                ext = body.find(".//wp:extent", ns)
                if ext is not None:
                    cw, ch = int(ext.get("cx")), int(ext.get("cy"))
                    fitted = ImageOps.fit(cover, (1600, max(1, int(1600 * ch / max(cw, 1)))), Image.LANCZOS)
                    b2 = io.BytesIO()
                    fitted.save(b2, "JPEG", quality=92)
                    data = b2.getvalue()
                part._blob = data
                doc.save(dst)
                return {"remplacee": True}
    # pas d'image de couverture trouvée (ou mode « ajouter ») : nouvelle page de couverture devant
    sec = doc.sections[0]
    width = sec.page_width - sec.left_margin - sec.right_margin
    height = sec.page_height - sec.top_margin - sec.bottom_margin
    p = doc.add_paragraph()
    run = p.add_run()
    w_img, h_img = cover.size
    scale = min(width / w_img, height / h_img)
    run.add_picture(io.BytesIO(data), width=Emu(int(w_img * scale)), height=Emu(int(h_img * scale)))
    run.add_break(WD_BREAK.PAGE)
    body.remove(p._p)
    body.insert(0, p._p)
    doc.save(dst)
    return {"remplacee": False}


def preview(path, pages=3, width=360):
    """Aperçu des premières pages (PDF) en JPEG."""
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(path)
    out = []
    for i in range(min(pages, len(doc))):
        pg = doc[i]
        w, _ = pg.get_size()
        im = pg.render(scale=width / max(w, 1)).to_pil().convert("RGB")
        b = io.BytesIO()
        im.save(b, "JPEG", quality=82)
        out.append(b.getvalue())
    return out

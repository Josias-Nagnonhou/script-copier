"""Détoure une personne (photo intacte, rien n'est redessiné) et repère son visage.

Lancé dans un processus séparé : la mémoire est rendue au serveur dès la fin.
Usage : python cutout.py photo.jpg sortie.png   → affiche un JSON {"w","h","face":[x0,y0,x1,y1],"head_top"}
"""
import json
import os
import sys

import numpy as np
from PIL import Image, ImageFilter, ImageOps

MODEL = os.environ.get("THUMB_SEG_MODEL") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "models",
                                                           "selfie_multiclass_256x256.tflite")


def main(src, dst):
    import mediapipe as mp
    from mediapipe.tasks import python as mpt
    from mediapipe.tasks.python import vision
    im = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
    if max(im.size) > 2000:
        im.thumbnail((2000, 2000), Image.LANCZOS)
    seg = vision.ImageSegmenter.create_from_options(vision.ImageSegmenterOptions(
        base_options=mpt.BaseOptions(model_asset_path=MODEL), output_confidence_masks=True,
        output_category_mask=False))
    r = seg.segment(mp.Image(image_format=mp.ImageFormat.SRGB, data=np.ascontiguousarray(np.array(im))))
    masks = [np.squeeze(m.numpy_view()).astype(np.float32) for m in r.confidence_masks]
    person = np.clip(1 - masks[0], 0, 1)
    person = np.clip((person - 0.45) / 0.35, 0, 1)            # pas de voile de l'ancien fond
    a = Image.fromarray((person * 255).astype(np.uint8)).filter(ImageFilter.MinFilter(3)).filter(
        ImageFilter.GaussianBlur(1.2))
    alpha = np.array(a)
    face = masks[3] > 0.5 if len(masks) > 3 else None          # 3 = peau du visage
    hair = masks[1] > 0.5 if len(masks) > 1 else None
    ys, xs = np.nonzero(alpha > 40)
    if len(xs) < 50:
        print(json.dumps({"error": "Aucune personne trouvée sur cette photo."}))
        return
    x0, x1, y0, y1 = int(xs.min()), int(xs.max()), int(ys.min()), int(ys.max())
    fbox = None
    if face is not None and face.sum() > 30:
        fy, fx = np.nonzero(face)
        fbox = [int(fx.min()), int(fy.min()), int(fx.max()), int(fy.max())]
        if hair is not None and hair.sum() > 30:  # le haut de la tête inclut les cheveux juste au-dessus du visage
            hy, hx = np.nonzero(hair & (np.arange(hair.shape[1])[None, :] > fbox[0] - (fbox[2] - fbox[0]))
                                & (np.arange(hair.shape[1])[None, :] < fbox[2] + (fbox[2] - fbox[0])))
            if len(hy):
                fbox[1] = int(min(fbox[1], hy.min()))
    rgba = im.copy()
    rgba.putalpha(Image.fromarray(alpha))
    rgba = rgba.crop((x0, y0, x1 + 1, y1 + 1))
    rgba.save(dst)
    out = {"w": rgba.width, "h": rgba.height,
           "face": [fbox[0] - x0, fbox[1] - y0, fbox[2] - x0, fbox[3] - y0] if fbox else None}
    print(json.dumps(out))


if __name__ == "__main__":
    try:
        main(sys.argv[1], sys.argv[2])
    except Exception as e:  # le serveur affiche ce message
        print(json.dumps({"error": f"Détourage impossible : {str(e)[:150]}"}))

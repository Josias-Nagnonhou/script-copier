"""Nettoyage audio : retire la musique et les bruits de fond, garde seulement la voix.

Moteur : DeepFilterNet 3 (libre, MIT/Apache), binaire `deep-filter` exécuté sur le processeur.
L'audio est traité par morceaux de 2 minutes (avec un léger chevauchement), donc un fichier
de 2 h ou plus passe sur un petit serveur de 512 Mo.
"""
import glob
import os
import shutil
import subprocess
import time

BIN_URL = "https://github.com/Rikorose/DeepFilterNet/releases/download/v0.5.6/deep-filter-0.5.6-x86_64-unknown-linux-musl"
SEG = 120        # secondes par morceau
OVER = 3         # secondes de chevauchement (le modèle « se met en route » au début de chaque morceau)
SR = 48000


def binary():
    for p in (os.environ.get("DEEP_FILTER_BIN"), shutil.which("deep-filter"), "/usr/local/bin/deep-filter",
              os.path.join(os.path.dirname(os.path.abspath(__file__)), "bin", "deep-filter")):
        if p and os.path.exists(p):
            return p
    return None


def available():
    return bool(binary())


def probe(path):
    """(durée en secondes, a_une_vidéo)"""
    try:
        out = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration:stream=codec_type",
                              "-of", "default=nw=1", path], capture_output=True, text=True, timeout=120).stdout
        dur = 0.0
        for line in out.splitlines():
            if line.startswith("duration="):
                try:
                    dur = float(line.split("=", 1)[1])
                except ValueError:
                    pass
        return dur, "codec_type=video" in out
    except Exception:
        return 0.0, False


def clean(src, workdir, voice_mp3, progress=None, cancel=None, strength=100, heavy=None, bitrate="192k"):
    """Écrit la voix seule dans `voice_mp3`. progress(fait_s, total_s, écoulé_s)."""
    exe = binary()
    if not exe:
        raise RuntimeError("Le moteur de nettoyage n'est pas installé sur le serveur.")
    total, _ = probe(src)
    if total <= 0:
        raise RuntimeError("Impossible de lire ce fichier audio.")
    segdir = os.path.abspath(os.path.join(workdir, "seg"))
    os.makedirs(segdir, exist_ok=True)
    pieces = []
    t0 = time.time()
    k = 0
    start = 0.0
    while start < total - 0.05:
        if cancel and cancel():
            raise RuntimeError("annulé")
        pre = OVER if start > 0 else 0
        a = start - pre
        length = min(SEG, total - start) + pre
        raw = os.path.join(segdir, f"in_{k:04d}.wav")
        subprocess.run(["ffmpeg", "-v", "error", "-y", "-threads", "1", "-ss", f"{a:.3f}", "-t", f"{length:.3f}",
                        "-i", src, "-vn", "-ac", "1", "-ar", str(SR), "-c:a", "pcm_s16le", raw], check=True)
        outd = os.path.join(segdir, f"o{k}")
        cmd = [exe, "-D", "--pf", "-a", str(int(strength)), "-o", outd, raw]
        if heavy:
            heavy.acquire()
        try:
            r = subprocess.run(cmd, capture_output=True, text=True)
        finally:
            if heavy:
                heavy.release()
        res = glob.glob(os.path.join(outd, "*.wav"))
        if r.returncode != 0 or not res:
            raise RuntimeError("Le nettoyage a échoué sur un morceau : " + (r.stderr or "")[-300:])
        cut = os.path.join(segdir, f"c_{k:04d}.wav")
        subprocess.run(["ffmpeg", "-v", "error", "-y", "-ss", f"{pre:.3f}", "-i", res[0], "-c:a", "pcm_s16le", cut],
                       check=True)
        os.remove(raw)
        shutil.rmtree(outd, ignore_errors=True)
        pieces.append(cut)
        start += SEG
        k += 1
        if progress:
            progress(min(start, total), total, time.time() - t0)
    lst = os.path.join(segdir, "list.txt")
    with open(lst, "w") as f:
        for p in pieces:
            f.write("file '" + p.replace("'", "'\\''") + "'\n")
    subprocess.run(["ffmpeg", "-v", "error", "-y", "-f", "concat", "-safe", "0", "-i", lst, "-af",
                    "loudnorm=I=-16:TP=-1.5:LRA=11", "-ar", "44100", "-c:a", "libmp3lame", "-b:a", bitrate, voice_mp3],
                   check=True)
    shutil.rmtree(segdir, ignore_errors=True)
    return total


def remux_video(video_src, voice_mp3, out_mp4):
    """Remet la voix nettoyée sur la vidéo d'origine (image copiée telle quelle, sans réencodage)."""
    subprocess.run(["ffmpeg", "-v", "error", "-y", "-i", video_src, "-i", voice_mp3, "-map", "0:v:0", "-map", "1:a:0",
                    "-c:v", "copy", "-c:a", "aac", "-b:a", "192k", "-shortest", "-movflags", "+faststart", out_mp4],
                   check=True)

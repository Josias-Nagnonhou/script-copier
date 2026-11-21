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
PASSES = 1   # une 2e écoute du fond n'apporte presque rien et double le temps
SEG_BG = 30     # séparation voix/fond : morceaux d'1 min (stéréo en mémoire, on reste léger)


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


def remove_voice(src, workdir, out_audio, progress=None, cancel=None, heavy=None):
    """Enlève la voix et garde le reste (musique, ambiance, bruits) : fond = original − voix détectée.
    La voix est isolée par DeepFilterNet (sans post-filtre, pour la retirer au maximum), alignée à
    l'échantillon près (-D), puis soustraite de chaque canal : la stéréo de la musique est conservée."""
    exe = binary()
    if not exe:
        raise RuntimeError("Le moteur de séparation de la voix n'est pas installé sur le serveur.")
    total, _ = probe(src)
    if total <= 0:
        raise RuntimeError("Impossible de lire le son de cette vidéo.")
    segdir = os.path.abspath(os.path.join(workdir, "bg"))
    os.makedirs(segdir, exist_ok=True)
    t0 = time.time()
    # le fond est encodé au fil de l'eau (pas de gros fichiers WAV sur le disque, même pour 2 h)
    enc = subprocess.Popen(["ffmpeg", "-v", "error", "-y", "-f", "f32le", "-ar", str(SR), "-ac", "2", "-i", "-",
                            "-c:a", "aac", "-b:a", "192k", out_audio], stdin=subprocess.PIPE)
    try:
        _remove_loop(src, segdir, enc, exe, total, progress, cancel, heavy, t0)
    except BaseException:
        enc.kill()
        raise
    enc.stdin.close()
    if enc.wait() != 0:
        raise RuntimeError("Encodage du fond sonore impossible.")
    shutil.rmtree(segdir, ignore_errors=True)
    return total


def _remove_loop(src, segdir, enc, exe, total, progress, cancel, heavy, t0):
    import numpy as np
    k, start = 0, 0.0
    while start < total - 0.05:
        if cancel and cancel():
            raise RuntimeError("annulé")
        pre = OVER if start > 0 else 0
        a = start - pre
        length = min(SEG_BG, total - start) + pre
        st = subprocess.run(["ffmpeg", "-v", "error", "-threads", "1", "-ss", f"{a:.3f}", "-t", f"{length:.3f}",
                             "-i", src, "-vn", "-ac", "2", "-ar", str(SR), "-f", "f32le", "-"],
                            capture_output=True, check=True).stdout
        orig = np.frombuffer(st, dtype=np.float32).reshape(-1, 2)
        voice = _dfn_voice(exe, orig.mean(1), segdir, f"{k:04d}a", heavy)
        bg = orig - voice[:, None]
        if PASSES > 1:  # 2e écoute du fond : le modèle retrouve la voix qu'il avait prise pour de la musique
            v2 = _dfn_voice(exe, bg.mean(1), segdir, f"{k:04d}b", heavy)
            bg -= v2[:, None]
            voice = voice + v2
        # 2e passe : là où la voix est encore présente dans le fond (ce que le 1er modèle a laissé),
        # on l'éteint fréquence par fréquence ; la musique seule n'est pas touchée
        bg = np.stack([_spectral_clean(bg[:, c], voice) for c in range(2)], 1)
        i0 = int(round(pre * SR))
        bg = bg[i0:]
        enc.stdin.write(np.clip(bg, -1, 1).astype(np.float32).tobytes())
        start += SEG_BG
        k += 1
        if progress:
            progress(min(start, total), total, time.time() - t0)


def _spectral_clean(bg, voice, alpha=30.0, floor=0.01, n=2048, hop=512):
    """Masque de Wiener : garde le fond, retire ce qui ressemble à la voix détectée (par blocs, mémoire légère)."""
    import numpy as np
    L = len(bg)
    if L < n:
        return bg
    win = np.hanning(n + 1)[:n].astype(np.float32)
    pad = n
    x = np.concatenate([np.zeros(pad, np.float32), bg.astype(np.float32), np.zeros(pad + n, np.float32)])
    v = np.concatenate([np.zeros(pad, np.float32), voice.astype(np.float32), np.zeros(pad + n, np.float32)])
    out = np.zeros_like(x)
    norm = np.zeros_like(x)
    nfr = (len(x) - n) // hop + 1
    blk = 400
    kern = np.ones(3, np.float32) / 3
    for f0 in range(0, nfr, blk):
        f1 = min(nfr, f0 + blk)
        # 1 trame de marge de chaque côté pour lisser le masque dans le temps sans couture
        g0, g1 = max(0, f0 - 1), min(nfr, f1 + 1)
        idx = (np.arange(g0, g1) * hop)[:, None] + np.arange(n)[None, :]
        X = np.fft.rfft(x[idx] * win, axis=1)
        V = np.fft.rfft(v[idx] * win, axis=1)
        pb, pv = np.abs(X) ** 2, np.abs(V) ** 2
        m = pb / (pb + alpha * pv + 1e-10)
        m = np.apply_along_axis(lambda r: np.convolve(r, kern, "same"), 1, m)   # lissage en fréquence
        m = (np.vstack([m[:1], m[:-1]]) + m + np.vstack([m[1:], m[-1:]])) / 3     # lissage dans le temps
        m = np.maximum(m, floor).astype(np.float32)
        y = np.fft.irfft(X * m, n=n, axis=1).astype(np.float32) * win
        a = f0 - g0
        for i, fr in enumerate(range(f0, f1)):
            s0 = fr * hop
            out[s0:s0 + n] += y[a + i]
            norm[s0:s0 + n] += win ** 2
        del X, V, pb, pv, m, y, idx
    out = out / np.maximum(norm, 1e-6)
    return out[pad:pad + L]


def _dfn_voice(exe, mono_arr, segdir, tag, heavy=None):
    """Voix détectée par DeepFilterNet dans un signal mono (même longueur, alignée à l'échantillon)."""
    import numpy as np
    mono = os.path.join(segdir, f"m_{tag}.wav")
    subprocess.run(["ffmpeg", "-v", "error", "-y", "-f", "f32le", "-ar", str(SR), "-ac", "1", "-i", "-",
                    "-c:a", "pcm_s16le", mono], input=np.ascontiguousarray(np.clip(mono_arr, -1, 1), np.float32).tobytes(),
                   check=True)
    outd = os.path.join(segdir, f"o{tag}")
    if heavy:
        heavy.acquire()
    try:
        r = subprocess.run([exe, "-D", "-o", outd, mono], capture_output=True, text=True)
    finally:
        if heavy:
            heavy.release()
    res = glob.glob(os.path.join(outd, "*.wav"))
    if r.returncode != 0 or not res:
        raise RuntimeError("La séparation de la voix a échoué : " + (r.stderr or "")[-300:])
    vraw = subprocess.run(["ffmpeg", "-v", "error", "-i", res[0], "-ac", "1", "-ar", str(SR), "-f", "f32le", "-"],
                          capture_output=True, check=True).stdout
    os.remove(mono)
    shutil.rmtree(outd, ignore_errors=True)
    v = np.frombuffer(vraw, dtype=np.float32)
    n = len(mono_arr)
    if len(v) < n:  # le modèle rend parfois quelques ms de moins : on complète (aucun décalage)
        v = np.concatenate([v, np.zeros(n - len(v), np.float32)])
    return v[:n].copy()

"""Kopi — récupère la transcription de plusieurs vidéos TikTok / YouTube
et les rassemble dans un document Word (Script 1, Script 2, ...)."""
import os, re, random, zipfile, json, time, uuid, glob, shutil, threading, tempfile, subprocess, datetime
from concurrent.futures import ThreadPoolExecutor

import requests
import yt_dlp
from flask import Flask, request, jsonify, send_file, render_template, session, redirect
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or uuid.uuid4().hex
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 ** 3  # 2 Go par fichier

def find_groq_key():
    """Trouve la clé Groq même si le nom de la variable est mal écrit
    (espaces, minuscules, guillemets) ou si la clé a été collée ailleurs."""
    def clean(v):
        return (v or "").strip().strip('"').strip("'").strip()
    for k, v in os.environ.items():
        if "GROQ" in k.strip().upper() and clean(v):
            return clean(v)
    for k, v in os.environ.items():
        if clean(v).startswith("gsk_"):
            return clean(v)
        if clean(k).startswith("gsk_"):  # clé collée dans la case du nom
            return clean(k)
    return ""


GROQ_KEY = find_groq_key()


def find_env_key(hint, prefix=None):
    """Clé tolérante : nom de variable approximatif accepté (espaces, minuscules, guillemets)."""
    for k, v in os.environ.items():
        v = (v or "").strip().strip('"').strip("'").strip()
        if hint in k.strip().upper() and v:
            return v
    if prefix:
        for v in os.environ.values():
            v = (v or "").strip().strip('"').strip("'")
            if v.startswith(prefix) and len(v) > 30:
                return v
    return ""


GEMINI_KEY = find_env_key("GEMINI") or find_env_key("GOOGLE") or find_env_key("NANO", "AIza")
PEXELS_KEY = find_env_key("PEXELS")
PIXABAY_KEY = find_env_key("PIXABAY")
print("Images : Google", "oui" if GEMINI_KEY else "non", "| Pexels", "oui" if PEXELS_KEY else "non",
      "| Pixabay", "oui" if PIXABAY_KEY else "non", flush=True)
print("Clé Groq :", "trouvée" if GROQ_KEY else "ABSENTE", "| variables :",
      sorted(k for k in os.environ if k.isupper() and not k.startswith(("RENDER", "PYTHON", "PATH", "HOME", "LANG", "GPG", "HOSTNAME"))), flush=True)
GROQ_MODEL = os.environ.get("GROQ_MODEL", "whisper-large-v3-turbo")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "").strip()
WORKERS = int(os.environ.get("WORKERS", "1"))  # serveur gratuit : 512 Mo, une vidéo à la fois
BASE = os.path.join(tempfile.gettempdir(), "script-copier")
os.makedirs(BASE, exist_ok=True)

COOKIES = {}
# état des cookies YouTube : absent / ok / expired (+ d'où ils viennent)
YT = {"status": "absent", "source": "", "updated": None}


def valid_yt_cookies(text):
    lines = [l for l in text.splitlines() if l.strip() and not l.startswith("#")]
    return any("youtube.com" in l and len(l.split("\t")) >= 7 for l in lines)


def install_yt_cookies(text, source):
    p = os.path.join(BASE, "yt_cookies.txt")  # copie modifiable (yt-dlp la met à jour)
    with open(p, "w") as f:
        f.write(text if text.startswith("# Netscape") else "# Netscape HTTP Cookie File\n" + text)
    COOKIES["YT_COOKIES"] = p
    YT.update(status="ok", source=source, updated=time.time())


def load_cookies():
    # 1) fichier secret Render (permanent) 2) variable d'environnement
    for p in sorted(glob.glob("/etc/secrets/*")):
        try:
            t = open(p, encoding="utf-8", errors="ignore").read()
        except OSError:
            continue
        if valid_yt_cookies(t):
            install_yt_cookies(t, "render")
            return
    for k, v in os.environ.items():
        if "COOKIE" in k.upper() and "TT" not in k.upper() and valid_yt_cookies(v):
            install_yt_cookies(v, "render")
            return


load_cookies()
if os.environ.get("TT_COOKIES"):
    _p = os.path.join(BASE, "tt_cookies.txt")
    with open(_p, "w") as f:
        f.write(os.environ["TT_COOKIES"])
    COOKIES["TT_COOKIES"] = _p
print("Cookies YouTube :", YT["status"], flush=True)

JOBS = {}
LOCK = threading.Lock()


# ---------------------------------------------------------------- helpers
def ydl_opts(platform, **extra):
    o = {"quiet": True, "no_warnings": True, "noprogress": True,
         "socket_timeout": 30, "retries": 3}
    ck = COOKIES.get("YT_COOKIES" if platform == "youtube" else "TT_COOKIES")
    if ck:
        o["cookiefile"] = ck
    o.update(extra)
    return o


def normalize_source(platform, raw, yt_tab):
    raw = raw.strip()
    if platform == "tiktok":
        if raw.startswith("http"):
            return raw
        return "https://www.tiktok.com/@" + raw.lstrip("@").strip()
    # youtube
    tab = "shorts" if yt_tab == "shorts" else "videos"
    if not raw.startswith("http"):
        return f"https://www.youtube.com/@{raw.lstrip('@').strip()}/{tab}"
    m = re.match(r"^(https?://(?:www\.|m\.)?youtube\.com/(?:@[^/?#]+|channel/[^/?#]+|c/[^/?#]+|user/[^/?#]+))/?(?:featured)?/?$", raw)
    if m:
        return f"{m.group(1)}/{tab}"
    return raw


def flatten(info):
    out = []
    for e in info.get("entries") or []:
        if not e:
            continue
        if e.get("entries"):
            out += flatten(e)
        else:
            out.append(e)
    return out


def entry_url(e, platform):
    u = e.get("webpage_url") or e.get("url") or ""
    if u.startswith("http"):
        return u
    vid = e.get("id") or u
    if platform == "youtube":
        return "https://www.youtube.com/watch?v=" + vid
    return u


def list_videos(platform, url, limit, sort):
    opts = ydl_opts(platform, extract_flat="in_playlist", skip_download=True, ignoreerrors=True)
    if sort == "recent" and limit:
        opts["playlistend"] = limit
    with yt_dlp.YoutubeDL(opts) as y:
        info = y.extract_info(url, download=False)
    if not info:  # relancer sans ignoreerrors pour obtenir la vraie raison
        opts.pop("ignoreerrors")
        try:
            with yt_dlp.YoutubeDL(opts) as y:
                info = y.extract_info(url, download=False)
        except Exception as e:
            raise RuntimeError(f"Compte ou lien introuvable ({e})")
    if not info:
        raise RuntimeError("Compte ou lien introuvable.")
    entries = flatten(info) if info.get("entries") is not None else [info]
    source_name = info.get("uploader") or info.get("channel") or info.get("title") or url
    vids = [{"url": entry_url(e, platform), "title": e.get("title") or e.get("description") or "",
             "views": e.get("view_count") or 0, "duration": e.get("duration")} for e in entries]
    vids = [v for v in vids if v["url"]]
    if sort == "popular":
        vids.sort(key=lambda v: v["views"] or 0, reverse=True)
    if limit:
        vids = vids[:limit]
    return source_name, vids


def clean_caption_text(raw, ext):
    if ext == "json3":
        try:
            data = json.loads(raw)
            parts = []
            for ev in data.get("events", []):
                t = "".join(s.get("utf8", "") for s in ev.get("segs") or [])
                if t.strip():
                    parts.append(t.strip())
            return re.sub(r"\s+", " ", " ".join(parts)).strip()
        except Exception:
            pass
    lines, recent = [], []
    for line in raw.splitlines():
        s = line.strip()
        if not s or "-->" in s or s.isdigit() or s.startswith(("WEBVTT", "Kind:", "Language:", "NOTE", "STYLE")):
            continue
        s = re.sub(r"<[^>]+>", "", s)
        s = re.sub(r"\{[^}]*\}", "", s).strip()
        if not s or s in recent:
            continue
        lines.append(s)
        recent = (recent + [s])[-3:]
    return re.sub(r"\s+", " ", " ".join(lines)).strip()


def pick_caption(info, lang):
    """Choisit la meilleure piste de sous-titres. Retourne (url, ext) ou None."""
    def best_fmt(tracks):
        by = {t.get("ext"): t for t in tracks if t.get("url")}
        for ext in ("json3", "vtt", "srt", "srv3", "ttml"):
            if ext in by:
                return by[ext]["url"], ext
        if tracks and tracks[0].get("url"):
            return tracks[0]["url"], tracks[0].get("ext", "vtt")
        return None

    subs = info.get("subtitles") or {}
    auto = info.get("automatic_captions") or {}
    subs = {k: v for k, v in subs.items() if k != "live_chat"}
    orig_lang = (info.get("language") or "").split("-")[0]
    order = []
    if lang:
        order += [k for k in subs if k.split("-")[0] == lang]
    if orig_lang:
        order += [k for k in subs if k.split("-")[0] == orig_lang]
    order += list(subs.keys())
    for k in order:
        r = best_fmt(subs[k])
        if r:
            return r
    # sous-titres auto : uniquement la langue d'origine (pas les traductions auto)
    keys = [k for k in auto if k.endswith("-orig")]
    if lang:
        keys += [k for k in auto if k.split("-")[0] == lang and not ("-" in k and not k.endswith("-orig"))]
    if orig_lang:
        keys += [k for k in auto if k == orig_lang]
    for k in keys:
        r = best_fmt(auto[k])
        if r:
            return r
    return None


def groq_transcribe(path, lang):
    data = {"model": GROQ_MODEL, "response_format": "text", "temperature": "0"}
    if lang:
        data["language"] = lang
    for attempt in range(6):
        with open(path, "rb") as f:
            r = requests.post("https://api.groq.com/openai/v1/audio/transcriptions",
                              headers={"Authorization": f"Bearer {GROQ_KEY}"},
                              data=data, files={"file": (os.path.basename(path), f)}, timeout=300)
        if r.status_code == 200:
            return r.text.strip()
        if r.status_code in (429, 500, 502, 503):
            wait = r.headers.get("retry-after")
            try:
                wait = float(wait)
            except (TypeError, ValueError):
                wait = 15 * (attempt + 1)
            time.sleep(min(max(wait, 3), 120))
            continue
        if r.status_code == 401:
            raise RuntimeError("Clé Groq invalide.")
        raise RuntimeError(f"Groq a répondu {r.status_code} : {r.text[:150]}")
    raise RuntimeError("Limite Groq atteinte, réessaie plus tard.")


def transcribe_audio(platform, url, lang, tmp):
    opts = ydl_opts(platform, format="bestaudio/worst[ext=mp4]/worst/best",
                    outtmpl=os.path.join(tmp, "media.%(ext)s"))
    with yt_dlp.YoutubeDL(opts) as y:
        y.extract_info(url, download=True)
    media = [p for p in glob.glob(os.path.join(tmp, "media.*")) if not p.endswith(".part")]
    if not media:
        raise RuntimeError("Téléchargement de l'audio impossible.")
    # audio mono 16 kHz très léger, découpé en morceaux de 20 min
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", media[0], "-vn", "-ac", "1", "-ar", "16000",
                    "-b:a", "32k", "-f", "segment", "-segment_time", "1200",
                    os.path.join(tmp, "part_%03d.mp3")], check=True)
    parts = sorted(glob.glob(os.path.join(tmp, "part_*.mp3")))
    if not parts:
        raise RuntimeError("Cette vidéo n'a pas de son.")
    return " ".join(t for t in (groq_transcribe(p, lang) for p in parts) if t).strip()


def is_bot_block(msg):
    m = msg.lower()
    return "sign in to confirm" in m or "not a bot" in m or "cookies" in m and "youtube" in m


def bot_message():
    if YT["status"] == "absent":
        return "YouTube bloque le serveur : ajoute tes cookies YouTube (bouton en haut)."
    YT["status"] = "expired"
    return "Tes cookies YouTube ont expiré : refais l'export et recolle-les."


def process_item(job, item):
    if job.get("blocked", 0) >= 3:  # YouTube bloque déjà : inutile d'insister
        item["status"], item["msg"] = "error", job["block_msg"]
        return
    item["status"] = "working"
    tmp = tempfile.mkdtemp(dir=BASE)
    try:
        with yt_dlp.YoutubeDL(ydl_opts(job["platform"], skip_download=True)) as y:
            info = y.extract_info(item["url"], download=False)
            item["title"] = info.get("title") or info.get("description") or item["title"]
            text = ""
            cap = pick_caption(info, job["lang"])
            if cap:
                try:
                    raw = y.urlopen(cap[0]).read().decode("utf-8", "ignore")
                    text = clean_caption_text(raw, cap[1])
                    item["method"] = "sous-titres"
                except Exception:
                    text = ""
        if len(text) < 20:
            if not GROQ_KEY:
                raise RuntimeError("Pas de sous-titres et aucune clé Groq configurée.")
            item["msg"] = "Transcription IA…"
            text = transcribe_audio(job["platform"], item["url"], job["lang"], tmp)
            item["method"] = "IA"
        if not text:
            item["status"], item["msg"] = "empty", "Aucune parole détectée"
        else:
            item["text"], item["status"], item["msg"] = text, "done", ""
        if job["platform"] == "youtube" and YT["status"] != "absent":
            YT["status"] = "ok"
    except Exception as e:
        msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")
        if job["platform"] == "youtube" and is_bot_block(msg):
            msg = bot_message()
            job["blocked"] = job.get("blocked", 0) + 1
            job["block_msg"] = msg
        item["status"], item["msg"] = "error", msg[:200]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def paragraphs(text, n=4):
    sentences = re.split(r"(?<=[.!?…])\s+", text.strip())
    if len(sentences) < 2:  # pas de ponctuation (sous-titres auto) : couper tous les ~70 mots
        words = text.split()
        return [" ".join(words[i:i + 70]) for i in range(0, len(words), 70)]
    return [" ".join(sentences[i:i + n]) for i in range(0, len(sentences), n)]


def build_docx(job, path):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Calibri", Pt(11)
    is_file = job["platform"] in ("file", "translate")
    kind = {"file": "Transcription", "translate": "Traduction"}.get(job["platform"], "Scripts")
    doc.add_heading(f"{kind} — {job['source_name']}", 0)
    done = [i for i in job["items"] if i["status"] == "done"]
    meta = doc.add_paragraph()
    label = {"tiktok": "TikTok", "youtube": "YouTube", "file": "Fichier importé",
             "translate": f"Traduit en {job.get('target', '')}"}.get(job["platform"], "")
    count = {"file": "transcription intégrale", "translate": f"{len(job['items'][0]['text'])} caractères"}.get(
        job["platform"], f"{len(done)} scripts")
    r = meta.add_run(f"{label} · {count} · "
                     f"{datetime.datetime.now().strftime('%d/%m/%Y')}")
    r.font.color.rgb, r.font.size = RGBColor(0x66, 0x66, 0x66), Pt(10)
    for n, it in enumerate(done, 1):
        if job["platform"] == "translate":  # garder les paragraphes du texte d'origine
            for p in [x.strip() for x in it["text"].split("\n") if x.strip()]:
                doc.add_paragraph(p)
            continue
        if is_file:
            for p in paragraphs(it["text"]):
                doc.add_paragraph(p)
            continue
        doc.add_heading(f"Script {n}", 1)
        info = doc.add_paragraph()
        t = info.add_run((it["title"] or "")[:200])
        t.italic = True
        info.add_run("\n" + it["url"]).font.size = Pt(9)
        for p in paragraphs(it["text"]):
            doc.add_paragraph(p)
    failed = [i for i in job["items"] if i["status"] != "done"]
    if failed:
        doc.add_page_break()
        doc.add_heading("Vidéos non transcrites", 1)
        for it in failed:
            doc.add_paragraph(f"{it['url']} — {it['msg']}", style="List Bullet")
    doc.save(path)


def run_job(job):
    print(f"[tâche {job['id']}] {job['platform']} {job['url']} (max {job['limit'] or 'toutes'})", flush=True)
    try:
        job["phase"] = "listing"
        name, vids = list_videos(job["platform"], job["url"], job["limit"], job["sort"])
        if not vids:
            raise RuntimeError("Aucune vidéo trouvée sur ce compte.")
        job["source_name"] = name
        job["items"] = [{"n": i + 1, "url": v["url"], "title": v["title"], "status": "pending",
                         "msg": "", "method": "", "text": ""} for i, v in enumerate(vids)]
        job["phase"] = "transcribing"
        with ThreadPoolExecutor(max_workers=WORKERS) as ex:
            list(ex.map(lambda it: process_item(job, it), job["items"]))
        if not any(i["status"] == "done" for i in job["items"]):
            raise RuntimeError(job.get("block_msg") or "Aucune vidéo n'a pu être transcrite.")
        safe = re.sub(r"[^\w-]+", "_", name)[:40].strip("_") or "scripts"
        job["filename"] = f"Scripts_{safe}.docx"
        job["file"] = os.path.join(BASE, job["id"] + ".docx")
        build_docx(job, job["file"])
        job["phase"] = "done"
        print(f"[tâche {job['id']}] terminée : {sum(i['status'] == 'done' for i in job['items'])}/{len(job['items'])} scripts", flush=True)
    except Exception as e:
        msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")
        print(f"[tâche {job['id']}] erreur : {msg[:200]}", flush=True)
        if job["platform"] == "youtube" and is_bot_block(msg):
            msg = bot_message()
        job["phase"], job["error"] = "error", msg[:300]
    job["finished"] = time.time()


def cleanup():
    now = time.time()
    with LOCK:
        for jid in list(JOBS):
            j = JOBS[jid]
            keep = {"download": 2, "media": 2, "clip": 3}.get(j["platform"], 6) * 3600
            if j.get("finished") and now - j["finished"] > keep:
                if j.get("file") and os.path.exists(j["file"]):
                    os.remove(j["file"])
                if j.get("dir"):
                    shutil.rmtree(j["dir"], ignore_errors=True)
                del JOBS[jid]


# ---------------------------------------------------------------- routes
def authed():
    return not APP_PASSWORD or session.get("ok")


@app.route("/login", methods=["GET", "POST"])
def login():
    err = ""
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["ok"] = True
            session.permanent = True
            return redirect("/")
        err = "Mot de passe incorrect"
    return render_template("login.html", err=err)


@app.route("/")
def index():
    if not authed():
        return redirect("/login")
    return render_template("index.html", groq=bool(GROQ_KEY), yt=YT["status"], yt_src=YT["source"],
                           gemini=bool(GEMINI_KEY), pexels=bool(PEXELS_KEY or PIXABAY_KEY))


@app.post("/api/start")
def start():
    if not authed():
        return jsonify(error="Non connecté"), 401
    cleanup()
    d = request.get_json(force=True)
    platform = "tiktok" if d.get("platform") == "tiktok" else "youtube"
    raw = (d.get("source") or "").strip()
    if not raw:
        return jsonify(error="Indique un compte ou un lien."), 400
    try:
        limit = int(d.get("limit") or 0)
    except ValueError:
        limit = 0
    job = {"id": uuid.uuid4().hex[:12], "platform": platform, "source": raw,
           "url": normalize_source(platform, raw, d.get("yt_tab")),
           "limit": max(0, min(limit, 1000)), "sort": "popular" if d.get("sort") == "popular" else "recent",
           "lang": (d.get("lang") or "").strip()[:5], "phase": "queued", "items": [],
           "source_name": raw, "error": "", "created": time.time()}
    with LOCK:
        JOBS[job["id"]] = job
    threading.Thread(target=run_job, args=(job,), daemon=True).start()
    return jsonify(job_id=job["id"])


@app.get("/api/job/<jid>")
def job_status(jid):
    if not authed():
        return jsonify(error="Non connecté"), 401
    j = JOBS.get(jid)
    if not j:
        return jsonify(error="Tâche introuvable (le serveur a peut-être redémarré)."), 404
    items = [{k: v for k, v in i.items() if k not in ("text", "path")} | {"chars": len(i["text"])} for i in j["items"]]
    return jsonify(id=j["id"], phase=j["phase"], error=j["error"], source_name=j["source_name"],
                   platform=j["platform"], items=items, has_file=bool(j.get("file")),
                   part=j.get("part", 0), parts=j.get("parts", 0), duration=j.get("duration", 0),
                   zip_size=j.get("zip_size", 0), step=j.get("step", ""), pct=j.get("pct", 0),
                   resumable=j.get("resumable", False))


@app.get("/api/job/<jid>/text")
def job_text(jid):
    if not authed():
        return "Non connecté", 401
    j = JOBS.get(jid)
    if not j:
        return "Introuvable", 404
    done = [i for i in j["items"] if i["status"] == "done"]
    out = "\n\n".join(f"Script {n}\n{i['text']}" for n, i in enumerate(done, 1))
    return out, 200, {"Content-Type": "text/plain; charset=utf-8"}


@app.get("/api/job/<jid>/docx")
def job_docx(jid):
    if not authed():
        return "Non connecté", 401
    j = JOBS.get(jid)
    if not j or not j.get("file") or not os.path.exists(j["file"]):
        return "Fichier introuvable", 404
    return send_file(j["file"], as_attachment=True, download_name=j["filename"])


# ------------------------------------------------ transcription d'un fichier
def probe_duration(path):
    try:
        out = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration",
                              "-of", "default=nw=1:nk=1", path], capture_output=True, text=True, timeout=60)
        return float(out.stdout.strip())
    except Exception:
        return 0.0


def run_file_job(job, src):
    print(f"[fichier {job['id']}] {job['source_name']}", flush=True)
    item = job["items"][0]
    tmp = os.path.dirname(src)
    try:
        job["phase"] = "transcribing"
        item["status"], item["msg"] = "working", "Extraction du son…"
        job["duration"] = probe_duration(src)
        # son mono 16 kHz léger, en morceaux de 10 min (bien sous la limite Groq)
        r = subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", src, "-vn", "-ac", "1", "-ar", "16000",
                            "-b:a", "32k", "-f", "segment", "-segment_time", "600",
                            os.path.join(tmp, "part_%03d.mp3")], capture_output=True, text=True)
        try:
            os.remove(src)  # libère le disque tout de suite
        except OSError:
            pass
        parts = sorted(glob.glob(os.path.join(tmp, "part_*.mp3")))
        if r.returncode != 0 or not parts:
            raise RuntimeError("Impossible de lire le son de ce fichier. Vérifie que c'est bien un audio ou une vidéo.")
        job["parts"] = len(parts)
        texts = []
        for i, p in enumerate(parts, 1):
            item["msg"] = f"Transcription IA · partie {i}/{len(parts)}"
            job["part"] = i - 1
            t = groq_transcribe(p, job["lang"])
            if t:
                texts.append(t)
            job["part"] = i
        text = " ".join(texts).strip()
        if not text:
            item["status"], item["msg"] = "empty", "Aucune parole détectée dans ce fichier."
            raise RuntimeError("Aucune parole détectée dans ce fichier.")
        item.update(text=text, status="done", msg="", method="IA")
        safe = re.sub(r"[^\w-]+", "_", os.path.splitext(job["source_name"])[0])[:40].strip("_") or "transcription"
        job["filename"] = f"Transcription_{safe}.docx"
        job["file"] = os.path.join(BASE, job["id"] + ".docx")
        build_docx(job, job["file"])
        job["phase"] = "done"
        print(f"[fichier {job['id']}] terminé : {len(text)} caractères", flush=True)
    except Exception as e:
        msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e))
        print(f"[fichier {job['id']}] erreur : {msg[:200]}", flush=True)
        if item["status"] == "working":
            item["status"], item["msg"] = "error", msg[:200]
        job["phase"], job["error"] = "error", msg[:300]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        job["finished"] = time.time()


@app.post("/api/file")
def upload_file():
    if not authed():
        return jsonify(error="Non connecté"), 401
    if not GROQ_KEY:
        return jsonify(error="La clé Groq n'est pas configurée : la transcription IA est impossible."), 400
    cleanup()
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="Choisis un fichier audio ou vidéo."), 400
    jid = uuid.uuid4().hex[:12]
    tmp = tempfile.mkdtemp(dir=BASE)
    ext = os.path.splitext(f.filename)[1][:10] or ".bin"
    src = os.path.join(tmp, "input" + ext)
    f.save(src)
    name = os.path.basename(f.filename)[:120]
    job = {"id": jid, "platform": "file", "source": name, "url": "", "limit": 1, "sort": "",
           "lang": (request.form.get("lang") or "").strip()[:5], "phase": "queued", "error": "",
           "source_name": name, "created": time.time(), "part": 0, "parts": 0, "duration": 0,
           "items": [{"n": 1, "url": "", "title": name, "status": "pending", "msg": "En attente",
                      "method": "", "text": ""}]}
    with LOCK:
        JOBS[jid] = job
    threading.Thread(target=run_file_job, args=(job, src), daemon=True).start()
    return jsonify(job_id=jid)


# ------------------------------------------------ téléchargement de vidéos TikTok
def safe_name(t, n=60):
    t = re.sub(r"#\S+", "", t or "")
    t = re.sub(r'[\\/:*?"<>|\n\r\t]+', " ", t)
    return re.sub(r"\s+", " ", t).strip()[:n].strip(" .") or "video"


def run_download_job(job):
    print(f"[téléchargement {job['id']}] {job['url']} (max {job['limit'] or 'toutes'})", flush=True)
    tmp = tempfile.mkdtemp(dir=BASE)
    zpath = os.path.join(BASE, job["id"] + ".zip")
    try:
        job["phase"] = "listing"
        name, vids = list_videos("tiktok", job["url"], job["limit"], job["sort"])
        if not vids:
            raise RuntimeError("Aucune vidéo trouvée sur ce compte.")
        job["source_name"] = name
        job["items"] = [{"n": i + 1, "url": v["url"], "title": v["title"], "status": "pending", "msg": "",
                         "method": "", "text": "", "size": 0, "arc": ""} for i, v in enumerate(vids)]
        job["phase"] = "transcribing"
        width = len(str(len(vids)))
        for it in job["items"]:
            if shutil.disk_usage(BASE).free < 400 * 1024 ** 2:
                it["status"], it["msg"] = "error", "Serveur plein : télécharge ce lot puis relance la suite."
                continue
            it["status"], it["msg"] = "working", "Téléchargement…"
            d = tempfile.mkdtemp(dir=tmp)
            try:
                # formats sans filigrane uniquement (TikTok marque « watermarked » les autres)
                opts = ydl_opts("tiktok", outtmpl=os.path.join(d, "v.%(ext)s"),
                                format="best[format_note!*=?atermark][vcodec!=?none][acodec!=?none]/"
                                       "best[format_note!*=?atermark]/best",
                                format_sort=["vcodec:h264"])  # H.264 : lisible partout (PC, CapCut, téléphone)
                with yt_dlp.YoutubeDL(opts) as y:
                    info = y.extract_info(it["url"], download=True)
                files = [f for f in glob.glob(os.path.join(d, "v.*")) if not f.endswith(".part")]
                if not files:
                    raise RuntimeError("Pas de vidéo (c'est peut-être un diaporama photo).")
                it["title"] = info.get("title") or info.get("description") or it["title"]
                ext = os.path.splitext(files[0])[1] or ".mp4"
                arc = f"{str(it['n']).zfill(width)} - {safe_name(it['title'])}{ext}"
                with zipfile.ZipFile(zpath, "a", zipfile.ZIP_STORED, allowZip64=True) as z:
                    z.write(files[0], arc)
                wm = "atermark" in (info.get("format_note") or "")
                it.update(status="done", msg="", size=os.path.getsize(files[0]), arc=arc,
                          method="avec filigrane" if wm else "sans filigrane")
                job["zip_size"] = os.path.getsize(zpath)
            except Exception as e:
                msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")
                if "photo" in msg.lower() or "slideshow" in msg.lower() or "requested format" in msg.lower():
                    msg = "Diaporama photo ou format indisponible : ignoré."
                it["status"], it["msg"] = "error", msg[:200]
            finally:
                shutil.rmtree(d, ignore_errors=True)
        ok = sum(i["status"] == "done" for i in job["items"])
        if not ok:
            raise RuntimeError(next((i["msg"] for i in job["items"] if i["msg"]), "Aucune vidéo n'a pu être téléchargée."))
        job["file"] = zpath
        job["filename"] = f"TikTok_{safe_name(name, 40).replace(' ', '_')}_{ok}_videos.zip"
        job["phase"] = "done"
        print(f"[téléchargement {job['id']}] terminé : {ok}/{len(job['items'])} vidéos, "
              f"{job.get('zip_size', 0) // 1024 ** 2} Mo", flush=True)
    except Exception as e:
        msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")
        print(f"[téléchargement {job['id']}] erreur : {msg[:200]}", flush=True)
        job["phase"], job["error"] = "error", msg[:300]
        if not job.get("file") and os.path.exists(zpath):
            os.remove(zpath)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        job["finished"] = time.time()


@app.post("/api/download")
def start_download():
    if not authed():
        return jsonify(error="Non connecté"), 401
    cleanup()
    d = request.get_json(force=True)
    raw = (d.get("source") or "").strip()
    if not raw:
        return jsonify(error="Indique un compte TikTok ou un lien."), 400
    if any(j["platform"] == "download" and j["phase"] in ("queued", "listing", "transcribing") for j in JOBS.values()):
        return jsonify(error="Un téléchargement est déjà en cours. Attends qu'il se termine."), 409
    try:
        limit = int(d.get("limit") or 0)
    except ValueError:
        limit = 0
    job = {"id": uuid.uuid4().hex[:12], "platform": "download", "source": raw,
           "url": normalize_source("tiktok", raw, None), "limit": max(0, min(limit, 1000)),
           "sort": "popular" if d.get("sort") == "popular" else "recent", "lang": "",
           "phase": "queued", "items": [], "source_name": raw, "error": "", "created": time.time(), "zip_size": 0}
    with LOCK:
        JOBS[job["id"]] = job
    threading.Thread(target=run_download_job, args=(job,), daemon=True).start()
    return jsonify(job_id=job["id"])


# ------------------------------------------------ photos / médias d'un post (X, TikTok)
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/126.0.0.0 Safari/537.36")


def http_get(url, **kw):
    h = {"User-Agent": UA, "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8"}
    h.update(kw.pop("headers", {}))
    return requests.get(url, headers=h, timeout=kw.pop("timeout", 30), **kw)


def tw_orig(u):
    """pbs.twimg.com/media/XXX.jpg → version d'origine (pleine qualité)."""
    m = re.match(r"(https://pbs\.twimg\.com/media/[^.?]+)(?:\.(\w+))?", u or "")
    if not m:
        return u
    ext = m.group(2) or (re.search(r"format=(\w+)", u).group(1) if "format=" in u else "jpg")
    return f"{m.group(1)}?format={ext}&name=orig"


def extract_twitter(url):
    m = re.search(r"/status(?:es)?/(\d+)", url)
    if not m:
        raise RuntimeError("Ce lien X/Twitter ne contient pas de numéro de post.")
    tid, errors = m.group(1), []
    # 1) fxtwitter
    try:
        t = http_get(f"https://api.fxtwitter.com/status/{tid}").json()["tweet"]
        media = (t.get("media") or {})
        out = [{"kind": "photo", "url": tw_orig(p["url"])} for p in media.get("photos") or []]
        out += [{"kind": "video", "url": v["url"]} for v in media.get("videos") or []]
        if out:
            return (t.get("author") or {}).get("screen_name", ""), t.get("text", ""), out
    except Exception as e:
        errors.append(f"fx {e}")
    # 2) vxtwitter
    try:
        t = http_get(f"https://api.vxtwitter.com/Twitter/status/{tid}").json()
        out = []
        for x in t.get("media_extended") or []:
            if x.get("type") == "image":
                out.append({"kind": "photo", "url": tw_orig(x["url"])})
            elif x.get("type") in ("video", "gif"):
                out.append({"kind": "video", "url": x["url"]})
        if out:
            return t.get("user_screen_name", ""), t.get("text", ""), out
    except Exception as e:
        errors.append(f"vx {e}")
    # 3) syndication officielle
    try:
        t = http_get(f"https://cdn.syndication.twimg.com/tweet-result?id={tid}&lang=fr&token=a").json()
        out = []
        for x in t.get("mediaDetails") or []:
            if x.get("type") == "photo":
                out.append({"kind": "photo", "url": tw_orig(x["media_url_https"])})
            else:
                vs = [v for v in (x.get("video_info") or {}).get("variants", []) if v.get("content_type") == "video/mp4"]
                if vs:
                    out.append({"kind": "video", "url": max(vs, key=lambda v: v.get("bitrate", 0))["url"]})
        if out:
            return (t.get("user") or {}).get("screen_name", ""), t.get("text", ""), out
    except Exception as e:
        errors.append(f"syn {e}")
    print(f"[médias] X {tid} : rien trouvé ({'; '.join(errors)[:300]})", flush=True)
    raise RuntimeError("Aucune photo ni vidéo trouvée dans ce post (post supprimé, privé ou réservé aux abonnés ?).")


def pick_img(urls):
    urls = [u for u in urls or [] if u]

    def rank(u):
        for i, pat in enumerate((r"\.jpe?g|format=jpe?g", r"\.png", r"\.webp")):
            if re.search(pat, u, re.I):
                return i
        return 5 if re.search(r"\.hei[cf]|\.avif", u, re.I) else 4
    return min(urls, key=rank) if urls else None


def find_key(obj, key, depth=0):
    """Cherche récursivement un dict qui contient `key` (structure TikTok changeante)."""
    if depth > 12:
        return None
    if isinstance(obj, dict):
        if key in obj and obj[key]:
            return obj
        vals = obj.values()
    elif isinstance(obj, list):
        vals = obj
    else:
        return None
    for v in vals:
        r = find_key(v, key, depth + 1)
        if r is not None:
            return r
    return None


def tt_item_to_media(item):
    author = (item.get("author") or {})
    author = author.get("uniqueId") or author.get("unique_id") or "" if isinstance(author, dict) else str(author)
    desc = item.get("desc") or item.get("title") or ""
    imgs = (item.get("imagePost") or item.get("image_post_info") or {}).get("images") or []
    out = []
    for i in imgs:
        urls = ((i.get("imageURL") or i.get("display_image") or {}).get("urlList")
                or (i.get("display_image") or {}).get("url_list") or [])
        u = pick_img(urls)
        if u:
            out.append({"kind": "photo", "url": u})
    if out:
        music = (item.get("music") or {})
        music = music.get("playUrl") or ((music.get("play_url") or {}).get("url_list") or [None])[0]
        if music:
            out.append({"kind": "audio", "url": music})
    return author, desc, out


def extract_tiktok(url):
    errors = []
    sess = requests.Session()
    sess.headers.update({"User-Agent": UA, "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
                         "Referer": "https://www.tiktok.com/"})
    # 1) résoudre les liens courts (vt.tiktok.com, vm.tiktok.com)
    try:
        r = sess.get(url, timeout=30, allow_redirects=True)
        url = r.url.split("?")[0]
    except Exception as e:
        errors.append(f"résolution {e}")
        r = None
    m = re.search(r"/(?:photo|video)/(\d+)", url)
    pid = m.group(1) if m else None
    is_photo = "/photo/" in url
    pages = [(url, r)]
    if pid:
        pages.append((url.replace("/photo/", "/video/"), None))
    # 2) données intégrées à la page (lien d'origine, puis variante /video/)
    for u, resp in pages:
        try:
            resp = resp if resp is not None and resp.url.split("?")[0] == u else sess.get(u, timeout=30)
            mm = re.search(r'<script[^>]+id="__UNIVERSAL_DATA_FOR_REHYDRATION__"[^>]*>(.*?)</script>', resp.text, re.S)
            if not mm:
                errors.append(f"page {u[-30:]} : {resp.status_code}, {len(resp.text)} o, pas de données")
                continue
            data = json.loads(mm.group(1))
            holder = find_key(data, "imagePost") or find_key(data, "itemStruct")
            item = holder.get("itemStruct", holder) if holder else None
            if item:
                author, desc, out = tt_item_to_media(item)
                if out:
                    print(f"[médias] TikTok via page ({len(out)} fichiers)", flush=True)
                    return author, desc, out
                if item.get("video") and not is_photo:
                    return author, desc, [{"kind": "ytdlp", "url": url}]
            scope = data.get("__DEFAULT_SCOPE__", {})
            st = [(k, (v or {}).get("statusCode")) for k, v in scope.items() if isinstance(v, dict) and "statusCode" in v]
            errors.append(f"page {u[-30:]} : clés {list(scope)[:6]} statuts {st[:3]}")
        except Exception as e:
            errors.append(f"page {u[-30:]} : {type(e).__name__} {str(e)[:80]}")
    # 3) API web de TikTok (avec les cookies de la session)
    if pid:
        try:
            j = sess.get("https://www.tiktok.com/api/item/detail/",
                         params={"itemId": pid, "aid": "1988", "app_language": "fr"}, timeout=30)
            data = j.json() if j.text.strip() else {}
            holder = find_key(data, "imagePost") or find_key(data, "itemStruct")
            item = holder.get("itemStruct", holder) if holder else None
            if item:
                author, desc, out = tt_item_to_media(item)
                if out:
                    print(f"[médias] TikTok via API web ({len(out)} fichiers)", flush=True)
                    return author, desc, out
            errors.append(f"api : {j.status_code}, {len(j.text)} o, statut {data.get('statusCode')}")
        except Exception as e:
            errors.append(f"api : {type(e).__name__} {str(e)[:80]}")
    # 4) services de secours
    full = url if not pid else f"https://www.tiktok.com/@i/{'photo' if is_photo else 'video'}/{pid}"
    for name, call in (
        ("tikwm", lambda: requests.post("https://www.tikwm.com/api/", headers={"User-Agent": UA},
                                        data={"url": full, "hd": "1", "web": "1"}, timeout=30)),
        ("tiklydown", lambda: http_get("https://api.tiklydown.eu.org/api/download", params={"url": full})),
    ):
        try:
            resp = call()
            d = resp.json()
            d = d.get("data", d) if isinstance(d, dict) else {}
            imgs = d.get("images") or []
            urls = [(x if isinstance(x, str) else x.get("url")) for x in imgs]
            urls = [("https://www.tikwm.com" + u) if u and u.startswith("/") else u for u in urls if u]
            if urls:
                out = [{"kind": "photo", "url": u} for u in urls]
                music = d.get("music")
                music = music if isinstance(music, str) else (music or {}).get("play_url")
                if music:
                    out.append({"kind": "audio", "url": ("https://www.tikwm.com" + music) if music.startswith("/") else music})
                a = d.get("author") or {}
                print(f"[médias] TikTok via {name} ({len(out)} fichiers)", flush=True)
                return (a.get("unique_id") or a.get("unique_id") or "") if isinstance(a, dict) else "", d.get("title", ""), out
            errors.append(f"{name} : {resp.status_code}, pas d'images ({str(d)[:80]})")
        except Exception as e:
            errors.append(f"{name} : {type(e).__name__} {str(e)[:60]}")
    print(f"[médias] TikTok {url} : rien trouvé | " + " | ".join(errors)[:1500], flush=True)
    if is_photo:
        raise RuntimeError("TikTok n'a pas voulu donner les photos de ce post. Réessaie dans une minute ; "
                           "si ça bloque encore, préviens Claude.")
    return "", "", [{"kind": "ytdlp", "url": url.replace("/photo/", "/video/")}]


def extract_post(url):
    host = re.sub(r"^www\.|^m\.|^mobile\.", "", requests.utils.urlparse(url).netloc.lower())
    if host in ("x.com", "twitter.com", "fxtwitter.com", "vxtwitter.com", "fixupx.com", "t.co"):
        if host == "t.co":
            url = http_get(url, allow_redirects=True).url
        return "x", *extract_twitter(url)
    if host.endswith("tiktok.com"):
        return "tiktok", *extract_tiktok(url)
    return "autre", "", "", [{"kind": "ytdlp", "url": url}]


EXT = {"image/jpeg": ".jpg", "image/png": ".png", "image/webp": ".webp", "image/heic": ".heic",
       "image/heif": ".heic", "image/avif": ".avif",
       "video/mp4": ".mp4", "audio/mpeg": ".mp3", "audio/mp4": ".m4a", "audio/x-m4a": ".m4a"}


def fetch_media(it, d, referer):
    if it["kind"] == "ytdlp":
        opts = ydl_opts("tiktok" if "tiktok" in it["url"] else "other", outtmpl=os.path.join(d, f"{it['n']:02d}.%(ext)s"),
                        format="best[format_note!*=?atermark][vcodec!=?none][acodec!=?none]/best[format_note!*=?atermark]/best",
                        format_sort=["vcodec:h264"])
        with yt_dlp.YoutubeDL(opts) as y:
            y.extract_info(it["url"], download=True)
        files = [f for f in glob.glob(os.path.join(d, f"{it['n']:02d}.*")) if not f.endswith(".part")]
        if not files:
            raise RuntimeError("Aucun média téléchargeable à ce lien.")
        it["kind"] = "video"
        return files[0]
    with http_get(it["url"], headers={"Referer": referer}, stream=True, timeout=120) as r:
        r.raise_for_status()
        ctype = r.headers.get("content-type", "").split(";")[0].strip()
        ext = EXT.get(ctype) or os.path.splitext(requests.utils.urlparse(it["url"]).path)[1][:5] or \
            {"photo": ".jpg", "video": ".mp4", "audio": ".mp3"}[it["kind"]]
        path = os.path.join(d, f"{it['n']:02d}{ext}")
        with open(path, "wb") as f:
            for chunk in r.iter_content(256 * 1024):
                f.write(chunk)
    if it["kind"] == "photo":
        path = to_jpeg(path)
    return path


def img_kind(path):
    with open(path, "rb") as f:
        h = f.read(32)
    if h[:3] == b"\xff\xd8\xff":
        return "jpeg"
    if h[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if h[:4] == b"RIFF" and h[8:12] == b"WEBP":
        return "webp"
    if h[4:8] == b"ftyp":
        return "avif" if b"avi" in h[8:32] else "heic"
    return "autre"


def to_jpeg(path):
    """Photos en JPEG pleine taille : visibles dans tous les navigateurs et dans la galerie."""
    kind = img_kind(path)
    if kind in ("jpeg", "png"):
        good = os.path.splitext(path)[0] + (".jpg" if kind == "jpeg" else ".png")
        if good != path:
            os.replace(path, good)
        return good
    try:
        from PIL import Image
        try:
            import pillow_heif
            pillow_heif.register_heif_opener()
            if hasattr(pillow_heif, "register_avif_opener"):
                pillow_heif.register_avif_opener()
        except Exception:
            pass
        jpg = os.path.splitext(path)[0] + ".jpg"
        Image.open(path).convert("RGB").save(jpg, "JPEG", quality=95, subsampling=0)
        if jpg != path:
            os.remove(path)
        return jpg
    except Exception as e:
        print(f"[médias] conversion impossible ({kind}) : {e}", flush=True)
        return path


def run_media_job(job):
    print(f"[médias {job['id']}] {job['url']}", flush=True)
    d = os.path.join(BASE, "m_" + job["id"])
    os.makedirs(d, exist_ok=True)
    job["dir"] = d
    try:
        job["phase"] = "listing"
        site, author, text, media = extract_post(job["url"])
        job["site"] = site
        job["source_name"] = ("@" + author if author else "Post") + (f" · {re.sub(chr(10), ' ', text)[:60]}" if text else "")
        labels = {"photo": "Photo", "video": "Vidéo", "audio": "Musique", "ytdlp": "Vidéo"}
        job["items"] = [{"n": i + 1, "url": m["url"], "kind": m["kind"], "title": f"{labels[m['kind']]} {i + 1}",
                         "status": "pending", "msg": "", "method": "", "text": "", "size": 0, "name": ""}
                        for i, m in enumerate(media)]
        job["phase"] = "transcribing"
        referer = "https://www.tiktok.com/" if site == "tiktok" else "https://x.com/"
        base = re.sub(r"[^\w-]+", "_", author or site)[:30] or "post"
        for it in job["items"]:
            it["status"], it["msg"] = "working", "Téléchargement…"
            try:
                path = fetch_media(it, d, referer)
                it["title"] = f"{labels[it['kind']]} {it['n']}"
                it["name"] = f"{base}_{it['n']:02d}{os.path.splitext(path)[1]}"
                it.update(path=path, size=os.path.getsize(path), status="done", msg="",
                          method="qualité d'origine" if it["kind"] == "photo" else "")
            except Exception as e:
                msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")
                it["status"], it["msg"] = "error", msg[:200]
        done = [i for i in job["items"] if i["status"] == "done"]
        if not done:
            raise RuntimeError(job["items"][0]["msg"] if job["items"] else "Rien à télécharger.")
        job["file"] = os.path.join(BASE, job["id"] + ".zip")
        with zipfile.ZipFile(job["file"], "w", zipfile.ZIP_STORED) as z:
            for it in done:
                z.write(it["path"], it["name"])
        job["filename"] = f"{base}_{len(done)}_medias.zip"
        job["zip_size"] = os.path.getsize(job["file"])
        job["phase"] = "done"
        print(f"[médias {job['id']}] terminé : {len(done)} fichiers ({site})", flush=True)
    except Exception as e:
        msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")
        if "Unsupported URL" in msg:
            msg = "Lien non pris en charge. Colle un lien de post X/Twitter ou TikTok."
        print(f"[médias {job['id']}] erreur : {msg[:200]}", flush=True)
        job["phase"], job["error"] = "error", msg[:300]
    finally:
        job["finished"] = time.time()


@app.post("/api/media")
def start_media():
    if not authed():
        return jsonify(error="Non connecté"), 401
    cleanup()
    url = ((request.get_json(force=True) or {}).get("url") or "").strip()
    m = re.search(r"https?://\S+", url)
    if not m:
        return jsonify(error="Colle le lien complet du post (il commence par https://)."), 400
    job = {"id": uuid.uuid4().hex[:12], "platform": "media", "source": m.group(0), "url": m.group(0),
           "limit": 0, "sort": "", "lang": "", "phase": "queued", "items": [], "source_name": "Post",
           "error": "", "created": time.time(), "zip_size": 0}
    with LOCK:
        JOBS[job["id"]] = job
    threading.Thread(target=run_media_job, args=(job,), daemon=True).start()
    return jsonify(job_id=job["id"])


@app.get("/api/job/<jid>/m/<int:n>")
def job_media(jid, n):
    if not authed():
        return "Non connecté", 401
    j = JOBS.get(jid)
    it = next((i for i in (j or {}).get("items", []) if i["n"] == n and i["status"] == "done"), None)
    if not it or not os.path.exists(it.get("path", "")):
        return "Fichier introuvable (le serveur a peut-être redémarré). Relance.", 404
    return send_file(it["path"], as_attachment=request.args.get("dl") == "1", download_name=it["name"], max_age=3600)


@app.get("/api/job/<jid>/zip")
def job_zip(jid):
    if not authed():
        return "Non connecté", 401
    j = JOBS.get(jid)
    if not j or j.get("phase") != "done" or not os.path.exists(j.get("file", "")):
        return "Fichier introuvable (le serveur a peut-être redémarré). Relance le téléchargement.", 404
    return send_file(j["file"], as_attachment=True, download_name=j["filename"], mimetype="application/zip")


@app.get("/api/job/<jid>/video/<int:n>")
def job_video(jid, n):
    if not authed():
        return "Non connecté", 401
    j = JOBS.get(jid)
    if not j or j.get("phase") != "done" or not os.path.exists(j.get("file", "")):
        return "Fichier introuvable (le serveur a peut-être redémarré). Relance le téléchargement.", 404
    it = next((i for i in j["items"] if i["n"] == n and i["status"] == "done"), None)
    if not it:
        return "Vidéo introuvable", 404
    z = zipfile.ZipFile(j["file"])
    return send_file(z.open(it["arc"]), as_attachment=True, download_name=it["arc"],
                     mimetype="video/mp4", max_age=0)


# ------------------------------------------------ traduction
CHAT_MODELS = [m.strip() for m in os.environ.get(
    "GROQ_CHAT_MODELS", "openai/gpt-oss-120b,llama-3.3-70b-versatile,openai/gpt-oss-20b,llama-3.1-8b-instant").split(",") if m.strip()]
GOOD_MODEL = {"name": None}


def groq_chat(messages, max_tokens=6000, json_mode=False, tag="traduction"):
    models = ([GOOD_MODEL["name"]] if GOOD_MODEL["name"] else []) + [m for m in CHAT_MODELS if m != GOOD_MODEL["name"]]
    last = ""
    for model in models:
        body = {"model": model, "messages": messages, "temperature": 0.3 if json_mode else 0.2, "max_tokens": max_tokens}
        if json_mode:
            body["response_format"] = {"type": "json_object"}
        if model.startswith("openai/gpt-oss"):
            body["reasoning_effort"] = "low"  # moins de jetons consommés sur la limite gratuite
        for attempt in range(6):
            r = requests.post("https://api.groq.com/openai/v1/chat/completions",
                              headers={"Authorization": f"Bearer {GROQ_KEY}"}, json=body, timeout=240)
            if r.status_code == 200:
                GOOD_MODEL["name"] = model
                return (r.json()["choices"][0]["message"].get("content") or "").strip()
            if r.status_code in (429, 500, 502, 503):
                try:
                    wait = float(r.headers.get("retry-after"))
                except (TypeError, ValueError):
                    wait = 10 * (attempt + 1)
                time.sleep(min(max(wait, 2), 90))
                continue
            if r.status_code == 401:
                raise RuntimeError("Clé Groq invalide.")
            if r.status_code == 400 and ("response_format" in body or "reasoning_effort" in body) and \
                    any(k in r.text for k in ("response_format", "reasoning_effort", "json")):
                body.pop("response_format", None)
                body.pop("reasoning_effort", None)
                continue  # même modèle, sans les options qu'il refuse
            last = f"{model} : {r.status_code} {r.text[:120]}"
            print(f"[{tag}] modèle refusé {last}", flush=True)
            break  # modèle indisponible ou requête trop grosse : on essaie le suivant
        else:
            raise RuntimeError("Limite Groq atteinte, réessaie dans quelques minutes.")
    raise RuntimeError(f"Aucun modèle d'IA disponible chez Groq ({last}).")


def chunk_text(text, size=2500):
    """Découpe en morceaux d'environ `size` caractères sans casser les paragraphes."""
    chunks, cur = [], ""
    for para in text.split("\n"):
        pieces = [para] if len(para) <= size else re.split(r"(?<=[.!?…])\s+", para)
        for i, piece in enumerate(pieces):
            sep = "\n" if i == 0 else " "
            if cur and len(cur) + len(piece) + 1 > size:
                chunks.append(cur)
                cur, sep = "", ""
            cur = (cur + sep + piece) if cur else piece
    if cur.strip():
        chunks.append(cur)
    return [c for c in chunks if c.strip()]


def run_translate_job(job, text):
    print(f"[traduction {job['id']}] {len(text)} caractères → {job['target']}", flush=True)
    item = job["items"][0]
    try:
        job["phase"] = "transcribing"
        chunks = chunk_text(text)
        job["parts"] = len(chunks)
        system = (f"Tu es un traducteur professionnel. Traduis le texte de l'utilisateur en {job['target']}. "
                  "Garde exactement le sens, le ton et l'émotion, ainsi que les retours à la ligne et la mise en forme. "
                  "Garde les noms propres, et traduis les citations religieuses avec leurs formulations reconnues "
                  "dans la langue cible quand elles existent. Si une partie est déjà dans la langue cible, laisse-la telle quelle. "
                  "Réponds uniquement avec la traduction, sans phrase d'introduction, sans note ni commentaire.")
        out = []
        for i, c in enumerate(chunks, 1):
            item["status"], item["msg"] = "working", f"Traduction · partie {i}/{len(chunks)}"
            job["part"] = i - 1
            out.append(groq_chat([{"role": "system", "content": system}, {"role": "user", "content": c}]))
            job["part"] = i
        result = "\n".join(out).strip()
        if not result:
            raise RuntimeError("La traduction est revenue vide. Réessaie.")
        item.update(text=result, status="done", msg="", method="IA")
        job["file"] = os.path.join(BASE, job["id"] + ".docx")
        safe = re.sub(r"[^\w-]+", "_", job["target"])[:20].strip("_") or "traduction"
        job["filename"] = f"Traduction_{safe}.docx"
        build_docx(job, job["file"])
        job["phase"] = "done"
        print(f"[traduction {job['id']}] terminée ({GOOD_MODEL['name']}) : {len(result)} caractères", flush=True)
    except Exception as e:
        msg = str(e)
        print(f"[traduction {job['id']}] erreur : {msg[:200]}", flush=True)
        item["status"], item["msg"] = "error", msg[:200]
        job["phase"], job["error"] = "error", msg[:300]
    finally:
        job["finished"] = time.time()


@app.post("/api/translate")
def start_translate():
    if not authed():
        return jsonify(error="Non connecté"), 401
    if not GROQ_KEY:
        return jsonify(error="La clé Groq n'est pas configurée : la traduction est impossible."), 400
    cleanup()
    d = request.get_json(force=True)
    text = (d.get("text") or "").replace("\r\n", "\n").strip()
    target = re.sub(r"\s+", " ", (d.get("target") or "")).strip()[:40]
    if not text:
        return jsonify(error="Colle d'abord un texte à traduire."), 400
    if not target:
        return jsonify(error="Choisis la langue de traduction."), 400
    if len(text) > 200000:
        return jsonify(error="Texte trop long (200 000 caractères maximum). Coupe-le en deux."), 400
    first = re.sub(r"\s+", " ", text)[:50].strip()
    job = {"id": uuid.uuid4().hex[:12], "platform": "translate", "source": first, "url": "", "limit": 1,
           "sort": "", "lang": "", "target": target, "phase": "queued", "error": "", "created": time.time(),
           "source_name": f"{first}…" if len(text) > 50 else first, "part": 0, "parts": 0,
           "items": [{"n": 1, "url": "", "title": first, "status": "pending", "msg": "En attente",
                      "method": "", "text": ""}]}
    with LOCK:
        JOBS[job["id"]] = job
    threading.Thread(target=run_translate_job, args=(job, text), daemon=True).start()
    return jsonify(job_id=job["id"])


@app.post("/api/extract")
def extract_text():
    if not authed():
        return jsonify(error="Non connecté"), 401
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify(error="Choisis un fichier."), 400
    ext = os.path.splitext(f.filename)[1].lower()
    try:
        if ext == ".docx":
            text = "\n".join(p.text for p in Document(f.stream).paragraphs)
        elif ext in (".txt", ".md", ".srt", ".vtt", ""):
            raw = f.read()
            text = raw.decode("utf-8") if raw[:3] != b"\xef\xbb\xbf" else raw[3:].decode("utf-8")
        else:
            return jsonify(error="Format non pris en charge : utilise un fichier .txt ou .docx (Word)."), 400
    except UnicodeDecodeError:
        text = raw.decode("latin-1")
    except Exception:
        return jsonify(error="Impossible de lire ce fichier."), 400
    return jsonify(text=re.sub(r"\n{3,}", "\n\n", text).strip())


# ------------------------------------------------ analyse d'audience et monétisation
AUDIT_SCHEMA = """{
 "resume": "3 phrases : de quoi parle la chaîne, pour qui, et la meilleure opportunité de vente",
 "niche": {"principale": "", "sous_niches": [""], "positionnement": "ce qui rend cette chaîne unique"},
 "audience": {"profil": "qui regarde (âge estimé, situation, centres d'intérêt)", "langue_pays": "", "niveau": "débutant / intermédiaire / avancé", "motivations": [""], "pouvoir_achat": "estimation honnête et ce que ça implique pour les prix"},
 "douleurs": [{"probleme": "", "preuve": "titre, extrait de script ou commentaire réel qui le montre"}],
 "desirs": [{"desir": "", "preuve": ""}],
 "sujets_porteurs": [{"sujet": "", "pourquoi": "", "exemples": ["titres réels"]}],
 "produits": [{"nom": "", "format": "ebook, mini-formation vidéo, programme X jours, guide audio, templates, coaching, communauté payante…", "promesse": "le résultat concret promis", "pour_qui": "", "pourquoi_ca_marchera": "lien direct avec les preuves", "contenu": ["modules ou chapitres"], "prix_fcfa": "", "prix_eur": "", "difficulte": "facile / moyen / difficile", "premier_pas": "action concrète pour le valider cette semaine"}],
 "meilleur_premier_produit": {"nom": "", "raison": ""},
 "plan_lancement": [{"etape": "", "detail": ""}],
 "strategie_vente": ["conseils concrets pour vendre via les vidéos, la bio, les lives, les commentaires épinglés, WhatsApp…"],
 "idees_videos_pour_vendre": ["titres de vidéos qui amènent naturellement vers le produit"],
 "autres_revenus": ["affiliation, partenariats, dons en live, abonnements…, adaptés à cette chaîne"],
 "pieges_a_eviter": [""]
}"""


def audit_listing(platform, url, limit):
    opts = ydl_opts(platform, extract_flat="in_playlist", skip_download=True, ignoreerrors=True, playlistend=limit)
    with yt_dlp.YoutubeDL(opts) as y:
        info = y.extract_info(url, download=False)
    if not info:
        opts.pop("ignoreerrors")
        with yt_dlp.YoutubeDL(opts) as y:
            info = y.extract_info(url, download=False)
    entries = flatten(info) if info.get("entries") is not None else [info]
    vids = [{"url": entry_url(e, platform), "title": (e.get("title") or e.get("description") or "").strip(),
             "views": e.get("view_count") or 0, "duration": e.get("duration") or 0} for e in entries]
    vids = [v for v in vids if v["url"]][:limit]
    meta = {"name": info.get("channel") or info.get("uploader") or info.get("title") or url,
            "followers": info.get("channel_follower_count") or 0,
            "description": (info.get("description") or "").strip()[:700]}
    return meta, vids


def audit_transcript(platform, url, lang, tmp):
    """Début du script d'une vidéo (sous-titres, sinon IA sur les 6 premières minutes)."""
    with yt_dlp.YoutubeDL(ydl_opts(platform, skip_download=True)) as y:
        info = y.extract_info(url, download=False)
        cap = pick_caption(info, lang)
        if cap:
            try:
                t = clean_caption_text(y.urlopen(cap[0]).read().decode("utf-8", "ignore"), cap[1])
                if len(t) > 80:
                    return t
            except Exception:
                pass
    if not GROQ_KEY:
        return ""
    d = tempfile.mkdtemp(dir=tmp)
    with yt_dlp.YoutubeDL(ydl_opts(platform, format="bestaudio/worst[ext=mp4]/worst/best",
                                   outtmpl=os.path.join(d, "a.%(ext)s"))) as y:
        y.extract_info(url, download=True)
    media = [f for f in glob.glob(os.path.join(d, "a.*")) if not f.endswith(".part")]
    if not media:
        return ""
    out = os.path.join(d, "clip.mp3")
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", media[0], "-t", "360", "-vn", "-ac", "1",
                    "-ar", "16000", "-b:a", "32k", out], check=True)
    return groq_transcribe(out, lang)


def audit_comments(url, n=30):
    opts = ydl_opts("youtube", skip_download=True, getcomments=True,
                    extractor_args={"youtube": {"max_comments": [str(n), str(n), "0", "0"], "comment_sort": ["top"]}})
    with yt_dlp.YoutubeDL(opts) as y:
        info = y.extract_info(url, download=False)
    cs = sorted(info.get("comments") or [], key=lambda c: c.get("like_count") or 0, reverse=True)
    return [(c.get("text") or "").strip().replace("\n", " ")[:220] for c in cs[:n] if (c.get("text") or "").strip()]


def parse_json_loose(t):
    t = re.sub(r"^```(?:json)?|```$", "", t.strip(), flags=re.M).strip()
    try:
        return json.loads(t)
    except Exception:
        a, b = t.find("{"), t.rfind("}")
        if a >= 0 and b > a:
            return json.loads(t[a:b + 1])
        raise


def fmt_num(n):
    n = int(n or 0)
    return f"{n / 1e6:.1f} M".replace(".", ",") if n >= 1e6 else f"{n / 1e3:.0f} k" if n >= 1e4 else str(n)


def run_audit_job(job):
    print(f"[analyse {job['id']}] {job['src_platform']} {job['url']} ({job['limit']} vidéos, approfondie={job['deep']})", flush=True)
    tmp = tempfile.mkdtemp(dir=BASE)
    try:
        job["phase"] = "listing"
        job["step"], job["pct"] = "Lecture des vidéos de la chaîne…", 5
        meta, vids = audit_listing(job["src_platform"], job["url"], job["limit"])
        if not vids:
            raise RuntimeError("Aucune vidéo trouvée sur cette chaîne ou ce compte.")
        job["source_name"] = meta["name"]
        job["phase"] = "transcribing"
        by_views = sorted(vids, key=lambda v: v["views"], reverse=True)
        has_views = any(v["views"] for v in vids)
        views = [v["views"] for v in vids if v["views"]]
        stats = {"videos": len(vids), "followers": meta["followers"],
                 "moyenne_vues": int(sum(views) / len(views)) if views else 0,
                 "top": [{"title": v["title"][:120], "views": v["views"], "url": v["url"]} for v in by_views[:5]] if has_views else []}
        job["stats"] = stats
        scripts, comments = [], []
        if job["deep"]:
            pick = (by_views if has_views else vids)[:5]
            for k, v in enumerate(pick, 1):
                job["step"], job["pct"] = f"Lecture des scripts · vidéo {k}/{len(pick)}", 10 + k * 9
                try:
                    t = audit_transcript(job["src_platform"], v["url"], job["lang"], tmp)
                    if t:
                        scripts.append((v["title"], t[:1100]))
                except Exception as e:
                    msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")[:120]
                    print(f"[analyse {job['id']}] script ignoré : {msg}", flush=True)
                    if job["src_platform"] == "youtube" and is_bot_block(msg):
                        break
            if job["src_platform"] == "youtube":
                for k, v in enumerate(pick[:3], 1):
                    job["step"], job["pct"] = f"Lecture des commentaires · vidéo {k}/3", 58 + k * 5
                    try:
                        comments += audit_comments(v["url"], 25)
                    except Exception as e:
                        print(f"[analyse {job['id']}] commentaires ignorés : {str(e)[:120]}", flush=True)
        job["step"], job["pct"] = "Analyse de l'audience par l'IA…", 78
        lines = [f"{v['title'][:110]} | {fmt_num(v['views']) + ' vues' if v['views'] else 'vues inconnues'}"
                 + (f" | {int(v['duration'] // 60)} min" if v["duration"] else "") for v in (by_views if has_views else vids)[:45]]
        data = [f"PLATEFORME : {'YouTube' if job['src_platform'] == 'youtube' else 'TikTok'}",
                f"NOM : {meta['name']}",
                f"ABONNÉS : {fmt_num(meta['followers']) if meta['followers'] else 'inconnu'}",
                f"VUES MOYENNES : {fmt_num(stats['moyenne_vues']) if stats['moyenne_vues'] else 'inconnues'}"]
        if meta["description"]:
            data.append(f"DESCRIPTION DE LA CHAÎNE : {meta['description']}")
        data.append(f"\nVIDÉOS ({len(lines)} sur {len(vids)} analysées, classées par vues décroissantes) :\n" + "\n".join(lines))
        if scripts:
            data.append("\nDÉBUT DES SCRIPTS DES VIDÉOS LES PLUS VUES :\n" + "\n\n".join(f"« {t[:90]} » : {x}" for t, x in scripts))
        if comments:
            data.append("\nCOMMENTAIRES LES PLUS AIMÉS :\n" + "\n".join(f"- {c}" for c in comments[:45]))
        system = (
            "Tu es un expert en monétisation d'audience et en création de produits digitaux pour les créateurs "
            "francophones (Afrique francophone et diaspora). On te donne les données réelles d'une chaîne. "
            "Analyse la niche, l'audience, ses problèmes et ses envies, puis propose des produits digitaux "
            "qu'elle achètera vraiment. Règles : appuie chaque affirmation sur les données fournies (cite des titres, "
            "des extraits ou des commentaires réels dans les champs « preuve » et « pourquoi_ca_marchera ») ; "
            "n'invente aucun chiffre ni aucun commentaire ; propose entre 5 et 7 produits variés, du plus simple "
            "au plus ambitieux ; donne des prix réalistes pour le pouvoir d'achat de l'audience (FCFA et euros) ; "
            "sois concret, direct et actionnable ; plan de lancement en 5 à 8 étapes ; écris tout en français. "
            "Réponds UNIQUEMENT avec un objet JSON valide qui suit exactement cette structure :\n" + AUDIT_SCHEMA)
        raw = groq_chat([{"role": "system", "content": system}, {"role": "user", "content": "\n".join(data)[:14000]}],
                        max_tokens=7000, json_mode=True, tag="analyse")
        try:
            report = parse_json_loose(raw)
        except Exception:
            job["step"] = "Mise en forme du rapport…"
            raw = groq_chat([{"role": "system", "content": "Corrige ce texte pour qu'il devienne un objet JSON valide. Réponds uniquement avec le JSON."},
                             {"role": "user", "content": raw[:12000]}], max_tokens=7000, json_mode=True, tag="analyse")
            report = parse_json_loose(raw)
        report["_sources"] = {"videos": len(vids), "scripts": len(scripts), "commentaires": len(comments)}
        job["report"] = report
        job["file"] = os.path.join(BASE, job["id"] + ".docx")
        job["filename"] = f"Monetisation_{re.sub(r'[^\w-]+', '_', meta['name'])[:30].strip('_') or 'chaine'}.docx"
        build_audit_docx(job, job["file"])
        job["phase"], job["pct"], job["step"] = "done", 100, "Rapport prêt"
        print(f"[analyse {job['id']}] terminée ({GOOD_MODEL['name']}) : {len(report.get('produits') or [])} produits, "
              f"{len(scripts)} scripts, {len(comments)} commentaires", flush=True)
    except Exception as e:
        msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")
        if job["src_platform"] == "youtube" and is_bot_block(msg):
            msg = bot_message()
        print(f"[analyse {job['id']}] erreur : {msg[:200]}", flush=True)
        job["phase"], job["error"] = "error", msg[:300]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        job["finished"] = time.time()


def build_audit_docx(job, path):
    r = job["report"]
    doc = Document()
    doc.styles["Normal"].font.name, doc.styles["Normal"].font.size = "Calibri", Pt(11)
    doc.add_heading(f"Monétisation — {job['source_name']}", 0)
    meta = doc.add_paragraph()
    src = r.get("_sources", {})
    m = meta.add_run(f"{'YouTube' if job['src_platform'] == 'youtube' else 'TikTok'} · {src.get('videos', 0)} vidéos, "
                     f"{src.get('scripts', 0)} scripts, {src.get('commentaires', 0)} commentaires analysés · "
                     f"{datetime.datetime.now().strftime('%d/%m/%Y')}")
    m.font.color.rgb, m.font.size = RGBColor(0x66, 0x66, 0x66), Pt(10)

    def h(t):
        doc.add_heading(t, 1)

    def bullets(items):
        for x in items or []:
            if x:
                doc.add_paragraph(str(x), style="List Bullet")

    def kv(label, val):
        if val:
            p = doc.add_paragraph()
            p.add_run(label + " : ").bold = True
            p.add_run(", ".join(val) if isinstance(val, list) else str(val))

    if r.get("resume"):
        doc.add_paragraph(r["resume"])
    best = r.get("meilleur_premier_produit") or {}
    if best.get("nom"):
        h("Le produit à lancer en premier")
        p = doc.add_paragraph()
        p.add_run(best["nom"]).bold = True
        doc.add_paragraph(best.get("raison", ""))
    n, a = r.get("niche") or {}, r.get("audience") or {}
    h("Niche et audience")
    kv("Niche", n.get("principale")); kv("Sous-niches", n.get("sous_niches")); kv("Positionnement", n.get("positionnement"))
    kv("Profil", a.get("profil")); kv("Langue et pays", a.get("langue_pays")); kv("Niveau", a.get("niveau"))
    kv("Motivations", a.get("motivations")); kv("Pouvoir d'achat", a.get("pouvoir_achat"))
    for title, key, field in (("Ce qui fait mal à ton audience", "douleurs", "probleme"), ("Ce que ton audience veut", "desirs", "desir")):
        if r.get(key):
            h(title)
            for d in r[key]:
                p = doc.add_paragraph(style="List Bullet")
                p.add_run(d.get(field, "")).bold = True
                if d.get("preuve"):
                    p.add_run(f" — {d['preuve']}")
    if r.get("sujets_porteurs"):
        h("Les sujets qui marchent")
        for t in r["sujets_porteurs"]:
            p = doc.add_paragraph(style="List Bullet")
            p.add_run(t.get("sujet", "")).bold = True
            p.add_run(f" — {t.get('pourquoi', '')}")
    if r.get("produits"):
        h("Produits digitaux recommandés")
        for i, pr in enumerate(r["produits"], 1):
            doc.add_heading(f"{i}. {pr.get('nom', '')}", 2)
            kv("Format", pr.get("format")); kv("Promesse", pr.get("promesse")); kv("Pour qui", pr.get("pour_qui"))
            kv("Pourquoi ça marchera", pr.get("pourquoi_ca_marchera"))
            kv("Prix conseillé", " / ".join(x for x in (pr.get("prix_fcfa"), pr.get("prix_eur")) if x))
            kv("Difficulté", pr.get("difficulte")); kv("Premier pas", pr.get("premier_pas"))
            if pr.get("contenu"):
                doc.add_paragraph().add_run("Contenu :").bold = True
                bullets(pr["contenu"])
    if r.get("plan_lancement"):
        h("Plan de lancement")
        for i, st in enumerate(r["plan_lancement"], 1):
            p = doc.add_paragraph()
            p.add_run(f"{i}. {st.get('etape', '')} : ").bold = True
            p.add_run(st.get("detail", ""))
    for title, key in (("Comment vendre", "strategie_vente"), ("Idées de vidéos pour vendre", "idees_videos_pour_vendre"),
                       ("Autres sources de revenus", "autres_revenus"), ("Pièges à éviter", "pieges_a_eviter")):
        if r.get(key):
            h(title)
            bullets(r[key])
    doc.save(path)


@app.post("/api/audit")
def start_audit():
    if not authed():
        return jsonify(error="Non connecté"), 401
    if not GROQ_KEY:
        return jsonify(error="La clé Groq n'est pas configurée : l'analyse est impossible."), 400
    cleanup()
    d = request.get_json(force=True)
    platform = "tiktok" if d.get("platform") == "tiktok" else "youtube"
    raw = (d.get("source") or "").strip()
    if not raw:
        return jsonify(error="Colle le lien ou le nom de ta chaîne ou de ton compte."), 400
    if any(j["platform"] == "audit" and j["phase"] in ("queued", "listing", "transcribing") for j in JOBS.values()):
        return jsonify(error="Une analyse est déjà en cours. Attends qu'elle se termine."), 409
    try:
        limit = max(5, min(int(d.get("limit") or 40), 150))
    except ValueError:
        limit = 40
    job = {"id": uuid.uuid4().hex[:12], "platform": "audit", "src_platform": platform, "source": raw,
           "url": normalize_source(platform, raw, "videos"), "limit": limit, "deep": bool(d.get("deep", True)),
           "lang": (d.get("lang") or "").strip()[:5], "phase": "queued", "items": [], "source_name": raw,
           "error": "", "created": time.time(), "step": "En attente…", "pct": 0}
    with LOCK:
        JOBS[job["id"]] = job
    threading.Thread(target=run_audit_job, args=(job,), daemon=True).start()
    return jsonify(job_id=job["id"])


# ------------------------------------------------ créateur d'ebook
import ebook as EB


def run_ebook_job(job, ctx):
    print(f"[ebook {job['id']}] « {ctx['title']} » · {job['pages']} pages · images={job['images']}", flush=True)
    d = os.path.join(BASE, "e_" + job["id"])
    os.makedirs(d, exist_ok=True)
    job["dir"] = d
    chat = lambda m, **k: groq_chat(m, tag="ebook", **k)
    try:
        job["phase"], job["error"], job["resumable"] = "transcribing", "", False
        st = job.setdefault("st", {})  # état conservé : une reprise repart d'ici
        if not st.get("plan"):
            job["step"], job["pct"] = "Conception du plan du livre…", 4
            raw, n_ch = EB.gen_plan(chat, ctx, job["pages"])
            plan = parse_json_loose(raw)
            chs = [c for c in (plan.get("chapitres") or []) if c.get("titre")][:20]
            if len(chs) < 3:
                raise RuntimeError("Le plan proposé par l'IA est incomplet. Relance la création.")
            plan["chapitres"] = chs
            if ctx.get("keep_title"):
                plan["titre"] = ctx["title"]  # le titre choisi par l'auteur est conservé tel quel
            st["plan"] = plan
        plan = st["plan"]
        chs = plan["chapitres"]
        job["toc"] = [c["titre"] for c in chs]
        job["book_title"] = plan.get("titre") or ctx["title"]
        total = job["pages"] * 230
        w_intro, w_bonus, w_concl = 380, max(300, min(900, total // 12)), 320
        w_ch = max(500, (total - w_intro - w_bonus - w_concl) // len(chs))
        steps = len(chs) + 3 + (len(chs) + 1 if job["images"] else 0)
        done = [sum(1 for k in ("intro", "bonus", "concl") if st.get(k)) + sum(1 for c in chs if c.get("texte"))]

        def tick(label):
            done[0] += 1
            job["step"], job["pct"] = label, 6 + int(done[0] / steps * 86)
        if not st.get("intro"):
            job["step"] = "Rédaction de l'introduction…"
            st["intro"] = EB.gen_section(chat, plan, ctx, "intro", w_intro); tick("Introduction écrite")
        intro = st["intro"]
        for k, c in enumerate(chs):
            if c.get("texte"):
                continue
            prev = chs[k - 1]["titre"] if k else ""
            secs = [x for x in (c.get("sections") or []) if x][:8]
            if w_ch > 1700 and len(secs) >= 2:
                # chapitre long : une partie à la fois, sinon l'IA raccourcit
                parts = c.setdefault("parts", [])
                for j, sec in enumerate(secs):
                    if j < len(parts):
                        continue
                    job["step"] = f"Chapitre {k + 1}/{len(chs)} · partie {j + 1}/{len(secs)} · {sec[:50]}"
                    parts.append(EB.gen_section(chat, plan, ctx, "part", w_ch // len(secs), idx=k, prev=prev, part=j))
                c["texte"] = "\n\n".join(parts)
            else:
                job["step"] = f"Rédaction du chapitre {k + 1}/{len(chs)} · {c['titre'][:60]}"
                c["texte"] = EB.gen_section(chat, plan, ctx, "chapter", w_ch, idx=k, prev=prev)
            tick(f"Chapitre {k + 1} écrit")
        if not st.get("bonus"):
            job["step"] = "Rédaction du bonus…"
            st["bonus"] = EB.gen_section(chat, plan, ctx, "bonus", w_bonus); tick("Bonus écrit")
        bonus = st["bonus"]
        if not st.get("concl"):
            job["step"] = "Rédaction de la conclusion…"
            st["concl"] = EB.gen_section(chat, plan, ctx, "conclusion", w_concl); tick("Conclusion écrite")
        concl = st["concl"]
        uni, pal = EB.resolve_design(plan.get("design"))
        job["univers"] = EB.STYLES[uni]["label"]
        seed = random.randint(1, 10 ** 6)
        cover = os.path.join(d, "cover.jpg")
        cw, chh = EB.cover_size(uni)
        imgs_ok = 0
        credits, used = [], set()
        src_count = {"google": 0, "pexels": 0, "pixabay": 0, "web": 0, "libre": 0}
        if job["images"]:
            job["step"] = "Création de la couverture…"
            cprompt = plan.get("cover_prompt") or ctx["title"]
            if GEMINI_KEY and EB.fetch_gemini(cprompt, cw, chh, cover, GEMINI_KEY, log=print, uni=uni):
                src_count["google"] += 1
            elif EB.fetch_image(cprompt, cw, chh, seed, cover, log=print, uni=uni):
                src_count["libre"] += 1
            elif PIXABAY_KEY and (cr := EB.fetch_pixabay(chs[0].get("photo_query") or ctx["title"], cover, PIXABAY_KEY, used,
                                                         log=print, orientation="vertical")):
                credits.append(cr); src_count["pixabay"] += 1
            else:
                EB.make_art(pal, cover, cw, chh, seed)
            tick("Couverture prête")
            for k, c in enumerate(chs):
                job["step"] = f"Image du chapitre {k + 1}/{len(chs)}…"
                pth = os.path.join(d, f"ch{k + 1}.jpg")
                ok = False
                q = c.get("photo_query") or c.get("image_prompt") or c["titre"]
                ih = 896 if uni in ("sacre", "luxe") else 720
                if job.get("chapter_images") == "google" and GEMINI_KEY:
                    if EB.fetch_gemini(c.get("image_prompt") or c["titre"], 1280, ih, pth, GEMINI_KEY, log=print, uni=uni, size="1K"):
                        src_count["google"] += 1; ok = True
                order = ["web", "pixabay", "pexels"] if job.get("chapter_images") == "web" else ["pixabay", "pexels", "web"]
                for prov in order:
                    if ok:
                        break
                    cr = None
                    if prov == "web":
                        cr = EB.fetch_openverse(q, pth, used, log=print)
                    elif prov == "pixabay" and PIXABAY_KEY:
                        cr = EB.fetch_pixabay(q, pth, PIXABAY_KEY, used, log=print)
                    elif prov == "pexels" and PEXELS_KEY:
                        cr = EB.fetch_pexels(q, pth, PEXELS_KEY, used, log=print)
                    if cr:
                        credits.append(cr); src_count[{"web": "web"}.get(prov, prov)] += 1; ok = True
                if not ok and job.get("chapter_images") != "google" and \
                        EB.fetch_image(c.get("image_prompt") or c["titre"], 1280, ih, seed + k + 1, pth, log=print, uni=uni):
                    src_count["libre"] += 1; ok = True
                c["img"] = pth if ok else None  # pas de visuel de remplissage : ouverture sobre sans image
                tick(f"Image {k + 1} prête")
        else:
            EB.make_art(pal, cover, cw, chh, seed)
        imgs_ok = sum(src_count.values())
        job["img_sources"] = src_count
        job["images_ok"] = imgs_ok
        job["step"], job["pct"] = "Mise en page du PDF…", 94
        others = [o for o in ctx.get("others") or [] if o.get("nom") and o.get("nom") != ctx.get("title")]
        book = {"titre": plan.get("titre") or ctx["title"], "sous_titre": plan.get("sous_titre", ""),
                "pour_qui": plan.get("pour_qui", ""), "benefices": plan.get("ce_que_tu_vas_obtenir") or [],
                "avertissement": plan.get("avertissement", ""), "univers": uni, "colors": pal,
                "etiquette": (plan.get("design") or {}).get("etiquette") or "Guide pratique",
                "auteur": ctx.get("author") or "", "boutique": ctx.get("shop") or "", "cover_img": cover,
                "intro": intro, "chapitres": chs, "bonus_titre": (plan.get("bonus") or {}).get("titre"),
                "bonus_texte": bonus, "conclusion": concl, "autres_produits": others, "credits": credits}
        safe = re.sub(r"[^\w-]+", "_", book["titre"])[:50].strip("_") or "ebook"
        job["pdf"] = os.path.join(d, safe + ".pdf")
        job["docx"] = os.path.join(d, safe + ".docx")
        job["safe"] = safe
        EB.build_pdf(book, job["pdf"])
        EB.build_docx(book, job["docx"])
        job["cover"] = cover
        try:  # aperçu = vraie première page du PDF (avec le titre)
            import pypdfium2 as pdfium
            pdf = pdfium.PdfDocument(job["pdf"])
            prev = os.path.join(d, "preview.jpg")
            pdf[0].render(scale=1.1).to_pil().convert("RGB").save(prev, "JPEG", quality=85)
            job["cover"] = prev
            hd = os.path.join(d, "cover_hd.jpg")
            pdf[0].render(scale=2.6).to_pil().convert("RGB").save(hd, "JPEG", quality=92)
            job["cover_hd"] = hd
            pdf.close()
        except Exception as e:
            print(f"[ebook {job['id']}] aperçu impossible : {e}", flush=True)
        try:
            from pypdf import PdfReader
            job["page_count"] = len(PdfReader(job["pdf"]).pages)
        except Exception:
            job["page_count"] = 0
        job["meta"] = {"title": book["titre"], "subtitle": book["sous_titre"], "benefits": book["benefices"],
                       "toc": job["toc"], "univers": uni, "colors": pal, "label": book["etiquette"],
                       "author": book["auteur"], "pages": job.get("page_count"), "pour_qui": book["pour_qui"],
                       "audience": ctx.get("audience") or {}, "niche": ctx.get("niche"),
                       "douleurs": ctx.get("douleurs") or [], "desirs": ctx.get("desirs") or []}
        job["phase"], job["pct"], job["step"] = "done", 100, "Ebook prêt"
        print(f"[ebook {job['id']}] terminé ({GOOD_MODEL['name']}) : {job.get('page_count')} pages, "
              f"{imgs_ok} illustrations IA, {os.path.getsize(job['pdf']) // 1024} Ko", flush=True)
    except Exception as e:
        msg = str(e)
        print(f"[ebook {job['id']}] erreur : {msg[:300]}", flush=True)
        if job.get("st", {}).get("plan"):
            job["resumable"] = True
            written = sum(1 for c in job["st"]["plan"]["chapitres"] if c.get("texte"))
            msg += (f" Rien n'est perdu : {written} chapitre(s) déjà écrit(s). "
                    "Clique sur « Reprendre la création » (si c'est la limite Groq du jour, attends quelques heures).")
        job["phase"], job["error"] = "error", msg[:500]
    finally:
        job["finished"] = time.time()


@app.post("/api/ebook")
def start_ebook():
    if not authed():
        return jsonify(error="Non connecté"), 401
    if not GROQ_KEY:
        return jsonify(error="La clé Groq n'est pas configurée : la rédaction est impossible."), 400
    cleanup()
    d = request.get_json(force=True) or {}
    title = re.sub(r"\s+", " ", (d.get("title") or "")).strip()[:120]
    if not title:
        return jsonify(error="Donne un titre ou un sujet à ton ebook."), 400
    if any(j["platform"] == "ebook" and j["phase"] in ("queued", "transcribing") for j in JOBS.values()):
        return jsonify(error="Un ebook est déjà en cours de création. Attends qu'il se termine."), 409
    try:
        pages = max(5, int(d.get("pages") or 30))
    except ValueError:
        pages = 30
    keep = lambda v, n: v[:n] if isinstance(v, str) else ""
    ctx = {"title": title, "product": d.get("product") or {}, "audience": d.get("audience") or {},
           "niche": keep(d.get("niche"), 200), "douleurs": [keep(x, 200) for x in (d.get("douleurs") or [])][:6],
           "desirs": [keep(x, 200) for x in (d.get("desirs") or [])][:6], "brief": keep(d.get("brief"), 800),
           "author": keep(d.get("author"), 60), "shop": keep(d.get("shop"), 200), "keep_title": bool(d.get("keep_title", True)),
           "others": [{"nom": keep(o.get("nom"), 120), "promesse": keep(o.get("promesse"), 200)}
                      for o in (d.get("others") or []) if isinstance(o, dict)][:6]}
    job = {"id": uuid.uuid4().hex[:12], "platform": "ebook", "source": title, "url": "", "limit": 0, "sort": "",
           "lang": "", "phase": "queued", "items": [], "source_name": title, "error": "", "created": time.time(),
           "pages": pages, "images": bool(d.get("images", True)), "step": "En attente…", "pct": 0,
           "chapter_images": d.get("chapter_images") if d.get("chapter_images") in ("google", "web", "photo") else "web"}
    job["ctx"] = ctx
    with LOCK:
        JOBS[job["id"]] = job
    threading.Thread(target=run_ebook_job, args=(job, ctx), daemon=True).start()
    return jsonify(job_id=job["id"])


@app.post("/api/job/<jid>/resume")
def resume_ebook(jid):
    if not authed():
        return jsonify(error="Non connecté"), 401
    j = JOBS.get(jid)
    if not j or j.get("platform") != "ebook" or not j.get("resumable") or not j.get("ctx"):
        return jsonify(error="Cette création ne peut pas être reprise (le serveur a peut-être redémarré). Relance-la."), 404
    if any(x["platform"] == "ebook" and x["phase"] in ("queued", "transcribing") for x in JOBS.values()):
        return jsonify(error="Un ebook est déjà en cours de création. Attends qu'il se termine."), 409
    j.pop("finished", None)
    j["phase"], j["step"] = "queued", "Reprise…"
    threading.Thread(target=run_ebook_job, args=(j, j["ctx"]), daemon=True).start()
    return jsonify(job_id=jid)


@app.post("/api/ebook/titles")
def ebook_titles():
    if not authed():
        return jsonify(error="Non connecté"), 401
    if not GROQ_KEY:
        return jsonify(error="La clé Groq n'est pas configurée."), 400
    d = request.get_json(force=True) or {}
    title = re.sub(r"\s+", " ", (d.get("title") or "")).strip()[:120]
    if not title:
        return jsonify(error="Écris d'abord un titre ou un sujet."), 400
    ctx = {"title": title, "product": d.get("product") or {}, "audience": d.get("audience") or {},
           "niche": (d.get("niche") or "")[:200] if isinstance(d.get("niche"), str) else "",
           "douleurs": [str(x)[:200] for x in (d.get("douleurs") or [])][:6],
           "desirs": [str(x)[:200] for x in (d.get("desirs") or [])][:6]}
    try:
        raw = EB.gen_titles(lambda m, **k: groq_chat(m, tag="titres", **k), ctx)
        titles = [re.sub(r"\s+", " ", str(t)).strip(" «»\"")[:90] for t in (parse_json_loose(raw).get("titres") or [])]
        titles = [t for t in titles if t][:6]
        if not titles:
            raise ValueError("vide")
        return jsonify(titles=titles)
    except Exception as e:
        print(f"[titres] erreur : {str(e)[:160]}", flush=True)
        return jsonify(error="L'IA n'a pas pu proposer de titres. Réessaie dans un instant."), 502


@app.get("/api/job/<jid>/ebook")
def job_ebook(jid):
    if not authed():
        return jsonify(error="Non connecté"), 401
    j = JOBS.get(jid)
    if not j or j.get("phase") != "done":
        return jsonify(error="Ebook introuvable"), 404
    return jsonify(title=j.get("book_title"), toc=j.get("toc", []), pages=j.get("page_count", 0),
                   images_ok=j.get("images_ok", 0), images=j.get("images"), univers=j.get("univers", ""),
                   img_sources=j.get("img_sources", {}))


@app.get("/api/job/<jid>/ebook/<kind>")
def job_ebook_file(jid, kind):
    if not authed():
        return "Non connecté", 401
    j = JOBS.get(jid)
    key = {"pdf": "pdf", "docx": "docx", "cover": "cover"}.get(kind)
    if not j or not key or not j.get(key) or not os.path.exists(j[key]):
        return "Fichier introuvable (le serveur a peut-être redémarré). Relance la création.", 404
    if kind == "cover":
        return send_file(j[key], mimetype="image/jpeg", max_age=3600)
    return send_file(j[key], as_attachment=True, download_name=f"{j['safe']}.{kind}")


# ------------------------------------------------ kit de vente (visuels + description)
import kit as KIT


def gen_sales_copy(meta, about=""):
    a = meta.get("audience") or {}
    facts = [f"Titre : {meta.get('title')}", f"Sous-titre : {meta.get('subtitle') or ''}",
             f"Pour qui : {meta.get('pour_qui') or a.get('profil') or ''}",
             f"Niche : {meta.get('niche') or ''}",
             f"Nombre de pages : {meta.get('pages') or 'non précisé'}",
             "Bénéfices : " + " ; ".join(meta.get("benefits") or []),
             "Chapitres : " + " | ".join(meta.get("toc") or []),
             "Douleurs de l'audience : " + " ; ".join(meta.get("douleurs") or []),
             "Désirs de l'audience : " + " ; ".join(meta.get("desirs") or [])]
    if meta.get("price"):
        facts.append(f"Prix : {meta['price']}")
    if about:
        facts.append(f"Contenu du livre (extrait ou résumé) : {about[:4000]}")
    facts = [f for f in facts if f.split(":", 1)[1].strip()]
    system = ("Tu es un copywriter francophone spécialiste de la vente de produits digitaux en Afrique francophone "
              "et dans la diaspora (boutiques en ligne type Chariow). Tu écris des descriptions qui donnent envie "
              "d'acheter tout en restant honnêtes : pas de faux témoignages, pas de chiffres inventés, pas de "
              "promesse de résultat garanti, pas de fausse urgence. Tutoiement, phrases courtes, ton chaleureux et sûr.")
    user = ("Écris la page de vente de cet ebook.\n" + "\n".join(facts) + "\n\n"
            "Réponds UNIQUEMENT en JSON : {\"titre_produit\": \"titre de la fiche produit (max 80 caractères)\", "
            "\"accroche\": \"1 phrase choc pour les réseaux sociaux\", "
            "\"points_forts\": [\"3 bénéfices très courts (max 42 caractères chacun) pour le visuel\"], "
            "\"description\": \"la description complète, 180 à 260 mots, en texte brut prêt à coller : une accroche qui parle "
            "du problème du lecteur, ce que le livre va changer pour lui, une liste « Dans ce guide, tu vas découvrir : » "
            "de 5 ou 6 points commençant par ✅, pour qui il est fait, ce qu'il reçoit (format PDF, nombre de pages si connu), "
            "et un appel à l'action final. Sauts de ligne entre les parties. 4 émojis maximum en dehors des ✅.\"}")
    raw = groq_chat([{"role": "system", "content": system}, {"role": "user", "content": user}],
                    max_tokens=2500, json_mode=True, tag="kit")
    out = parse_json_loose(raw)
    out["points_forts"] = [p.strip()[:48] for p in (out.get("points_forts") or []) if isinstance(p, str) and p.strip()][:3]
    out["description"] = (out.get("description") or "").replace("\\n", "\n").strip()
    return out


def run_kit_job(job, meta, cover_src, about):
    print(f"[kit {job['id']}] « {meta.get('title')} »", flush=True)
    d = os.path.join(BASE, "k_" + job["id"])
    os.makedirs(d, exist_ok=True)
    job["dir"] = d
    try:
        job["phase"] = "transcribing"
        job["step"], job["pct"] = "Préparation de la couverture…", 8
        cover = os.path.join(d, "cover.jpg")
        KIT.cover_from_upload(cover_src, cover)
        pal = meta.get("colors") or KIT.palette_from_image(cover)
        uni = meta.get("univers") or "moderne"
        copy = {}
        if GROQ_KEY:
            job["step"], job["pct"] = "Écriture de la description de vente…", 25
            try:
                copy = gen_sales_copy(meta, about)
            except Exception as e:
                print(f"[kit {job['id']}] description impossible : {str(e)[:160]}", flush=True)
                job["copy_error"] = "La description n'a pas pu être écrite (limite Groq ?). Relance dans quelques minutes."
        points = [p for p in (meta.get("points") or []) if p][:3] or copy.get("points_forts") or \
            [b[:48] for b in (meta.get("benefits") or [])][:3]
        safe = re.sub(r"[^\w-]+", "_", meta.get("title") or "ebook")[:40].strip("_") or "ebook"
        job["step"], job["pct"] = "Création du visuel 16:9…", 70
        wide = os.path.join(d, f"{safe}_16x9.jpg")
        KIT.wide_visual(cover, wide, {"title": meta.get("title") or "", "subtitle": meta.get("subtitle") or "",
                                      "points": points, "price": meta.get("price") or "", "label": meta.get("label") or "Ebook PDF",
                                      "author": meta.get("author") or ""}, pal, uni)
        job["step"], job["pct"] = "Création du visuel carré…", 88
        square = os.path.join(d, f"{safe}_carre.jpg")
        KIT.square_visual(cover, square, pal, uni)
        zp = os.path.join(d, f"{safe}_kit.zip")
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_STORED) as z:
            z.write(wide, os.path.basename(wide)); z.write(square, os.path.basename(square))
            if copy.get("description"):
                z.writestr(f"{safe}_description.txt", (copy.get("titre_produit", "") + "\n\n" + copy["description"]
                                                       + ("\n\n" + copy["accroche"] if copy.get("accroche") else "")).encode("utf-8"))
        job.update(kit={"wide": wide, "square": square, "zip": zp}, copy=copy)
        job["phase"], job["pct"], job["step"] = "done", 100, "Kit prêt"
        print(f"[kit {job['id']}] terminé (description : {'oui' if copy.get('description') else 'non'})", flush=True)
    except Exception as e:
        msg = str(e)
        print(f"[kit {job['id']}] erreur : {msg[:200]}", flush=True)
        job["phase"], job["error"] = "error", msg[:300]
    finally:
        job["finished"] = time.time()


@app.post("/api/kit")
def start_kit():
    if not authed():
        return jsonify(error="Non connecté"), 401
    cleanup()
    form = request.form if request.files or request.form else (request.get_json(force=True, silent=True) or {})
    price = re.sub(r"\s+", " ", (form.get("price") or "")).strip()[:30]
    label = re.sub(r"\s+", " ", (form.get("label") or "")).strip()[:28]
    points = [p.strip()[:48] for p in re.split(r"[\n;]+", form.get("points") or "") if p.strip()][:3]
    about = ""
    eid = form.get("ebook_job")
    if eid:
        ej = JOBS.get(eid)
        if not ej or ej.get("phase") != "done" or not os.path.exists(ej.get("cover_hd", "")):
            return jsonify(error="Cet ebook n'est plus sur le serveur (il a peut-être redémarré). Importe sa couverture à la place."), 404
        meta = dict(ej["meta"])
        cover_src = ej["cover_hd"]
    else:
        f = request.files.get("file")
        if not f or not f.filename:
            return jsonify(error="Importe la couverture de ton ebook (image ou PDF)."), 400
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in (".jpg", ".jpeg", ".png", ".webp", ".pdf", ".heic"):
            return jsonify(error="Format non pris en charge : utilise une image (JPG, PNG) ou le PDF de l'ebook."), 400
        title = re.sub(r"\s+", " ", (form.get("title") or "")).strip()[:120]
        if not title:
            return jsonify(error="Indique le titre de l'ebook."), 400
        up = tempfile.mkdtemp(dir=BASE)
        cover_src = os.path.join(up, "upload" + ext)
        f.save(cover_src)
        if ext == ".heic":
            cover_src = to_jpeg(cover_src)
        about = (form.get("about") or "").strip()[:3000]
        if ext == ".pdf":
            about = (about + " " + KIT.pdf_text(cover_src)).strip()
        meta = {"title": title, "subtitle": (form.get("subtitle") or "").strip()[:160],
                "author": (form.get("author") or "").strip()[:60]}
    meta.update({k: v for k, v in (("price", price), ("label", label), ("points", points)) if v})
    if not meta.get("label"):
        meta["label"] = "Ebook PDF"
    job = {"id": uuid.uuid4().hex[:12], "platform": "kit", "source": meta["title"], "url": "", "limit": 0, "sort": "",
           "lang": "", "phase": "queued", "items": [], "source_name": meta["title"], "error": "", "created": time.time(),
           "step": "En attente…", "pct": 0}
    with LOCK:
        JOBS[job["id"]] = job
    threading.Thread(target=run_kit_job, args=(job, meta, cover_src, about), daemon=True).start()
    return jsonify(job_id=job["id"])


@app.get("/api/job/<jid>/kit")
def job_kit(jid):
    if not authed():
        return jsonify(error="Non connecté"), 401
    j = JOBS.get(jid)
    if not j or j.get("phase") != "done":
        return jsonify(error="Kit introuvable"), 404
    return jsonify(title=j["source_name"], copy=j.get("copy") or {}, copy_error=j.get("copy_error", ""))


@app.get("/api/job/<jid>/kit/<kind>")
def job_kit_file(jid, kind):
    if not authed():
        return "Non connecté", 401
    j = JOBS.get(jid)
    k = (j or {}).get("kit") or {}
    path = k.get(kind)
    if not path or not os.path.exists(path):
        return "Fichier introuvable (le serveur a peut-être redémarré). Relance.", 404
    return send_file(path, as_attachment=request.args.get("dl") == "1", download_name=os.path.basename(path), max_age=600)


@app.get("/api/job/<jid>/report")
def job_report(jid):
    if not authed():
        return jsonify(error="Non connecté"), 401
    j = JOBS.get(jid)
    if not j or j.get("phase") != "done":
        return jsonify(error="Rapport introuvable"), 404
    return jsonify(report=j["report"], stats=j.get("stats", {}), name=j["source_name"], platform=j.get("src_platform"))


@app.get("/api/job/<jid>/full")
def job_full(jid):
    if not authed():
        return jsonify(error="Non connecté"), 401
    j = JOBS.get(jid)
    if not j:
        return jsonify(error="Tâche introuvable"), 404
    it = j["items"][0] if j["items"] else {}
    return jsonify(text=it.get("text", ""))


@app.get("/api/cookies")
def cookies_status():
    if not authed():
        return jsonify(error="Non connecté"), 401
    return jsonify(status=YT["status"], source=YT["source"], updated=YT["updated"])


@app.post("/api/cookies")
def cookies_upload():
    if not authed():
        return jsonify(error="Non connecté"), 401
    f = request.files.get("file")
    text = f.read().decode("utf-8", "ignore") if f else (request.form.get("text") or "")
    lines = [l for l in text.splitlines() if l.strip() and not l.startswith("#")]
    domains = sorted({l.split("\t")[0] for l in lines if "\t" in l})[:8]
    print(f"Dépôt cookies : {len(text)} caractères, {len(lines)} lignes, domaines {domains}, "
          f"1re ligne : {text[:40]!r}", flush=True)
    if text.lstrip().startswith(("[", "{")):
        return jsonify(error="Ce fichier est au format JSON. Dans l'extension, choisis le format « Netscape » puis refais Export."), 400
    if lines and not any("youtube.com" in l for l in lines):
        return jsonify(error="Ce fichier ne contient pas de cookies YouTube : fais l'export en étant sur youtube.com."), 400
    if not valid_yt_cookies(text):
        return jsonify(error="Ce fichier ne contient pas de cookies YouTube. Exporte-le depuis youtube.com avec l'extension."), 400
    install_yt_cookies(text, "upload")
    return jsonify(status="ok")


# ------------------------------------------------ clipping : moments viraux d'une longue vidéo
# durée visée (min, max) : c'est une cible, le clip va toujours jusqu'à la fin de la phrase
CLIP_LENGTHS = {"auto": (20, 90), "court": (15, 35), "moyen": (30, 60), "long": (55, 90)}
AUTO_MIN_SCORE, AUTO_MAX = 75, 15
CLIP_W, CLIP_H = 720, 1280
FONTS_DIR = os.path.join(app.root_path, "fonts")
FF_THREADS = os.environ.get("FFMPEG_THREADS", "2")
WINDOW, STEP = 600, 510  # analyse par tranches de 10 min qui se chevauchent


def fmt_ts(t):
    t = int(max(0, t))
    h, m, s = t // 3600, t % 3600 // 60, t % 60
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"


def clean_words(words):
    """[début, fin, mot] triés, fins bornées par le mot suivant, sans [Musique] & co."""
    words = [w for w in words if w[2] and not re.fullmatch(r"[\[(].*[\])]|>>|-", w[2])]
    words.sort(key=lambda w: w[0])
    for i, w in enumerate(words):
        nxt = words[i + 1][0] if i + 1 < len(words) else w[0] + 1.0
        end = w[1] if w[1] and w[1] > w[0] else w[0] + 0.6
        w[1] = max(w[0] + 0.05, min(end, nxt, w[0] + 1.5))
    return words


def words_from_json3(raw):
    data = json.loads(raw)
    words = []
    for ev in data.get("events", []):
        segs = ev.get("segs") or []
        t0 = (ev.get("tStartMs") or 0) / 1000
        dur = (ev.get("dDurationMs") or 0) / 1000
        for k, s in enumerate(segs):
            ws = (s.get("utf8") or "").split()
            if not ws:
                continue
            a = t0 + (s.get("tOffsetMs") or 0) / 1000
            b = t0 + (segs[k + 1].get("tOffsetMs") or 0) / 1000 if k + 1 < len(segs) else t0 + (dur or 2.0)
            if b <= a:
                b = a + 0.35 * len(ws)
            step = (b - a) / len(ws)
            words += [[a + i * step, a + (i + 1) * step, w] for i, w in enumerate(ws)]
    return clean_words(words)


def groq_words(path, lang):
    """Transcription Whisper avec l'heure de chaque mot."""
    data = {"model": GROQ_MODEL, "response_format": "verbose_json", "temperature": "0",
            "timestamp_granularities[]": ["word", "segment"]}
    if lang:
        data["language"] = lang
    for attempt in range(6):
        with open(path, "rb") as f:
            r = requests.post("https://api.groq.com/openai/v1/audio/transcriptions",
                              headers={"Authorization": f"Bearer {GROQ_KEY}"},
                              data=data, files={"file": (os.path.basename(path), f)}, timeout=300)
        if r.status_code == 200:
            j = r.json()
            ws = [[w["start"], w["end"], (w.get("word") or "").strip()] for w in j.get("words") or []]
            if not ws:  # pas de mots : on répartit le texte de chaque segment
                for sg in j.get("segments") or []:
                    t = (sg.get("text") or "").split()
                    if t:
                        st = (sg["end"] - sg["start"]) / len(t)
                        ws += [[sg["start"] + i * st, sg["start"] + (i + 1) * st, x] for i, x in enumerate(t)]
            return ws
        if r.status_code in (429, 500, 502, 503):
            try:
                wait = float(r.headers.get("retry-after"))
            except (TypeError, ValueError):
                wait = 15 * (attempt + 1)
            time.sleep(min(max(wait, 3), 120))
            continue
        if r.status_code == 401:
            raise RuntimeError("Clé Groq invalide.")
        raise RuntimeError(f"Groq a répondu {r.status_code} : {r.text[:150]}")
    raise RuntimeError("Limite Groq atteinte, réessaie plus tard.")


def sentence_ends(words):
    """ends[k] = vrai si on peut couper juste après le mot k sans couper une phrase :
    ponctuation de fin, ou vraie pause dans la parole (les sous-titres auto n'ont pas de points)."""
    punct = sum(bool(re.search(r"[.!?…]$", w[2])) for w in words) > len(words) / 60
    pause = 0.8 if punct else 0.45
    ends = []
    for k, w in enumerate(words):
        gap = words[k + 1][0] - w[1] if k + 1 < len(words) else 99
        ends.append(gap >= pause or (punct and bool(re.search(r"[.!?…]$", w[2])) and gap >= 0.05))
    return ends


def build_lines(words, ends):
    """Lignes horodatées qui finissent toujours sur une fin de phrase ou une pause."""
    lines, i = [], 0
    for k, w in enumerate(words):
        dur = w[1] - words[i][0]
        gap = words[k + 1][0] - w[1] if k + 1 < len(words) else 99
        if (ends[k] and dur >= 2.5) or (dur >= 14 and gap >= 0.25) or dur >= 20 or k == len(words) - 1:
            lines.append({"s": words[i][0], "e": w[1], "wi": i, "wj": k,
                          "text": " ".join(x[2] for x in words[i:k + 1])})
            i = k + 1
    return lines


def clip_transcript(job, info, tmp):
    """Mots horodatés : sous-titres YouTube d'abord, sinon transcription IA."""
    if job["site"] == "youtube":
        cap = pick_caption(info, job["lang"])
        if cap:
            url = cap[0] if cap[1] == "json3" else re.sub(r"([?&]fmt=)[^&]*", r"\1json3", cap[0])
            if "fmt=json3" not in url:
                url += "&fmt=json3"
            try:
                with yt_dlp.YoutubeDL(ydl_opts("youtube")) as y:
                    raw = y.urlopen(url).read().decode("utf-8", "ignore")
                words = words_from_json3(raw)
                if len(words) > 30:
                    job["method"] = "sous-titres"
                    return words
            except Exception as e:
                print(f"[clip {job['id']}] sous-titres illisibles : {str(e)[:120]}", flush=True)
    if not GROQ_KEY:
        raise RuntimeError("Cette vidéo n'a pas de sous-titres et la clé Groq n'est pas configurée.")
    job["msg"] = "Téléchargement du son…"
    opts = ydl_opts(job["site"], format="bestaudio/worst[ext=mp4]/worst/best",
                    outtmpl=os.path.join(tmp, "media.%(ext)s"))
    with yt_dlp.YoutubeDL(opts) as y:
        y.extract_info(job["url"], download=True)
    media = [p for p in glob.glob(os.path.join(tmp, "media.*")) if not p.endswith(".part")]
    if not media:
        raise RuntimeError("Téléchargement du son impossible.")
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", media[0], "-vn", "-ac", "1", "-ar", "16000",
                    "-b:a", "32k", "-f", "segment", "-segment_time", "600",
                    os.path.join(tmp, "part_%03d.mp3")], check=True)
    os.remove(media[0])
    parts = sorted(glob.glob(os.path.join(tmp, "part_*.mp3")))
    if not parts:
        raise RuntimeError("Cette vidéo n'a pas de son.")
    words, offset = [], 0.0
    for i, p in enumerate(parts, 1):
        job["msg"] = f"Transcription IA · partie {i}/{len(parts)}"
        words += [[a + offset, b + offset, w] for a, b, w in groq_words(p, job["lang"])]
        offset += probe_duration(p) or 600
        os.remove(p)
    job["method"] = "IA"
    return clean_words(words)


def parse_json_obj(txt):
    m = re.search(r"\{.*\}", txt or "", re.S)
    if not m:
        return {}
    try:
        return json.loads(m.group(0))
    except Exception:
        try:
            return json.loads(re.sub(r",\s*([\]}])", r"\1", m.group(0)))
        except Exception:
            return {}


def clip_prompt(job):
    mn, mx = CLIP_LENGTHS[job["length"]]
    if job["length"] == "auto":
        duree = (f"- dure le temps qu'il faut pour que le moment soit complet et percutant : choisis toi-même la durée "
                 f"idéale, entre {mn} et {mx} secondes (calcule avec les heures des lignes) ;\n")
    else:
        duree = (f"- dure environ {mn} à {mx} secondes (calcule avec les heures des lignes). C'est une cible, pas une limite : "
                 "le clip doit TOUJOURS aller jusqu'au bout de la phrase et de l'idée en cours. Mieux vaut dépasser de "
                 "15 ou 20 secondes que couper quelqu'un au milieu de ce qu'il dit ;\n")
    combien = ("Propose tous les passages vraiment forts de cet extrait, au maximum 5"
               if job["count"] == "auto" else "Maximum 4 clips dans cet extrait")
    return (
        "Tu es un monteur expert en clips viraux pour TikTok, Reels et YouTube Shorts. On te donne un extrait "
        "horodaté de la transcription d'une longue vidéo (podcast, conférence, interview, prêche, débat…). "
        "Chaque ligne commence par son numéro, puis l'heure où elle est dite. Chaque ligne se termine sur une "
        "pause ou une fin de phrase.\n\n"
        "Ta mission : repérer les passages qui feraient d'excellents clips courts, capables de devenir viraux.\n"
        "Un bon clip :\n"
        "- commence au début d'une phrase forte qui accroche dès les 3 premières secondes (question, affirmation choc, "
        "début d'histoire, vérité qui dérange), jamais par « euh », « donc », « et », une fin de phrase ou une salutation ;\n"
        "- parle d'un vrai sujet de fond : une idée forte, une leçon, une histoire, un conseil concret, une révélation, "
        "une émotion forte, une punchline, un désaccord ;\n"
        "- se comprend tout seul, sans avoir vu le reste de la vidéo ;\n"
        "- se termine quand la personne a fini son message : sur une conclusion, une chute ou une phrase marquante. "
        "Jamais au milieu d'une phrase, d'une explication ou d'une histoire ;\n"
        + duree +
        "À éviter absolument : introductions, présentation des invités, remerciements, appels à s'abonner, "
        "publicités et partenariats, logistique, passages flous, bavardage sans fond.\n\n"
        "Note chaque clip de 0 à 100 selon son potentiel viral (90 et plus : exceptionnel, 75-89 : très bon, "
        "60-74 : correct). Sois exigeant et honnête. Ne propose rien en dessous de 60. Mieux vaut 0 ou 1 clip "
        f"excellent que 4 moyens. {combien}, qui ne se chevauchent pas.\n\n"
        "Réponds UNIQUEMENT avec un objet JSON, sans aucun texte autour, au format :\n"
        '{"clips":[{"debut":12,"fin":19,"score":87,"titre":"...","accroche":"...","pourquoi":"...","type":"..."}]}\n'
        "- debut : numéro de la ligne où le clip commence ; fin : numéro de la ligne où la personne termine son message ;\n"
        "- titre : titre court et percutant en français pour publier le clip (70 caractères maximum, sans hashtag) ;\n"
        "- accroche : la première phrase du clip, telle qu'elle est dite ;\n"
        "- pourquoi : en une phrase en français, pourquoi ce passage peut devenir viral ;\n"
        "- type : un seul mot parmi histoire, conseil, révélation, émotion, punchline, débat, humour, motivation, foi.\n"
        'S\'il n\'y a aucun bon passage, réponds {"clips":[]}.')


def fit_moment(words, ends, wi, wj, mn, mx):
    """Cale le début au début d'une phrase et la fin à la fin du message en cours.
    La durée choisie est une cible : on ne coupe jamais quelqu'un en pleine phrase pour la respecter."""
    n = len(words)
    cap = max(mx * 1.5, mx + 25)

    def dur(a, b):
        return words[b][1] - words[a][0]

    # début : remonter au début de la phrase si l'IA a démarré en plein milieu
    k = wi
    while k > 0 and not ends[k - 1] and words[wi][0] - words[k - 1][0] < 6:
        k -= 1
    if k == 0 or ends[k - 1]:
        wi = k
    # fin : aller jusqu'au bout de la phrase en cours
    k = wj
    while not ends[k] and k + 1 < n and words[k + 1][1] - words[wj][1] < 15:
        k += 1
    wj = k
    # beaucoup trop long : revenir à une fin de phrase plus tôt (jamais en dessous de la durée minimale)
    if dur(wi, wj) > cap:
        ok = [k for k in range(wi, wj) if ends[k] and mn <= dur(wi, k) <= cap]
        if ok:
            wj = ok[-1]
    # trop court : ajouter la phrase suivante tant que la personne continue de parler
    while dur(wi, wj) < mn and wj + 1 < n and words[wj + 1][0] - words[wj][1] < 2.5:
        k = wj + 1
        while not ends[k] and k + 1 < n:
            k += 1
        if dur(wi, k) > mx + 10:
            break
        wj = k
    gap_b = words[wi][0] - words[wi - 1][1] if wi > 0 else 1.0
    gap_a = words[wj + 1][0] - words[wj][1] if wj + 1 < n else 1.0
    s = max(0.0, words[wi][0] - min(0.3, max(0.05, gap_b * 0.6)))
    e = words[wj][1] + min(0.7, max(0.2, gap_a * 0.6))
    return wi, wj, s, e


def find_moments(job, words, ends, lines):
    mn, mx = CLIP_LENGTHS[job["length"]]
    total = lines[-1]["e"]
    windows, ws = [], 0.0
    while ws < total:
        idx = [i for i, l in enumerate(lines) if ws <= l["s"] < ws + WINDOW]
        if idx:
            windows.append(idx)
        if ws + WINDOW >= total:
            break
        ws += STEP
    system = clip_prompt(job)
    found = []
    job["parts"] = len(windows)
    for k, idx in enumerate(windows, 1):
        job["part"], job["msg"] = k - 1, f"Analyse IA · partie {k}/{len(windows)}"
        body = "\n".join(f"{i} [{fmt_ts(lines[i]['s'])}] {lines[i]['text']}" for i in idx)
        try:
            out = groq_chat([{"role": "system", "content": system},
                             {"role": "user", "content": f"Vidéo : {job['title']}\n\n{body}"}], max_tokens=4000)
        except RuntimeError as e:
            if "invalide" in str(e):
                raise
            print(f"[clip {job['id']}] analyse {k} ratée : {e}", flush=True)
            continue
        for c in (parse_json_obj(out).get("clips") or [])[:6]:
            try:
                ld, lf, score = int(c.get("debut")), int(c.get("fin")), int(float(c.get("score") or 0))
            except (TypeError, ValueError):
                continue
            if not (idx[0] <= ld <= idx[-1]) or score < 55:
                continue
            lf = max(ld, min(lf, len(lines) - 1))
            wi, wj, s, e = fit_moment(words, ends, lines[ld]["wi"], lines[lf]["wj"], mn, mx)
            if e - s < max(8, mn * 0.5):
                continue
            found.append({"s": round(s, 2), "e": round(e, 2), "score": min(score, 100),
                          "titre": str(c.get("titre") or "").strip()[:90],
                          "accroche": str(c.get("accroche") or lines[ld]["text"]).strip()[:220],
                          "pourquoi": str(c.get("pourquoi") or "").strip()[:260],
                          "type": str(c.get("type") or "").strip().lower()[:20]})
        job["part"] = k
    found.sort(key=lambda m: m["score"], reverse=True)
    kept = []
    for m in found:  # pas de doublons entre tranches qui se chevauchent
        if all(min(m["e"], o["e"]) - max(m["s"], o["s"]) < 0.3 * min(m["e"] - m["s"], o["e"] - o["s"]) for o in kept):
            kept.append(m)
    if job["count"] == "auto":
        best = [m for m in kept if m["score"] >= AUTO_MIN_SCORE][:AUTO_MAX]
        # vidéo faible : on garde quand même les 2 ou 3 meilleurs passages corrects
        return best if len(best) >= 2 else [m for m in kept if m["score"] >= 65][:3] or kept[:1]
    return kept[:job["count"]]


def ass_text(t):
    return t.replace("\\", "").replace("{", "(").replace("}", ")")


def clip_ass(words, m, job, path):
    """Sous-titres animés mot à mot + titre en haut (format ASS)."""
    blur = job["style"] == "blur"
    sub_mv, title_mv = (300, 250) if blur else (330, 170)
    head = (f"[Script Info]\nScriptType: v4.00+\nPlayResX: {CLIP_W}\nPlayResY: {CLIP_H}\nWrapStyle: 0\n"
            "ScaledBorderAndShadow: yes\n\n[V4+ Styles]\n"
            "Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, "
            "Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, "
            "MarginR, MarginV, Encoding\n"
            f"Style: Sub,Anton,68,&H00FFFFFF,&H00FFFFFF,&H00000000,&H78000000,0,0,0,0,100,100,1,0,1,5,2,2,40,40,{sub_mv},1\n"
            f"Style: Title,Anton,58,&H0000E5FF,&H0000E5FF,&H00000000,&H78000000,0,0,0,0,100,100,0.5,0,1,4,2,8,50,50,{title_mv},1\n\n"
            "[Events]\nFormat: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text\n")

    def ts(t):
        t = max(0.0, t)
        return f"{int(t // 3600)}:{int(t % 3600 // 60):02d}:{t % 60:05.2f}"

    dur = m["e"] - m["s"]
    out = []
    if job["show_title"] and m["titre"]:
        out.append(f"Dialogue: 1,{ts(0)},{ts(dur)},Title,,0,0,0,,{{\\fad(250,0)}}{ass_text(m['titre'].upper())}")
    if job["subs"]:
        ws = [[a - m["s"], b - m["s"], w] for a, b, w in words if m["s"] - 0.05 <= a < m["e"]]
        chunks, cur = [], []
        for w in ws:
            if cur and (len(cur) >= 3 or w[0] - cur[-1][1] > 0.5 or len(" ".join(x[2] for x in cur + [w])) > 18):
                chunks.append(cur)
                cur = []
            cur.append(w)
            if re.search(r"[.!?,;:…]$", w[2]):
                chunks.append(cur)
                cur = []
        if cur:
            chunks.append(cur)
        for i, c in enumerate(chunks):
            a = c[0][0]
            b = min(chunks[i + 1][0][0] if i + 1 < len(chunks) else dur, c[-1][1] + 0.4, dur)
            if b - a < 0.08:
                continue
            txt = re.sub(r"[.,;:]+$", "", " ".join(x[2] for x in c)).upper()
            out.append(f"Dialogue: 0,{ts(a)},{ts(b)},Sub,,0,0,0,,"
                       f"{{\\fscx86\\fscy86\\t(0,90,\\fscx100\\fscy100)}}{ass_text(txt)}")
    with open(path, "w", encoding="utf-8") as f:
        f.write(head + "\n".join(out) + "\n")


def clip_sources(job):
    h = 720 if job["style"] == "blur" else 1080  # recadrage : source en 1080p pour garder une image nette
    fmt = (f"bv*[height<={h}][vcodec^=avc1]+ba[ext=m4a]/bv*[height<={h}][vcodec^=avc1]+ba/"
           f"bv*[height<={h}]+ba/b[height<={h}]/b")
    with yt_dlp.YoutubeDL(ydl_opts(job["site"], format=fmt, skip_download=True)) as y:
        info = y.extract_info(job["url"], download=False)
    fs = info.get("requested_formats") or [info]
    v = next((f for f in fs if f.get("vcodec") not in (None, "none")), fs[0])
    srcs = [(f["url"], f.get("http_headers") or {}) for f in fs if f.get("url")]
    return srcs, int(v.get("width") or info.get("width") or 0), int(v.get("height") or info.get("height") or 0)


def ff_input(u, hd, s, d, extra=()):
    a = []
    if u.startswith("http"):
        a += ["-reconnect", "1", "-reconnect_streamed", "1", "-reconnect_delay_max", "5"]
        if hd:
            a += ["-headers", "".join(f"{k}: {v}\r\n" for k, v in hd.items())]
    return a + list(extra) + ["-ss", f"{s:.3f}", "-t", f"{d:.3f}", "-i", u]


def sample_frames(src, s, d, keyonly, sw=480):
    """Petites images couleur du passage, avec leur heure."""
    import numpy as np
    extra = ["-skip_frame", "nokey"] if keyonly else []
    vf = ("" if keyonly else "fps=2,") + f"scale={sw}:-2,showinfo"
    cmd = (["ffmpeg", "-nostdin", "-hide_banner", "-loglevel", "info"] + ff_input(src[0], src[1], s, d, extra)
           + ["-an", "-vf", vf, "-fps_mode", "passthrough", "-f", "rawvideo", "-pix_fmt", "bgr24", "pipe:1"])
    r = subprocess.run(cmd, capture_output=True, timeout=1200)
    times = [float(x) for x in re.findall(rb"pts_time:\s*([-\d.]+)", r.stderr)]
    if not times or not r.stdout:
        return []
    sh = len(r.stdout) // (sw * 3 * len(times))
    if sh < 10:
        return []
    arr = np.frombuffer(r.stdout[:sw * sh * 3 * len(times)], dtype=np.uint8).reshape(len(times), sh, sw, 3)
    return list(zip(times, arr))


FACE = {}
FACE_MODEL = os.path.join(app.root_path, "models", "face_yunet.onnx")


def detect_faces(img):
    """Visages (x, y, l, h) dans une petite image couleur, détecteur YuNet (repère aussi les profils)."""
    import cv2
    h, w = img.shape[:2]
    if "yn" not in FACE:
        FACE["yn"] = cv2.FaceDetectorYN.create(FACE_MODEL, "", (w, h), 0.6, 0.3, 50)
    FACE["yn"].setInputSize((w, h))
    _, res = FACE["yn"].detect(img)
    if res is None:
        return []
    return [(int(r[0]), int(r[1]), int(r[2]), int(r[3])) for r in res if r[2] * r[3] >= (h / 18) ** 2]


def face_track(job, m, src, vw, vh):
    """Positions du cadre 9:16 au fil du clip : il suit le visage le plus visible,
    saute d'un coup aux changements de plan et glisse en douceur quand la personne bouge."""
    import numpy as np
    import cv2
    cw = int(vh * 9 / 16) // 2 * 2
    maxx = vw - cw
    center = maxx / 2
    d = m["e"] - m["s"]
    frames = sample_frames(src, m["s"], d, keyonly=True)
    if len(frames) < max(4, d / 4):  # images clés trop rares : on échantillonne 2 images par seconde
        frames = sample_frames(src, m["s"], d, keyonly=False)
    if not frames:
        return cw, [(0.0, center, 0.0)], 0
    sw = frames[0][1].shape[1]
    k = vw / sw
    keys, cur, pending, prev, prevg, seen = [], None, None, None, None, 0
    for t, img in frames:
        g = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        hist = cv2.calcHist([g], [0], None, [32], [0, 256])
        cv2.normalize(hist, hist)
        # changement de plan = toute l'image change (pas seulement la personne qui bouge)
        cut = prev is not None and (cv2.compareHist(prev, hist, cv2.HISTCMP_CORREL) < 0.7 or
                                    float(np.mean(np.abs(g.astype(np.int16) - prevg.astype(np.int16)))) > 60)
        prev, prevg = hist, g
        faces = detect_faces(img)
        if not faces:
            continue
        seen += 1
        big = max(w * h for _, _, w, h in faces)
        cand = [f for f in faces if f[2] * f[3] >= 0.6 * big]
        if cur is not None and not cut:  # deux visages aussi visibles : on reste sur celui qu'on filme
            pick = min(cand, key=lambda f: abs((f[0] + f[2] / 2) * k - (cur + cw / 2)))
        else:
            pick = max(cand, key=lambda f: f[2] * f[3])
        x = min(max((pick[0] + pick[2] / 2) * k - cw / 2, 0), maxx)
        t = max(0.0, t)
        if cur is None:
            cur = x
            keys.append((0.0, x, 0.0))
            continue
        if abs(x - cur) < 0.06 * vw:
            pending = None
            continue
        if cut:
            keys.append((t, x, 0.0))
            cur, pending = x, None
        elif pending and (x - cur) * (pending[1] - cur) > 0:  # la personne s'est bien déplacée : on glisse
            keys.append((pending[0], x, min(max(0.6, t - pending[0]), 2.5)))
            cur, pending = x, None
        else:
            pending = (t, x)
    if not keys:
        keys = [(0.0, center, 0.0)]
    clean = []
    for kk in sorted(keys):
        if clean and kk[0] - clean[-1][0] < 0.7:
            clean[-1] = (clean[-1][0], kk[1], clean[-1][2])
        else:
            clean.append(kk)
    return cw, clean[:60], seen


def x_expr(keys):
    """Expression ffmpeg de la position x du cadre en fonction du temps t."""
    def seg(i):
        t, x, p = keys[i]
        if p <= 0 or i == 0:
            return f"{x:.1f}"
        x0 = keys[i - 1][1]
        return f"if(lt(t,{t + p:.3f}),{x0:.1f}+({x - x0:.1f})*(t-{t:.3f})/{p:.3f},{x:.1f})"
    expr = seg(len(keys) - 1)
    for i in range(len(keys) - 2, -1, -1):
        expr = f"if(lt(t,{keys[i + 1][0]:.3f}),{seg(i)},{expr})"
    return expr


def render_clip(job, m, path, src_info):
    srcs, vw, vh = src_info
    d = m["e"] - m["s"]
    ass = os.path.join(job["dir"], f"c{m['n']:02d}.ass")
    clip_ass(job["words"], m, job, ass)
    W, H = CLIP_W, CLIP_H
    if job["style"] == "blur":
        fc = (f"[0:v]scale={W}:-2,setsar=1,split[a][b];"
              f"[a]scale=144:256:force_original_aspect_ratio=increase,crop=144:256,boxblur=6:2,"
              f"scale={W}:{H},eq=brightness=-0.12:saturation=1.2[bg];"
              f"[bg][b]overlay=(W-w)/2:(H-h)/2,fps=30")
    elif vw and vh and vw * 16 > vh * 9 + 16:  # vidéo horizontale : cadrage intelligent sur le visage
        m["msg"] = "Repérage des visages…"
        try:
            cw, keys, seen = face_track(job, m, srcs[0], vw, vh)
        except Exception as e:
            print(f"[clip {job['id']}] repérage des visages raté : {str(e)[:200]}", flush=True)
            cw, keys, seen = int(vh * 9 / 16) // 2 * 2, [(0.0, (vw - int(vh * 9 / 16) // 2 * 2) / 2, 0.0)], 0
        print(f"[clip {job['id']}] clip {m['n']} : {seen} images avec visage, {len(keys)} positions de cadre", flush=True)
        m["msg"] = "Montage…"
        fc = (f"[0:v]crop={cw}:{vh}:x='{x_expr(keys)}':y=0,scale={W}:{H}:flags=lanczos,setsar=1,fps=30")
    else:  # vidéo déjà verticale (ou taille inconnue) : on remplit l'écran
        fc = f"[0:v]scale={W}:{H}:force_original_aspect_ratio=increase,crop={W}:{H},setsar=1,fps=30"
    fc += f",ass=filename={ass}:fontsdir={FONTS_DIR},format=yuv420p[v]"
    cmd = ["ffmpeg", "-y", "-loglevel", "error", "-nostdin"]
    for u, hd in srcs:
        cmd += ff_input(u, hd, m["s"], d)
    audio = "1:a:0" if len(srcs) > 1 else "0:a:0?"
    cmd += ["-filter_complex", fc, "-map", "[v]", "-map", audio,
            "-af", "loudnorm=I=-14:TP=-1.5:LRA=11", "-c:v", "libx264", "-preset", "veryfast", "-crf", "22",
            "-c:a", "aac", "-b:a", "128k", "-ar", "44100", "-movflags", "+faststart",
            "-threads", FF_THREADS, "-t", f"{d:.3f}", path]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=2400)
    if r.returncode != 0 or not os.path.exists(path) or os.path.getsize(path) < 10000:
        raise RuntimeError((r.stderr or "montage impossible").strip()[-300:])


def run_clip_job(job):
    print(f"[clip {job['id']}] {job['url']} (clips : {job['count']}, {job['length']}, {job['style']})", flush=True)
    tmp = tempfile.mkdtemp(dir=BASE)
    try:
        job["phase"], job["msg"] = "listing", "Lecture de la vidéo…"
        with yt_dlp.YoutubeDL(ydl_opts(job["site"], skip_download=True)) as y:
            info = y.extract_info(job["url"], download=False)
        if info.get("entries"):
            raise RuntimeError("C'est une playlist ou une chaîne : colle le lien d'une seule vidéo.")
        job["title"] = info.get("title") or "Vidéo"
        job["duration"] = info.get("duration") or 0
        job["vid"] = info.get("id") or ""
        if info.get("is_live") or info.get("live_status") in ("is_live", "is_upcoming"):
            raise RuntimeError("Cette vidéo est un live en cours : attends la fin du direct puis relance.")
        if job["duration"] and job["duration"] > 4 * 3600:
            raise RuntimeError("Vidéo trop longue (4 h maximum).")
        if job["duration"] and job["duration"] < 90:
            raise RuntimeError("Vidéo trop courte : colle une longue vidéo (au moins quelques minutes).")
        job["phase"], job["msg"] = "transcribing", "Récupération du texte…"
        words = clip_transcript(job, info, tmp)
        if len(words) < 40:
            raise RuntimeError("Pas assez de paroles dans cette vidéo pour trouver des moments forts.")
        job["words"] = words
        ends = sentence_ends(words)
        lines = build_lines(words, ends)
        job["phase"] = "analyzing"
        moments = find_moments(job, words, ends, lines)
        if not moments:
            raise RuntimeError("L'IA n'a trouvé aucun passage assez fort dans cette vidéo. Essaie une autre durée de clip.")
        for n, m in enumerate(moments, 1):
            m.update(n=n, status="pending", msg="", size=0, file="")
        job["moments"] = moments
        job["phase"] = "rendering"
        srcs = None
        for m in moments:
            if shutil.disk_usage(BASE).free < 300 * 1024 ** 2:
                m["status"], m["msg"] = "error", "Serveur plein : télécharge les clips prêts puis relance."
                continue
            m["status"], m["msg"] = "working", "Montage…"
            path = os.path.join(job["dir"], f"clip_{m['n']:02d}.mp4")
            for attempt in range(2):
                try:
                    if srcs is None or attempt:
                        srcs = clip_sources(job)
                    render_clip(job, m, path, srcs)
                    safe = safe_name(m["titre"] or job["title"], 50).replace(" ", "_")
                    m.update(status="done", msg="", size=os.path.getsize(path), file=path,
                             name=f"Clip_{m['n']:02d}_{safe}.mp4")
                    break
                except Exception as e:
                    msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")
                    print(f"[clip {job['id']}] clip {m['n']} essai {attempt + 1} : {msg[:300]}", flush=True)
                    if is_bot_block(msg):
                        msg = bot_message()
                    m["status"], m["msg"] = "error", ("Montage raté : " + msg)[:200]
        if not any(m["status"] == "done" for m in moments):
            raise RuntimeError(moments[0]["msg"] or "Aucun clip n'a pu être monté.")
        job["phase"], job["msg"] = "done", ""
        print(f"[clip {job['id']}] terminé : {sum(m['status'] == 'done' for m in moments)}/{len(moments)} clips", flush=True)
    except Exception as e:
        msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")
        print(f"[clip {job['id']}] erreur : {msg[:300]}", flush=True)
        if job["site"] == "youtube" and is_bot_block(msg):
            msg = bot_message()
        if "Unsupported URL" in msg:
            msg = "Lien non reconnu. Colle le lien d'une vidéo YouTube."
        job["phase"], job["error"] = "error", msg[:300]
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        job["finished"] = time.time()


@app.post("/api/clip")
def start_clip():
    if not authed():
        return jsonify(error="Non connecté"), 401
    cleanup()
    d = request.get_json(force=True) or {}
    m = re.search(r"https?://\S+", (d.get("url") or "").strip())
    if not m:
        return jsonify(error="Colle le lien complet de la vidéo (il commence par https://)."), 400
    if any(j["platform"] == "clip" and not j.get("finished") for j in JOBS.values()):
        return jsonify(error="Un clipping est déjà en cours. Attends qu'il se termine."), 409
    url = m.group(0)
    try:
        count = "auto" if d.get("count") in (None, "", "auto") else max(1, min(int(d.get("count")), 15))
    except (TypeError, ValueError):
        count = "auto"
    jid = uuid.uuid4().hex[:12]
    job = {"id": jid, "platform": "clip", "url": url, "source": url,
           "site": "youtube" if re.search(r"youtu\.?be", url) else "other",
           "count": count, "length": d.get("length") if d.get("length") in CLIP_LENGTHS else "auto",
           "style": "blur" if d.get("style") == "blur" else "smart",
           "subs": d.get("subs", True) is not False, "show_title": d.get("title", True) is not False,
           "lang": (d.get("lang") or "").strip()[:5], "phase": "queued", "msg": "En attente…", "error": "",
           "title": "", "duration": 0, "vid": "", "method": "", "part": 0, "parts": 0, "moments": [],
           "words": [], "items": [], "created": time.time(), "dir": os.path.join(BASE, "c_" + jid)}
    os.makedirs(job["dir"], exist_ok=True)
    with LOCK:
        JOBS[jid] = job
    threading.Thread(target=run_clip_job, args=(job,), daemon=True).start()
    return jsonify(job_id=jid)


@app.get("/api/clip/<jid>")
def clip_status(jid):
    if not authed():
        return jsonify(error="Non connecté"), 401
    j = JOBS.get(jid)
    if not j or j.get("platform") != "clip":
        return jsonify(error="Clipping introuvable (le serveur a peut-être redémarré). Relance-le."), 404
    keys = ("n", "s", "e", "score", "titre", "accroche", "pourquoi", "type", "status", "msg", "size")
    return jsonify(phase=j["phase"], msg=j["msg"], error=j["error"], title=j["title"], duration=j["duration"],
                   vid=j["vid"], url=j["url"], method=j["method"], part=j["part"], parts=j["parts"],
                   style=j["style"], moments=[{k: m.get(k) for k in keys} for m in j["moments"]])


@app.get("/api/clip/<jid>/v/<int:n>")
def clip_video(jid, n):
    if not authed():
        return "Non connecté", 401
    j = JOBS.get(jid)
    m = next((x for x in (j or {}).get("moments", []) if x["n"] == n and x["status"] == "done"), None)
    if not m or not os.path.exists(m["file"]):
        return "Clip introuvable (le serveur a peut-être redémarré). Relance le clipping.", 404
    return send_file(m["file"], mimetype="video/mp4", as_attachment=request.args.get("dl") == "1",
                     download_name=m["name"], max_age=3600, conditional=True)


@app.get("/api/clip/<jid>/zip")
def clip_zip(jid):
    if not authed():
        return "Non connecté", 401
    j = JOBS.get(jid)
    done = [m for m in (j or {}).get("moments", []) if m["status"] == "done" and os.path.exists(m["file"])]
    if not done:
        return "Aucun clip prêt (le serveur a peut-être redémarré).", 404
    zpath = os.path.join(j["dir"], "clips.zip")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_STORED) as z:
        for m in done:
            z.write(m["file"], m["name"])
        z.writestr("titres.txt", "\n\n".join(
            f"Clip {m['n']:02d} · {fmt_ts(m['s'])} → {fmt_ts(m['e'])} · note {m['score']}/100\n{m['titre']}\n{m['pourquoi']}"
            for m in done))
    safe = safe_name(j["title"], 40).replace(" ", "_")
    return send_file(zpath, as_attachment=True, download_name=f"Clips_{safe}.zip", mimetype="application/zip")


@app.get("/favicon.ico")
def favicon():
    return send_file(os.path.join(app.root_path, "static", "favicon.ico"), mimetype="image/x-icon")


@app.errorhandler(413)
def too_big(_):
    return jsonify(error="Fichier trop lourd (2 Go maximum)."), 413


@app.get("/health")
def health():
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)

"""Kopi — récupère la transcription de plusieurs vidéos TikTok / YouTube
et les rassemble dans un document Word (Script 1, Script 2, ...)."""
import os, re, zipfile, json, time, uuid, glob, shutil, threading, tempfile, subprocess, datetime
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
            keep = 2 * 3600 if j["platform"] in ("download", "media") else 6 * 3600
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
    return render_template("index.html", groq=bool(GROQ_KEY), yt=YT["status"], yt_src=YT["source"])


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
                   zip_size=j.get("zip_size", 0))


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
    return next((u for u in urls if re.search(r"\.jpe?g|format=jpe?g", u, re.I)), urls[0] if urls else None)


def extract_tiktok(url):
    errors = []
    try:
        r = http_get(url, headers={"Referer": "https://www.tiktok.com/"}, allow_redirects=True)
        url = r.url
        m = re.search(r'<script[^>]+id="__UNIVERSAL_DATA_FOR_REHYDRATION__"[^>]*>(.*?)</script>', r.text, re.S)
        if m:
            item = json.loads(m.group(1))["__DEFAULT_SCOPE__"]["webapp.video-detail"]["itemInfo"]["itemStruct"]
            author = (item.get("author") or {}).get("uniqueId", "")
            imgs = (item.get("imagePost") or {}).get("images") or []
            out = [{"kind": "photo", "url": pick_img((i.get("imageURL") or {}).get("urlList"))} for i in imgs]
            out = [o for o in out if o["url"]]
            if out:
                music = (item.get("music") or {}).get("playUrl")
                if music:
                    out.append({"kind": "audio", "url": music})
                return author, item.get("desc", ""), out
            if item.get("video"):
                return author, item.get("desc", ""), [{"kind": "ytdlp", "url": url}]
        errors.append("page sans données")
    except Exception as e:
        errors.append(f"page {e}")
    # secours : tikwm
    try:
        d = http_get("https://www.tikwm.com/api/", params={"url": url, "hd": 1}).json().get("data") or {}
        out = [{"kind": "photo", "url": u} for u in d.get("images") or []]
        if out:
            if d.get("music"):
                out.append({"kind": "audio", "url": d["music"]})
            return (d.get("author") or {}).get("unique_id", ""), d.get("title", ""), out
        v = d.get("hdplay") or d.get("play")
        if v:
            return (d.get("author") or {}).get("unique_id", ""), d.get("title", ""), [{"kind": "video", "url": v}]
    except Exception as e:
        errors.append(f"tikwm {e}")
    print(f"[médias] TikTok {url} : rien trouvé ({'; '.join(errors)[:300]})", flush=True)
    return "", "", [{"kind": "ytdlp", "url": url}]


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
    if it["kind"] == "photo" and ext in (".webp", ".heic"):
        try:  # JPEG : s'ouvre partout et s'enregistre dans la galerie de l'iPhone
            from PIL import Image
            jpg = os.path.splitext(path)[0] + ".jpg"
            Image.open(path).convert("RGB").save(jpg, "JPEG", quality=95, subsampling=0)
            os.remove(path)
            path = jpg
        except Exception:
            pass
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


def groq_chat(messages, max_tokens=6000):
    models = ([GOOD_MODEL["name"]] if GOOD_MODEL["name"] else []) + [m for m in CHAT_MODELS if m != GOOD_MODEL["name"]]
    last = ""
    for model in models:
        for attempt in range(6):
            r = requests.post("https://api.groq.com/openai/v1/chat/completions",
                              headers={"Authorization": f"Bearer {GROQ_KEY}"},
                              json={"model": model, "messages": messages, "temperature": 0.2,
                                    "max_tokens": max_tokens}, timeout=180)
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
            last = f"{model} : {r.status_code} {r.text[:120]}"
            print(f"[traduction] modèle refusé {last}", flush=True)
            break  # modèle indisponible : on essaie le suivant
        else:
            raise RuntimeError("Limite Groq atteinte, réessaie dans quelques minutes.")
    raise RuntimeError(f"Aucun modèle de traduction disponible chez Groq ({last}).")


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

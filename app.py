"""Kopi — récupère la transcription de plusieurs vidéos TikTok / YouTube
et les rassemble dans un document Word (Script 1, Script 2, ...)."""
import os, re, json, time, uuid, glob, shutil, threading, tempfile, subprocess, datetime
from concurrent.futures import ThreadPoolExecutor

import requests
import yt_dlp
from flask import Flask, request, jsonify, send_file, render_template, session, redirect
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or uuid.uuid4().hex

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
WORKERS = int(os.environ.get("WORKERS", "3"))
BASE = os.path.join(tempfile.gettempdir(), "script-copier")
os.makedirs(BASE, exist_ok=True)

COOKIES = {}
for name in ("YT_COOKIES", "TT_COOKIES"):
    if os.environ.get(name):
        p = os.path.join(BASE, name.lower() + ".txt")
        with open(p, "w") as f:
            f.write(os.environ[name])
        COOKIES[name] = p

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


def process_item(job, item):
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
    except Exception as e:
        msg = str(e)
        msg = re.sub(r"\x1b\[[0-9;]*m", "", msg).replace("ERROR: ", "")
        if "Sign in to confirm" in msg or "bot" in msg.lower():
            msg = "YouTube bloque le serveur (cookies nécessaires)."
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
    doc.add_heading(f"Scripts — {job['source_name']}", 0)
    done = [i for i in job["items"] if i["status"] == "done"]
    meta = doc.add_paragraph()
    r = meta.add_run(f"{'TikTok' if job['platform'] == 'tiktok' else 'YouTube'} · {len(done)} scripts · "
                     f"{datetime.datetime.now().strftime('%d/%m/%Y')}")
    r.font.color.rgb, r.font.size = RGBColor(0x66, 0x66, 0x66), Pt(10)
    for n, it in enumerate(done, 1):
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
            raise RuntimeError("Aucune vidéo n'a pu être transcrite.")
        safe = re.sub(r"[^\w-]+", "_", name)[:40].strip("_") or "scripts"
        job["filename"] = f"Scripts_{safe}.docx"
        job["file"] = os.path.join(BASE, job["id"] + ".docx")
        build_docx(job, job["file"])
        job["phase"] = "done"
    except Exception as e:
        msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "")
        if "Sign in to confirm" in msg:
            msg = "YouTube bloque le serveur. Ajoute tes cookies YouTube (voir le guide)."
        job["phase"], job["error"] = "error", msg[:300]
    job["finished"] = time.time()


def cleanup():
    now = time.time()
    with LOCK:
        for jid in list(JOBS):
            j = JOBS[jid]
            if j.get("finished") and now - j["finished"] > 6 * 3600:
                if j.get("file") and os.path.exists(j["file"]):
                    os.remove(j["file"])
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
    return render_template("index.html", groq=bool(GROQ_KEY))


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
    items = [{k: v for k, v in i.items() if k != "text"} | {"chars": len(i["text"])} for i in j["items"]]
    return jsonify(id=j["id"], phase=j["phase"], error=j["error"], source_name=j["source_name"],
                   platform=j["platform"], items=items, has_file=bool(j.get("file")))


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


@app.get("/favicon.ico")
def favicon():
    return send_file(os.path.join(app.root_path, "static", "favicon.ico"), mimetype="image/x-icon")


@app.get("/health")
def health():
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)

"""Snapchat : stories, moments forts et Spotlight publics, en fichier d'origine (sans filigrane).

Snapchat met dans chaque page publique un bloc JSON (__NEXT_DATA__) qui contient l'adresse directe
des vidéos et photos telles qu'elles ont été publiées : on les récupère sans passer par l'appli,
donc sans le logo ni le nom d'utilisateur incrustés.
"""
import html as htmlmod
import json
import re

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/131.0.0.0 Safari/537.36")
HEADERS = {"User-Agent": UA, "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
           "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"}


def is_snap(url):
    return bool(re.search(r"(^|\.)snapchat\.com/|snap\.com/", url or "", re.I))


def _val(x):
    return x.get("value") if isinstance(x, dict) else x


def _walk(o, path=()):
    if isinstance(o, dict):
        yield o, path
        for k, v in o.items():
            yield from _walk(v, path + (str(k),))
    elif isinstance(o, list):
        for i, v in enumerate(o):
            yield from _walk(v, path + (str(i),))


def _group(path):
    p = "/".join(path).lower()
    if "spotlight" in p:
        return "Spotlight"
    if "highlight" in p:
        return "Moments forts"
    if "story" in p:
        return "Story"
    return "Médias"


def snap_extract(url):
    """→ (titre, [médias]) ; chaque média = {kind, url, thumb, id, group, ts, duration}."""
    s = requests.Session()
    r = s.get(url.strip(), headers=HEADERS, timeout=25, allow_redirects=True)
    final = r.url
    if r.status_code == 404:
        raise RuntimeError("Ce lien Snapchat n'existe plus (story expirée ou supprimée).")
    if r.status_code >= 400:
        raise RuntimeError(f"Snapchat refuse l'accès à cette page (erreur {r.status_code}). Réessaie dans un instant.")
    page = r.text
    m = re.search(r'<script id="__NEXT_DATA__" type="application/json">(.+?)</script>', page, re.S)
    data = {}
    if m:
        try:
            data = json.loads(m.group(1))
        except ValueError:
            data = {}
    props = (data.get("props") or {}).get("pageProps") or {}
    want = ""
    mm = re.search(r"/spotlight/([\w-]{20,})", final) or re.search(r"/spotlight/([\w-]{20,})", url)
    if mm:
        want = mm.group(1)
    q = data.get("query") or {}
    pp = q.get("profileParams") or []
    snap_want = pp[1] if isinstance(pp, list) and len(pp) > 1 else ""

    items, seen = [], set()

    def add(kind, murl, thumb, sid, group, ts=None, dur=None, title=""):
        murl = htmlmod.unescape(murl or "")
        if not murl.startswith("http") or murl in seen:
            return
        seen.add(murl)
        items.append({"kind": kind, "url": murl, "thumb": htmlmod.unescape(thumb or ""), "id": sid or "",
                      "group": group, "ts": ts, "duration": dur, "title": title})

    for d, path in _walk(props):
        su = d.get("snapUrls")
        if isinstance(su, dict) and su.get("mediaUrl"):
            kind = "image" if d.get("snapMediaType") == 0 else "video"
            add(kind, su.get("mediaUrl"), _val(su.get("mediaPreviewUrl")), _val(d.get("snapId")), _group(path),
                ts=_val(d.get("timestampInSec")))
        vm = d.get("videoMetadata")
        if not (isinstance(vm, dict) and vm.get("contentUrl")) and isinstance(d.get("metadata"), dict):
            vm = d["metadata"].get("videoMetadata")   # Spotlight : {story: {storyId}, metadata: {videoMetadata}}
        if isinstance(vm, dict) and vm.get("contentUrl"):
            sid = ""
            st = d.get("story") if isinstance(d.get("story"), dict) else None
            if st:
                sid = _val(st.get("storyId")) or ""
            add("video", vm.get("contentUrl"), vm.get("thumbnailUrl"), sid, _group(path) if not want else "Spotlight",
                dur=(vm.get("durationMs") or 0) / 1000 or None, title=vm.get("name") or "")
    # le Spotlight demandé est dans une liste de suggestions : on ne garde que lui
    if want:
        exact = [i for i in items if want in (i["id"] or "") or want in i["url"]]
        if not exact:
            pre = re.search(r'<link[^>]+rel="preload"[^>]+href="([^"]+)"[^>]+as="video"', page) or \
                re.search(r'<link[^>]+href="([^"]+)"[^>]+as="video"', page)
            if pre:
                items.insert(0, {"kind": "video", "url": htmlmod.unescape(pre.group(1)), "thumb": "", "id": want,
                                 "group": "Spotlight", "ts": None, "duration": None, "title": ""})
                exact = items[:1]
        items = exact or items[:1]
    elif snap_want:
        exact = [i for i in items if i["id"] == snap_want]
        if exact:
            items = exact + [i for i in items if i["id"] != snap_want]
    if not items:  # dernière chance : balises de partage de la page
        for pat, kind in ((r'<meta[^>]+property="og:video(?::secure_url)?"[^>]+content="([^"]+)"', "video"),
                          (r'"contentUrl"\s*:\s*"([^"]+)"', "video")):
            for u in re.findall(pat, page):
                add(kind, u.replace("\\u002F", "/").replace("\\/", "/"), "", "", "Médias")
    title = ""
    for key in ("userProfile", "publicProfileInfo"):
        info = props.get(key) or {}
        if isinstance(info, dict):
            info = info.get("publicProfileInfo") or info.get("userInfo") or info
            title = info.get("title") or info.get("displayName") or info.get("username") or ""
            if title:
                break
    if not title:
        tm = re.search(r"<title>([^<]+)</title>", page)
        title = htmlmod.unescape(tm.group(1)).split("|")[0].strip() if tm else "Snapchat"
    if not items:
        if re.search(r"/add/|/@", final) and "story" not in "/".join(props.keys()).lower():
            raise RuntimeError("Ce compte n'a aucune story publique en ce moment (les stories disparaissent "
                               "au bout de 24 h, et les stories privées ne sont pas accessibles).")
        raise RuntimeError("Kopi ne trouve aucune vidéo sur ce lien Snapchat. Vérifie que la story ou le "
                           "Spotlight est public.")
    return title, items

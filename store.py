"""Sauvegarde permanente des tâches de Kopi : rien ne doit être perdu.

- Base principale : Cloudflare D1 (gratuite, permanente, hors du serveur Render).
- Sans D1 (permission manquante), repli sur un fichier SQLite local : ça survit aux coupures réseau
  mais pas à la mise en veille du serveur.
Les écritures partent dans une file d'attente en arrière-plan : la copie des scripts n'est jamais ralentie.
"""
import json
import os
import queue
import sqlite3
import threading
import time

import requests

SCHEMA = [
    """CREATE TABLE IF NOT EXISTS jobs (id TEXT PRIMARY KEY, kind TEXT, platform TEXT, title TEXT, source TEXT,
       url TEXT, params TEXT, status TEXT, created REAL, updated REAL, total INTEGER DEFAULT 0, done INTEGER DEFAULT 0,
       error TEXT DEFAULT '')""",
    """CREATE TABLE IF NOT EXISTS items (job_id TEXT, n INTEGER, url TEXT, title TEXT, status TEXT, method TEXT,
       msg TEXT, text TEXT, PRIMARY KEY (job_id, n))""",
    "CREATE INDEX IF NOT EXISTS items_job ON items (job_id)",
]


class Store:
    def __init__(self, local_path, cf=lambda: ("", "")):
        self.local_path = local_path
        self.cf = cf                      # fonction qui renvoie (account, token) au moment voulu
        self.db_id = None
        self.mode = "local"
        self.note = ""
        self.q = queue.Queue()
        self.lock = threading.Lock()
        self._local = None
        threading.Thread(target=self._writer, daemon=True).start()

    # ------------------------------------------------------------ connexion
    def connect(self, log=print):
        """Essaie D1, sinon SQLite local. Peut être rappelée (ex. quand l'Account ID arrive plus tard)."""
        acc, tok = self.cf()
        if acc and tok:
            try:
                h = {"Authorization": f"Bearer {tok}"}
                base = f"https://api.cloudflare.com/client/v4/accounts/{acc}/d1/database"
                r = requests.get(base, headers=h, params={"name": "kopi"}, timeout=20)
                if r.status_code == 200:
                    found = [d for d in r.json().get("result") or [] if d.get("name") == "kopi"]
                    if found:
                        self.db_id = found[0]["uuid"]
                    else:
                        c = requests.post(base, headers=h, json={"name": "kopi"}, timeout=30)
                        if c.status_code != 200:
                            raise RuntimeError(f"création refusée ({c.status_code}) {c.text[:120]}")
                        self.db_id = c.json()["result"]["uuid"]
                    self.mode = "d1"
                    for sql in SCHEMA:
                        self._d1(sql, [])
                    self.note = ""
                    log("Historique : sauvegarde permanente Cloudflare D1 ✓")
                    return True
                self.note = ("Le jeton Cloudflare n'a pas la permission « D1 ». Ajoute-la pour que la sauvegarde soit "
                             "permanente.")
                log(f"Historique : D1 refusé ({r.status_code}) → sauvegarde locale")
            except Exception as e:
                self.note = f"Cloudflare D1 injoignable : {str(e)[:100]}"
                log(f"Historique : D1 impossible ({str(e)[:100]}) → sauvegarde locale")
        else:
            self.note = "Cloudflare n'est pas encore relié (Account ID manquant)."
        self.mode = "local"
        self._local_conn()
        return False

    def _local_conn(self):
        if self._local is None:
            self._local = sqlite3.connect(self.local_path, check_same_thread=False)
            for sql in SCHEMA:
                self._local.execute(sql)
            self._local.commit()
        return self._local

    def _d1(self, sql, params):
        acc, tok = self.cf()
        for attempt in range(4):
            try:
                r = requests.post(f"https://api.cloudflare.com/client/v4/accounts/{acc}/d1/database/{self.db_id}/query",
                                  headers={"Authorization": f"Bearer {tok}"}, json={"sql": sql, "params": params},
                                  timeout=40)
                if r.status_code == 200:
                    res = r.json().get("result") or [{}]
                    return res[0].get("results") or []
                if r.status_code in (400, 401, 403):
                    raise RuntimeError(f"D1 {r.status_code} : {r.text[:150]}")
            except requests.RequestException:
                pass
            time.sleep(1.5 * (attempt + 1))
        raise RuntimeError("D1 ne répond pas")

    def query(self, sql, params=()):
        """Lecture immédiate (renvoie une liste de dictionnaires)."""
        if self.mode == "d1":
            return self._d1(sql, list(params))
        with self.lock:
            cur = self._local_conn().execute(sql, list(params))
            cols = [c[0] for c in cur.description or []]
            return [dict(zip(cols, row)) for row in cur.fetchall()]

    def write(self, sql, params=()):
        """Écriture en arrière-plan (ne bloque jamais la tâche)."""
        self.q.put((sql, list(params)))

    def _writer(self):
        while True:
            sql, params = self.q.get()
            for attempt in range(6):
                try:
                    if self.mode == "d1":
                        self._d1(sql, params)
                    else:
                        with self.lock:
                            self._local_conn().execute(sql, params)
                            self._local.commit()
                    break
                except Exception as e:
                    if attempt == 5:
                        print(f"Historique : écriture perdue ({str(e)[:100]})", flush=True)
                    time.sleep(2 * (attempt + 1))

    def flush(self, timeout=60):
        t = time.time()
        while not self.q.empty() and time.time() - t < timeout:
            time.sleep(.2)

    # ------------------------------------------------------------ tâches
    def save_job(self, job, kind, status, params=None):
        items = job.get("items") or []
        self.write("""INSERT INTO jobs (id, kind, platform, title, source, url, params, status, created, updated, total,
                      done, error) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                      ON CONFLICT(id) DO UPDATE SET title=excluded.title, status=excluded.status,
                      updated=excluded.updated, total=excluded.total, done=excluded.done, error=excluded.error,
                      params=excluded.params""",
                   [job["id"], kind, job.get("platform", ""), (job.get("source_name") or "")[:300],
                    (job.get("source") or "")[:500], (job.get("url") or "")[:500],
                    json.dumps(params or {}, ensure_ascii=False), status, job.get("created") or time.time(),
                    time.time(), len(items), sum(1 for i in items if i.get("status") == "done"),
                    (job.get("error") or "")[:300]])

    def save_items(self, job):
        """Toute la liste d'un coup (par paquets : D1 accepte 100 valeurs par requête)."""
        rows = job.get("items") or []
        for k in range(0, len(rows), 12):
            chunk = rows[k:k + 12]
            vals, params = [], []
            for it in chunk:
                vals.append("(?,?,?,?,?,?,?,?)")
                params += [job["id"], it["n"], it.get("url", ""), (it.get("title") or "")[:500], it.get("status", ""),
                           it.get("method", ""), (it.get("msg") or "")[:300], it.get("text", "")]
            self.write("INSERT OR REPLACE INTO items (job_id, n, url, title, status, method, msg, text) VALUES "
                       + ",".join(vals), params)

    def save_item(self, job, it):
        self.write("INSERT OR REPLACE INTO items (job_id, n, url, title, status, method, msg, text) VALUES "
                   "(?,?,?,?,?,?,?,?)",
                   [job["id"], it["n"], it.get("url", ""), (it.get("title") or "")[:500], it.get("status", ""),
                    it.get("method", ""), (it.get("msg") or "")[:300], it.get("text", "")])
        if it["n"] % 10 == 0:  # le compteur de la tâche suit, sans écrire à chaque vidéo
            done = sum(1 for i in job.get("items") or [] if i.get("status") == "done")
            self.write("UPDATE jobs SET done=?, updated=? WHERE id=?", [done, time.time(), job["id"]])

    def jobs(self, limit=200):
        return self.query("SELECT id, kind, platform, title, source, url, params, status, created, updated, total, "
                          "done, error FROM jobs ORDER BY created DESC LIMIT ?", [limit])

    def job(self, jid):
        r = self.query("SELECT * FROM jobs WHERE id=?", [jid])
        return r[0] if r else None

    def items(self, jid):
        return self.query("SELECT n, url, title, status, method, msg, text FROM items WHERE job_id=? ORDER BY n", [jid])

    def delete(self, jid):
        self.write("DELETE FROM items WHERE job_id=?", [jid])
        self.write("DELETE FROM jobs WHERE id=?", [jid])

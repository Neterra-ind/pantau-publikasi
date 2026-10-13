"""Penyimpanan berkas JSON: klien, periode (postingan + target per bulan), pengguna, alias nama.

data/klien.json                  daftar klien (nama, istilah kelompok, warna, logo, target bawaan, alias)
data/pengguna.json               pengguna (email, peran admin/operator, klien yang diikuti, hash sandi)
data/periode/<klien>/<YYYY-MM>.json   {targets, posts}
data/logo/<klien>.png            logo klien
"""
import hashlib
import json
import os
import re
import secrets
import shutil
import threading
import uuid

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.environ.get("PANTAU_DATA") or os.path.join(ROOT, "data")
KUNCI = threading.RLock()


def _p(*a):
    return os.path.join(DATA, *a)


def baca(path, bawaan):
    if not os.path.exists(path):
        return bawaan
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def tulis(path, isi):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(isi, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)


def id_baru(awalan=""):
    return awalan + uuid.uuid4().hex[:10]


# ------------------------------------------------------------------ klien
def semua_klien():
    return baca(_p("klien.json"), {"klien": []})["klien"]


def simpan_semua_klien(daftar):
    tulis(_p("klien.json"), {"klien": daftar})


def ambil_klien(kid):
    return next((k for k in semua_klien() if k["id"] == kid), None)


def simpan_klien(k):
    daftar = semua_klien()
    for i, x in enumerate(daftar):
        if x["id"] == k["id"]:
            daftar[i] = k
            break
    else:
        daftar.append(k)
    simpan_semua_klien(daftar)


def hapus_klien(kid):
    simpan_semua_klien([k for k in semua_klien() if k["id"] != kid])
    shutil.rmtree(_p("periode", kid), ignore_errors=True)
    for f in (_p("logo", kid + ".png"),):
        if os.path.exists(f):
            os.remove(f)
    pengguna = semua_pengguna()
    for u in pengguna:
        u["klien"] = [x for x in u.get("klien", []) if x != kid]
    simpan_semua_pengguna(pengguna)


def jalur_logo(kid):
    f = _p("logo", kid + ".png")
    return f if os.path.exists(f) else None


def simpan_logo(kid, biner):
    """Logo diperkecil ke maks 256 px dan disimpan sebagai PNG."""
    from io import BytesIO
    from PIL import Image
    im = Image.open(BytesIO(biner))
    im = im.convert("RGBA")
    im.thumbnail((256, 256))
    os.makedirs(_p("logo"), exist_ok=True)
    im.save(_p("logo", kid + ".png"), "PNG")


def hapus_logo(kid):
    f = _p("logo", kid + ".png")
    if os.path.exists(f):
        os.remove(f)


# ------------------------------------------------------------------ periode
def daftar_periode(kid):
    d = _p("periode", kid)
    if not os.path.isdir(d):
        return []
    return sorted(f[:7] for f in os.listdir(d) if re.fullmatch(r"\d{4}-\d{2}\.json", f))


def ambil_periode(kid, ym, buat=False):
    f = _p("periode", kid, ym + ".json")
    if not os.path.exists(f):
        if not buat:
            return None
        return {"ym": ym, "targets": None, "posts": []}
    return baca(f, None)


def simpan_periode(kid, per):
    tulis(_p("periode", kid, per["ym"] + ".json"), per)


# ------------------------------------------------------------------ pengguna
def semua_pengguna():
    return baca(_p("pengguna.json"), {"pengguna": []})["pengguna"]


def simpan_semua_pengguna(daftar):
    tulis(_p("pengguna.json"), {"pengguna": daftar})


def ambil_pengguna(email):
    email = (email or "").strip().lower()
    return next((u for u in semua_pengguna() if u["email"] == email), None)


def hash_sandi(sandi, garam=None):
    garam = garam or secrets.token_hex(16)
    return garam, _pbkdf2(sandi.encode("utf-8"), bytes.fromhex(garam), 200_000).hex()


def _pbkdf2(sandi, garam, n):
    if hasattr(hashlib, "pbkdf2_hmac"):
        return hashlib.pbkdf2_hmac("sha256", sandi, garam, n)
    import hmac  # cadangan murni Python (mis. Python di browser tanpa OpenSSL)
    u = hmac.new(sandi, garam + b"\x00\x00\x00\x01", hashlib.sha256).digest()
    hasil = int.from_bytes(u, "big")
    for _ in range(n - 1):
        u = hmac.new(sandi, u, hashlib.sha256).digest()
        hasil ^= int.from_bytes(u, "big")
    return hasil.to_bytes(32, "big")


def cek_sandi(u, sandi):
    if not u or not u.get("hash"):
        return False
    _, h = hash_sandi(sandi, u["garam"])
    return secrets.compare_digest(h, u["hash"])


def set_sandi(u, sandi):
    u["garam"], u["hash"] = hash_sandi(sandi)


def publik(u):
    return {k: u.get(k) for k in ("email", "nama", "peran", "klien")}

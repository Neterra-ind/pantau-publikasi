"""
Server lokal Pantau Publikasi.

    python scripts/server.py            # http://localhost:8082
    python scripts/server.py 9000       # port lain

Menyajikan tampilan + API JSON (login, klien, periode, postingan, target, impor, ekspor).
Hanya menerima koneksi dari komputer ini (127.0.0.1). Perlu: pip install openpyxl python-pptx pillow
"""
import base64
import json
import os
import re
import secrets
import sys
import threading
import time
from datetime import date, datetime
from http.cookies import SimpleCookie
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from urllib.parse import parse_qs, quote, urlparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import inti  # noqa: E402
import penyimpanan as ps  # noqa: E402

BATAS_JSON = 2_000_000
BATAS_FILE = 40_000_000
SESI = {}  # token -> (email, kedaluwarsa)
SESI_LAMA = 12 * 3600
KUNCI_SESI = threading.Lock()


class Galat(Exception):
    def __init__(self, pesan, kode=400):
        super().__init__(pesan)
        self.kode = kode


def hari_ini():
    t = os.environ.get("PANTAU_HARI_INI")  # untuk uji: tanggal tetap
    return date.fromisoformat(t) if t else datetime.now(inti.WIB).date()


def cek_ym(ym):
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", str(ym or "")):
        raise Galat("periode harus berformat YYYY-MM")
    return ym


# ------------------------------------------------------------------ akses
def klien_untuk(u, kid, perlu_admin=False):
    k = ps.ambil_klien(kid)
    if not k or kid not in (u.get("klien") or []):
        raise Galat("klien tidak ditemukan atau Anda bukan anggotanya", 404)
    if perlu_admin and u["peran"] != "admin":
        raise Galat("hanya admin yang boleh melakukan ini", 403)
    return k


def muat_periode(k, ym):
    per = ps.ambil_periode(k["id"], ym)
    if per is None:
        raise Galat("periode belum ada; tambahkan dulu dari pemilih bulan", 404)
    return per


def ringkasan_periode(k, per):
    targets, _ = inti.target_efektif(k, per)
    lalu = ps.ambil_periode(k["id"], inti.periode_sebelum(per["ym"]))
    pakai = k.get("pakai_kelompok", True)
    flags = inti.tandai(per["posts"], per["ym"], pakai)
    r = inti.ringkas(per["posts"], per["ym"], targets, (lalu or {}).get("posts") or None, hari_ini(),
                     k.get("label_kelompok") or "Kelompok", pakai, flags)
    return r, flags


def paket_periode(k, per):
    r, flags = ringkasan_periode(k, per)
    targets, sumber = inti.target_efektif(k, per)
    return {"ym": per["ym"], "posts": per["posts"], "flags": flags, "ringkasan": r,
            "targets": targets, "sumber_target": sumber, "flag_info": inti.FLAG}


def klien_publik(k, u):
    return {**{x: k.get(x) for x in ("id", "nama", "label_kelompok", "pakai_kelompok", "warna", "targets", "alias")},
            "logo": bool(ps.jalur_logo(k["id"])), "periode": ps.daftar_periode(k["id"]),
            "peran": "owner" if u["peran"] == "admin" else "editor"}


# ------------------------------------------------------------------ handler
class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=ps.ROOT, **kw)

    def log_message(self, fmt, *args):
        if args and "/api/" in str(args[0]):
            sys.stderr.write("%s %s\n" % (self.log_date_time_string(), fmt % args))

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    # ---- util
    def _pengguna(self):
        c = SimpleCookie(self.headers.get("Cookie") or "")
        tok = c["sesi"].value if "sesi" in c else None
        with KUNCI_SESI:
            s = SESI.get(tok)
            if not s or s[1] < time.time():
                return None
            SESI[tok] = (s[0], time.time() + SESI_LAMA)
        return ps.ambil_pengguna(s[0])

    def _wajib(self):
        u = self._pengguna()
        if not u:
            raise Galat("silakan masuk dulu", 401)
        return u

    def _json_masuk(self, batas=BATAS_JSON):
        n = int(self.headers.get("Content-Length") or 0)
        if n > batas:
            raise Galat("data terlalu besar", 413)
        try:
            return json.loads(self.rfile.read(n) or b"{}")
        except json.JSONDecodeError:
            raise Galat("format data tidak valid")

    def _balas(self, obj, kode=200, cookie=None):
        b = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(kode)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        if cookie:
            self.send_header("Set-Cookie", cookie)
        self.end_headers()
        self.wfile.write(b)

    def _berkas(self, biner, nama, tipe):
        self.send_response(200)
        self.send_header("Content-Type", tipe)
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(nama)}")
        self.send_header("Content-Length", str(len(biner)))
        self.end_headers()
        self.wfile.write(biner)

    def _jalankan(self, fn):
        try:
            fn()
        except Galat as e:
            self._balas({"galat": str(e)}, e.kode)
        except Exception as e:  # noqa: BLE001
            import traceback
            traceback.print_exc()
            self._balas({"galat": f"kesalahan server: {e}"}, 500)

    # ---- GET
    def do_GET(self):
        u = urlparse(self.path)
        if u.path.startswith("/data/") or u.path.startswith("/scripts/") or u.path.startswith("/tests/"):
            self.send_error(404)
            return
        if not u.path.startswith("/api/"):
            return super().do_GET()
        q = {k: v[0] for k, v in parse_qs(u.query).items()}
        self._jalankan(lambda: self._get(u.path, q))

    def _get(self, path, q):
        if path == "/api/saya":
            usr = self._pengguna()
            if not usr:
                return self._balas({"pengguna": None})
            klien = [klien_publik(k, usr) for k in ps.semua_klien() if k["id"] in (usr.get("klien") or [])]
            return self._balas({"pengguna": ps.publik(usr), "klien": klien, "hari_ini": hari_ini().isoformat()})
        usr = self._wajib()
        if path == "/api/periode":
            k = klien_untuk(usr, q.get("klien"))
            per = muat_periode(k, cek_ym(q.get("ym")))
            return self._balas(paket_periode(k, per))
        if path == "/api/logo":
            k = klien_untuk(usr, q.get("klien"))
            f = ps.jalur_logo(k["id"])
            if not f:
                raise Galat("tanpa logo", 404)
            with open(f, "rb") as fh:
                b = fh.read()
            self.send_response(200)
            self.send_header("Content-Type", "image/png")
            self.send_header("Content-Length", str(len(b)))
            self.end_headers()
            self.wfile.write(b)
            return
        if path == "/api/rapikan":
            k = klien_untuk(usr, q.get("klien"))
            per = muat_periode(k, cek_ym(q.get("ym")))
            kolom = q.get("kolom") if q.get("kolom") in ("kelompok", "akun") else "kelompok"
            rencana = inti.rencana_rapikan(per["posts"], kolom, (k.get("alias") or {}).get(kolom, {}))
            del rencana["peta"]
            return self._balas(rencana)
        if path.startswith("/api/ekspor/"):
            jenis = path.rsplit("/", 1)[1]
            k = klien_untuk(usr, q.get("klien"))
            per = muat_periode(k, cek_ym(q.get("ym")))
            r, flags = ringkasan_periode(k, per)
            nama = f"Laporan {k['nama']} {r['nama_periode']}"
            if jenis == "xlsx":
                import ekspor_xlsx
                return self._berkas(ekspor_xlsx.buat_laporan(k, r, per["posts"], flags), nama + ".xlsx",
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            if jenis == "pptx":
                import ekspor_pptx
                return self._berkas(ekspor_pptx.buat(k, r, per["posts"], ps.jalur_logo(k["id"])), nama + ".pptx",
                                    "application/vnd.openxmlformats-officedocument.presentationml.presentation")
            raise Galat("jenis ekspor tidak dikenal", 404)
        if path == "/api/pengguna":
            if usr["peran"] != "admin":
                raise Galat("hanya admin", 403)
            return self._balas({"pengguna": [ps.publik(x) for x in ps.semua_pengguna()]})
        if path == "/api/template":
            import impor_workbook
            buf = os.path.join(ps.DATA, "_template.xlsx")
            impor_workbook.buat_template(buf)
            with open(buf, "rb") as fh:
                b = fh.read()
            os.remove(buf)
            return self._berkas(b, "Template Pantau Publikasi.xlsx",
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        raise Galat("alamat tidak dikenal", 404)

    # ---- POST
    def do_POST(self):
        u = urlparse(self.path)
        self._jalankan(lambda: self._post(u.path))

    def _post(self, path):
        if path == "/api/masuk":
            d = self._json_masuk(10_000)
            usr = ps.ambil_pengguna(d.get("email"))
            if not ps.cek_sandi(usr, str(d.get("sandi") or "")):
                time.sleep(0.5)
                raise Galat("email atau kata sandi salah", 401)
            tok = secrets.token_urlsafe(32)
            with KUNCI_SESI:
                SESI[tok] = (usr["email"], time.time() + SESI_LAMA)
            return self._balas({"ok": True}, cookie=f"sesi={tok}; HttpOnly; SameSite=Strict; Path=/")
        if path == "/api/keluar":
            c = SimpleCookie(self.headers.get("Cookie") or "")
            if "sesi" in c:
                with KUNCI_SESI:
                    SESI.pop(c["sesi"].value, None)
            return self._balas({"ok": True}, cookie="sesi=; Max-Age=0; Path=/")

        usr = self._wajib()
        batas = BATAS_FILE if path in ("/api/impor", "/api/klien/logo") else BATAS_JSON
        d = self._json_masuk(batas)
        with ps.KUNCI:
            hasil = self._post_data(path, usr, d)
        if isinstance(hasil, tuple):
            return self._berkas(*hasil)
        return self._balas(hasil)

    def _post_data(self, path, usr, d):
        admin = usr["peran"] == "admin"
        # ---------------- klien
        if path == "/api/klien":
            if not admin:
                raise Galat("hanya admin yang boleh mengelola klien", 403)
            nama = str(d.get("nama") or "").strip()
            if not nama:
                raise Galat("nama klien wajib diisi")
            warna = str(d.get("warna") or "#C00000").strip()
            if not re.fullmatch(r"#[0-9A-Fa-f]{6}", warna):
                raise Galat("warna harus berformat #RRGGBB")
            if d.get("id"):
                k = klien_untuk(usr, d["id"], True)
            else:
                k = {"id": ps.id_baru("k"), "targets": {}, "alias": {"kelompok": {}, "akun": {}}}
                semua = ps.semua_pengguna()
                for x in semua:
                    if x["email"] == usr["email"]:
                        x["klien"] = (x.get("klien") or []) + [k["id"]]
                ps.simpan_semua_pengguna(semua)
            k.update({"nama": nama, "warna": warna.upper(),
                      "label_kelompok": str(d.get("label_kelompok") or "Kelompok").strip()[:30] or "Kelompok",
                      "pakai_kelompok": bool(d.get("pakai_kelompok", True))})
            ps.simpan_klien(k)
            if not ps.daftar_periode(k["id"]):
                ps.simpan_periode(k["id"], ps.ambil_periode(k["id"], hari_ini().strftime("%Y-%m"), buat=True))
            return {"ok": True, "id": k["id"]}
        if path == "/api/klien/hapus":
            k = klien_untuk(usr, d.get("id"), True)
            if str(d.get("konfirmasi") or "").strip() != k["nama"]:
                raise Galat("ketik nama klien persis untuk mengonfirmasi penghapusan")
            ps.hapus_klien(k["id"])
            return {"ok": True}
        if path == "/api/klien/logo":
            k = klien_untuk(usr, d.get("klien"), True)
            if d.get("hapus"):
                ps.hapus_logo(k["id"])
            else:
                try:
                    ps.simpan_logo(k["id"], base64.b64decode(d.get("data") or ""))
                except Exception:
                    raise Galat("gambar logo tidak bisa dibaca (gunakan PNG/JPG)")
            return {"ok": True}
        if path == "/api/alias":
            k = klien_untuk(usr, d.get("klien"), True)
            kolom = d.get("kolom") if d.get("kolom") in ("kelompok", "akun") else "kelompok"
            al = k.setdefault("alias", {}).setdefault(kolom, {})
            for x in d.get("hapus") or []:
                al.pop(inti.name_key(x), None)
            for x in d.get("tambah") or []:
                lama, baku = str(x.get("dari") or "").strip(), str(x.get("ke") or "").strip()
                if lama and baku and inti.name_key(lama) != inti.name_key(baku):
                    al[inti.name_key(lama)] = baku
            ps.simpan_klien(k)
            return {"ok": True, "alias": k["alias"]}

        # ---------------- periode & target
        if path == "/api/periode/tambah":
            k = klien_untuk(usr, d.get("klien"))
            ym = cek_ym(d.get("ym"))
            if ps.ambil_periode(k["id"], ym) is None:
                ps.simpan_periode(k["id"], ps.ambil_periode(k["id"], ym, buat=True))
            return {"ok": True, "periode": ps.daftar_periode(k["id"])}
        if path == "/api/target":
            k = klien_untuk(usr, d.get("klien"), True)
            t = bersihkan_target(d.get("targets"))
            if d.get("bawaan"):
                k["targets"] = t
                ps.simpan_klien(k)
            else:
                per = muat_periode(k, cek_ym(d.get("ym")))
                per["targets"] = t if t else None
                ps.simpan_periode(k["id"], per)
            return {"ok": True}
        if path == "/api/target/salin":
            k = klien_untuk(usr, d.get("klien"), True)
            per = muat_periode(k, cek_ym(d.get("ym")))
            lalu = ps.ambil_periode(k["id"], inti.periode_sebelum(per["ym"]))
            if not lalu:
                raise Galat("bulan lalu belum punya data")
            t, _ = inti.target_efektif(k, lalu)
            if not t:
                raise Galat("bulan lalu belum punya target")
            per["targets"] = json.loads(json.dumps(t))
            ps.simpan_periode(k["id"], per)
            return {"ok": True}

        # ---------------- postingan
        if path in ("/api/posting/simpan", "/api/posting/tambah", "/api/posting/hapus", "/api/tempel", "/api/impor",
                    "/api/rapikan", "/api/ekspor/data"):
            k = klien_untuk(usr, d.get("klien"))
            per = muat_periode(k, cek_ym(d.get("ym")))
            alias = k.get("alias") or {}
            hasil = {"ok": True}
            if path == "/api/posting/simpan":
                peta = {p["id"]: p for p in per["posts"]}
                tidak_ada = []
                for ubah in d.get("ubah") or []:
                    p = peta.get(ubah.get("id"))
                    if not p:
                        tidak_ada.append(ubah.get("id"))
                        continue
                    for f, v in (ubah.get("isi") or {}).items():
                        if f in ("platform", "tanggal", "kelompok", "akun", "jenis", "judul", "link", "catatan") + tuple(inti.METRIK):
                            p[f] = v
                    inti.bersihkan_post(p)
                hasil["tidak_ada"] = tidak_ada
                hasil["posts"] = [peta[x["id"]] for x in d.get("ubah") or [] if x.get("id") in peta]
            elif path == "/api/posting/tambah":
                pl = d.get("platform") if d.get("platform") in inti.PLATFORM else "ig"
                tgl = hari_ini().isoformat() if hari_ini().strftime("%Y-%m") == per["ym"] else per["ym"] + "-01"
                p = inti.bersihkan_post({"id": ps.id_baru("p"), "platform": pl, "tanggal": tgl, "link": ""})
                per["posts"].insert(0, p)
                hasil["post"] = p
            elif path == "/api/posting/hapus":
                ids = set(d.get("ids") or [])
                sebelum = len(per["posts"])
                per["posts"] = [p for p in per["posts"] if p["id"] not in ids]
                hasil["dihapus"] = sebelum - len(per["posts"])
            elif path == "/api/tempel":
                rows, dilewati, ket = inti.parse_tempel(d.get("teks"))
                siap, koreksi = [], []
                for row in rows:
                    row, info = inti.isi_dari_link(row, hari_ini())
                    if not info:
                        dilewati += 1
                        ket.append(f"Link tidak valid: {str(row.get('link'))[:60]}")
                        continue
                    if row.get("tanggal"):
                        row["tanggal"] = inti.baca_tanggal(row["tanggal"])
                        row["tanggal"], kor = inti.koreksi_tahun(row["tanggal"], per["ym"])
                        if kor:
                            koreksi.append(row["link"])
                    for kol in ("kelompok", "akun"):
                        if row.get(kol):
                            row[kol] = inti.terapkan_alias(row[kol], alias.get(kol, {}))
                    if row.get("jenis"):
                        row["jenis"] = inti.jenis_dari(row["jenis"], row["platform"])
                    siap.append(row)
                baru, upd = inti.gabung(per["posts"], siap, lambda: ps.id_baru("p"))
                # baris baru ke atas
                hasil.update({"baru": baru, "diperbarui": upd, "dilewati": dilewati, "keterangan": ket[:20],
                              "tanggal_dikoreksi": len(koreksi)})
            elif path == "/api/impor":
                import impor_workbook
                try:
                    b = base64.b64decode(d.get("data") or "")
                    rows, lap = impor_workbook.baca_workbook(BytesIO(b), per["ym"], alias, hari_ini())
                except Exception as e:  # noqa: BLE001
                    raise Galat(f"file Excel tidak bisa dibaca: {e}")
                baru, upd = inti.gabung(per["posts"], rows, lambda: ps.id_baru("p"), pertahankan_ganda=True)
                hasil.update({"laporan": lap, "baru": baru, "diperbarui": upd})
            elif path == "/api/rapikan":
                kolom = d.get("kolom") if d.get("kolom") in ("kelompok", "akun") else "kelompok"
                al = k.setdefault("alias", {}).setdefault(kolom, {})
                rencana = inti.rencana_rapikan(per["posts"], kolom, al)
                for p in per["posts"]:
                    v = str(p.get(kolom) or "").strip()
                    if v in rencana["peta"]:
                        p[kolom] = rencana["peta"][v]
                for dari, ke in rencana["peta"].items():  # simpan aturan "nama lama = nama baku"
                    if inti.name_key(dari) != inti.name_key(ke):
                        al[inti.name_key(dari)] = ke
                ps.simpan_klien(k)
                hasil["sel"] = rencana["sel"]
            elif path == "/api/ekspor/data":
                import ekspor_xlsx
                flags = inti.tandai(per["posts"], per["ym"], k.get("pakai_kelompok", True))
                f = d.get("filter") or {}
                rows = inti.saring(per["posts"], f, flags)
                teks = inti.keterangan_filter(f, k.get("label_kelompok") or "Kelompok")
                b = ekspor_xlsx.buat_data(k, per["ym"], rows, flags, teks)
                return (b, f"Data {k['nama']} {inti.nama_periode(per['ym'])}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            if path != "/api/posting/simpan" or d.get("ubah"):
                ps.simpan_periode(k["id"], per)
            r, flags = ringkasan_periode(k, per)
            hasil.update({"flags": flags, "ringkasan": r})
            return hasil

        # ---------------- pengguna
        if path == "/api/pengguna":
            if not admin:
                raise Galat("hanya admin", 403)
            email = str(d.get("email") or "").strip().lower()
            if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
                raise Galat("email tidak valid")
            semua = ps.semua_pengguna()
            x = next((y for y in semua if y["email"] == email), None)
            if d.get("hapus"):
                if email == usr["email"]:
                    raise Galat("tidak bisa menghapus akun sendiri")
                ps.simpan_semua_pengguna([y for y in semua if y["email"] != email])
                return {"ok": True}
            if not x:
                if len(str(d.get("sandi") or "")) < 8:
                    raise Galat("kata sandi awal minimal 8 karakter")
                x = {"email": email}
                semua.append(x)
            milik = set(usr.get("klien") or [])
            x["nama"] = str(d.get("nama") or email.split("@")[0]).strip()
            x["peran"] = "admin" if d.get("peran") == "admin" else "operator"
            # admin hanya bisa membagikan klien yang ia ikuti sendiri
            lain = [c for c in (x.get("klien") or []) if c not in milik]
            x["klien"] = lain + [c for c in (d.get("klien") or []) if c in milik]
            if d.get("sandi"):
                if len(str(d["sandi"])) < 8:
                    raise Galat("kata sandi minimal 8 karakter")
                ps.set_sandi(x, str(d["sandi"]))
            ps.simpan_semua_pengguna(semua)
            return {"ok": True}
        if path == "/api/sandi":
            lama, baru = str(d.get("lama") or ""), str(d.get("baru") or "")
            if not ps.cek_sandi(usr, lama):
                raise Galat("kata sandi lama salah")
            if len(baru) < 8:
                raise Galat("kata sandi baru minimal 8 karakter")
            semua = ps.semua_pengguna()
            for x in semua:
                if x["email"] == usr["email"]:
                    ps.set_sandi(x, baru)
            ps.simpan_semua_pengguna(semua)
            return {"ok": True}
        raise Galat("alamat tidak dikenal", 404)


def bersihkan_target(t):
    """Target dari form -> bentuk jsonb PRD. Kosong berarti tidak dinilai."""
    t = t or {}

    def angka(v):
        n = inti.baca_angka(v)
        return n if n and n > 0 else None

    def er(v):  # form mengirim persen, mis. "3" atau "2,5"
        if v in (None, ""):
            return None
        try:
            x = float(str(v).replace(",", "."))
        except ValueError:
            return None
        return x / 100 if x > 0 else None

    hasil = {}
    posts = {pl: angka((t.get("posts") or {}).get(pl)) for pl in inti.PLATFORM}
    if any(posts.values()):
        hasil["posts"] = {k: v for k, v in posts.items() if v}
    for kk in ("views", "interactions"):
        if angka(t.get(kk)):
            hasil[kk] = angka(t.get(kk))
    if er(t.get("er")):
        hasil["er"] = er(t.get("er"))
    ep = {pl: er((t.get("erPlatform") or {}).get(pl)) for pl in inti.SOSMED}
    if any(ep.values()):
        hasil["erPlatform"] = {k: v for k, v in ep.items() if v}
    for kk in ("perGroup", "perAccount"):
        m = {str(n).strip(): angka(v) for n, v in (t.get(kk) or {}).items() if str(n).strip() and angka(v)}
        if m:
            hasil[kk] = m
    return hasil


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8082
    if not ps.semua_pengguna():
        print("Belum ada pengguna. Jalankan dulu: python scripts/buat_simulasi.py")
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"Pantau Publikasi berjalan di http://localhost:{port}  (Ctrl+C untuk berhenti)")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()

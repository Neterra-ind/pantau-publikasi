"""Uji server: login, batas akses per klien (pengganti uji RLS), simpan, dan ekspor — memakai folder data sementara."""
import json
import os
import sys
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer

AKAR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TMP = tempfile.mkdtemp(prefix="pantau_uji_")
os.environ["PANTAU_DATA"] = TMP
os.environ["PANTAU_HARI_INI"] = "2026-10-13"
sys.path.insert(0, os.path.join(AKAR, "scripts"))
import penyimpanan as ps  # noqa: E402
import server  # noqa: E402


class Klien:
    def __init__(self, port):
        self.base = f"http://127.0.0.1:{port}"
        self.cookie = ""

    def req(self, path, body=None):
        data = json.dumps(body).encode() if body is not None else None
        r = urllib.request.Request(self.base + path, data=data, headers={"Content-Type": "application/json", "Cookie": self.cookie})
        try:
            with urllib.request.urlopen(r) as res:
                sc = res.headers.get("Set-Cookie")
                if sc:
                    self.cookie = sc.split(";")[0]
                isi = res.read()
                return res.status, (json.loads(isi) if "json" in res.headers.get("Content-Type", "") else isi)
        except urllib.error.HTTPError as e:
            return e.code, json.loads(e.read() or b"{}")


class UjiServer(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        ps.simpan_semua_klien([{"id": "kA", "nama": "Klien A", "label_kelompok": "DPC", "pakai_kelompok": True, "warna": "#C00000", "targets": {}},
                               {"id": "kB", "nama": "Klien B", "label_kelompok": "Kota", "pakai_kelompok": True, "warna": "#0000C0", "targets": {}}])
        for kid in ("kA", "kB"):
            ps.simpan_periode(kid, {"ym": "2026-10", "targets": None, "posts": []})
        pengguna = []
        for email, peran, kl in (("admin@uji.test", "admin", ["kA", "kB"]), ("op@uji.test", "operator", ["kA"])):
            u = {"email": email, "nama": email, "peran": peran, "klien": kl}
            ps.set_sandi(u, "sandi-uji-123")
            pengguna.append(u)
        ps.simpan_semua_pengguna(pengguna)
        cls.srv = ThreadingHTTPServer(("127.0.0.1", 0), server.Handler)
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        cls.port = cls.srv.server_address[1]

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()

    def masuk(self, email):
        k = Klien(self.port)
        st, _ = k.req("/api/masuk", {"email": email, "sandi": "sandi-uji-123"})
        self.assertEqual(st, 200)
        return k

    def test_wajib_login(self):
        st, _ = Klien(self.port).req("/api/periode?klien=kA&ym=2026-10")
        self.assertEqual(st, 401)
        st, _ = Klien(self.port).req("/api/masuk", {"email": "op@uji.test", "sandi": "salah"})
        self.assertEqual(st, 401)

    def test_operator_tidak_bisa_membaca_klien_lain(self):
        op = self.masuk("op@uji.test")
        _, saya = op.req("/api/saya")
        self.assertEqual([k["id"] for k in saya["klien"]], ["kA"])
        self.assertEqual(op.req("/api/periode?klien=kB&ym=2026-10")[0], 404)
        self.assertEqual(op.req("/api/ekspor/xlsx?klien=kB&ym=2026-10")[0], 404)
        self.assertEqual(op.req("/api/tempel", {"klien": "kB", "ym": "2026-10", "teks": "https://x.id/a"})[0], 404)
        self.assertEqual(op.req("/api/target", {"klien": "kA", "ym": "2026-10", "targets": {}})[0], 403)

    def test_data_tidak_bisa_diunduh_langsung(self):
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{self.port}/data/pengguna.json")
            self.fail("data/ seharusnya tidak tersaji")
        except urllib.error.HTTPError as e:
            self.assertEqual(e.code, 404)

    def test_alur_tempel_simpan_ekspor(self):
        op = self.masuk("op@uji.test")
        st, j = op.req("/api/tempel", {"klien": "kA", "ym": "2026-10",
                                       "teks": "https://www.instagram.com/reel/DbfGWf-TjoP/\nhttps://vt.tiktok.com/ZS4P9uf1H/\nbukan link"})
        self.assertEqual((st, j["baru"], j["dilewati"]), (200, 2, 1))
        _, d = op.req("/api/periode?klien=kA&ym=2026-10")
        pid = d["posts"][0]["id"]
        st, j = op.req("/api/posting/simpan", {"klien": "kA", "ym": "2026-10", "ubah": [{"id": pid, "isi": {"views": "1.500", "like": "30"}}]})
        self.assertEqual((st, j["posts"][0]["views"], j["ringkasan"]["total"]["interaksi"]), (200, 1500, 30))
        admin = self.masuk("admin@uji.test")
        st, _ = admin.req("/api/target", {"klien": "kA", "ym": "2026-10", "targets": {"posts": {"ig": "10"}, "er": "3"}})
        self.assertEqual(st, 200)
        _, d = admin.req("/api/periode?klien=kA&ym=2026-10")
        self.assertEqual((d["targets"]["posts"]["ig"], d["targets"]["er"]), (10, 0.03))
        for jenis in ("xlsx", "pptx"):
            st, b = admin.req(f"/api/ekspor/{jenis}?klien=kA&ym=2026-10")
            self.assertEqual(st, 200)
            self.assertEqual(b[:2], b"PK")


if __name__ == "__main__":
    unittest.main()

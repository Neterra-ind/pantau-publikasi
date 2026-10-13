"""Uji aturan bisnis (scripts/inti.py) + regresi sampel Agustus 2026.

    python -m unittest discover -s tests -v

Uji regresi memakai workbook asli bila ada (lihat WORKBOOK di bawah, atau variabel PANTAU_SAMPEL); bila tidak ada, dilewati.
"""
import itertools
import os
import sys
import unittest
from datetime import date

AKAR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(AKAR, "scripts"))
import inti  # noqa: E402

HARI = date(2026, 10, 13)
WORKBOOK = os.environ.get("PANTAU_SAMPEL") or os.path.join(
    os.path.dirname(AKAR), "Dashboard Pekerjaan", "5. Agustus 2026 - Monitoring Konten DPD PDIP Jateng .xlsx")


def post(**kw):
    p = {"id": kw.pop("id", "x"), "platform": "ig", "tanggal": "2026-08-10", "kelompok": "A", "akun": "a", "jenis": "Video",
         "judul": "", "link": "https://www.instagram.com/reel/DbfGWf-TjoP/", "views": 100, "like": 5, "comment": 1,
         "repost": 0, "share": 0}
    p.update(kw)
    return p


class UjiLink(unittest.TestCase):
    def test_tanggal_instagram_dari_kode(self):
        self.assertEqual(inti.ig_date("DbfGWf-TjoP", HARI), "2026-08-01")

    def test_tanggal_tiktok_dari_id(self):
        vid = str((1786000000 << 32) + 123)  # ID digeser 32 bit ke kanan = detik Unix
        info = inti.parse_link(f"https://www.tiktok.com/@akunku/video/{vid}", HARI)
        self.assertEqual(info["platform"], "tt")
        self.assertEqual(info["akun"], "akunku")
        self.assertEqual(info["tanggal"], inti.ymd_wib(1786000000 * 1000, HARI))
        self.assertEqual(info["kunci"], "tt:" + vid)

    def test_link_pendek_tiktok_tanpa_tanggal(self):
        info = inti.parse_link("https://vt.tiktok.com/ZS4P9uf1H/", HARI)
        self.assertEqual((info["platform"], info["tanggal"]), ("tt", ""))

    def test_dedupe_abaikan_parameter_pelacakan(self):
        a = inti.parse_link("https://www.instagram.com/reel/DbfGWf-TjoP/?utm_source=ig_web_copy_link&igsh=MzRl")
        b = inti.parse_link("instagram.com/reel/DbfGWf-TjoP")
        self.assertEqual(a["kunci"], b["kunci"])
        y1 = inti.parse_link("https://youtube.com/shorts/gtVu-KJu9Jo?si=OWYu7r")
        y2 = inti.parse_link("https://www.youtube.com/watch?v=gtVu-KJu9Jo&t=3")
        self.assertEqual(y1["kunci"], y2["kunci"])
        w1 = inti.parse_link("https://www.bacaaja.co/judul-berita/?utm=x")
        w2 = inti.parse_link("https://bacaaja.co/judul-berita")
        self.assertEqual((w1["kunci"], w1["akun"]), (w2["kunci"], "bacaaja.co"))

    def test_platform_dari_domain(self):
        self.assertEqual(inti.parse_link("fb.watch/abc/")["platform"], "fb")
        self.assertEqual(inti.parse_link("https://youtu.be/xyz")["kunci"], "yt:xyz")
        self.assertIsNone(inti.parse_link("bukan link"))


class UjiAngkaNama(unittest.TestCase):
    def test_angka_indonesia(self):
        for s, n in [("1.234", 1234), ("22.100", 22100), ("3,7k", 3700), ("12rb", 12000), ("1,2 jt", 1_200_000), ("", None), ("-", None)]:
            self.assertEqual(inti.baca_angka(s), n, s)

    def test_name_key(self):
        self.assertEqual(inti.name_key("Kab.Kendal"), inti.name_key("kab.  kendal"))

    def test_rapikan_pilih_paling_sering_dan_alias_lebih_dulu(self):
        posts = [post(id=str(i), kelompok=k) for i, k in enumerate(["Kab. Kendal", "Kab.Kendal", "kab. Kendal", "Kab. Kendal", "Kota Solo"])]
        r = inti.rencana_rapikan(posts, "kelompok", {"kota solo": "Kota Surakarta"})
        peta = r["peta"]
        self.assertEqual(peta["Kab.Kendal"], "Kab. Kendal")
        self.assertEqual(peta["kab. Kendal"], "Kab. Kendal")
        self.assertEqual(peta["Kota Solo"], "Kota Surakarta")
        self.assertEqual(r["sel"], 3)

    def test_seri_menang_kapital(self):
        r = inti.rencana_rapikan([post(id="1", kelompok="kab. demak"), post(id="2", kelompok="Kab. demak")], "kelompok")
        self.assertEqual(r["peta"], {"kab. demak": "Kab. demak"})

    def test_jenis(self):
        self.assertEqual(inti.jenis_dari("Reels", "ig"), "Video")
        self.assertEqual(inti.jenis_dari("Carousel foto", "ig"), "Flyer")
        self.assertEqual(inti.jenis_dari("", "web"), "Artikel")
        self.assertEqual(inti.jenis_dari("", "tt"), "Video")


class UjiStatusRingkasan(unittest.TestCase):
    def test_periode_lewat(self):
        self.assertEqual(inti.status(100, 100, None), "tercapai")
        self.assertEqual(inti.status(80, 100, None), "hampir")
        self.assertEqual(inti.status(79, 100, None), "tertinggal")
        self.assertIsNone(inti.status(5, None, None))

    def test_bulan_berjalan_pakai_laju(self):
        porsi = inti.porsi_hari("2026-10", HARI)  # 13/31
        self.assertAlmostEqual(porsi, 13 / 31)
        self.assertEqual(inti.status(40, 100, porsi), "sesuai")       # 0,40/0,419 = 95%
        self.assertEqual(inti.status(32, 100, porsi), "dikejar")      # 76%
        self.assertEqual(inti.status(20, 100, porsi), "tertinggal")   # 48%
        self.assertEqual(inti.status(100, 100, porsi), "tercapai")

    def test_er_minimum_tanpa_laju(self):
        self.assertEqual(inti.status(0.025, 0.03, 0.4, minimum=True), "hampir")

    def test_er_agregat_bukan_rata_rata_dan_berita_tanpa_views(self):
        posts = [post(id="1", views=1000, like=100), post(id="2", views=9000, like=90, link="https://instagram.com/p/AAAAAAAAAAA/"),
                 post(id="3", platform="web", views=None, like=None, comment=None, link="https://x.id/a")]
        r = inti.ringkas(posts, "2026-08", hari_ini=HARI)
        self.assertEqual(r["total"]["views"], 10000)
        self.assertAlmostEqual(r["total"]["er"], (101 + 91) / 10000)  # bukan rata-rata (0,101 + 0,0101)/2
        self.assertEqual(r["total"]["publikasi"], 3)

    def test_delapan_flag(self):
        posts = [post(id="1"), post(id="2"), post(id="3", link="tidak valid", tanggal="", views=None, akun="", kelompok=""),
                 post(id="4", link="https://instagram.com/p/BBBBBBBBBBB/", tanggal="2026-09-01", views=10, like=20)]
        f = inti.tandai(posts, "2026-08")
        self.assertEqual(f["1"], ["ganda"])
        self.assertEqual(set(f["3"]), {"link", "tgl", "views", "akun", "kelompok"})
        self.assertEqual(set(f["4"]), {"luar", "janggal"})

    def test_koreksi_tahun(self):
        self.assertEqual(inti.koreksi_tahun("2028-08-26", "2026-08"), ("2026-08-26", True))
        self.assertEqual(inti.koreksi_tahun("2026-07-31", "2026-08"), ("2026-07-31", False))

    def test_alias_gabungkan_dua_kelompok(self):
        posts = [post(id="1", kelompok="Kota Solo"), post(id="2", kelompok="Kota Surakarta", link="https://instagram.com/p/CCCCCCCCCCC/")]
        r = inti.rencana_rapikan(posts, "kelompok", {"kota solo": "Kota Surakarta"})
        for p in posts:
            p["kelompok"] = r["peta"].get(p["kelompok"], p["kelompok"])
        ring = inti.ringkas(posts, "2026-08", hari_ini=HARI)
        self.assertEqual([(b["nama"], b["total"]) for b in ring["kelompok"]], [("Kota Surakarta", 2)])


class UjiTempel(unittest.TestCase):
    def test_tempel_dengan_judul_dan_tanpa_judul(self):
        teks = "DPC\tTanggal\tNama Akun\tJenis\tJudul\tLink\tViews\tLike\nKab. Pati\t05/08/2026\tkabar\tVideo\tJudul\thttps://instagram.com/reel/DbfGWf-TjoP/\t3,7k\t1.234"
        rows, lewat, _ = inti.parse_tempel(teks)
        self.assertEqual((len(rows), lewat), (1, 0))
        self.assertEqual(rows[0]["views"], "3,7k")
        teks2 = "1\tKab. Pati\t2026-08-05\takun\tVideo\tJudul\thttps://instagram.com/reel/DbfGWf-TjoP/\t22.100\t10\t2\t0\t1\ntanpa link"
        rows, lewat, _ = inti.parse_tempel(teks2)
        self.assertEqual((rows[0]["kelompok"], rows[0]["views"], rows[0]["share"], lewat), ("Kab. Pati", "22.100", "1", 1))

    def test_gabung_memperbarui_bukan_menggandakan(self):
        posts = [inti.bersihkan_post(post(id="1"))]
        c = itertools.count(2)
        baru, upd = inti.gabung(posts, [{"platform": "ig", "link": "https://instagram.com/reel/DbfGWf-TjoP/?igsh=x", "views": "500"}],
                                lambda: str(next(c)))
        self.assertEqual((baru, upd, len(posts), posts[0]["views"]), (0, 1, 1, 500))

    def test_filter_dan_total_subset(self):
        posts = [post(id="1", platform="tt", kelompok="Kab. Pati", tanggal="2026-08-03", link="https://tiktok.com/@a/video/1"),
                 post(id="2", platform="tt", kelompok="kab.pati", tanggal="2026-08-20", link="https://tiktok.com/@a/video/2"),
                 post(id="3", platform="ig", kelompok="Kab. Pati")]
        rows = inti.saring(posts, {"platform": "tt", "kelompok": "KAB. PATI", "dari": "2026-08-01", "sampai": "2026-08-10"})
        self.assertEqual([p["id"] for p in rows], ["1"])
        self.assertEqual(sum(inti.interaksi(p) for p in rows), 6)


@unittest.skipUnless(os.path.exists(WORKBOOK), "workbook sampel Agustus 2026 tidak ditemukan")
class UjiRegresiSampel(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import impor_workbook
        rows, cls.lap = impor_workbook.baca_workbook(WORKBOOK, "2026-08")
        cls.posts = []
        c = itertools.count()
        inti.gabung(cls.posts, rows, lambda: "p%d" % next(c), pertahankan_ganda=True)
        cls.r = inti.ringkas(cls.posts, "2026-08", hari_ini=HARI)

    def test_jumlah_dan_angka_instagram(self):
        self.assertEqual(self.r["total"]["publikasi"], 774)
        ig = self.r["per"]["ig"]
        self.assertEqual((ig["posts"], ig["views"], ig["interaksi"]), (348, 5_228_263, 114_606))

    def test_rumus_bukan_angka_pdf(self):
        # Interaksi Facebook menurut rumus = 28.083 (PDF: 25.580); total = 159.149 (PDF: 156.524)
        self.assertEqual(self.r["per"]["fb"]["interaksi"], 28_083)
        self.assertEqual(self.r["total"]["interaksi"], 159_149)

    def test_artikel_hanya_bulan_aktif_dan_tanggal_dikoreksi(self):
        self.assertEqual(self.r["per"]["web"]["posts"], 14)
        self.assertEqual(len(self.lap["dikoreksi"]), 2)

    def test_impor_ulang_tidak_menggandakan(self):
        import impor_workbook
        rows, _ = impor_workbook.baca_workbook(WORKBOOK, "2026-08")
        posts = [dict(p) for p in self.posts]
        baru, upd = inti.gabung(posts, rows, lambda: "q", pertahankan_ganda=True)
        self.assertEqual((baru, upd, len(posts)), (0, 774, 774))


if __name__ == "__main__":
    unittest.main()

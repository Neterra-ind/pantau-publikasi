"""
Membuat data simulasi (nama akun & klien fiktif) dan akun awal admin + operator.

    python scripts/buat_simulasi.py            # hanya bila data/ masih kosong
    python scripts/buat_simulasi.py --ulang    # hapus data lama lalu buat ulang

Kata sandi akun awal ditulis ke data/akun_awal.txt (ganti setelah masuk).
"""
import os
import random
import secrets
import shutil
import sys
from datetime import date, datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import inti  # noqa: E402
import penyimpanan as ps  # noqa: E402

B64 = inti.B64
WILAYAH = ["Kab. Semarang", "Kota Semarang", "Kab. Kendal", "Kab. Demak", "Kab. Kudus", "Kab. Pati", "Kab. Jepara",
           "Kab. Grobogan", "Kab. Boyolali", "Kab. Klaten", "Kota Surakarta", "Kab. Sukoharjo", "Kab. Wonogiri",
           "Kab. Banyumas", "Kab. Cilacap", "Kab. Tegal", "Kota Tegal", "Kab. Brebes", "Kab. Magelang", "Kota Salatiga"]
AKUN = {"ig": ["suarawarga.id", "kabarbaik_jateng", "jelajahdesa", "ruangsapa", "warta_nusa", "lensa.kita"],
        "tt": ["ceritakampung", "jatengbergerak", "obrolanwarga", "kilasdaerah"],
        "fb": ["Kabar Warga Jateng", "Info Pesisir Utara", "Seputar Karesidenan"],
        "yt": ["SuaraWargaTV", "JelajahDesaChannel"],
        "web": ["beritasimulasi.id", "kabarcontoh.com", "portalwarga.net"]}
TEMA = ["Musyawarah warga membahas program air bersih", "Kerja bakti bersama relawan", "Pelatihan UMKM digital",
        "Bakti sosial dan cek kesehatan gratis", "Dialog kebangsaan bersama pemuda", "Lomba kebersihan lingkungan",
        "Penanaman seribu pohon", "Festival budaya desa", "Sosialisasi pencegahan banjir", "Jalan sehat bersama warga"]


def ig_kode(d):
    ms = int(datetime(d.year, d.month, d.day, 10, 0, tzinfo=inti.WIB).timestamp() * 1000) + random.randint(0, 3_600_000 * 8)
    n = (ms - 1314220021721) << 23 | random.getrandbits(23)
    s = ""
    while n:
        s = B64[n % 64] + s
        n //= 64
    return s[:11]


def tt_id(d):
    detik = int(datetime(d.year, d.month, d.day, 12, tzinfo=inti.WIB).timestamp()) + random.randint(0, 30000)
    return str((detik << 32) | random.getrandbits(32))


def buat_post(pl, d, wilayah):
    akun = random.choice(AKUN[pl])
    tema = random.choice(TEMA)
    judul = f"{tema} di {wilayah}"
    if pl == "ig":
        link = f"https://www.instagram.com/reel/{ig_kode(d)}/?utm_source=ig_web_copy_link"
    elif pl == "tt":
        link = f"https://www.tiktok.com/@{akun}/video/{tt_id(d)}"
    elif pl == "fb":
        link = f"https://www.facebook.com/reel/{random.randint(10**14, 10**15)}"
    elif pl == "yt":
        link = "https://youtube.com/shorts/" + "".join(random.choice(B64) for _ in range(11))
    else:
        slug = judul.lower().replace(" ", "-").replace(".", "")
        link = f"https://{akun}/{d.isoformat()}/{slug}-{random.randint(100, 999)}"
    p = {"id": ps.id_baru("p"), "platform": pl, "tanggal": d.isoformat(), "kelompok": wilayah, "akun": akun,
         "jenis": "Artikel" if pl == "web" else ("Flyer" if random.random() < 0.12 else "Video"), "judul": judul, "link": link}
    if pl != "web":
        dasar = {"ig": 15000, "tt": 2500, "fb": 4000, "yt": 600}[pl]
        views = int(random.lognormvariate(0, 0.9) * dasar)
        er = random.uniform(0.008, 0.06)
        it = int(views * er)
        like = int(it * 0.88)
        comment = int(it * 0.08)
        share = int(it * 0.03)
        p.update({"views": views, "like": like, "comment": comment, "repost": it - like - comment - share if pl != "fb" else 0,
                  "share": share})
    return inti.bersihkan_post(p)


def buat_periode(ym, sampai, jumlah, rusak=True):
    y, m = int(ym[:4]), int(ym[5:7])
    awal = date(y, m, 1)
    hari = (sampai - awal).days + 1
    bobot = {"ig": 0.42, "tt": 0.3, "fb": 0.16, "yt": 0.04, "web": 0.08}
    posts = []
    for _ in range(jumlah):
        pl = random.choices(list(bobot), list(bobot.values()))[0]
        d = awal + timedelta(days=random.randrange(hari))
        posts.append(buat_post(pl, d, random.choice(WILAYAH)))
    if rusak:  # masalah data yang sengaja dibuat agar fitur "Perlu dicek" & "Rapikan nama" terlihat
        sos = [p for p in posts if p["platform"] in ("ig", "tt")]
        dup = dict(sos[0], id=ps.id_baru("p"))
        dup["link"] += "&igsh=abc" if "?" in dup["link"] else "?is_from_webapp=1"
        posts.append(dup)
        sos[1]["views"] = None
        sos[2]["akun"] = ""
        sos[3]["kelompok"] = ""
        sos[4]["tanggal"] = f"{y + 2}-{m:02d}-{min(sampai.day, 20):02d}"
        sos[5]["like"] = (sos[5]["views"] or 100) * 2
        for p in posts[10:16]:
            p["kelompok"] = random.choice(["kab. Kendal", "Kab.Kendal", "kab.  Kendal"])
        posts[20]["kelompok"] = "Kota Solo"  # perlu aturan alias "Kota Solo = Kota Surakarta"
    return posts


def main():
    ulang = "--ulang" in sys.argv
    if ps.semua_pengguna() and not ulang:
        print("Data sudah ada. Pakai --ulang untuk menghapus dan membuat ulang.")
        return
    if ulang and os.path.isdir(ps.DATA):
        shutil.rmtree(ps.DATA)
    random.seed(2026)
    hari_ini = datetime.now(inti.WIB).date()

    k1 = {"id": "k-contoh", "nama": "Klien Contoh (simulasi)", "label_kelompok": "DPC", "pakai_kelompok": True,
          "warna": "#B3261E", "alias": {"kelompok": {}, "akun": {}},
          "targets": {"posts": {"ig": 160, "tt": 110, "fb": 60, "yt": 12, "web": 25}, "views": 2500000,
                      "interactions": 70000, "er": 0.03, "erPlatform": {"ig": 0.025, "tt": 0.03},
                      "perGroup": {w: 14 for w in WILAYAH[:8]}, "perAccount": {"suarawarga.id": 30, "ceritakampung": 25}}}
    k2 = {"id": "k-kopi", "nama": "Komunitas Kopi Nusantara (simulasi)", "label_kelompok": "Kota", "pakai_kelompok": True,
          "warna": "#6D4C41", "alias": {"kelompok": {}, "akun": {}},
          "targets": {"posts": {"ig": 60, "tt": 40}, "views": 600000, "er": 0.035}}
    ps.simpan_semua_klien([k1, k2])

    bulan = []
    d = date(hari_ini.year, hari_ini.month, 1)
    for _ in range(3):
        bulan.insert(0, d)
        d = (d - timedelta(days=1)).replace(day=1)
    for i, b in enumerate(bulan):
        ym = b.strftime("%Y-%m")
        akhir = (b.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
        sampai = min(akhir, hari_ini)
        n = int(370 * (sampai.day / akhir.day)) if sampai < akhir else 330 + i * 40
        ps.simpan_periode(k1["id"], {"ym": ym, "targets": None, "posts": buat_periode(ym, sampai, n)})
        posts2 = [p for p in buat_periode(ym, sampai, n // 3, rusak=False) if p["platform"] in ("ig", "tt", "fb")]
        for p in posts2:
            p["kelompok"] = random.choice(["Bandung", "Yogyakarta", "Malang", "Medan", "Makassar"])
        ps.simpan_periode(k2["id"], {"ym": ym, "targets": None, "posts": posts2})

    sandi = {"admin@pantau.test": secrets.token_urlsafe(9), "operator@pantau.test": secrets.token_urlsafe(9)}
    pengguna = []
    for email, peran, nama, klien in (("admin@pantau.test", "admin", "Admin Vendor", [k1["id"], k2["id"]]),
                                      ("operator@pantau.test", "operator", "Operator Satu", [k1["id"]])):
        u = {"email": email, "nama": nama, "peran": peran, "klien": klien}
        ps.set_sandi(u, sandi[email])
        pengguna.append(u)
    ps.simpan_semua_pengguna(pengguna)
    with open(os.path.join(ps.DATA, "akun_awal.txt"), "w", encoding="utf-8") as f:
        f.write("Akun awal Pantau Publikasi (data simulasi). Ganti kata sandi setelah masuk (menu Pengaturan).\n\n")
        for e, s in sandi.items():
            f.write(f"{e}\t{s}\n")
    print("Data simulasi dibuat:", ", ".join(b.strftime("%Y-%m") for b in bulan))
    print("Akun awal ditulis ke data/akun_awal.txt")


if __name__ == "__main__":
    main()

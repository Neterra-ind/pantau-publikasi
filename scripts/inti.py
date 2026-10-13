"""
Paket inti Pantau Publikasi: semua aturan bisnis sebagai fungsi murni (tanpa file, server, atau tampilan).

Dipakai bersama oleh layar (lewat server), ekspor Excel, dan ekspor PPTX, sehingga angkanya selalu sama.
Diuji oleh tests/test_inti.py. Bila ada perbedaan dengan perilaku lain, PRD yang berlaku.
"""
import calendar
import re
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta, timezone
from urllib.parse import parse_qs, urlparse

PLATFORM = ["ig", "tt", "fb", "yt", "web"]
NAMA_PLATFORM = {"ig": "Instagram", "tt": "TikTok", "fb": "Facebook", "yt": "YouTube", "web": "Berita"}
SOSMED = ["ig", "tt", "fb", "yt"]
JENIS = ["Video", "Flyer", "Artikel"]
METRIK = ["views", "like", "comment", "repost", "share"]
INTERAKSI = ["like", "comment", "repost", "share"]
WIB = timezone(timedelta(hours=7))
BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
         "September", "Oktober", "November", "Desember"]

# Ambang status (usulan awal PRD; fase 2: diatur per klien)
AMBANG_HAMPIR = 0.8
AMBANG_LAJU = 0.9
AMBANG_KEJAR = 0.7

STATUS_LABEL = {"tercapai": "Tercapai", "hampir": "Hampir", "tertinggal": "Tertinggal",
                "sesuai": "Sesuai laju", "dikejar": "Perlu dikejar", "belum": "Belum mulai"}

FLAG = {
    "link": ("Link tidak valid", "Link tidak bisa dibaca atau bukan dari platform baris ini."),
    "ganda": ("Link ganda", "Link yang sama dipakai baris lain di platform yang sama; postingan dan views bisa terhitung dua kali."),
    "luar": ("Tanggal di luar periode", "Tanggal posting tidak berada di bulan periode ini."),
    "tgl": ("Tanggal kosong", "Tanggal posting belum diisi."),
    "views": ("Views kosong", "Angka views belum diisi."),
    "akun": ("Akun kosong", "Nama akun belum diisi."),
    "kelompok": ("Kelompok kosong", "Kelompok belum diisi."),
    "janggal": ("Metrik janggal", "Interaksi lebih besar dari views, atau ER di atas 50%."),
}


# ------------------------------------------------------------------ angka & tanggal
def baca_angka(v):
    """'22.100' -> 22100, '3,7k' -> 3700, '1,2 jt' -> 1200000. Kosong/tidak terbaca -> None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(round(v))
    s = str(v).strip().lower().replace(" ", "").replace(" ", "")
    if not s or s in ("-", "—"):
        return None
    kali = 1
    for akhiran, k in (("jt", 1_000_000), ("m", 1_000_000), ("rb", 1_000), ("k", 1_000)):
        if s.endswith(akhiran):
            s, kali = s[: -len(akhiran)], k
            break
    if kali > 1:
        s = s.replace(",", ".")
        if s.count(".") > 1:
            s = s.replace(".", "")
    else:
        s = s.replace(".", "").replace(",", ".")
    try:
        return int(round(float(s) * kali))
    except ValueError:
        return None


def ymd_wib(ms, hari_ini=None):
    """Milidetik Unix -> 'YYYY-MM-DD' WIB. Tanggal di luar 2015..besok ditolak ('')."""
    try:
        d = datetime.fromtimestamp(ms / 1000, WIB).date()
    except (OverflowError, OSError, ValueError):
        return ""
    besok = (hari_ini or datetime.now(WIB).date()) + timedelta(days=1)
    if d < date(2015, 1, 1) or d > besok:
        return ""
    return d.isoformat()


B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"


def ig_date(kode, hari_ini=None):
    if len(kode) < 10:
        return ""
    n = 0
    for ch in kode[:11]:
        k = B64.find(ch)
        if k < 0:
            return ""
        n = n * 64 + k
    return ymd_wib((n >> 23) + 1314220021721, hari_ini)


def tt_date(vid, hari_ini=None):
    if not vid.isdigit():
        return ""
    return ymd_wib((int(vid) >> 32) * 1000, hari_ini)


def baca_tanggal(v):
    """Tanggal dari sel Excel / teks tempel -> 'YYYY-MM-DD' atau ''."""
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, (int, float)):  # nomor seri Excel
        if 30000 < v < 80000:
            return (date(1899, 12, 30) + timedelta(days=int(v))).isoformat()
        return ""
    s = str(v).strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$", s)
        if m:
            d, mo, y = map(int, m.groups())
            y += 2000 if y < 100 else 0
        else:
            m = re.match(r"^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$", s)
            if not m:
                return ""
            nama = m.group(2).lower()[:3]
            cocok = [i for i, b in enumerate(BULAN) if b.lower()[:3] == nama or (nama == "agt" and i == 7)]
            if not cocok:
                return ""
            d, mo, y = int(m.group(1)), cocok[0] + 1, int(m.group(3))
    try:
        return date(y, mo, d).isoformat()
    except ValueError:
        return ""


def koreksi_tahun(tgl, ym):
    """Tanggal salah tahun pada bulan yang sama (2028-08-26 di periode 2026-08) -> dikoreksi. Hasil: (tgl, dikoreksi?)."""
    if tgl and tgl[:7] != ym and tgl[5:7] == ym[5:7]:
        try:
            return date(int(ym[:4]), int(tgl[5:7]), int(tgl[8:10])).isoformat(), True
        except ValueError:
            pass
    return tgl, False


def tgl_panjang(t):
    """'2026-10-13' -> '13 Oktober 2026'."""
    if not t:
        return ""
    y, m, d = t[:10].split("-")
    return f"{int(d)} {BULAN[int(m) - 1]} {y}"


def nama_periode(ym):
    return f"{BULAN[int(ym[5:7]) - 1]} {ym[:4]}"


def periode_sebelum(ym):
    y, m = int(ym[:4]), int(ym[5:7])
    y, m = (y - 1, 12) if m == 1 else (y, m - 1)
    return f"{y:04d}-{m:02d}"


def ribu(n):
    """Format id-ID: 1234567 -> '1.234.567'."""
    if n is None:
        return "–"
    return f"{int(round(n)):,}".replace(",", ".")


def persen(x, digit=2):
    if x is None:
        return "–"
    return f"{x * 100:.{digit}f}".replace(".", ",") + "%"


# ------------------------------------------------------------------ link
def _host(u):
    h = (u.hostname or "").lower()
    for awal in ("www.", "m.", "mobile.", "web."):
        if h.startswith(awal):
            h = h[len(awal):]
    return h


def parse_link(url, hari_ini=None):
    """Link -> {platform, akun, tanggal, kunci, url} atau None bila tidak valid."""
    s = str(url or "").strip()
    if not s or " " in s:
        return None
    if not re.match(r"^https?://", s, re.I):
        s = "https://" + s
    try:
        u = urlparse(s)
    except ValueError:
        return None
    host = _host(u)
    if "." not in host or not re.fullmatch(r"[a-z0-9.\-]+", host):
        return None
    bagian = [p for p in u.path.split("/") if p]
    q = parse_qs(u.query)
    hasil = {"platform": "web", "akun": "", "tanggal": "", "kunci": "", "url": s}

    if host == "instagram.com" or host.endswith(".instagram.com") or host == "instagr.am":
        hasil["platform"] = "ig"
        kode = ""
        for i, p in enumerate(bagian):
            if p.lower() in ("p", "reel", "reels", "tv") and i + 1 < len(bagian):
                kode = bagian[i + 1]
                if i == 1:
                    hasil["akun"] = bagian[0]
                break
        if kode:
            hasil["kunci"] = "ig:" + kode
            hasil["tanggal"] = ig_date(kode, hari_ini)
        else:
            hasil["kunci"] = "ig:" + "/".join(bagian).lower()
            if len(bagian) == 1:
                hasil["akun"] = bagian[0]
    elif host == "tiktok.com" or host.endswith(".tiktok.com"):
        hasil["platform"] = "tt"
        if bagian and bagian[0].startswith("@"):
            hasil["akun"] = bagian[0][1:]
        vid = ""
        for i, p in enumerate(bagian):
            if p in ("video", "photo") and i + 1 < len(bagian):
                vid = bagian[i + 1]
        if vid.isdigit():
            hasil["kunci"] = "tt:" + vid
            hasil["tanggal"] = tt_date(vid, hari_ini)
        else:
            hasil["kunci"] = "tt:" + host + "/" + "/".join(bagian)  # link pendek: kode peka huruf
    elif host in ("facebook.com", "fb.watch", "fb.com") or host.endswith(".facebook.com"):
        hasil["platform"] = "fb"
        jalur = "/".join(bagian).lower()
        tambahan = ""
        for k in ("v", "story_fbid", "fbid", "id"):
            if q.get(k):
                tambahan += f"?{k}={q[k][0]}"
        hasil["kunci"] = "fb:" + ("fb.watch/" if host == "fb.watch" else "") + jalur + tambahan
    elif host in ("youtube.com", "youtu.be", "music.youtube.com") or host.endswith(".youtube.com"):
        hasil["platform"] = "yt"
        if bagian and bagian[0].startswith("@"):
            hasil["akun"] = bagian[0][1:]
        vid = ""
        if q.get("v"):
            vid = q["v"][0]
        elif host == "youtu.be" and bagian:
            vid = bagian[0]
        else:
            for i, p in enumerate(bagian):
                if p in ("shorts", "live", "embed", "v") and i + 1 < len(bagian):
                    vid = bagian[i + 1]
        hasil["kunci"] = "yt:" + (vid or "/".join(bagian).lower())
    else:
        hasil["akun"] = host
        hasil["kunci"] = "web:" + host + ("/" + "/".join(bagian) if bagian else "")
    return hasil


# ------------------------------------------------------------------ nama & jenis
def name_key(s):
    """Kunci pembanding nama: huruf kecil, titik diberi spasi sesudahnya, spasi ganda dipadatkan."""
    s = str(s or "").lower().replace(".", ". ")
    return re.sub(r"\s+", " ", s).strip()


def jenis_dari(teks, platform):
    t = str(teks or "").strip().lower()
    if t.startswith(("vid", "reel", "short")):
        return "Video"
    if any(k in t for k in ("fly", "poster", "foto", "gambar", "carousel")):
        return "Flyer"
    if any(k in t for k in ("art", "berita")):
        return "Artikel"
    return "Artikel" if platform == "web" else "Video"


def terapkan_alias(nilai, alias):
    if not nilai:
        return nilai
    return alias.get(name_key(nilai), nilai)


def rencana_rapikan(posts, kolom, alias=None):
    """Rencana 'Rapikan nama' untuk kolom 'kelompok' atau 'akun'.

    Aturan alias (nama lama = baku) dijalankan dulu; sisanya disatukan per name_key ke penulisan
    paling sering (seri: yang berawalan kapital, lalu abjad). Hasil: {ubah: [{dari, ke, jumlah}], sel, peta}.
    """
    alias = alias or {}
    nilai = [str(p.get(kolom) or "").strip() for p in posts]
    setelah_alias = [terapkan_alias(v, alias) for v in nilai]
    per_kunci = defaultdict(Counter)
    for v in setelah_alias:
        if v:
            per_kunci[name_key(v)][v] += 1
    baku = {}
    for k, c in per_kunci.items():
        baku[k] = sorted(c.items(), key=lambda x: (-x[1], not x[0][:1].isupper(), x[0]))[0][0]
    peta = {}
    for asli, a in zip(nilai, setelah_alias):
        if not asli:
            continue
        ke = baku[name_key(a)]
        if ke != asli:
            peta[asli] = ke
    hitung = Counter(v for v in nilai if v in peta)
    ubah = sorted(({"dari": d, "ke": peta[d], "jumlah": n} for d, n in hitung.items()), key=lambda x: (x["ke"], x["dari"]))
    return {"ubah": ubah, "sel": sum(hitung.values()), "peta": peta}


def nama_tampil(posts, kolom):
    """name_key -> penulisan baku (paling sering), dipakai untuk mengelompokkan di layar dan di Excel."""
    per_kunci = defaultdict(Counter)
    for p in posts:
        v = str(p.get(kolom) or "").strip()
        if v:
            per_kunci[name_key(v)][v] += 1
    return {k: sorted(c.items(), key=lambda x: (-x[1], not x[0][:1].isupper(), x[0]))[0][0] for k, c in per_kunci.items()}


# ------------------------------------------------------------------ metrik per baris
def interaksi(p):
    if p.get("platform") == "web":
        return 0
    return sum(int(p.get(k) or 0) for k in INTERAKSI)


def er_baris(p):
    v = p.get("views")
    if p.get("platform") == "web" or not v:
        return None
    return interaksi(p) / v


def bersihkan_post(p):
    """Normalisasi satu postingan sebelum disimpan (berita tidak punya metrik)."""
    p["platform"] = p.get("platform") if p.get("platform") in PLATFORM else "web"
    for k in METRIK:
        p[k] = None if p["platform"] == "web" else baca_angka(p.get(k))
        if p[k] is not None and p[k] < 0:
            p[k] = None
    for k in ("kelompok", "akun", "judul", "link", "catatan"):
        p[k] = re.sub(r"\s+", " ", str(p.get(k) or "")).strip() if k != "judul" else str(p.get(k) or "").strip()
    p["tanggal"] = baca_tanggal(p.get("tanggal"))
    if p.get("jenis") not in JENIS:
        p["jenis"] = jenis_dari(p.get("jenis"), p["platform"])
    return p


# ------------------------------------------------------------------ validasi (8 flag)
def tandai(posts, ym, pakai_kelompok=True):
    """id -> [kode flag]. Hanya baris bermasalah yang muncul."""
    kunci = {}
    hitung = Counter()
    for p in posts:
        info = parse_link(p.get("link"))
        k = (p["platform"], info["kunci"]) if info else None
        kunci[p["id"]] = (info, k)
        if k:
            hitung[k] += 1
    hasil = {}
    for p in posts:
        f = []
        info, k = kunci[p["id"]]
        if not info or (p["platform"] in SOSMED and info["platform"] != p["platform"]) or (p["platform"] == "web" and info["platform"] != "web"):
            f.append("link")
        if k and hitung[k] > 1:
            f.append("ganda")
        if not p.get("tanggal"):
            f.append("tgl")
        elif p["tanggal"][:7] != ym:
            f.append("luar")
        if p["platform"] != "web" and p.get("views") is None:
            f.append("views")
        if not p.get("akun"):
            f.append("akun")
        if pakai_kelompok and not p.get("kelompok"):
            f.append("kelompok")
        if p["platform"] != "web" and p.get("views"):
            it = interaksi(p)
            if it > p["views"] or it / p["views"] > 0.5:
                f.append("janggal")
        if f:
            hasil[p["id"]] = f
    return hasil


# ------------------------------------------------------------------ status
def porsi_hari(ym, hari_ini):
    """Porsi hari berlalu: None = periode sudah lewat, 0 = belum mulai, selain itu tanggal ÷ jumlah hari."""
    y, m = int(ym[:4]), int(ym[5:7])
    akhir = date(y, m, calendar.monthrange(y, m)[1])
    if hari_ini > akhir:
        return None
    if hari_ini < date(y, m, 1):
        return 0.0
    return hari_ini.day / akhir.day


def status(realisasi, target, porsi, minimum=False):
    """Status capaian. porsi None = periode lewat. minimum=True untuk target ER (tanpa laju)."""
    if target in (None, "", 0) or realisasi is None:
        return None
    cap = realisasi / target
    if porsi is None or minimum:
        if cap >= 1:
            return "tercapai"
        return "hampir" if cap >= AMBANG_HAMPIR else "tertinggal"
    if cap >= 1:
        return "tercapai"
    if porsi == 0:
        return "belum"
    laju = cap / porsi
    if laju >= AMBANG_LAJU:
        return "sesuai"
    return "dikejar" if laju >= AMBANG_KEJAR else "tertinggal"


def target_efektif(klien, periode):
    """Target periode sendiri, atau bawaan klien bila periode tidak punya target."""
    if periode.get("targets"):
        return periode["targets"], "periode"
    return klien.get("targets") or {}, "bawaan"


def _num(v):
    if v in (None, ""):
        return None
    try:
        v = float(v)
    except (TypeError, ValueError):
        return None
    return v if v > 0 else None


# ------------------------------------------------------------------ ringkasan
def _agregat(posts):
    per = {}
    for pl in PLATFORM:
        rows = [p for p in posts if p["platform"] == pl]
        views = sum(int(p.get("views") or 0) for p in rows) if pl != "web" else 0
        it = sum(interaksi(p) for p in rows)
        jenis = Counter(p.get("jenis") for p in rows)
        per[pl] = {
            "posts": len(rows), "video": jenis["Video"], "flyer": jenis["Flyer"], "artikel": jenis["Artikel"],
            "views": views, "interaksi": it,
            "like": sum(int(p.get("like") or 0) for p in rows), "comment": sum(int(p.get("comment") or 0) for p in rows),
            "repost": sum(int(p.get("repost") or 0) for p in rows), "share": sum(int(p.get("share") or 0) for p in rows),
            "er": (it / views) if (pl != "web" and views) else None,
        }
    sos = [per[pl] for pl in SOSMED]
    views = sum(x["views"] for x in sos)
    it = sum(x["interaksi"] for x in sos)
    total = {"publikasi": len(posts), "views": views, "interaksi": it, "er": (it / views) if views else None,
             "video": sum(x["video"] for x in per.values()), "flyer": sum(x["flyer"] for x in per.values()),
             "artikel": sum(x["artikel"] for x in per.values())}
    return per, total


def ringkas(posts, ym, targets=None, posts_lalu=None, hari_ini=None, label_kelompok="Kelompok", pakai_kelompok=True, flags=None):
    """Satu fungsi ringkasan untuk layar, Excel, dan PPTX."""
    hari_ini = hari_ini or datetime.now(WIB).date()
    targets = targets or {}
    porsi = porsi_hari(ym, hari_ini)
    per, total = _agregat(posts)
    lalu = None
    if posts_lalu:
        per_l, total_l = _agregat(posts_lalu)
        lalu = {"per": per_l, "total": total_l}

    # ---- indikator target
    tp = targets.get("posts") or {}
    tep = targets.get("erPlatform") or {}
    ind = []

    def tambah(kunci, label, real, target, prev, minimum=False, turunan=False, jenis="angka"):
        t = _num(target)
        if jenis == "er" and t is not None and t > 1:
            t = t / 100  # target ER boleh ditulis dalam persen
        st = status(real, t, porsi, minimum)
        ind.append({"kunci": kunci, "label": label, "realisasi": real, "target": t,
                    "capaian": (real / t) if (t and real is not None) else None, "status": st,
                    "lalu": prev, "turunan": turunan, "jenis": jenis})

    for pl in PLATFORM:
        tambah("posts." + pl, f"Postingan {NAMA_PLATFORM[pl]}", per[pl]["posts"], tp.get(pl),
               lalu["per"][pl]["posts"] if lalu else None)
    tot_t = [_num(tp.get(pl)) for pl in PLATFORM]
    tambah("publikasi", "Total publikasi", total["publikasi"], sum(t for t in tot_t if t) or None,
           lalu["total"]["publikasi"] if lalu else None, turunan=True)
    tambah("views", "Total views", total["views"], targets.get("views"), lalu["total"]["views"] if lalu else None)
    tambah("interaksi", "Total interaksi", total["interaksi"], targets.get("interactions"), lalu["total"]["interaksi"] if lalu else None)
    tambah("er", "ER total", total["er"], targets.get("er"), lalu["total"]["er"] if lalu else None, minimum=True, jenis="er")
    for pl in SOSMED:
        tambah("er." + pl, f"ER {NAMA_PLATFORM[pl]}", per[pl]["er"], tep.get(pl), lalu["per"][pl]["er"] if lalu else None,
               minimum=True, jenis="er")
    dinilai = [i for i in ind if i["status"] and not i["turunan"]]

    # ---- per kelompok & per akun
    def per_kunci(kolom, kuota):
        tampil = nama_tampil(posts, kolom)
        kuota_k = {name_key(k): _num(v) for k, v in (kuota or {}).items()}
        baris = {}
        for p in posts:
            k = name_key(p.get(kolom))
            if not k:
                continue
            b = baris.setdefault(k, {"nama": tampil[k], **{pl: 0 for pl in PLATFORM}, "total": 0, "views": 0,
                                     "interaksi": 0, "platform": Counter()})
            b[p["platform"]] += 1
            b["total"] += 1
            b["platform"][p["platform"]] += 1
            if p["platform"] != "web":
                b["views"] += int(p.get("views") or 0)
                b["interaksi"] += interaksi(p)
        for k, t in kuota_k.items():  # kelompok yang punya kuota tapi belum ada data
            if t and k not in baris:
                nama = next((n for n in (kuota or {}) if name_key(n) == k), k)
                baris[k] = {"nama": nama, **{pl: 0 for pl in PLATFORM}, "total": 0, "views": 0, "interaksi": 0, "platform": Counter()}
        hasil = []
        for k, b in baris.items():
            t = kuota_k.get(k)
            b["er"] = (b["interaksi"] / b["views"]) if b["views"] else None
            b["kuota"] = t
            b["capaian"] = (b["total"] / t) if t else None
            b["status"] = status(b["total"], t, porsi)
            b["platform_utama"] = b["platform"].most_common(1)[0][0] if b["platform"] else ""
            del b["platform"]
            hasil.append(b)
        hasil.sort(key=lambda b: b["nama"].lower())
        return hasil

    kelompok = per_kunci("kelompok", targets.get("perGroup")) if pakai_kelompok else []
    akun = per_kunci("akun", targets.get("perAccount"))
    kuota_kel = [b for b in kelompok if b["kuota"]]
    kuota_akun = [b for b in akun if b["kuota"]]

    # ---- laju kumulatif harian
    y, m = int(ym[:4]), int(ym[5:7])
    nhari = calendar.monthrange(y, m)[1]
    harian = {k: [0] * nhari for k in ("publikasi", "views", "interaksi")}
    tanpa_tgl = 0
    for p in posts:
        t = p.get("tanggal") or ""
        if t[:7] != ym:
            tanpa_tgl += 1
            continue
        d = int(t[8:10]) - 1
        harian["publikasi"][d] += 1
        if p["platform"] != "web":
            harian["views"][d] += int(p.get("views") or 0)
            harian["interaksi"][d] += interaksi(p)
    kum = {}
    for k, v in harian.items():
        acc, out = 0, []
        for x in v:
            acc += x
            out.append(acc)
        kum[k] = out
    batas_hari = nhari if porsi is None else int(round((porsi or 0) * nhari))
    target_kum = {"publikasi": sum(t for t in tot_t if t) or None, "views": _num(targets.get("views")),
                  "interaksi": _num(targets.get("interactions"))}

    # ---- teratas
    sos_posts = [p for p in posts if p["platform"] != "web"]
    teratas = sorted(sos_posts, key=lambda p: -(p.get("views") or 0))[:10]

    def top_akun(rows, kunci_urut, n=10):
        return sorted([a for a in rows if a["total"]], key=lambda a: (-(a[kunci_urut] or 0), a["nama"].lower()))[:n]

    akun_sos = [a for a in akun if a["platform_utama"] != "web"]
    per_platform_detail = {}
    for pl in SOSMED:
        rows = [p for p in posts if p["platform"] == pl]
        if not rows:
            continue
        akun_pl = per_kunci_sederhana(rows)
        per_platform_detail[pl] = {
            "aktif": sorted(akun_pl, key=lambda a: (-a["posts"], a["nama"].lower()))[:10],
            "views": sorted(akun_pl, key=lambda a: (-a["views"], a["nama"].lower()))[:10],
            "konten": [ringkas_post(p) for p in sorted(rows, key=lambda p: -(p.get("views") or 0))[:3]],
        }
    flags = flags if flags is not None else tandai(posts, ym, pakai_kelompok)
    hasil = {
        "ym": ym, "nama_periode": nama_periode(ym), "porsi": porsi, "berjalan": porsi is not None,
        "hari_ini": hari_ini.isoformat(), "label_kelompok": label_kelompok,
        "per": per, "total": total, "lalu": lalu, "indikator": ind,
        "skor": {"tercapai": sum(1 for i in dinilai if i["status"] == "tercapai"), "dinilai": len(dinilai)},
        "kelompok": kelompok, "akun": akun,
        "kuota_kelompok": {"terpenuhi": sum(1 for b in kuota_kel if b["status"] == "tercapai"), "dinilai": len(kuota_kel)},
        "kuota_akun": {"terpenuhi": sum(1 for b in kuota_akun if b["status"] == "tercapai"), "dinilai": len(kuota_akun)},
        "harian": {"kumulatif": kum, "hari": nhari, "sampai": batas_hari, "target": target_kum, "di_luar": tanpa_tgl},
        "teratas": [ringkas_post(p) for p in teratas],
        "akun_teratas": {"posts": top_akun(akun_sos, "total"), "views": top_akun(akun_sos, "views"),
                         "interaksi": top_akun(akun_sos, "interaksi"),
                         "er": sorted([a for a in akun_sos if a["views"] >= 1000 and a["er"] is not None],
                                      key=lambda a: -a["er"])[:10]},
        "detail_platform": per_platform_detail,
        "perlu_dicek": len(flags),
        "flag_hitung": dict(Counter(f for fs in flags.values() for f in fs)),
    }
    return hasil


def per_kunci_sederhana(rows):
    tampil = nama_tampil(rows, "akun")
    b = {}
    for p in rows:
        k = name_key(p.get("akun"))
        if not k:
            continue
        x = b.setdefault(k, {"nama": tampil[k], "posts": 0, "views": 0, "interaksi": 0})
        x["posts"] += 1
        x["views"] += int(p.get("views") or 0)
        x["interaksi"] += interaksi(p)
    return list(b.values())


def ringkas_post(p):
    return {"id": p.get("id"), "platform": p["platform"], "tanggal": p.get("tanggal"), "akun": p.get("akun"),
            "kelompok": p.get("kelompok"), "judul": p.get("judul"), "link": p.get("link"), "jenis": p.get("jenis"),
            "views": p.get("views"), "interaksi": interaksi(p), "er": er_baris(p)}


def selisih(sekarang, lalu):
    """Perubahan relatif terhadap bulan lalu; None bila bulan lalu kosong."""
    if lalu in (None, 0) or sekarang is None:
        return None
    return sekarang / lalu - 1


# ------------------------------------------------------------------ tempel baris
KOLOM_ALIAS = {
    "no": ["no", "nomor"],
    "kelompok": ["dpc", "kelompok", "wilayah", "grup", "group", "kabkota", "daerah"],
    "tanggal": ["tanggal", "tgl", "date", "tanggalposting", "tanggalunggah"],
    "akun": ["namaakun", "akun", "account", "username", "sumber", "media", "namamedia", "portal"],
    "jenis": ["videoflyer", "jenis", "materi", "tipe", "type", "format", "jeniskonten"],
    "judul": ["judul", "caption", "title", "keterangan", "judulberita"],
    "link": ["link", "url", "tautan", "linkposting", "linkberita"],
    "views": ["views", "view", "tayangan", "viewer", "plays", "putar"],
    "like": ["like", "likes", "suka"],
    "comment": ["comment", "comments", "komentar", "komen"],
    "repost": ["repost", "reposts"],
    "share": ["share", "shares", "bagikan", "dibagikan"],
}
URUT_RELATIF = {-5: "kelompok", -4: "tanggal", -3: "akun", -2: "jenis", -1: "judul",
                1: "views", 2: "like", 3: "comment", 4: "repost", 5: "share"}


def kunci_kolom(h):
    return re.sub(r"[^a-z0-9]", "", str(h or "").lower())


def peta_kolom(kepala):
    """Judul kolom -> {field: indeks}. Kolom yang tidak dikenali dilewati (mis. Total, ER)."""
    peta = {}
    for i, h in enumerate(kepala):
        k = kunci_kolom(h)
        for field, alias in KOLOM_ALIAS.items():
            if field not in peta and k in alias:
                peta[field] = i
                break
    return peta


def tampak_link(s):
    s = str(s or "").strip()
    return bool(re.match(r"^(https?://|www\.)", s, re.I) or re.match(r"^[a-z0-9\-]+(\.[a-z0-9\-]+)+/\S*$", s, re.I))


def parse_tempel(teks):
    """Teks tempel (daftar link atau baris Excel ber-tab) -> (baris, dilewati, keterangan)."""
    baris_teks = [b for b in str(teks or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    sel = [b.split("\t") for b in baris_teks if b.strip()]
    if not sel:
        return [], 0, []
    peta = None
    awal = 0
    kepala = peta_kolom(sel[0])
    if "link" in kepala and not any(tampak_link(c) for c in sel[0]):
        peta, awal = kepala, 1
    hasil, dilewati, ket = [], 0, []
    for n, cells in enumerate(sel[awal:], start=awal + 1):
        cells = [c.strip() for c in cells]
        if peta:
            li = peta["link"]
            if li >= len(cells) or not cells[li]:
                dilewati += 1
                ket.append(f"Baris {n}: tidak ada link")
                continue
            row = {f: cells[i] for f, i in peta.items() if i < len(cells) and f != "no"}
        else:
            idx = [i for i, c in enumerate(cells) if tampak_link(c)]
            if not idx:
                dilewati += 1
                ket.append(f"Baris {n}: tidak ada link")
                continue
            li = idx[0]
            row = {"link": cells[li]}
            if len(cells) > 1:
                for off, f in URUT_RELATIF.items():
                    j = li + off
                    if 0 <= j < len(cells) and cells[j]:
                        row[f] = cells[j]
        hasil.append(row)
    return hasil, dilewati, ket


def isi_dari_link(row, hari_ini=None):
    """Lengkapi platform, akun, tanggal dari link. Tanggal dari link hanya saran (yang diketik menang)."""
    info = parse_link(row.get("link"), hari_ini)
    if not info:
        return row, None
    row.setdefault("platform", info["platform"])
    if not row.get("akun") and info["akun"]:
        row["akun"] = info["akun"]
    if not row.get("tanggal") and info["tanggal"]:
        row["tanggal"] = info["tanggal"]
    return row, info


def gabung(posts, rows, buat_id, pertahankan_ganda=False):
    """Masukkan baris ke postingan periode. Link yang sudah ada diperbarui, bukan digandakan.

    pertahankan_ganda=True (impor workbook): kemunculan ke-n link di file mencocokkan kemunculan ke-n di data,
    sehingga baris ganda di file tetap utuh (dan ditandai 'link ganda'), dan impor ulang tidak menggandakan.
    """
    indeks = defaultdict(list)
    for p in posts:
        info = parse_link(p.get("link"))
        indeks[(p["platform"], info["kunci"] if info else p.get("link"))].append(p)
    pakai = Counter()
    baru = diperbarui = 0
    for row in rows:
        info = parse_link(row.get("link"))
        k = (row["platform"], info["kunci"] if info else row.get("link"))
        cocok = indeks.get(k, [])
        n = pakai[k] if pertahankan_ganda else 0
        if n < len(cocok):
            p = cocok[n]
            for f, v in row.items():
                if f != "platform" and v not in (None, ""):
                    p[f] = v
            bersihkan_post(p)
            diperbarui += 1
        else:
            p = bersihkan_post({"id": buat_id(), **row})
            posts.append(p)
            indeks[k].append(p)
            baru += 1
        pakai[k] += 1
    return baru, diperbarui


# ------------------------------------------------------------------ filter (sama dengan tabel Data)
def saring(posts, f, flags=None):
    f = f or {}
    cari = str(f.get("cari") or "").strip().lower()
    hasil = []
    for p in posts:
        if f.get("platform") and p["platform"] != f["platform"]:
            continue
        if f.get("kelompok") and name_key(p.get("kelompok")) != name_key(f["kelompok"]):
            continue
        if f.get("akun") and name_key(p.get("akun")) != name_key(f["akun"]):
            continue
        if f.get("jenis") and p.get("jenis") != f["jenis"]:
            continue
        if f.get("dari") and (p.get("tanggal") or "") < f["dari"]:
            continue
        if f.get("sampai") and (not p.get("tanggal") or p["tanggal"] > f["sampai"]):
            continue
        if f.get("dicek") and not (flags or {}).get(p["id"]):
            continue
        if cari and cari not in " ".join(str(p.get(k) or "") for k in ("judul", "akun", "kelompok", "link")).lower():
            continue
        hasil.append(p)
    return hasil


def keterangan_filter(f, label_kelompok="Kelompok"):
    f = f or {}
    bagian = []
    if f.get("platform"):
        bagian.append("Platform " + NAMA_PLATFORM.get(f["platform"], f["platform"]))
    if f.get("kelompok"):
        bagian.append(f"{label_kelompok} {f['kelompok']}")
    if f.get("akun"):
        bagian.append("Akun " + f["akun"])
    if f.get("jenis"):
        bagian.append("Jenis " + f["jenis"])
    if f.get("dari") or f.get("sampai"):
        bagian.append(f"Tanggal {tgl_panjang(f.get('dari')) or 'awal'} s.d. {tgl_panjang(f.get('sampai')) or 'akhir'}")
    if f.get("dicek"):
        bagian.append("Hanya perlu dicek")
    if f.get("cari"):
        bagian.append(f"Pencarian \"{f['cari']}\"")
    return "; ".join(bagian) if bagian else "Tidak ada (semua baris)"

/* Pantau Publikasi — tampilan. Semua angka ringkasan dihitung server (scripts/inti.py), sama dengan Excel & PPTX.
   Tabel Data menyaring di browser agar instan; aturan saring & total sama dengan inti.saring. */
(() => {
  "use strict";
  const PL = ["ig", "tt", "fb", "yt", "web"];
  const SOSMED = ["ig", "tt", "fb", "yt"];
  const NAMA_PL = { ig: "Instagram", tt: "TikTok", fb: "Facebook", yt: "YouTube", web: "Berita" };
  const WARNA_PL = { ig: "#c13584", tt: "#3d3d3d", fb: "#1877f2", yt: "#e00000", web: "#6b7a8f" };
  const STATUS = { tercapai: "Tercapai", hampir: "Hampir", tertinggal: "Tertinggal", sesuai: "Sesuai laju", dikejar: "Perlu dikejar", belum: "Belum mulai" };
  const BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const METRIK = ["views", "like", "comment", "repost", "share"];
  const PER_HAL = 50;

  const S = {
    saya: null, klien: [], kid: null, ym: null, data: null, flags: {}, r: null,
    filter: {}, hal: 1, pilih: new Set(), antre: {}, timer: null, kirimJalan: false, coba: null,
    lajuM: "publikasi", akunM: "posts", chart: null, tab: "ringkasan",
  };

  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const fmt = new Intl.NumberFormat("id-ID");
  const ribu = (n) => (n === null || n === undefined || n === "" ? "–" : fmt.format(Math.round(n)));
  const persen = (x, d = 2) => (x === null || x === undefined ? "–" : (x * 100).toLocaleString("id-ID", { minimumFractionDigits: d, maximumFractionDigits: d }) + "%");
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const tglPanjang = (t) => (t ? `${+t.slice(8, 10)} ${BULAN[+t.slice(5, 7) - 1]} ${t.slice(0, 4)}` : "");
  const namaPeriode = (ym) => `${BULAN[+ym.slice(5, 7) - 1]} ${ym.slice(0, 4)}`;
  const nameKey = (s) => String(s || "").toLowerCase().replace(/\./g, ". ").replace(/\s+/g, " ").trim();
  const stChip = (st) => (st ? `<span class="st st-${st}">${STATUS[st]}</span>` : '<span class="muted">–</span>');
  const plChip = (pl) => `<span class="pl"><i style="background:${WARNA_PL[pl]}"></i>${NAMA_PL[pl]}</span>`;
  const klien = () => S.klien.find((k) => k.id === S.kid);
  const admin = () => S.saya && S.saya.peran === "admin";

  function toast(t, lama = 3500) {
    const el = $("#toast");
    el.textContent = t;
    el.hidden = false;
    clearTimeout(toast.t);
    toast.t = setTimeout(() => (el.hidden = true), lama);
  }

  async function api(path, body) {
    const opt = body === undefined ? {} : { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) };
    const res = await fetch(path, opt);
    const ct = res.headers.get("Content-Type") || "";
    const isi = ct.includes("json") ? await res.json() : null;
    if (res.status === 401 && path !== "/api/masuk") { tampilMasuk(); throw new Error("silakan masuk dulu"); }
    if (!res.ok) throw new Error((isi && isi.galat) || `galat ${res.status}`);
    return isi;
  }

  async function unduhPost(path, body, namaCadangan) {
    const res = await fetch(path, body === undefined ? {} : { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!res.ok) { const j = await res.json().catch(() => ({})); throw new Error(j.galat || "gagal mengunduh"); }
    const cd = res.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename\*=UTF-8''([^;]+)/);
    const blob = await res.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = m ? decodeURIComponent(m[1]) : namaCadangan;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // ------------------------------------------------------------------ tema
  function pasangTema() {
    let t = null;
    try { t = localStorage.getItem("pp-tema"); } catch (e) { /* abaikan */ }
    if (t) document.documentElement.dataset.theme = t;
    $("#btn-tema").onclick = () => {
      const gelap = document.documentElement.dataset.theme
        ? document.documentElement.dataset.theme === "dark"
        : matchMedia("(prefers-color-scheme: dark)").matches;
      document.documentElement.dataset.theme = gelap ? "light" : "dark";
      try { localStorage.setItem("pp-tema", document.documentElement.dataset.theme); } catch (e) { /* abaikan */ }
      if (S.r) gambarLaju();
    };
  }
  const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();

  // ------------------------------------------------------------------ masuk
  function tampilMasuk() {
    $("#layar-app").hidden = true;
    $("#layar-masuk").hidden = false;
  }

  $("#form-masuk").addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = new FormData(e.target);
    $("#galat-masuk").hidden = true;
    try {
      await api("/api/masuk", { email: f.get("email"), sandi: f.get("sandi") });
      e.target.reset();
      await mulai();
    } catch (err) {
      $("#galat-masuk").textContent = err.message;
      $("#galat-masuk").hidden = false;
    }
  });

  $("#btn-keluar").onclick = async () => {
    await kirimAntre();
    await api("/api/keluar", {});
    S.saya = null;
    tampilMasuk();
  };

  async function mulai(kidPilih, ymPilih) {
    const j = await api("/api/saya");
    if (!j.pengguna) return tampilMasuk();
    S.saya = j.pengguna;
    S.klien = j.klien;
    S.hariIni = j.hari_ini;
    $("#layar-masuk").hidden = true;
    $("#layar-app").hidden = false;
    $("#nama-pengguna").textContent = `${S.saya.nama} · ${S.saya.peran === "admin" ? "Admin" : "Operator"}`;
    let simpan = {};
    try { simpan = JSON.parse(localStorage.getItem("pp-pilih") || "{}"); } catch (e) { /* abaikan */ }
    S.kid = kidPilih || (S.klien.some((k) => k.id === simpan.kid) ? simpan.kid : S.klien[0] && S.klien[0].id);
    const k = klien();
    if (!k) {
      isiPemilih();
      tab("pengaturan");
      toast(admin() ? "Belum ada klien. Buat klien baru di Pengaturan." : "Anda belum menjadi anggota klien mana pun. Hubungi admin.", 6000);
      return;
    }
    const ymNow = S.hariIni.slice(0, 7);
    S.ym = ymPilih || (k.periode.includes(simpan.ym) && simpan.kid === S.kid ? simpan.ym : k.periode.includes(ymNow) ? ymNow : k.periode[k.periode.length - 1]);
    isiPemilih();
    await muat();
  }

  // ------------------------------------------------------------------ pemilih klien & periode
  function isiPemilih() {
    const sk = $("#pilih-klien");
    sk.innerHTML = S.klien.map((k) => `<option value="${k.id}">${esc(k.nama)}</option>`).join("") || "<option>(belum ada klien)</option>";
    sk.value = S.kid || "";
    const k = klien();
    const sp = $("#pilih-periode");
    sp.innerHTML = k ? k.periode.slice().reverse().map((ym) => `<option value="${ym}">${namaPeriode(ym)}</option>`).join("") : "";
    sp.value = S.ym || "";
    document.documentElement.style.setProperty("--brand", (k && k.warna) || "#B3261E");
    $("#badge-sim").hidden = !(k && /simulasi/i.test(k.nama));
    try { localStorage.setItem("pp-pilih", JSON.stringify({ kid: S.kid, ym: S.ym })); } catch (e) { /* abaikan */ }
  }

  $("#pilih-klien").onchange = async (e) => {
    await kirimAntre();
    S.kid = e.target.value;
    const k = klien();
    const ymNow = S.hariIni.slice(0, 7);
    S.ym = k.periode.includes(S.ym) ? S.ym : k.periode.includes(ymNow) ? ymNow : k.periode[k.periode.length - 1];
    S.filter = {}; S.hal = 1; S.pilih.clear();
    isiPemilih();
    await muat();
  };
  $("#pilih-periode").onchange = async (e) => {
    await kirimAntre();
    S.ym = e.target.value;
    S.hal = 1; S.pilih.clear();
    isiPemilih();
    await muat();
  };
  $("#btn-periode").onclick = () => {
    dialog("Tambah periode", `<label>Bulan<input type="month" id="in-ym" value="${S.hariIni.slice(0, 7)}"></label>
      <p class="muted kecil">Periode baru memakai target bawaan klien sampai diberi target sendiri.</p>`,
      [["Tambah", async () => {
        const ym = $("#in-ym").value;
        if (!ym) return false;
        const j = await api("/api/periode/tambah", { klien: S.kid, ym });
        klien().periode = j.periode;
        await kirimAntre();
        S.ym = ym;
        isiPemilih();
        await muat();
      }]]);
  };

  async function muat() {
    if (!S.kid || !S.ym) return;
    const d = await api(`/api/periode?klien=${encodeURIComponent(S.kid)}&ym=${S.ym}`);
    S.data = d;
    S.flags = d.flags;
    S.r = d.ringkasan;
    pulihkanAntre();
    gambarSemua();
  }

  function gambarSemua() {
    gambarPil();
    if (S.tab === "ringkasan") gambarRingkasan();
    if (S.tab === "data") gambarData();
    if (S.tab === "target") gambarTarget();
    if (S.tab === "ekspor") gambarEkspor();
    if (S.tab === "pengaturan") gambarPengaturan();
  }

  function gambarPil() {
    const n = Object.keys(S.flags || {}).length;
    $("#pil-cek").hidden = !n;
    $("#pil-cek").textContent = ribu(n);
    $("#pil-cek").title = `${n} baris perlu dicek`;
  }

  // ------------------------------------------------------------------ tab
  function tab(nama) {
    S.tab = nama;
    $$(".tabs button").forEach((b) => b.classList.toggle("aktif", b.dataset.tab === nama));
    $$(".tab").forEach((s) => (s.hidden = s.id !== "tab-" + nama));
    if (S.data || nama === "pengaturan") gambarSemua();
  }
  $$(".tabs button").forEach((b) => (b.onclick = () => tab(b.dataset.tab)));

  // ------------------------------------------------------------------ Ringkasan
  function gambarRingkasan() {
    const r = S.r, k = klien();
    const t = r.total, lalu = r.lalu && r.lalu.total;
    const ind = Object.fromEntries(r.indikator.map((i) => [i.kunci, i]));

    const beda = (now, prev) => {
      if (!lalu || prev === null || prev === undefined || !prev || now === null) return "";
      const d = now / prev - 1;
      return `<span class="${d >= 0 ? "naik" : "turun"}">${d >= 0 ? "▲" : "▼"} ${persen(Math.abs(d), 1)}</span><span class="muted">vs bulan lalu${r.berjalan ? " (sebulan penuh)" : ""}</span>`;
    };
    const tgtTeks = (i, f) => (i && i.target ? `<span>Target ${f(i.target)}</span>${stChip(i.status)}` : "");
    const kartu = [
      ["Publikasi", ribu(t.publikasi), tgtTeks(ind.publikasi, ribu), beda(t.publikasi, lalu && lalu.publikasi)],
      ["Views", ribu(t.views), tgtTeks(ind.views, ribu), beda(t.views, lalu && lalu.views)],
      ["Interaksi", ribu(t.interaksi), tgtTeks(ind.interaksi, ribu), beda(t.interaksi, lalu && lalu.interaksi)],
      ["Engagement rate", persen(t.er), tgtTeks(ind.er, (x) => "min. " + persen(x)), beda(t.er, lalu && lalu.er)],
    ];
    $("#kpi").innerHTML = kartu.map(([l, v, tg, b]) => `<div class="kartu"><div class="lbl">${l}</div><div class="nilai">${v}</div>
      <div class="sub">${tg || ""}</div><div class="sub">${b || (lalu ? "" : '<span class="muted">Bulan lalu belum ada data</span>')}</div></div>`).join("");

    // per platform
    $("#tbl-platform").innerHTML = `<thead><tr><th>Platform</th><th class="n">Postingan</th><th class="n">Video</th><th class="n">Flyer</th>
      <th class="n">Artikel</th><th class="n">Views</th><th class="n">Interaksi</th><th class="n">ER</th></tr></thead><tbody>
      ${PL.map((p) => { const x = r.per[p]; return `<tr><td>${plChip(p)}</td><td class="n">${ribu(x.posts)}</td><td class="n">${ribu(x.video)}</td>
        <td class="n">${ribu(x.flyer)}</td><td class="n">${ribu(x.artikel)}</td><td class="n">${p === "web" ? "–" : ribu(x.views)}</td>
        <td class="n">${p === "web" ? "–" : ribu(x.interaksi)}</td><td class="n">${persen(x.er)}</td></tr>`; }).join("")}
      <tr class="tot"><td>Total</td><td class="n">${ribu(t.publikasi)}</td><td class="n">${ribu(t.video)}</td><td class="n">${ribu(t.flyer)}</td>
        <td class="n">${ribu(t.artikel)}</td><td class="n">${ribu(t.views)}</td><td class="n">${ribu(t.interaksi)}</td><td class="n">${persen(t.er)}</td></tr></tbody>`;

    // target & realisasi
    const f = (i, v) => (i.jenis === "er" ? persen(v) : ribu(v));
    $("#tbl-target").innerHTML = `<thead><tr><th>Indikator</th><th class="n">Realisasi</th><th class="n">Target</th><th class="n">Capaian</th>
      <th>Status</th><th class="n">Bulan lalu</th><th class="n">Perubahan</th></tr></thead><tbody>
      ${r.indikator.map((i) => { const d = i.lalu ? (i.realisasi ?? 0) / i.lalu - 1 : null;
        return `<tr${i.turunan ? ' class="tot"' : ""}><td>${esc(i.label)}</td><td class="n">${f(i, i.realisasi)}</td>
        <td class="n">${i.target ? f(i, i.target) : '<span class="muted">tidak dinilai</span>'}</td>
        <td class="n">${i.capaian !== null ? persen(i.capaian, 0) : "–"}</td><td>${stChip(i.status)}</td>
        <td class="n">${i.lalu !== null && i.lalu !== undefined ? f(i, i.lalu) : "–"}</td>
        <td class="n">${d === null || i.realisasi === null ? "–" : `<span class="${d >= 0 ? "naik" : "turun"}">${d >= 0 ? "+" : "−"}${persen(Math.abs(d), 1)}</span>`}</td></tr>`; }).join("")}</tbody>`;

    // per kelompok
    $("#judul-kelompok").textContent = `Per ${r.label_kelompok}`;
    const kel = r.kelompok;
    $("#tbl-kelompok").innerHTML = !k.pakai_kelompok ? '<tbody><tr><td class="kosong">Klien ini tidak memakai kelompok.</td></tr></tbody>'
      : `<thead><tr><th>${esc(r.label_kelompok)}</th>${PL.map((p) => `<th class="n">${p.toUpperCase()}</th>`).join("")}<th class="n">Total</th>
        <th class="n">Views</th><th class="n">Kuota</th><th>Status</th></tr></thead><tbody>
        ${kel.map((b) => `<tr><td><a href="#" data-kel="${esc(b.nama)}">${esc(b.nama)}</a></td>${PL.map((p) => `<td class="n">${ribu(b[p])}</td>`).join("")}
          <td class="n"><b>${ribu(b.total)}</b></td><td class="n">${ribu(b.views)}</td><td class="n">${b.kuota ? ribu(b.kuota) : "–"}</td>
          <td>${stChip(b.status)}</td></tr>`).join("") || '<tr><td colspan="11" class="kosong">Belum ada data.</td></tr>'}</tbody>`;
    $$("#tbl-kelompok a[data-kel]").forEach((a) => (a.onclick = (e) => { e.preventDefault(); S.filter = { kelompok: a.dataset.kel }; S.hal = 1; tab("data"); }));

    gambarAkun();
    // konten teratas
    $("#tbl-teratas").innerHTML = `<thead><tr><th>#</th><th>Platform</th><th>Tanggal</th><th>Akun</th><th>Judul</th><th class="n">Views</th>
      <th class="n">Interaksi</th><th class="n">ER</th></tr></thead><tbody>
      ${r.teratas.map((p, i) => `<tr><td>${i + 1}</td><td>${plChip(p.platform)}</td><td class="n">${tglPanjang(p.tanggal)}</td><td>${esc(p.akun)}</td>
        <td class="judul"><a class="potong" href="${esc(urlAman(p.link))}" target="_blank" rel="noopener">${esc(p.judul || "(tanpa judul)")}</a></td>
        <td class="n">${ribu(p.views)}</td><td class="n">${ribu(p.interaksi)}</td><td class="n">${persen(p.er)}</td></tr>`).join("") || '<tr><td colspan="8" class="kosong">Belum ada data.</td></tr>'}</tbody>`;
    gambarLaju();
  }

  function urlAman(u) {
    u = String(u || "").trim();
    if (!u) return "#";
    if (!/^https?:\/\//i.test(u)) u = "https://" + u;
    return u;
  }

  function gambarAkun() {
    const m = S.akunM, rows = S.r.akun_teratas[m] || [];
    $("#tbl-akun").innerHTML = `<thead><tr><th>#</th><th>Akun</th><th>Platform utama</th><th class="n">Postingan</th><th class="n">Views</th>
      <th class="n">Interaksi</th><th class="n">ER</th></tr></thead><tbody>
      ${rows.map((a, i) => `<tr><td>${i + 1}</td><td><a href="#" data-akun="${esc(a.nama)}">${esc(a.nama)}</a></td><td>${a.platform_utama ? plChip(a.platform_utama) : ""}</td>
        <td class="n">${ribu(a.total)}</td><td class="n">${ribu(a.views)}</td><td class="n">${ribu(a.interaksi)}</td><td class="n">${persen(a.er)}</td></tr>`).join("")
        || '<tr><td colspan="7" class="kosong">Belum ada data.</td></tr>'}</tbody>`;
    if (m === "er") $("#tbl-akun tbody").insertAdjacentHTML("beforeend", '<tr><td colspan="7" class="muted kecil">ER hanya untuk akun dengan minimal 1.000 views.</td></tr>');
    $$("#tbl-akun a[data-akun]").forEach((a) => (a.onclick = (e) => { e.preventDefault(); S.filter = { akun: a.dataset.akun }; S.hal = 1; tab("data"); }));
  }
  $$("#seg-akun button").forEach((b) => (b.onclick = () => { S.akunM = b.dataset.m; $$("#seg-akun button").forEach((x) => x.classList.toggle("aktif", x === b)); gambarAkun(); }));
  $$("#seg-laju button").forEach((b) => (b.onclick = () => { S.lajuM = b.dataset.m; $$("#seg-laju button").forEach((x) => x.classList.toggle("aktif", x === b)); gambarLaju(); }));

  function gambarLaju() {
    const h = S.r.harian, m = S.lajuM;
    const label = Array.from({ length: h.hari }, (_, i) => String(i + 1));
    const nyata = h.kumulatif[m].map((v, i) => (i < h.sampai ? v : null));
    const tgt = h.target[m];
    const jalur = tgt ? label.map((_, i) => (tgt * (i + 1)) / h.hari) : null;
    const brand = css("--brand"), ink2 = css("--ink-2"), grid = css("--grid");
    const ds = [{ label: "Realisasi kumulatif", data: nyata, borderColor: brand, backgroundColor: brand + "22", fill: true, pointRadius: 0, borderWidth: 2.5, tension: 0.15 }];
    if (jalur) ds.push({ label: "Jalur target", data: jalur, borderColor: ink2, borderDash: [6, 5], pointRadius: 0, borderWidth: 1.5, fill: false });
    if (S.chart) S.chart.destroy();
    S.chart = new Chart($("#chart-laju"), {
      type: "line",
      data: { labels: label, datasets: ds },
      options: {
        responsive: true, maintainAspectRatio: false, animation: false, interaction: { mode: "index", intersect: false },
        plugins: { legend: { labels: { color: ink2, boxWidth: 14 } },
          tooltip: { callbacks: { title: (it) => `Tanggal ${it[0].label}`, label: (c) => `${c.dataset.label}: ${ribu(c.parsed.y)}` } } },
        scales: { x: { grid: { display: false }, ticks: { color: ink2, maxRotation: 0, autoSkipPadding: 8 } },
          y: { beginAtZero: true, grid: { color: grid }, ticks: { color: ink2, callback: (v) => ribu(v) } } },
      },
    });
    const sekarang = h.kumulatif[m][Math.max(0, h.sampai - 1)] || 0;
    $("#ket-laju").textContent = (tgt
      ? `Target ${ribu(tgt)} · realisasi ${ribu(sekarang)} (${persen(sekarang / tgt, 0)}) · jalur target sampai hari ini ${ribu((tgt * h.sampai) / h.hari)}.`
      : "Belum ada target untuk metrik ini.") + (h.di_luar ? ` ${h.di_luar} baris tanpa tanggal di periode ini tidak masuk grafik.` : "");
  }

  function bukaPerluDicek() { S.filter = { dicek: true }; S.hal = 1; tab("data"); }

  // ------------------------------------------------------------------ Data: filter
  function saring() {
    const f = S.filter, cari = (f.cari || "").trim().toLowerCase();
    return S.data.posts.filter((p) => {
      if (f.platform && p.platform !== f.platform) return false;
      if (f.kelompok && nameKey(p.kelompok) !== nameKey(f.kelompok)) return false;
      if (f.akun && nameKey(p.akun) !== nameKey(f.akun)) return false;
      if (f.jenis && p.jenis !== f.jenis) return false;
      if (f.dari && (p.tanggal || "") < f.dari) return false;
      if (f.sampai && (!p.tanggal || p.tanggal > f.sampai)) return false;
      if (f.dicek && !S.flags[p.id]) return false;
      if (cari && ![p.judul, p.akun, p.kelompok, p.link].join(" ").toLowerCase().includes(cari)) return false;
      return true;
    });
  }

  function opsiNama(kolom, semua) {
    const per = new Map();
    for (const p of S.data.posts) {
      const v = (p[kolom] || "").trim();
      if (!v) continue;
      const k = nameKey(v);
      const c = per.get(k) || new Map();
      c.set(v, (c.get(v) || 0) + 1);
      per.set(k, c);
    }
    const nama = [...per.values()].map((c) => [...c.entries()].sort((a, b) => b[1] - a[1] || (/^[A-Z]/.test(b[0]) - /^[A-Z]/.test(a[0])) || a[0].localeCompare(b[0]))[0][0]);
    nama.sort((a, b) => a.localeCompare(b, "id"));
    return `<option value="">${semua}</option>` + nama.map((n) => `<option>${esc(n)}</option>`).join("");
  }

  function isiFilter() {
    const k = klien(), f = S.filter;
    $("#f-platform").innerHTML = '<option value="">Semua platform</option>' + PL.map((p) => `<option value="${p}">${NAMA_PL[p]}</option>`).join("");
    $("#f-kelompok").innerHTML = opsiNama("kelompok", `Semua ${k.label_kelompok || "kelompok"}`);
    $("#f-kelompok").hidden = !k.pakai_kelompok;
    $("#f-akun").innerHTML = opsiNama("akun", "Semua akun");
    const pasang = (id, v) => {
      const el = $(id);
      if (v && el.tagName === "SELECT" && ![...el.options].some((o) => nameKey(o.value) === nameKey(v))) el.insertAdjacentHTML("beforeend", `<option>${esc(v)}</option>`);
      if (el.tagName === "SELECT" && v) el.value = [...el.options].find((o) => nameKey(o.value) === nameKey(v)).value;
      else el.value = v || "";
    };
    pasang("#f-platform", f.platform); pasang("#f-kelompok", f.kelompok); pasang("#f-akun", f.akun); pasang("#f-jenis", f.jenis);
    $("#f-dari").value = f.dari || ""; $("#f-sampai").value = f.sampai || "";
    $("#f-dicek").checked = !!f.dicek; $("#f-cari").value = f.cari || "";
  }

  const ubahFilter = () => {
    S.filter = {
      platform: $("#f-platform").value, kelompok: $("#f-kelompok").value, akun: $("#f-akun").value, jenis: $("#f-jenis").value,
      dari: $("#f-dari").value, sampai: $("#f-sampai").value, dicek: $("#f-dicek").checked, cari: $("#f-cari").value,
    };
    S.hal = 1;
    gambarBaris();
  };
  ["#f-platform", "#f-kelompok", "#f-akun", "#f-jenis", "#f-dari", "#f-sampai", "#f-dicek"].forEach((id) => ($(id).onchange = ubahFilter));
  let tCari;
  $("#f-cari").oninput = () => { clearTimeout(tCari); tCari = setTimeout(ubahFilter, 120); };
  $("#btn-reset").onclick = () => { S.filter = {}; S.hal = 1; isiFilter(); gambarBaris(); };

  // ------------------------------------------------------------------ Data: tabel
  function gambarData() {
    const k = klien();
    isiFilter();
    $("#tbl-data thead").innerHTML = `<tr><th class="col-cek"><input type="checkbox" id="cek-semua" title="Pilih semua di halaman ini" aria-label="Pilih semua"></th>
      <th class="col-flag" title="Penanda masalah">!</th><th class="col-pl">Platform</th><th class="col-tgl">Tanggal</th>
      ${k.pakai_kelompok ? `<th class="col-kel">${esc(k.label_kelompok)}</th>` : ""}<th class="col-akun">Akun</th><th class="col-jenis">Jenis</th>
      <th class="col-judul">Judul</th><th class="col-link">Link</th>${METRIK.map((m) => `<th class="n col-m">${m[0].toUpperCase() + m.slice(1)}</th>`).join("")}
      <th class="n col-it">Interaksi</th><th class="n col-er">ER</th></tr>`;
    $("#cek-semua").onchange = (e) => {
      $$("#tbl-data tbody input.pilih").forEach((c) => { c.checked = e.target.checked; S.pilih[e.target.checked ? "add" : "delete"](c.closest("tr").dataset.id); });
      $("#btn-hapus").disabled = !S.pilih.size;
    };
    gambarBaris();
  }

  function selFlag(id) {
    const f = S.flags[id];
    if (!f) return "";
    const info = S.data.flag_info;
    return `<span class="flag" tabindex="0" aria-label="${f.length} masalah">${f.length}<span class="tip">${f.map((x) => `<b>${esc(info[x][0])}</b><span>${esc(info[x][1])}</span>`).join("")}</span></span>`;
  }

  function barisHTML(p) {
    const k = klien(), web = p.platform === "web";
    const it = web ? null : ["like", "comment", "repost", "share"].reduce((a, m) => a + (+p[m] || 0), 0);
    const er = !web && p.views ? it / p.views : null;
    const ubah = S.antre[p.id] ? " ubah" : "";
    return `<tr data-id="${p.id}" class="${ubah}">
      <td><input type="checkbox" class="pilih" ${S.pilih.has(p.id) ? "checked" : ""} aria-label="Pilih baris"></td>
      <td class="flagcell">${selFlag(p.id)}</td>
      <td><select data-f="platform">${PL.map((x) => `<option value="${x}"${x === p.platform ? " selected" : ""}>${NAMA_PL[x]}</option>`).join("")}</select></td>
      <td><input type="date" data-f="tanggal" value="${esc(p.tanggal)}"></td>
      ${k.pakai_kelompok ? `<td><input data-f="kelompok" value="${esc(p.kelompok)}" list="dl-kel"></td>` : ""}
      <td><input data-f="akun" value="${esc(p.akun)}"></td>
      <td><select data-f="jenis">${["Video", "Flyer", "Artikel"].map((x) => `<option${x === p.jenis ? " selected" : ""}>${x}</option>`).join("")}</select></td>
      <td><input data-f="judul" value="${esc(p.judul)}" title="${esc(p.judul)}"></td>
      <td><div class="linkcell"><input data-f="link" value="${esc(p.link)}" title="${esc(p.link)}" placeholder="https://…">${p.link ? `<a href="${esc(urlAman(p.link))}" target="_blank" rel="noopener" title="Buka link">↗</a>` : ""}</div></td>
      ${METRIK.map((m) => `<td><input class="n" data-f="${m}" inputmode="numeric" value="${web || p[m] === null || p[m] === undefined ? "" : esc(p[m])}"${web ? ' disabled title="Berita tidak memiliki metrik"' : ""}></td>`).join("")}
      <td class="n it">${web ? "–" : ribu(it)}</td><td class="n er">${persen(er)}</td></tr>`;
  }

  function gambarBaris() {
    const rows = saring();
    const k = klien();
    const nhal = Math.max(1, Math.ceil(rows.length / PER_HAL));
    S.hal = Math.min(S.hal, nhal);
    const hal = rows.slice((S.hal - 1) * PER_HAL, S.hal * PER_HAL);
    const tb = $("#tbl-data tbody");
    const kol = 15 + (k.pakai_kelompok ? 1 : 0);
    tb.innerHTML = hal.map(barisHTML).join("") || `<tr><td colspan="${kol}" class="kosong">${S.data.posts.length ? "Tidak ada baris yang cocok dengan filter." : "Belum ada data. Tempel daftar link atau impor workbook."}</td></tr>`;
    if (!$("#dl-kel")) document.body.insertAdjacentHTML("beforeend", '<datalist id="dl-kel"></datalist>');
    $("#dl-kel").innerHTML = opsiNama("kelompok", "").replace('<option value=""></option>', "");
    // total (semua halaman, mengikuti filter)
    const sos = rows.filter((p) => p.platform !== "web");
    const jum = Object.fromEntries(METRIK.map((m) => [m, sos.reduce((a, p) => a + (+p[m] || 0), 0)]));
    const it = jum.like + jum.comment + jum.repost + jum.share;
    $("#tbl-data tfoot").innerHTML = `<tr><td colspan="${k.pakai_kelompok ? 9 : 8}">Total ${ribu(rows.length)} baris${rows.length !== S.data.posts.length ? ` (dari ${ribu(S.data.posts.length)})` : ""}</td>
      ${METRIK.map((m) => `<td class="n">${ribu(jum[m])}</td>`).join("")}<td class="n">${ribu(it)}</td><td class="n">${persen(jum.views ? it / jum.views : null)}</td></tr>`;
    $("#info-data").textContent = S.filter.dicek ? `${ribu(rows.length)} baris perlu dicek` : "";
    // halaman
    const h = $("#halaman");
    h.innerHTML = nhal <= 1 ? "" : `<button class="btn kecil" data-h="${S.hal - 1}" ${S.hal === 1 ? "disabled" : ""}>‹ Sebelumnya</button>
      <span class="muted">Halaman ${S.hal} dari ${nhal} · ${PER_HAL} baris per halaman</span>
      <button class="btn kecil" data-h="${S.hal + 1}" ${S.hal === nhal ? "disabled" : ""}>Berikutnya ›</button>`;
    $$("#halaman button").forEach((b) => (b.onclick = () => { S.hal = +b.dataset.h; gambarBaris(); $(".data-wrap").scrollTop = 0; }));
    $("#cek-semua").checked = false;
    $("#btn-hapus").disabled = !S.pilih.size;
  }

  // edit langsung
  const tbData = $("#tbl-data tbody");
  tbData.addEventListener("change", (e) => {
    const el = e.target;
    const tr = el.closest("tr");
    if (!tr) return;
    const id = tr.dataset.id;
    if (el.classList.contains("pilih")) {
      S.pilih[el.checked ? "add" : "delete"](id);
      $("#btn-hapus").disabled = !S.pilih.size;
      return;
    }
    if (!el.dataset.f) return;
    ubahSel(id, el.dataset.f, el.value, tr, el.dataset.f === "platform");
  });
  tbData.addEventListener("input", (e) => {
    const el = e.target;
    if (!el.dataset.f || el.tagName === "SELECT" || el.type === "date") return;
    const tr = el.closest("tr");
    ubahSel(tr.dataset.id, el.dataset.f, el.value, tr, false);
  });

  function bacaAngka(v) {
    let s = String(v ?? "").trim().toLowerCase().replace(/\s/g, "");
    if (!s || s === "-") return null;
    let kali = 1;
    for (const [a, k] of [["jt", 1e6], ["m", 1e6], ["rb", 1e3], ["k", 1e3]]) if (s.endsWith(a)) { s = s.slice(0, -a.length); kali = k; break; }
    if (kali > 1) { s = s.replace(",", "."); if ((s.match(/\./g) || []).length > 1) s = s.replace(/\./g, ""); }
    else s = s.replace(/\./g, "").replace(",", ".");
    const n = Number(s);
    return Number.isFinite(n) ? Math.round(n * kali) : null;
  }

  function ubahSel(id, f, v, tr, gambarUlang) {
    const p = S.data.posts.find((x) => x.id === id);
    if (!p) return;
    p[f] = METRIK.includes(f) ? bacaAngka(v) : v;
    (S.antre[id] = S.antre[id] || {})[f] = v;
    tr.classList.add("ubah");
    if (gambarUlang) { tr.outerHTML = barisHTML(p); }
    else if (METRIK.includes(f)) {
      const it = ["like", "comment", "repost", "share"].reduce((a, m) => a + (+p[m] || 0), 0);
      $(".it", tr).textContent = ribu(it);
      $(".er", tr).textContent = persen(p.views ? it / p.views : null);
      gambarTotal();
    }
    jadwalKirim();
  }

  function gambarTotal() {
    const rows = saring().filter((p) => p.platform !== "web");
    const jum = Object.fromEntries(METRIK.map((m) => [m, rows.reduce((a, p) => a + (+p[m] || 0), 0)]));
    const it = jum.like + jum.comment + jum.repost + jum.share;
    const tds = $$("#tbl-data tfoot td");
    METRIK.forEach((m, i) => (tds[1 + i].textContent = ribu(jum[m])));
    tds[6].textContent = ribu(it);
    tds[7].textContent = persen(jum.views ? it / jum.views : null);
  }

  // ------------------------------------------------------------------ simpan otomatis (antrean + coba ulang)
  const kunciAntre = () => `pp-antre:${S.kid}:${S.ym}`;
  function statusSimpan(t, kelas = "") {
    const el = $("#status-simpan");
    el.textContent = t;
    el.className = "simpan " + kelas;
  }
  function cadangkan() {
    try {
      if (Object.keys(S.antre).length) localStorage.setItem(kunciAntre(), JSON.stringify(S.antre));
      else localStorage.removeItem(kunciAntre());
    } catch (e) { /* abaikan */ }
  }
  function pulihkanAntre() {
    let ada = {};
    try { ada = JSON.parse(localStorage.getItem(kunciAntre()) || "{}"); } catch (e) { /* abaikan */ }
    if (!Object.keys(ada).length) return;
    S.antre = ada;
    for (const [id, isi] of Object.entries(ada)) {
      const p = S.data.posts.find((x) => x.id === id);
      if (p) for (const [f, v] of Object.entries(isi)) p[f] = METRIK.includes(f) ? bacaAngka(v) : v;
    }
    toast("Perubahan yang belum tersimpan dipulihkan dan sedang dikirim ulang.");
    jadwalKirim(100);
  }
  function jadwalKirim(jeda = 1000) {
    cadangkan();
    statusSimpan("Belum tersimpan…");
    clearTimeout(S.timer);
    S.timer = setTimeout(kirimAntre, jeda);
  }
  async function kirimAntre() {
    clearTimeout(S.timer);
    if (S.kirimJalan) { await S.kirimJalan; }
    const ids = Object.keys(S.antre);
    if (!ids.length) return;
    const kirim = S.antre;
    S.antre = {};
    const kid = S.kid, ym = S.ym;
    statusSimpan("Menyimpan…");
    S.kirimJalan = (async () => {
      try {
        const j = await api("/api/posting/simpan", { klien: kid, ym, ubah: ids.map((id) => ({ id, isi: kirim[id] })) });
        if (kid !== S.kid || ym !== S.ym) return;
        S.flags = j.flags; S.r = j.ringkasan;
        for (const q of j.posts) {
          const p = S.data.posts.find((x) => x.id === q.id);
          if (!p) continue;
          for (const f of Object.keys(q)) if (!(S.antre[q.id] && f in S.antre[q.id])) p[f] = q[f];
          const tr = $(`#tbl-data tbody tr[data-id="${q.id}"]`);
          if (tr && !tr.contains(document.activeElement) && !S.antre[q.id]) tr.outerHTML = barisHTML(p);
          else if (tr) { tr.querySelector(".flagcell").innerHTML = selFlag(q.id); if (!S.antre[q.id]) tr.classList.remove("ubah"); }
        }
        cadangkan();
        statusSimpan(Object.keys(S.antre).length ? "Belum tersimpan…" : "✓ Tersimpan", "ok");
        gambarPil();
        clearInterval(S.coba); S.coba = null;
      } catch (err) {
        for (const [id, isi] of Object.entries(kirim)) S.antre[id] = { ...isi, ...(S.antre[id] || {}) };
        cadangkan();
        statusSimpan("⚠ Gagal menyimpan — mencoba lagi", "gagal");
        if (!S.coba) S.coba = setInterval(() => kirimAntre(), 5000);
      }
    })();
    await S.kirimJalan;
    S.kirimJalan = false;
  }
  window.addEventListener("beforeunload", (e) => {
    if (Object.keys(S.antre).length) { cadangkan(); e.preventDefault(); e.returnValue = ""; }
  });

  // ------------------------------------------------------------------ Data: aksi
  async function terapkanHasil(j) {
    await muat();
  }

  $("#btn-baris").onclick = async () => {
    await kirimAntre();
    const j = await api("/api/posting/tambah", { klien: S.kid, ym: S.ym, platform: S.filter.platform || "ig" });
    S.data.posts.unshift(j.post);
    S.flags = j.flags; S.r = j.ringkasan;
    S.filter = {}; S.hal = 1;
    isiFilter(); gambarBaris(); gambarPil();
    const tr = $(`#tbl-data tbody tr[data-id="${j.post.id}"]`);
    if (tr) { tr.classList.add("baru"); $('input[data-f="link"]', tr).focus(); }
  };

  $("#btn-hapus").onclick = () => {
    const n = S.pilih.size;
    dialog("Hapus baris", `<p>Hapus <b>${n}</b> baris terpilih? Total akan diperbarui. Tindakan ini tidak bisa dibatalkan.</p>`, [["Hapus", async () => {
      await kirimAntre();
      const ids = [...S.pilih];
      const j = await api("/api/posting/hapus", { klien: S.kid, ym: S.ym, ids });
      const set = new Set(ids);
      S.data.posts = S.data.posts.filter((p) => !set.has(p.id));
      S.flags = j.flags; S.r = j.ringkasan; S.pilih.clear();
      gambarBaris(); gambarPil();
      toast(`${j.dihapus} baris dihapus.`);
    }, "bahaya"]]);
  };

  $("#btn-tempel").onclick = () => {
    dialog(`Tempel ke ${namaPeriode(S.ym)}`, `
      <p class="muted kecil">Tempel daftar link (satu per baris), atau salin baris dari Excel (dengan atau tanpa baris judul).
      Platform dikenali dari domain; tanggal Instagram/TikTok diisi dari link bila kosong; link yang sudah ada diperbarui, bukan digandakan.
      Angka seperti "1.234" atau "3,7k" terbaca benar.</p>
      <textarea id="in-tempel" rows="12" placeholder="https://www.instagram.com/reel/…&#10;https://vt.tiktok.com/…&#10;&#10;atau: DPC [tab] Tanggal [tab] Akun [tab] Jenis [tab] Judul [tab] Link [tab] Views [tab] Like …"></textarea>
      <div id="hasil-tempel"></div>`,
      [["Proses", async () => {
        const teks = $("#in-tempel").value;
        if (!teks.trim()) return false;
        await kirimAntre();
        const j = await api("/api/tempel", { klien: S.kid, ym: S.ym, teks });
        $("#hasil-tempel").innerHTML = `<div class="hasil"><b>${j.baru}</b> baru · <b>${j.diperbarui}</b> diperbarui · <b>${j.dilewati}</b> dilewati
          ${j.tanggal_dikoreksi ? ` · ${j.tanggal_dikoreksi} tanggal salah tahun dikoreksi` : ""}
          ${j.keterangan.length ? `<ul>${j.keterangan.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>` : ""}</div>`;
        $("#in-tempel").value = "";
        await terapkanHasil(j);
        return false;
      }]], "Tutup");
    setTimeout(() => $("#in-tempel").focus(), 50);
  };

  $("#btn-impor").onclick = () => {
    dialog(`Impor workbook ke ${namaPeriode(S.ym)}`, `
      <p class="muted kecil">Lembar dikenali dari namanya (IG/Instagram, TikTok, Facebook, YouTube, Artikel/Berita); kolom dibaca dari judulnya.
      Hanya baris bertanggal di ${namaPeriode(S.ym)} yang masuk. Tanggal salah tahun pada bulan yang sama dikoreksi dan dilaporkan.
      Impor ulang file yang sama memperbarui baris, tidak menggandakan.</p>
      <input type="file" id="in-xlsx" accept=".xlsx">
      <div id="hasil-impor"></div>`,
      [["Impor", async () => {
        const f = $("#in-xlsx").files[0];
        if (!f) return false;
        $("#hasil-impor").innerHTML = '<p class="muted">Membaca workbook…</p>';
        const data = await bacaBase64(f);
        await kirimAntre();
        const j = await api("/api/impor", { klien: S.kid, ym: S.ym, data, nama: f.name });
        const l = j.laporan;
        $("#hasil-impor").innerHTML = `<div class="hasil"><b>${j.baru}</b> baru · <b>${j.diperbarui}</b> diperbarui
          <ul>${l.lembar.map((x) => `<li>${esc(x.nama)} → ${NAMA_PL[x.platform]}: ${x.masuk} masuk${x.luar_periode ? `, ${x.luar_periode} di luar periode` : ""}${x.tanpa_tanggal ? `, ${x.tanpa_tanggal} tanpa tanggal` : ""}${x.dikoreksi ? `, ${x.dikoreksi} tanggal dikoreksi` : ""}${x.galat ? ` — ${esc(x.galat)}` : ""}</li>`).join("")}</ul>
          ${l.dikoreksi.length ? `<p>Koreksi tanggal:</p><ul>${l.dikoreksi.slice(0, 10).map((x) => `<li>${esc(x)}</li>`).join("")}</ul>` : ""}
          ${l.diabaikan.length ? `<p class="muted">Lembar dilewati: ${esc(l.diabaikan.join(", "))}</p>` : ""}</div>`;
        await terapkanHasil(j);
        return false;
      }]], "Tutup");
  };

  function bacaBase64(file) {
    return new Promise((ok, gagal) => {
      const fr = new FileReader();
      fr.onload = () => ok(String(fr.result).split(",")[1]);
      fr.onerror = gagal;
      fr.readAsDataURL(file);
    });
  }

  $("#btn-rapikan").onclick = async () => {
    await kirimAntre();
    const k = klien();
    const kolom = k.pakai_kelompok ? "kelompok" : "akun";
    dialog("Rapikan nama", `
      <div class="seg" id="seg-rapi">${k.pakai_kelompok ? `<button type="button" data-k="kelompok" class="aktif">${esc(k.label_kelompok)}</button>` : ""}
        <button type="button" data-k="akun"${k.pakai_kelompok ? "" : ' class="aktif"'}>Akun</button></div>
      <p class="muted kecil">Menyatukan variasi huruf besar-kecil dan spasi ke penulisan yang paling sering. Aturan "nama lama = nama baku" dijalankan lebih dulu
      dan disimpan per klien.</p>
      <div id="pratinjau-rapi"></div>
      <h3>Tambah aturan nama lama = nama baku</h3>
      <div class="baris-tambah"><input id="al-dari" placeholder="Nama lama, mis. Kota Solo"><input id="al-ke" placeholder="Nama baku, mis. Kota Surakarta">
        <button class="btn kecil" type="button" id="al-tambah">Tambah aturan</button></div>`,
      [["Terapkan", async () => {
        const kk = $("#seg-rapi .aktif").dataset.k;
        const j = await api("/api/rapikan", { klien: S.kid, ym: S.ym, kolom: kk });
        toast(`${ribu(j.sel)} sel dirapikan; aturan nama disimpan.`);
        await muatSaya();
        await muat();
      }]], "Batal");
    let aktif = kolom;
    const pratinjau = async () => {
      const j = await api(`/api/rapikan?klien=${encodeURIComponent(S.kid)}&ym=${S.ym}&kolom=${aktif}`);
      $("#pratinjau-rapi").innerHTML = j.sel
        ? `<div class="hasil"><b>${ribu(j.sel)}</b> sel akan berubah:<div class="tabel-wrap tinggi"><table class="tbl"><thead><tr><th>Dari</th><th>Menjadi</th><th class="n">Sel</th></tr></thead><tbody>
          ${j.ubah.map((u) => `<tr><td>${esc(u.dari)}</td><td>${esc(u.ke)}</td><td class="n">${ribu(u.jumlah)}</td></tr>`).join("")}</tbody></table></div></div>`
        : '<div class="hasil">Tidak ada nama yang perlu dirapikan.</div>';
    };
    $$("#seg-rapi button").forEach((b) => (b.onclick = () => { aktif = b.dataset.k; $$("#seg-rapi button").forEach((x) => x.classList.toggle("aktif", x === b)); pratinjau(); }));
    $("#al-tambah").onclick = async () => {
      const dari = $("#al-dari").value.trim(), ke = $("#al-ke").value.trim();
      if (!dari || !ke) return;
      if (!admin()) return toast("Hanya admin yang bisa menambah aturan nama.");
      await api("/api/alias", { klien: S.kid, kolom: aktif, tambah: [{ dari, ke }] });
      $("#al-dari").value = ""; $("#al-ke").value = "";
      await muatSaya();
      pratinjau();
    };
    pratinjau();
  };

  $("#btn-unduh").onclick = async () => {
    await kirimAntre();
    try {
      await unduhPost("/api/ekspor/data", { klien: S.kid, ym: S.ym, filter: S.filter }, "Data.xlsx");
    } catch (e) { toast(e.message); }
  };

  // ------------------------------------------------------------------ Target
  function gambarTarget() {
    const k = klien(), d = S.data, t = d.targets || {};
    $("#sumber-target").textContent = d.sumber_target === "periode" ? `Target khusus ${namaPeriode(S.ym)}` : "Memakai target bawaan klien";
    const nilai = (v) => (v === undefined || v === null ? "" : String(v));
    const pc = (v) => (v ? String(+(v * 100).toFixed(3)).replace(".", ",") : "");
    $("#t-posts").innerHTML = PL.map((p) => `<label>${NAMA_PL[p]}<input name="posts.${p}" inputmode="numeric" value="${nilai((t.posts || {})[p])}"></label>`).join("");
    $("#t-er").innerHTML = SOSMED.map((p) => `<label>${NAMA_PL[p]}<input name="erPlatform.${p}" inputmode="decimal" value="${pc((t.erPlatform || {})[p])}"></label>`).join("");
    const f = $("#form-target");
    f.views.value = nilai(t.views); f.interactions.value = nilai(t.interactions); f.er.value = pc(t.er);
    $("#leg-kuota-kel").textContent = `Kuota per ${k.label_kelompok}`;
    $("#leg-kuota-kel").parentElement.hidden = !k.pakai_kelompok;
    const namaKel = new Map();
    S.r.kelompok.forEach((b) => namaKel.set(nameKey(b.nama), b.nama));
    Object.keys(t.perGroup || {}).forEach((n) => { if (!namaKel.has(nameKey(n))) namaKel.set(nameKey(n), n); });
    const kuotaKel = Object.fromEntries(Object.entries(t.perGroup || {}).map(([n, v]) => [nameKey(n), v]));
    $("#t-kelompok").innerHTML = [...namaKel.values()].sort((a, b) => a.localeCompare(b, "id"))
      .map((n) => `<label>${esc(n)}<input data-kuota="perGroup" data-nama="${esc(n)}" inputmode="numeric" value="${nilai(kuotaKel[nameKey(n)])}"></label>`).join("")
      || '<span class="muted">Belum ada kelompok di data.</span>';
    $("#t-akun").innerHTML = Object.entries(t.perAccount || {}).map(([n, v]) => `<label>${esc(n)}<input data-kuota="perAccount" data-nama="${esc(n)}" inputmode="numeric" value="${nilai(v)}"></label>`).join("")
      || '<span class="muted kosong-akun">Belum ada kuota akun.</span>';
    $("#dl-akun").innerHTML = S.r.akun.map((a) => `<option>${esc(a.nama)}</option>`).join("");
    const ro = !admin();
    $$("#form-target input").forEach((i) => (i.disabled = ro));
    $("#aksi-target").hidden = ro;
    $("#target-ro").hidden = !ro;
    $("#btn-target-reset").hidden = d.sumber_target !== "periode";
  }

  function tambahKuota(wadah, jenis, nama) {
    nama = nama.trim();
    if (!nama) return;
    if ($$(`input[data-kuota="${jenis}"]`).some((i) => nameKey(i.dataset.nama) === nameKey(nama))) return toast("Nama sudah ada di daftar.");
    $(".muted", wadah) && $(".muted", wadah).remove();
    wadah.insertAdjacentHTML("beforeend", `<label>${esc(nama)}<input data-kuota="${jenis}" data-nama="${esc(nama)}" inputmode="numeric"></label>`);
    wadah.lastElementChild.querySelector("input").focus();
  }
  $("#t-kel-tambah").onclick = () => { tambahKuota($("#t-kelompok"), "perGroup", $("#t-kel-nama").value); $("#t-kel-nama").value = ""; };
  $("#t-akun-tambah").onclick = () => { tambahKuota($("#t-akun"), "perAccount", $("#t-akun-nama").value); $("#t-akun-nama").value = ""; };

  function bacaFormTarget() {
    const f = $("#form-target");
    const t = { posts: {}, erPlatform: {}, perGroup: {}, perAccount: {}, views: f.views.value, interactions: f.interactions.value, er: f.er.value };
    $$("#t-posts input").forEach((i) => (t.posts[i.name.split(".")[1]] = i.value));
    $$("#t-er input").forEach((i) => (t.erPlatform[i.name.split(".")[1]] = i.value));
    $$("input[data-kuota]").forEach((i) => { if (i.value.trim()) t[i.dataset.kuota][i.dataset.nama] = i.value; });
    return t;
  }
  async function simpanTarget(bawaan) {
    await api("/api/target", { klien: S.kid, ym: S.ym, bawaan, targets: bacaFormTarget() });
    if (bawaan) await muatSaya();
    await muat();
    toast(bawaan ? "Target bawaan klien disimpan." : `Target ${namaPeriode(S.ym)} disimpan.`);
  }
  $("#form-target").onsubmit = (e) => { e.preventDefault(); simpanTarget(false).catch((x) => toast(x.message)); };
  $("#btn-target-bawaan").onclick = () => simpanTarget(true).catch((x) => toast(x.message));
  $("#btn-target-salin").onclick = async () => {
    try { await api("/api/target/salin", { klien: S.kid, ym: S.ym }); await muat(); toast("Target bulan lalu disalin. Ubah lalu simpan bila perlu."); }
    catch (e) { toast(e.message); }
  };
  $("#btn-target-reset").onclick = async () => {
    await api("/api/target", { klien: S.kid, ym: S.ym, bawaan: false, targets: {} });
    await muat();
    toast("Periode ini kembali memakai target bawaan klien.");
  };

  // ------------------------------------------------------------------ Ekspor
  function gambarEkspor() {
    const q = `?klien=${encodeURIComponent(S.kid)}&ym=${S.ym}`;
    $("#dl-xlsx").dataset.url = "/api/ekspor/xlsx" + q;
    $("#dl-pptx").dataset.url = "/api/ekspor/pptx" + q;
    const n = S.r.perlu_dicek;
    const pw = $("#peringatan-ekspor");
    pw.hidden = !n;
    if (n) {
      const rinci = Object.entries(S.r.flag_hitung).map(([f, c]) => `${S.data.flag_info[f][0].toLowerCase()} ${c}`).join(", ");
      pw.innerHTML = `⚠ Masih ada <b>${ribu(n)}</b> baris perlu dicek (${esc(rinci)}). Laporan tetap bisa diunduh, tetapi sebaiknya diperbaiki dulu.
        <a href="#" id="pw-cek">Buka baris tersebut</a>`;
      $("#pw-cek").onclick = (e) => { e.preventDefault(); bukaPerluDicek(); };
    }
  }
  const unduhTombol = (b, teks, nama) => (b.onclick = async () => {
    b.disabled = true;
    toast(teks);
    try { await kirimAntre(); await unduhPost(b.dataset.url, undefined, nama); } catch (e) { toast(e.message); }
    b.disabled = false;
  });
  unduhTombol($("#dl-xlsx"), "Menyiapkan Excel…", "Laporan.xlsx");
  unduhTombol($("#dl-pptx"), "Menyiapkan PowerPoint…", "Laporan.pptx");
  unduhTombol($("#dl-template"), "Menyiapkan template…", "Template Pantau Publikasi.xlsx");
  $("#btn-ke-data").onclick = () => tab("data");

  // ------------------------------------------------------------------ Pengaturan
  async function muatSaya() {
    const j = await api("/api/saya");
    S.klien = j.klien;
    S.saya = j.pengguna;
  }

  function isiFormKlien(k) {
    const f = $("#form-klien");
    f.id.value = k ? k.id : "";
    f.nama.value = k ? k.nama : "";
    f.label_kelompok.value = k ? k.label_kelompok : "Kelompok";
    f.pakai_kelompok.checked = k ? k.pakai_kelompok !== false : true;
    f.warna.value = k ? k.warna : "#B3261E";
    const img = $("#logo-pratinjau");
    img.hidden = !(k && k.logo);
    if (k && k.logo) {
      fetch(`/api/logo?klien=${encodeURIComponent(k.id)}&t=${Date.now()}`).then((r) => (r.ok ? r.blob() : null))
        .then((b) => { if (b) { URL.revokeObjectURL(img.src); img.src = URL.createObjectURL(b); } });
    }
    $("#logo-file").parentElement.hidden = $("#logo-hapus").hidden = $("#btn-klien-hapus").hidden = !k;
  }

  function gambarPengaturan() {
    const k = klien();
    const ad = admin();
    isiFormKlien(k);
    $$("#form-klien input, #form-klien button").forEach((x) => (x.disabled = !ad));
    $("#btn-klien-baru").hidden = !ad;
    const al = (k && k.alias) || {};
    const baris = [];
    for (const kol of ["kelompok", "akun"]) for (const [dari, ke] of Object.entries(al[kol] || {})) baris.push([kol, dari, ke]);
    $("#daftar-alias").innerHTML = baris.length
      ? baris.map(([kol, dari, ke]) => `<div class="alias-baris"><span class="muted">${kol === "kelompok" ? esc(k.label_kelompok) : "Akun"}</span>
        <code>${esc(dari)}</code> → <b>${esc(ke)}</b>${ad ? `<button class="btn kecil" data-kol="${kol}" data-dari="${esc(dari)}">Hapus</button>` : ""}</div>`).join("")
      : '<p class="muted">Belum ada aturan. Aturan tercipta saat "Rapikan nama" diterapkan atau ditambah manual.</p>';
    $$("#daftar-alias button").forEach((b) => (b.onclick = async () => {
      await api("/api/alias", { klien: S.kid, kolom: b.dataset.kol, hapus: [b.dataset.dari] });
      await muatSaya(); gambarPengaturan();
    }));
    $("#kartu-pengguna").hidden = !ad;
    if (ad) gambarPengguna();
  }

  async function gambarPengguna() {
    const j = await api("/api/pengguna");
    const namaK = Object.fromEntries(S.klien.map((k) => [k.id, k.nama]));
    $("#tbl-pengguna").innerHTML = `<thead><tr><th>Email</th><th>Nama</th><th>Peran</th><th>Klien</th><th></th></tr></thead><tbody>
      ${j.pengguna.map((u) => `<tr><td>${esc(u.email)}</td><td>${esc(u.nama)}</td><td>${u.peran === "admin" ? "Admin" : "Operator"}</td>
        <td>${esc((u.klien || []).map((x) => namaK[x] || "(klien lain)").join(", "))}</td>
        <td><button class="btn kecil" data-ubah="${esc(u.email)}">Ubah</button></td></tr>`).join("")}</tbody>`;
    const f = $("#form-pengguna");
    const isiCek = (dipilih) => {
      $("#pengguna-klien").innerHTML = "<span class=\"muted kecil\">Klien yang boleh dilihat:</span>" + S.klien.map((k) =>
        `<label class="cek"><input type="checkbox" value="${k.id}"${dipilih.includes(k.id) ? " checked" : ""}> ${esc(k.nama)}</label>`).join("");
    };
    isiCek(S.kid ? [S.kid] : []);
    $$("#tbl-pengguna button[data-ubah]").forEach((b) => (b.onclick = () => {
      const u = j.pengguna.find((x) => x.email === b.dataset.ubah);
      f.email.value = u.email; f.nama.value = u.nama; f.peran.value = u.peran; f.sandi.value = "";
      isiCek(u.klien || []);
      f.email.focus();
    }));
  }

  $("#form-pengguna").onsubmit = async (e) => {
    e.preventDefault();
    const f = e.target;
    try {
      await api("/api/pengguna", { email: f.email.value, nama: f.nama.value, peran: f.peran.value, sandi: f.sandi.value,
        klien: $$("#pengguna-klien input:checked").map((i) => i.value) });
      f.reset();
      toast("Pengguna disimpan. Berikan email dan kata sandi awal kepada yang bersangkutan.");
      gambarPengguna();
    } catch (err) { toast(err.message); }
  };
  $("#form-sandi").onsubmit = async (e) => {
    e.preventDefault();
    try { await api("/api/sandi", { lama: e.target.lama.value, baru: e.target.baru.value }); e.target.reset(); toast("Kata sandi diganti."); }
    catch (err) { toast(err.message); }
  };

  $("#btn-klien-baru").onclick = () => { isiFormKlien(null); $("#form-klien").nama.focus(); };
  $("#form-klien").onsubmit = async (e) => {
    e.preventDefault();
    const f = e.target;
    try {
      const j = await api("/api/klien", { id: f.id.value || undefined, nama: f.nama.value, label_kelompok: f.label_kelompok.value,
        pakai_kelompok: f.pakai_kelompok.checked, warna: f.warna.value });
      toast("Klien disimpan.");
      await mulai(j.id, f.id.value ? S.ym : undefined);
      tab("pengaturan");
    } catch (err) { toast(err.message); }
  };
  $("#logo-file").onchange = async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      await api("/api/klien/logo", { klien: S.kid, data: await bacaBase64(f) });
      await muatSaya(); gambarPengaturan(); toast("Logo disimpan (diperkecil ke maks. 256 px).");
    } catch (err) { toast(err.message); }
    e.target.value = "";
  };
  $("#logo-hapus").onclick = async () => { await api("/api/klien/logo", { klien: S.kid, hapus: true }); await muatSaya(); gambarPengaturan(); };
  $("#btn-klien-hapus").onclick = () => {
    const k = klien();
    dialog("Hapus klien", `<p>Langkah 1 dari 2. Menghapus <b>${esc(k.nama)}</b> juga menghapus <b>semua periode, postingan, target, logo, dan aturan nama</b> klien ini.
      Tindakan ini tidak bisa dibatalkan.</p>`, [["Lanjut", () => {
      setTimeout(() => dialog("Konfirmasi hapus klien", `<p>Langkah 2 dari 2. Ketik nama klien persis untuk menghapus:</p><p><code>${esc(k.nama)}</code></p>
        <input id="in-konfirm" autocomplete="off">`, [["Hapus permanen", async () => {
        await api("/api/klien/hapus", { id: k.id, konfirmasi: $("#in-konfirm").value });
        toast("Klien dihapus.");
        S.kid = null;
        await mulai();
      }, "bahaya"]]), 0);
    }, "bahaya"]]);
  };

  // ------------------------------------------------------------------ dialog umum
  function dialog(judul, html, tombol = [], batal = "Batal") {
    const d = $("#dlg");
    $("#dlg-judul").textContent = judul;
    $("#dlg-badan").innerHTML = html;
    const aksi = $("#dlg-aksi");
    aksi.innerHTML = "";
    for (const [label, fn, kelas] of tombol) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "btn " + (kelas || "utama");
      b.textContent = label;
      b.onclick = async () => {
        b.disabled = true;
        try {
          const tetap = await fn();
          if (tetap !== false) d.close();
        } catch (err) { toast(err.message, 6000); }
        b.disabled = false;
      };
      aksi.appendChild(b);
    }
    const bt = document.createElement("button");
    bt.className = "btn";
    bt.value = "batal";
    bt.textContent = batal;
    aksi.appendChild(bt);
    d.showModal();
  }

  window.addEventListener("unhandledrejection", (e) => { toast(e.reason && e.reason.message ? e.reason.message : "Terjadi kesalahan"); });

  pasangTema();
  Promise.resolve(window.PANTAU_SIAP || { demo: false }).then((m) => {
    S.demo = m.demo;
    if (m.demo) {
      $("#btn-keluar").hidden = true;
      const b = $("#btn-reset-demo");
      b.hidden = false;
      b.onclick = () => dialog("Reset data demo", "<p>Kembalikan semua data demo di browser ini ke data simulasi awal?</p>", [["Reset", async () => {
        S.antre = {};
        await window.PANTAU_RESET_DEMO();
        await mulai();
        toast("Data demo dikembalikan ke awal.");
      }, "bahaya"]]);
    }
    return mulai();
  }).catch((e) => { if (!S.demo) tampilMasuk(); console.error(e); });
})();

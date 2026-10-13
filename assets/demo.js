/* Mode demo (GitHub Pages / hosting statis).
   Bila tidak ada server Python, scripts/server.py dijalankan di browser lewat Pyodide: permintaan ke /api/... dialihkan
   ke Handler yang sama, data simulasi disimpan di IndexedDB browser pengunjung (tidak dibagikan ke siapa pun). */
(() => {
  "use strict";
  const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
  const SKRIP = ["inti.py", "penyimpanan.py", "impor_workbook.py", "ekspor_xlsx.py", "ekspor_pptx.py", "server.py", "buat_simulasi.py"];
  const fetchAsli = window.fetch.bind(window);
  let py = null, tangani = null, antrean = Promise.resolve();

  function layar(teks, persen) {
    let el = document.getElementById("demo-muat");
    if (!el) {
      el = document.createElement("div");
      el.id = "demo-muat";
      el.className = "demo-muat";
      el.innerHTML = '<div class="kartu"><div class="brand"><span class="brand-mark"></span><h1>Pantau Publikasi</h1></div>' +
        '<p class="muted">Mode demo: menyiapkan aplikasi di browser Anda. Pemuatan pertama ±10–30 detik.</p>' +
        '<div class="meter"><i></i></div><p class="kecil muted" id="demo-teks"></p></div>';
      document.body.appendChild(el);
    }
    document.getElementById("demo-teks").textContent = teks;
    el.querySelector(".meter i").style.width = persen + "%";
  }

  function muatSkrip(src) {
    return new Promise((ok, gagal) => {
      const s = document.createElement("script");
      s.src = src;
      s.onload = ok;
      s.onerror = () => gagal(new Error("gagal memuat " + src));
      document.head.appendChild(s);
    });
  }

  const simpanFS = () => new Promise((ok) => py.FS.syncfs(false, () => ok()));

  async function siapkanPython() {
    layar("Memuat Python…", 10);
    await muatSkrip(PYODIDE + "pyodide.js");
    py = await loadPyodide({ indexURL: PYODIDE });
    layar("Memuat pustaka Excel & PowerPoint…", 35);
    await py.loadPackage(["micropip", "lxml", "pillow", "hashlib"]);
    await py.runPythonAsync(`
import micropip
await micropip.install(["openpyxl", "python-pptx"])
`);
    layar("Memuat aplikasi…", 75);
    py.FS.mkdir("/app");
    py.FS.mkdir("/app/scripts");
    for (const f of SKRIP) {
      const r = await fetchAsli("scripts/" + f, { cache: "no-store" });
      if (!r.ok) throw new Error("gagal memuat scripts/" + f);
      py.FS.writeFile("/app/scripts/" + f, await r.text());
    }
    py.FS.mkdir("/data");
    py.FS.mount(py.FS.filesystems.IDBFS, {}, "/data");
    await new Promise((ok) => py.FS.syncfs(true, () => ok()));
    layar("Menyiapkan data simulasi…", 90);
    await py.runPythonAsync(`
import os, sys, json, secrets, time
from io import BytesIO
os.environ["PANTAU_DATA"] = "/data/pantau"
sys.path.insert(0, "/app/scripts")
import penyimpanan as ps, server, buat_simulasi
if not ps.semua_pengguna():
    sys.argv = ["buat_simulasi"]
    buat_simulasi.main()
TOKEN = secrets.token_urlsafe(24)
server.SESI[TOKEN] = ("admin@pantau.test", time.time() + 10**9)

class Kepala(dict):
    def get(self, k, d=None):
        return dict.get(self, k.lower(), d)

def tangani(metode, jalur, isi):
    h = server.Handler.__new__(server.Handler)
    h.command, h.path, h.request_version = metode, jalur, "HTTP/1.1"
    h.requestline = f"{metode} {jalur} HTTP/1.1"
    h.client_address = ("127.0.0.1", 0)
    h.close_connection = True
    h.directory = ps.ROOT
    b = bytes(isi.to_py()) if isi is not None else b""
    h.headers = Kepala({"cookie": "sesi=" + TOKEN, "content-length": str(len(b)), "content-type": "application/json"})
    h.rfile, h.wfile = BytesIO(b), BytesIO()
    (h.do_POST if metode == "POST" else h.do_GET)()
    mentah = h.wfile.getvalue()
    kepala, _, badan = mentah.partition(b"\\r\\n\\r\\n")
    baris = kepala.decode("latin-1").split("\\r\\n")
    status = int(baris[0].split(" ")[1])
    hdr = {}
    for x in baris[1:]:
        k, _, v = x.partition(": ")
        if k.lower() not in ("set-cookie", "content-length"):
            hdr[k] = v
    return status, json.dumps(hdr), badan
`);
    tangani = py.globals.get("tangani");
  }

  async function panggil(metode, jalur, badan) {
    const kerja = antrean.then(async () => {
      const isi = badan === undefined || badan === null ? null : typeof badan === "string" ? new TextEncoder().encode(badan) : new Uint8Array(badan);
      const hasil = tangani(metode, jalur, isi);
      const [status, hdr, b] = hasil.toJs();
      hasil.destroy();
      const bytes = b.toJs ? b.toJs() : b;
      if (metode === "POST") await simpanFS();
      return new Response(new Uint8Array(bytes), { status, headers: JSON.parse(hdr) });
    });
    antrean = kerja.catch(() => {});
    return kerja;
  }

  window.PANTAU_SIAP = (async () => {
    let server = false;
    try {
      const r = await fetchAsli("api/saya", { cache: "no-store" });
      server = r.ok && (r.headers.get("Content-Type") || "").includes("json");
    } catch (e) { /* tidak ada server */ }
    if (server) return { demo: false };
    try {
      await siapkanPython();
    } catch (e) {
      layar("Gagal menyiapkan mode demo: " + e.message + ". Muat ulang halaman untuk mencoba lagi.", 100);
      throw e;
    }
    window.fetch = (input, opt = {}) => {
      const url = new URL(typeof input === "string" ? input : input.url, location.href);
      const i = url.pathname.indexOf("/api/");
      if (i < 0) return fetchAsli(input, opt);
      return panggil((opt.method || "GET").toUpperCase(), url.pathname.slice(i) + url.search, opt.body);
    };
    window.PANTAU_RESET_DEMO = async () => {
      py.runPython(`
import shutil
shutil.rmtree(ps.DATA, ignore_errors=True)
sys.argv = ["buat_simulasi"]
buat_simulasi.main()
server.SESI[TOKEN] = ("admin@pantau.test", time.time() + 10**9)
`);
      await simpanFS();
    };
    document.getElementById("demo-muat").remove();
    return { demo: true };
  })();
})();

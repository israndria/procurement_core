"""
order_bot.py — Playwright automation untuk katalog.inaproc.id.

Strategi: User buka Edge sendiri (via 'Buka Edge untuk Order.bat') dengan
--remote-debugging-port=9222, login manual, lalu bot nempel via CDP.
Bot tidak pernah membuka atau mengelola browser — hanya mengontrol Edge
yang sudah berjalan dan sudah login.
"""

import json
import os
import sys
import asyncio

from playwright.sync_api import sync_playwright, Page, BrowserContext

if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

from matcher import cocokkan_variasi, ada_variasi

URL_HOME        = "https://katalog.inaproc.id"
CDP_URL         = "http://127.0.0.1:9222"

SEL_VARIANT_BTN = "button.text-caption-sm-semibold"
SEL_QTY_INPUT   = "input[id*='quantity']"
SEL_TAMBAH_BTN  = "button:has-text('Tambah Keranjang')"
SEL_TIDAK_AKTIF = "text=Tidak Aktif"

LOGIN_OK          = "ok"
LOGIN_EXPIRED     = "expired"
LOGIN_BELUM_LOGIN = "belum_login"


class OrderBot:
    def __init__(self, on_log=None):
        self._playwright = None
        self._browser    = None
        self._page: Page | None = None
        self.on_log = on_log or print

    def _log(self, msg: str):
        self.on_log(msg)

    # ------------------------------------------------------------------ #

    def hubungkan(self):
        """
        Hubungkan ke Edge yang sudah dibuka user via CDP.
        Edge harus sudah berjalan dengan --remote-debugging-port=9222.
        """
        self._log("Menghubungkan ke Edge...")
        self._playwright = sync_playwright().start()
        self._browser = self._playwright.chromium.connect_over_cdp(CDP_URL)

        # Ambil tab yang sedang aktif atau tab pertama yang ada
        ctx = self._browser.contexts[0] if self._browser.contexts else None
        if not ctx:
            raise RuntimeError("Tidak ada context di Edge. Pastikan Edge sudah terbuka.")

        # Cari tab katalog.inaproc.id, atau pakai tab pertama
        target_page = None
        for pg in ctx.pages:
            if "katalog.inaproc.id" in pg.url or "inaproc" in pg.url:
                target_page = pg
                break
        if target_page is None:
            target_page = ctx.pages[0] if ctx.pages else ctx.new_page()

        self._page = target_page
        self._log("Terhubung ke Edge.")

    def diagnosa_session(self) -> str:
        """
        Cek status login via /api/auth/session.
        Return: LOGIN_OK | LOGIN_EXPIRED | LOGIN_BELUM_LOGIN
        """
        try:
            # Pastikan kita di halaman katalog.inaproc.id
            if "katalog.inaproc.id" not in self._page.url:
                self._page.goto(URL_HOME, wait_until="domcontentloaded", timeout=15000)
                self._page.wait_for_timeout(1500)

            result = self._page.evaluate("""
            async () => {
                try {
                    const resp = await fetch('/api/auth/session', {credentials: 'include'});
                    const data = await resp.json();
                    return JSON.stringify(data);
                } catch(e) {
                    return '{"error":"fetch_failed"}';
                }
            }
            """)
            data = json.loads(result)

            if data.get("error") == "ERROR_TOKEN_REFRESH":
                self._log("Status: Session expired — login ulang di Edge lalu klik 'Cek Ulang'.")
                return LOGIN_EXPIRED

            token = data.get("token", {})
            if token and isinstance(token, dict) and any(
                k in token for k in ("accessToken", "sub", "userId", "name")
            ):
                self._log("Status: Sudah login.")
                return LOGIN_OK

            # Fallback: cek tombol Masuk di DOM
            masuk = self._page.query_selector("button:has-text('Masuk')")
            if masuk and masuk.is_visible():
                self._log("Status: Belum login (tombol Masuk terlihat).")
                return LOGIN_BELUM_LOGIN

            self._log("Status: Sudah login.")
            return LOGIN_OK

        except Exception as e:
            self._log(f"Diagnosa error: {e}")
            return LOGIN_BELUM_LOGIN

    def cek_login(self) -> bool:
        return self.diagnosa_session() == LOGIN_OK

    def proses_item(self, item: dict) -> dict:
        nama = item["nama_barang"]
        qty  = item["kuantitas"]
        link = item["link"]
        base = {"nama_barang": nama, "kuantitas": qty, "link": link}

        self._log(f"Memproses: {nama[:50]}")
        try:
            self._page.goto(link, wait_until="domcontentloaded", timeout=20000)
            self._page.wait_for_timeout(1500)

            # Deteksi redirect ke login (session expired)
            if "login" in self._page.url or "signin" in self._page.url:
                return {**base, "status": "session_expired", "pesan": "Session expired — login ulang di Edge"}

            if self._cek_tidak_aktif():
                self._log("  ⏭ Skip — produk tidak aktif")
                return {**base, "status": "skip_tidak_aktif", "pesan": "Produk tidak aktif"}

            variasi_tersedia = self._ambil_variasi()
            if ada_variasi(variasi_tersedia):
                cocok = cocokkan_variasi(nama, variasi_tersedia)
                if cocok is None:
                    self._log(f"  ⚠️  Variasi tidak cocok. Tersedia: {variasi_tersedia}")
                    return {
                        **base,
                        "status": "variasi_tidak_cocok",
                        "pesan": f"Variasi tidak cocok. Tersedia: {', '.join(variasi_tersedia)}",
                    }
                self._pilih_variasi(cocok)
                self._log(f"  Variasi: {cocok}")
                self._page.wait_for_timeout(800)

            self._isi_kuantitas(qty)
            self._log(f"  Qty: {qty}")
            self._page.wait_for_timeout(500)

            tambah_btn = self._page.query_selector(SEL_TAMBAH_BTN)
            if tambah_btn is None:
                return {**base, "status": "error", "pesan": "Tombol 'Tambah Keranjang' tidak ditemukan"}
            if tambah_btn.is_disabled():
                return {**base, "status": "skip_tidak_aktif", "pesan": "Tombol disabled"}

            tambah_btn.click()
            self._page.wait_for_timeout(1500)
            self._log("  ✅ Berhasil")
            return {**base, "status": "berhasil", "pesan": "Berhasil ditambahkan ke keranjang"}

        except Exception as e:
            self._log(f"  ❌ Error: {e}")
            return {**base, "status": "error", "pesan": str(e)}

    def tutup(self):
        """Putuskan koneksi CDP (tidak menutup Edge milik user)."""
        try:
            if self._browser:
                self._browser.close()
        except Exception:
            pass
        try:
            if self._playwright:
                self._playwright.stop()
        except Exception:
            pass
        self._browser = None
        self._page = None
        self._playwright = None
        self._log("Koneksi CDP diputus (Edge tetap berjalan).")

    # ------------------------------------------------------------------ #
    def _cek_tidak_aktif(self) -> bool:
        # Tunggu sampai tombol muncul (max 5 detik) — halaman Next.js butuh waktu hydrate
        try:
            self._page.wait_for_selector(
                f"{SEL_TAMBAH_BTN}, {SEL_TIDAK_AKTIF}",
                timeout=5000,
            )
        except Exception:
            pass  # timeout → lanjut cek manual

        if self._page.query_selector(SEL_TIDAK_AKTIF):
            return True
        return self._page.query_selector(SEL_TAMBAH_BTN) is None

    def _ambil_variasi(self) -> list[str]:
        btns = self._page.query_selector_all(SEL_VARIANT_BTN)
        return [b.inner_text().strip() for b in btns if b.inner_text().strip()]

    def _pilih_variasi(self, nama_variasi: str):
        for btn in self._page.query_selector_all(SEL_VARIANT_BTN):
            if btn.inner_text().strip().lower() == nama_variasi.lower():
                btn.click()
                return
        raise ValueError(f"Tombol variasi '{nama_variasi}' tidak ditemukan")

    def _isi_kuantitas(self, qty: int):
        inp = self._page.query_selector(SEL_QTY_INPUT)
        if inp is None:
            return
        inp.click()
        inp.select_text()
        inp.type(str(qty))

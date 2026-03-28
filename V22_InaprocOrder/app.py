"""
V22 Inaproc Order Bot
Otomatisasi pemesanan di katalog.inaproc.id.

Flow:
1. Upload Excel pesanan dari PPK
2. Buka Edge sendiri via 'Buka Edge untuk Order.bat', login manual
3. Klik 'Hubungkan ke Edge' — bot nempel via CDP
4. Mulai proses — bot navigasi ke setiap link, pilih variasi, isi qty, tambah ke keranjang
"""

import os
import sys
import tempfile

import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_reader import baca_pesanan
from order_bot import OrderBot, LOGIN_EXPIRED, LOGIN_BELUM_LOGIN, LOGIN_OK

st.set_page_config(page_title="V22 Inaproc Order Bot", page_icon="🛒", layout="centered")
st.title("🛒 Inaproc Order Bot")
st.caption("Otomatisasi tambah ke keranjang di katalog.inaproc.id")

for key, default in [
    ("pesanan", []),
    ("terhubung", False),
    ("hasil", []),
    ("proses_selesai", False),
    ("session_expired_midprocess", False),
    ("item_terakhir", 0),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ------------------------------------------------------------------ #
# BAGIAN 1 — Upload Excel
st.markdown("### 1. Upload File Excel Pesanan")

uploaded = st.file_uploader("Pilih file Excel dari PPK (.xlsx)", type=["xlsx"])

if uploaded:
    # Hanya parse ulang jika file berbeda — mencegah reset state saat rerun
    if uploaded.name != st.session_state.get("uploaded_filename"):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded.read())
            tmp_path = tmp.name
        try:
            pesanan = baca_pesanan(tmp_path)
            st.session_state.uploaded_filename = uploaded.name
            st.session_state.pesanan = pesanan
            st.session_state.hasil = []
            st.session_state.proses_selesai = False
            st.session_state.session_expired_midprocess = False
            st.session_state.item_terakhir = 0
        except Exception as e:
            st.error(f"Gagal membaca Excel: {e}")
        finally:
            os.unlink(tmp_path)

    if st.session_state.pesanan:
        st.success(f"✅ {len(st.session_state.pesanan)} item ditemukan")
        import pandas as pd
        df = pd.DataFrame([
            {"No": i+1, "Nama Barang": p["nama_barang"].replace("\n"," "), "Qty": p["kuantitas"], "Link": p["link"]}
            for i, p in enumerate(st.session_state.pesanan)
        ])
        st.dataframe(df, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------ #
# BAGIAN 2 — Hubungkan ke Edge
if st.session_state.pesanan:
    st.markdown("---")
    st.markdown("### 2. Hubungkan ke Edge")

    if not st.session_state.terhubung:
        st.info(
            "**Langkah sebelum klik tombol di bawah:**\n"
            "1. Jalankan **`Buka Edge untuk Order.bat`** (ada di folder ini)\n"
            "2. Edge akan terbuka dan menuju katalog.inaproc.id\n"
            "3. **Login** seperti biasa — pastikan nama/avatar sudah muncul\n"
            "4. Biarkan Edge tetap terbuka, lalu klik tombol di bawah"
        )

        col1, col2 = st.columns([2, 1])
        with col1:
            if st.button("🔌 Cek Koneksi & Login Edge", type="primary"):
                with st.spinner("Mengecek koneksi ke Edge..."):
                    try:
                        bot = OrderBot(on_log=lambda m: None)
                        bot.hubungkan()
                        status = bot.diagnosa_session()
                        bot.tutup()  # Tutup setelah cek — bot dibuat ulang saat proses

                        if status == LOGIN_OK:
                            st.session_state.terhubung = True
                            st.rerun()

                        elif status == LOGIN_EXPIRED:
                            st.error(
                                "**Session Expired (ERROR_TOKEN_REFRESH)**\n\n"
                                "Access token sudah kedaluwarsa. "
                                "Di Edge, logout lalu login ulang ke katalog.inaproc.id, "
                                "pastikan nama/avatar muncul, lalu klik tombol ini lagi."
                            )

                        else:  # LOGIN_BELUM_LOGIN
                            st.error(
                                "**Belum Login**\n\n"
                                "Belum terdeteksi sesi aktif di Edge. "
                                "Login dulu di Edge, pastikan nama/avatar muncul, lalu klik tombol ini lagi."
                            )

                    except Exception as e:
                        st.error(
                            f"Gagal terhubung ke Edge: {e}\n\n"
                            "Pastikan:\n"
                            "- `Buka Edge untuk Order.bat` sudah dijalankan\n"
                            "- Edge sedang terbuka (jangan ditutup)\n"
                            "- Port 9222 tidak dipakai aplikasi lain"
                        )

        with col2:
            bat_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Buka Edge untuk Order.bat")
            if os.path.exists(bat_path):
                if st.button("📂 Buka .bat sekarang"):
                    os.startfile(bat_path)
                    st.info("Edge sedang dibuka... tunggu sebentar lalu login.")

    else:
        st.success("✅ Edge terverifikasi login — siap memproses")

        col_cek, col_putus = st.columns([1, 1])
        with col_cek:
            if st.button("🔄 Cek Ulang Status Login"):
                with st.spinner("Mengecek..."):
                    try:
                        bot = OrderBot(on_log=lambda m: None)
                        bot.hubungkan()
                        status = bot.diagnosa_session()
                        bot.tutup()
                        if status == LOGIN_OK:
                            st.success("✅ Masih login.")
                        elif status == LOGIN_EXPIRED:
                            st.error("Session expired — login ulang di Edge lalu Cek Ulang.")
                        else:
                            st.error("Belum login — login di Edge lalu Cek Ulang.")
                    except Exception as e:
                        st.error(f"Gagal cek: {e}")

        with col_putus:
            if st.button("🔌 Reset Koneksi"):
                st.session_state.terhubung = False
                st.rerun()

# ------------------------------------------------------------------ #
# BAGIAN 3 — Mulai Proses
if st.session_state.pesanan and st.session_state.terhubung and not st.session_state.proses_selesai:
    st.markdown("---")
    st.markdown("### 3. Mulai Proses")

    pesanan = st.session_state.pesanan
    mulai_dari = st.session_state.item_terakhir

    if st.session_state.session_expired_midprocess:
        st.warning(
            f"⚠️ **Session expired di tengah proses** — item ke-{mulai_dari + 1} dari {len(pesanan)}\n\n"
            "Login ulang di Edge (pastikan nama/avatar muncul), lalu klik tombol di bawah."
        )
        if st.button("✅ Saya sudah login, Lanjutkan Proses"):
            with st.spinner("Mengecek login..."):
                try:
                    bot = OrderBot(on_log=lambda m: None)
                    bot.hubungkan()
                    status = bot.diagnosa_session()
                    bot.tutup()
                    if status == LOGIN_OK:
                        st.session_state.session_expired_midprocess = False
                        st.rerun()
                    else:
                        st.error("Masih belum terdeteksi login. Login ulang dulu di Edge.")
                except Exception as e:
                    st.error(f"Gagal cek login: {e}")
    else:
        if mulai_dari == 0:
            st.info(f"Siap memproses **{len(pesanan)} item** ke keranjang.")
        else:
            st.info(f"Lanjutkan dari item ke-**{mulai_dari + 1}** (dari {len(pesanan)} item).")

        label_btn = "▶️ Mulai Proses Semua Item" if mulai_dari == 0 else f"▶️ Lanjutkan dari Item {mulai_dari + 1}"
        if st.button(label_btn, type="primary"):
            total = len(pesanan)
            hasil_baru = list(st.session_state.hasil)

            progress_bar = st.progress(mulai_dari / total if total else 0, text="Memulai...")
            status_box   = st.empty()
            log_box      = st.empty()
            log_lines    = []

            def log_fn(msg):
                log_lines.append(msg)
                log_box.code("\n".join(log_lines[-8:]))

            expired_di = None

            try:
                # Buat bot baru di thread ini — Playwright sync API harus dibuat & dipakai di thread yang sama
                bot = OrderBot(on_log=log_fn)
                bot.hubungkan()

                for i in range(mulai_dari, total):
                    item = pesanan[i]
                    nama_pendek = item["nama_barang"].replace("\n", " ")[:45]
                    progress_bar.progress(i / total, text=f"Item {i+1}/{total}: {nama_pendek}")
                    status_box.caption(f"Memproses: {nama_pendek}")

                    result = bot.proses_item(item)
                    hasil_baru.append(result)

                    if result["status"] == "session_expired":
                        expired_di = i
                        break

            except Exception as e:
                import traceback
                st.error(f"❌ Error tidak terduga:\n\n```\n{traceback.format_exc()}\n```")
                st.stop()
            finally:
                try: bot.tutup()
                except: pass

            progress_bar.progress(1.0 if expired_di is None else expired_di / total,
                                   text="Selesai" if expired_di is None else "Terhenti — session expired")
            st.session_state.hasil = hasil_baru

            if expired_di is not None:
                st.session_state.item_terakhir = expired_di
                st.session_state.session_expired_midprocess = True
            else:
                st.session_state.proses_selesai = True

            st.rerun()

# ------------------------------------------------------------------ #
# BAGIAN 4 — Hasil
if st.session_state.proses_selesai and st.session_state.hasil:
    st.markdown("---")
    st.markdown("### ✅ Hasil Proses")

    hasil = st.session_state.hasil
    n_berhasil = sum(1 for h in hasil if h["status"] == "berhasil")
    n_skip     = sum(1 for h in hasil if h["status"] == "skip_tidak_aktif")
    n_variasi  = sum(1 for h in hasil if h["status"] == "variasi_tidak_cocok")
    n_error    = sum(1 for h in hasil if h["status"] == "error")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("✅ Berhasil", n_berhasil)
    c2.metric("⏭ Skip", n_skip)
    c3.metric("⚠️ Variasi ?", n_variasi)
    c4.metric("❌ Error", n_error)

    IKON = {"berhasil":"✅","skip_tidak_aktif":"⏭","variasi_tidak_cocok":"⚠️","error":"❌","session_expired":"🔒"}
    import pandas as pd
    df = pd.DataFrame([
        {
            "Status": IKON.get(h["status"],"?"),
            "Nama Barang": h["nama_barang"].replace("\n"," ")[:60],
            "Qty": h["kuantitas"],
            "Keterangan": h["pesan"],
        }
        for h in hasil
    ])
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    if st.button("🔄 Proses Ulang / File Baru"):
        for k in ["pesanan","terhubung","hasil","proses_selesai","session_expired_midprocess","item_terakhir","uploaded_filename"]:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

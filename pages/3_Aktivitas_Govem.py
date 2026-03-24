import streamlit as st
import os
import sys

# Pastikan Aktivitas_Govem bisa diimport
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from Aktivitas_Govem import script_manager

st.title("Aktivitas Govem")
st.markdown("### Pusat Kontrol Otomatisasi Aktivitas")
st.markdown("---")

st.sidebar.info("Pastikan Anda tidak sedang menggunakan mouse/keyboard saat skrip berjalan.")

DIR_SUAMI = os.path.join(BASE_DIR, "Aktivitas_Govem", "Aktivitas")
DIR_ISTRI = os.path.join(BASE_DIR, "Aktivitas_Govem", "Aktivitas Istri")

tab1, tab2 = st.tabs(["Suami", "Istri"])

def render_activity_tab(directory, label, chrome_profile):
    st.header(f"{label}")

    scripts = script_manager.get_scripts(directory)

    if not scripts:
        st.warning(f"Tidak ada skrip ditemukan di folder: {directory}")
        return

    script_map = {script_manager.format_display_name(s): s for s in scripts}

    selected_name = st.selectbox(
        f"Pilih Aktivitas:",
        list(script_map.keys()),
        key=f"akt_select_{label}"
    )

    if selected_name:
        selected_script = script_map[selected_name]
        script_path = os.path.join(directory, selected_script)

        st.caption(f"File: `{selected_script}`")

        with st.expander("Lihat Detail Aktivitas", expanded=True):
            descriptions = script_manager.extract_description(script_path)
            st.markdown("**Aksi yang akan dilakukan (Teks Input):**")
            for desc in descriptions:
                st.code(desc, language="text")

        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button(f"1. Buka Browser ({chrome_profile})", key=f"akt_btn_open_{label}"):
                c_success, c_msg = script_manager.open_chrome_profile(chrome_profile)
                if c_success:
                    st.success(c_msg)
                    st.info("Browser terbuka! Silakan login/siapkan halaman Govem sampai siap.")
                    st.session_state[f"akt_ready_{label}"] = True
                else:
                    st.error(c_msg)

        with col2:
            if st.session_state.get(f"akt_ready_{label}"):
                st.markdown("**Klik ini jika website sudah siap!**")
                if st.button(f"2. MULAI INPUT (Siap!)", key=f"akt_btn_start_{label}", type="primary"):
                    with st.spinner(f"Menjalankan skrip... JANGAN GERAKKAN MOUSE!"):
                        success, msg = script_manager.run_script_process(script_path)
                        if success:
                            st.success(f"Berhasil: {msg}")
                            st.balloons()
                            st.session_state[f"akt_ready_{label}"] = False
                        else:
                            st.error(f"Error: {msg}")
            else:
                st.caption(f"*Tombol Mulai akan muncul setelah Anda membuka browser.*")

with tab1:
    render_activity_tab(DIR_SUAMI, "Aktivitas Suami", "Default")

with tab2:
    render_activity_tab(DIR_ISTRI, "Aktivitas Istri", "Profile 5")

st.markdown("---")
st.caption("Govem Automation System v1.0")

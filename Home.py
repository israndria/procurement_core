import os
os.environ["POKJA_MULTIPAGE"] = "1"

import streamlit as st

st.set_page_config(page_title="POKJA 2026 Toolkit", page_icon="🏛️", layout="wide")

st.title("POKJA 2026 Toolkit")
st.markdown("Pilih menu di sidebar untuk mulai.")

st.markdown("""
### Menu

- **Jadwal Tender** — Kelola jadwal tender + sync Google Calendar
- **SPSE Scraper** — Scrape data LPSE ke Supabase
- **Aktivitas Govem** — Input aktivitas harian otomatis
""")

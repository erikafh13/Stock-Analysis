"""
app.py  ←  Entry point utama aplikasi Streamlit.

Struktur project:
    stock_app/
    ├── app.py                  ← file ini (jalankan: streamlit run app.py)
    ├── utils/
    │   ├── __init__.py
    │   ├── gdrive.py           ← koneksi & operasi Google Drive
    │   └── analysis.py         ← semua fungsi kalkulasi & analisis
    └── pages/
        ├── input_data.py       ← halaman Input Data
        ├── stock_analysis.py   ← halaman Hasil Analisa Stock
        └── abc_analysis.py     ← halaman Hasil Analisa ABC
"""

import streamlit as st

# ── Konfigurasi Halaman ────────────────────────────────────────────────────────
st.set_page_config(layout="wide", page_title="Analisis Stock & ABC")

# ── Sidebar ────────────────────────────────────────────────────────────────────
st.sidebar.image(
    "https://eq-cdn.equiti-me.com/website/images/What_does_a_stock_split_mean.2e16d0ba.fill-1600x900.jpg",
    use_container_width=True,
)
st.sidebar.title("Analisis Stock dan ABC")

page = st.sidebar.radio(
    "Menu Navigasi:",
    ("Input Data", "Hasil Analisa Stock", "Hasil Analisa ABC", "Hasil Analisis Margin"),
    help="Pilih halaman untuk ditampilkan.",
)
st.sidebar.markdown("---")

# ── Inisialisasi Session State ─────────────────────────────────────────────────
_defaults = {
    "df_penjualan":         __import__("pandas").DataFrame(),
    "produk_ref":           __import__("pandas").DataFrame(),
    "df_stock":             __import__("pandas").DataFrame(),
    "stock_filename":       "",
    "stock_analysis_result": None,
    "abc_analysis_result":  None,
    "bulan_columns_stock":  [],
    "df_portal_analyzed":   __import__("pandas").DataFrame(),
}
for key, default in _defaults.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ── Koneksi Google Drive (dilakukan sekali di sini) ────────────────────────────
from utils.gdrive import init_drive_service

drive_service, DRIVE_AVAILABLE = init_drive_service()

if not DRIVE_AVAILABLE:
    if page != "Input Data":
        st.warning("Koneksi Google Drive tidak tersedia. Harap periksa kredensial.")

# ── Routing Halaman ────────────────────────────────────────────────────────────
if page == "Input Data":
    if not DRIVE_AVAILABLE:
        st.warning("Tidak dapat melanjutkan karena koneksi ke Google Drive gagal.")
        st.stop()
    from pages.input_data import render
    render(drive_service)

elif page == "Hasil Analisa Stock":
    from pages.stock_analysis import render
    render()

elif page == "Hasil Analisa ABC":
    from pages.abc_analysis import render
    render()

elif page == "Hasil Analisis Margin":
    st.title("💰 Hasil Analisis Margin (Placeholder)")
    st.info("Halaman ini adalah placeholder untuk analisis margin yang akan dikembangkan selanjutnya.")

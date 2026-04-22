"""
pages/stock_donor.py
Halaman Analisis Donor Stock — distribusi lateral antar cabang.

Halaman ini TIDAK menjalankan ulang analisis dari nol.
Ia menggunakan hasil dari Analisa Stock V2 yang sudah ada di session_state,
lalu menghitung siapa yang bisa jadi donor dan siapa yang menerima.
"""

import numpy as np
import pandas as pd
import streamlit as st
from io import BytesIO

from utils.analysis import (
    calculate_donor_distribution,
    highlight_status_stock,
    highlight_kategori_abc_log,
)

# Warna skenario donor
_SKENARIO_COLOR = {
    "0 - TIDAK ADA KEBUTUHAN":     "#d4edda",
    "2 - SBY CUKUP":               "#cce5ff",
    "3 - SBY TERBATAS + ADA DONOR":"#fff3cd",
    "4 - HANYA DONOR CABANG":      "#ffe5b4",
    "5 - POOL TIDAK CUKUP":        "#f8d7da",
    "1 - KURANG (TIDAK ADA POOL)": "#f5c6cb",
}

def _highlight_skenario(val: str) -> str:
    return f"background-color: {_SKENARIO_COLOR.get(val, '')}"

def _highlight_donor_ke(val: str) -> str:
    if val and val != "-":
        return "background-color: #d4edda; font-weight: bold"
    return ""

def _highlight_terima_dari(val: str) -> str:
    if val and val != "-":
        return "background-color: #cce5ff; font-weight: bold"
    return ""


def render():
    st.title("🔄 Analisis Donor Stock — Distribusi Lateral Antar Cabang")
    st.info("""
    **Cara kerja halaman ini:**
    - Menggunakan hasil **Analisa Stock V2** yang sudah dijalankan.
    - Mengidentifikasi cabang **Overstock** (donor) dan **Understock** (penerima) per SKU.
    - Menghitung berapa unit yang bisa dikirim dari donor ke penerima, **sebelum PO ke supplier**.
    - **Urutan sumber**: Surabaya (sisa stock di atas Min) → Cabang donor lain → Supplier.
    
    ⚠️ **Harap jalankan Analisa Stock V2 terlebih dahulu sebelum menggunakan halaman ini.**
    """)

    # ── Cek data tersedia ──────────────────────────────────────────────────────
    result_v2 = st.session_state.get("stock_v2_result")
    if result_v2 is None or result_v2.empty:
        st.warning("⚠️ Data V2 belum tersedia. Silakan jalankan **Hasil Analisa Stock V2** terlebih dahulu.")
        st.stop()

    df = result_v2.copy()
    df = df[df["City"] != "OTHERS"]

    KAT_COL = "Kategori ABC (Log-Benchmark - WMA)"

    # Pastikan kolom Persentase Stock ada
    if "Persentase Stock" not in df.columns:
        df["Persentase Stock"] = np.where(
            df["Min Stock"] > 0,
            (df["Stock Cabang"] / df["Min Stock"]) * 100,
            np.where(df["Stock Cabang"] > 0, 10000, 0)
        ).round(1)

    # ── Filter ─────────────────────────────────────────────────────────────────
    st.markdown("---")
    st.header("🔍 Filter Produk")
    col1, col2, col3 = st.columns(3)
    sel_kat   = col1.multiselect("Kategori Barang:", sorted(df["Kategori Barang"].dropna().unique().astype(str)), key="donor_kat")
    sel_brand = col2.multiselect("Brand:",           sorted(df["BRAND Barang"].dropna().unique().astype(str)),    key="donor_brand")
    sel_prod  = col3.multiselect("Nama Produk:",     sorted(df["Nama Barang"].dropna().unique().astype(str)),     key="donor_prod")

    if sel_kat:   df = df[df["Kategori Barang"].astype(str).isin(sel_kat)]
    if sel_brand: df = df[df["BRAND Barang"].astype(str).isin(sel_brand)]
    if sel_prod:  df = df[df["Nama Barang"].astype(str).isin(sel_prod)]

    col4, col5 = st.columns(2)
    sel_abc    = col4.multiselect("Kategori ABC:", sorted(df[KAT_COL].dropna().unique().astype(str)), key="donor_abc")
    only_active = col5.checkbox("Hanya tampilkan SKU yang ada aktivitas donor/terima", value=True, key="donor_active")

    if sel_abc:
        df = df[df[KAT_COL].isin(sel_abc)]

    # ── Hitung distribusi donor ────────────────────────────────────────────────
    st.markdown("---")
    with st.spinner("Menghitung distribusi donor..."):
        donor_df = calculate_donor_distribution(df)

    if donor_df.empty:
        st.info("Tidak ada data yang bisa diproses.")
        st.stop()

    # Filter hanya SKU dengan aktivitas jika checkbox aktif
    if only_active:
        active_skus = donor_df[
            (donor_df["Donor_Ke"] != "-") | (donor_df["Terima_Dari"] != "-")
        ]["No. Barang"].unique()
        donor_df_display = donor_df[donor_df["No. Barang"].isin(active_skus)]
    else:
        donor_df_display = donor_df.copy()

    # ── Summary Metrics ────────────────────────────────────────────────────────
    st.header("📊 Ringkasan")

    total_sku_donor    = donor_df[donor_df["Donor_Ke"] != "-"]["No. Barang"].nunique()
    total_sku_terima   = donor_df[donor_df["Terima_Dari"] != "-"]["No. Barang"].nunique()
    total_qty_donor    = donor_df["Qty_Donor"].sum()
    total_qty_terima   = donor_df["Qty_Terima"].sum()
    total_sisa_sup     = donor_df.groupby("No. Barang")["Sisa_Need_Supplier"].max().sum()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("SKU dengan Donor",       f"{total_sku_donor} SKU")
    c2.metric("SKU Menerima Transfer",  f"{total_sku_terima} SKU")
    c3.metric("Total Unit Didonorkan",  f"{int(total_qty_donor)} unit")
    c4.metric("Total Unit Diterima",    f"{int(total_qty_terima)} unit")
    c5.metric("Sisa Butuh Supplier",    f"{int(total_sisa_sup)} unit")

    # ── Distribusi Skenario ────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Distribusi Skenario per SKU")
    sku_skenario = donor_df.groupby("No. Barang")["Skenario_Donor"].first()
    skenario_counts = sku_skenario.value_counts()
    cols_sk = st.columns(len(skenario_counts))
    for i, (label, count) in enumerate(skenario_counts.items()):
        color = _SKENARIO_COLOR.get(label, "#eee")
        cols_sk[i].markdown(
            f"<div style='background:{color};padding:10px;border-radius:8px;text-align:center'>"
            f"<b>{label}</b><br>{count} SKU</div>", unsafe_allow_html=True
        )

    st.markdown("---")

    # ── Tab: Per Cabang & Detail SKU ───────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs(["📋 Detail per Cabang", "📦 Rekap per SKU", "🔀 Matriks Transfer"])

    DISPLAY_COLS = [
        "No. Barang", "Kategori Barang", "BRAND Barang", "Nama Barang",
        "City", "Kategori ABC",
        "Stock Cabang", "Min Stock", "Max Stock",
        "Status Stock", "Persentase Stock",
        "Add Stock", "Qty_Bisa_Donor",
        "Donor_Ke", "Qty_Donor",
        "Terima_Dari", "Qty_Terima",
        "Sisa_Need_Supplier", "Skenario_Donor",
    ]
    exist_cols = [c for c in DISPLAY_COLS if c in donor_df_display.columns]

    with tab1:
        st.subheader("Detail Distribusi Donor per Cabang")
        for city in sorted(donor_df_display["City"].unique()):
            city_df = donor_df_display[donor_df_display["City"] == city][exist_cols].copy()
            if city_df.empty:
                continue
            has_activity = (city_df["Donor_Ke"] != "-") | (city_df["Terima_Dari"] != "-")
            n_active = has_activity.sum()
            label = f"📍 {city}  —  {n_active} SKU aktif"
            with st.expander(label, expanded=(n_active > 0)):
                fmt = {
                    "Persentase Stock": "{:.1f}%",
                    "Stock Cabang": "{:.0f}", "Min Stock": "{:.0f}",
                    "Max Stock": "{:.0f}", "Add Stock": "{:.0f}",
                    "Qty_Bisa_Donor": "{:.0f}", "Qty_Donor": "{:.0f}",
                    "Qty_Terima": "{:.0f}", "Sisa_Need_Supplier": "{:.0f}",
                }
                style_cols = {}
                for c in city_df.columns:
                    if c == "Status Stock":      style_cols[c] = highlight_status_stock
                    elif c == "Kategori ABC":    style_cols[c] = highlight_kategori_abc_log
                    elif c == "Skenario_Donor":  style_cols[c] = _highlight_skenario
                    elif c == "Donor_Ke":        style_cols[c] = _highlight_donor_ke
                    elif c == "Terima_Dari":     style_cols[c] = _highlight_terima_dari

                styled = city_df.style.format(fmt, na_rep="-")
                for col, fn in style_cols.items():
                    if col in city_df.columns:
                        styled = styled.map(fn, subset=[col])
                st.dataframe(styled, use_container_width=True)

    with tab2:
        st.subheader("Rekap per SKU — Siapa Donor, Siapa Penerima")
        # Pivot: 1 baris per SKU, ringkasan donor & penerima
        sku_summary_rows = []
        KEYS = ["No. Barang", "Kategori Barang", "BRAND Barang", "Nama Barang"]
        for no_barang, grp in donor_df_display.groupby("No. Barang"):
            donors_list   = grp[grp["Donor_Ke"] != "-"][["City", "Qty_Donor", "Donor_Ke"]].copy()
            recv_list     = grp[grp["Terima_Dari"] != "-"][["City", "Qty_Terima", "Terima_Dari"]].copy()
            skenario      = grp["Skenario_Donor"].iloc[0]
            total_need    = grp["Total_Need"].iloc[0]
            total_pool    = grp["Total_Pool"].iloc[0]
            sisa_sup      = grp["Sisa_Need_Supplier"].sum()

            donor_str = " | ".join(
                f"{r['City']} kirim {r['Qty_Donor']} → {r['Donor_Ke']}"
                for _, r in donors_list.iterrows() if r["Qty_Donor"] > 0
            ) or "-"

            recv_str = " | ".join(
                f"{r['City']} terima {r['Qty_Terima']} dari {r['Terima_Dari']}"
                for _, r in recv_list.iterrows() if r["Qty_Terima"] > 0
            ) or "-"

            first = grp.iloc[0]
            row   = {k: first[k] for k in KEYS if k in grp.columns}
            row.update({
                "Skenario":        skenario,
                "Total Need":      int(total_need),
                "Total Pool":      int(total_pool),
                "Sisa → Supplier": int(sisa_sup),
                "Donor (dari siapa → ke siapa)": donor_str,
                "Penerima (kota → dari mana)":   recv_str,
            })
            sku_summary_rows.append(row)

        sku_summary = pd.DataFrame(sku_summary_rows)
        if not sku_summary.empty:
            def _color_ske(val):
                return _highlight_skenario(val)
            st.dataframe(
                sku_summary.style.map(_color_ske, subset=["Skenario"]),
                use_container_width=True,
                height=500,
            )
        else:
            st.info("Tidak ada SKU dengan aktivitas donor/terima pada filter ini.")

    with tab3:
        st.subheader("🔀 Matriks Transfer Antar Cabang")
        st.caption("Baris = pengirim, Kolom = penerima, nilai = total unit yang ditransfer")
        cities = sorted(donor_df_display["City"].unique())
        matrix = pd.DataFrame(0, index=cities, columns=cities)

        for _, row in donor_df_display.iterrows():
            if row["Donor_Ke"] != "-" and row["Qty_Donor"] > 0:
                dests = [d.strip() for d in row["Donor_Ke"].split(",")]
                # Distribusi merata per penerima dari tabel ringkasan
                each = row["Qty_Donor"] // max(len(dests), 1)
                for dest in dests:
                    if dest in matrix.columns:
                        matrix.loc[row["City"], dest] += each

        # Tambah total
        matrix["TOTAL KIRIM"] = matrix.sum(axis=1)
        matrix.loc["TOTAL TERIMA"] = matrix.sum(axis=0)

        def _color_matrix(val):
            if isinstance(val, (int, float)) and val > 0:
                return "background-color: #cce5ff; font-weight: bold"
            return ""
        st.dataframe(matrix.style.map(_color_matrix), use_container_width=True)

    # ── Download Excel ─────────────────────────────────────────────────────────
    st.markdown("---")
    st.header("💾 Unduh Hasil Analisis Donor")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Sheet 1: Detail lengkap per cabang per SKU
        donor_df_display[exist_cols].to_excel(writer, sheet_name="Detail Donor", index=False)

        # Sheet 2: Rekap per SKU
        if not sku_summary.empty:
            sku_summary.to_excel(writer, sheet_name="Rekap per SKU", index=False)

        # Sheet 3: Matriks transfer
        matrix.to_excel(writer, sheet_name="Matriks Transfer")

        # Sheet 4: Per kota (sheet terpisah)
        for city in sorted(donor_df_display["City"].unique()):
            city_df = donor_df_display[donor_df_display["City"] == city][exist_cols]
            if not city_df.empty:
                city_df.to_excel(writer, sheet_name=city[:31], index=False)

    st.download_button(
        label="📥 Unduh Excel (Analisis Donor)",
        data=output.getvalue(),
        file_name="Analisis_Donor_Stock.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ── Legenda ────────────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📖 Legenda Skenario Donor")
    cols_leg = st.columns(3)
    cols_leg[0].success("**2 - SBY CUKUP**\nSisa stock SBY ≥ total kebutuhan cabang.\nSemua dipenuhi dari SBY.")
    cols_leg[1].warning("**3 - SBY TERBATAS + ADA DONOR**\nSisa SBY ada tapi tidak cukup.\nKekurangan ditutup dari donor cabang lain.")
    cols_leg[2].warning("**4 - HANYA DONOR CABANG**\nSisa SBY = 0.\nDistribusi murni dari cabang overstock ke understock.")
    cols_leg2 = st.columns(3)
    cols_leg2[0].error("**1 - KURANG (TIDAK ADA POOL)**\nTidak ada SBY sisa, tidak ada donor.\nSemua harus PO ke supplier.")
    cols_leg2[1].error("**5 - POOL TIDAK CUKUP**\nAda pool (SBY + donor) tapi tidak cukup.\nDistribusi sebagian + sisa PO ke supplier.")
    cols_leg2[2].info("**0 - TIDAK ADA KEBUTUHAN**\nSemua cabang Balance/Overstock.\nTidak ada yang perlu menerima transfer.")

"""
pages/stock_donor.py  —  Analisis Demand & Pool Stock Antar Cabang
(versi sederhana: tanpa prioritas jarak, tanpa aturan kirim, tanpa alokasi
 otomatis siapa-kirim-ke-siapa — hanya menghitung total kebutuhan vs total
 ketersediaan per SKU di seluruh cabang)
"""

import numpy as np
import pandas as pd
import streamlit as st
from io import BytesIO

from utils.analysis import (
    highlight_status_stock,
    highlight_kategori_abc_log,
)


def _hl_kesimpulan(val):
    if val == "✅ CUKUP":
        return "background-color: #d4edda; font-weight: bold"
    if val == "⚠️ KURANG":
        return "background-color: #f8d7da; font-weight: bold"
    return ""


# ── Kalkulasi: Demand & Pool per SKU per Cabang ─────────────────────────────────
def _run_demand_summary(df):
    """
    Menghitung total kebutuhan (Add Stock) dan total yang bisa didonorkan
    (Qty Bisa Donor) per cabang per SKU. Tidak ada proses alokasi/pencocokan
    siapa kirim ke siapa — cuma gambaran besar demand vs pool.
    """
    KAT_COL = "Kategori ABC (Log-Benchmark - WMA)"
    KEYS    = ["No. Barang", "Kategori Barang", "BRAND Barang", "Nama Barang"]

    detail_rows = []
    agg_rows    = []

    for sku, group in df.groupby("No. Barang"):
        group = group.copy().reset_index(drop=True)
        mask_sby      = group["City"] == "SURABAYA"
        sby_rows      = group[mask_sby]
        stock_sby     = float(sby_rows["Stock Cabang"].sum())
        min_stock_sby = float(sby_rows["Min Stock"].sum())
        # Surabaya boleh donor kalau stoknya masih di atas Min Stock miliknya
        # sendiri (lebih longgar dibanding cabang lain, karena posisinya
        # sebagai hub pusat).
        sisa_sby      = max(0.0, stock_sby - min_stock_sby)

        first = group.iloc[0]
        meta  = {k: first.get(k, "") for k in KEYS}

        total_need  = 0.0
        total_avail = 0.0

        for _, row in group.iterrows():
            city      = row["City"]
            kat       = row.get(KAT_COL, "-")
            excluded  = (kat == "F")   # kategori F tidak ikut proses donor sama sekali

            need  = 0.0
            avail = 0.0
            if not excluded:
                if row["Status Stock"] == "Understock":
                    need = float(row["Add Stock"])

                if city == "SURABAYA":
                    avail = sisa_sby
                elif row["Status Stock"] == "Overstock":
                    avail = max(0.0, float(row["Stock Cabang"]) - float(row["Max Stock"]))

            total_need  += need
            total_avail += avail

            detail_rows.append({
                **meta,
                "City":                      city,
                "Kategori ABC":              kat,
                "Stock Cabang":              int(row["Stock Cabang"]),
                "Min Stock":                 int(row["Min Stock"]),
                "Max Stock":                 int(row["Max Stock"]),
                "Status Stock":              row["Status Stock"],
                "% Stock":                   round(float(row.get("Persentase Stock", 0)), 1),
                "Add Stock (Butuh)":         int(need),
                "Qty Bisa Donor (Tersedia)": int(avail),
            })

        selisih     = total_avail - total_need
        kesimpulan  = "✅ CUKUP" if (total_need == 0 or selisih >= 0) else "⚠️ KURANG"
        agg_rows.append({
            **meta,
            "Total Butuh Semua Cabang":     int(total_need),
            "Total Tersedia Semua Cabang":  int(total_avail),
            "Selisih (Tersedia − Butuh)":   int(selisih),
            "Kesimpulan":                   kesimpulan,
        })

    detail_df = pd.DataFrame(detail_rows)
    agg_df    = pd.DataFrame(agg_rows)
    return detail_df, agg_df


# ── Render ─────────────────────────────────────────────────────────────────────
def render():
    st.title("📦 Analisis Demand & Pool Stock Antar Cabang")
    st.caption("Lihat total kebutuhan vs total ketersediaan tiap cabang — sebelum order ke supplier")

    result_v2 = st.session_state.get("stock_v2_result")
    if result_v2 is None or (isinstance(result_v2, pd.DataFrame) and result_v2.empty):
        st.warning("⚠️ Jalankan dulu **Hasil Analisa Stock V2**, kemudian kembali ke halaman ini.")
        st.stop()

    df = result_v2.copy()
    df = df[df["City"] != "OTHERS"]
    KAT_COL = "Kategori ABC (Log-Benchmark - WMA)"

    if "Persentase Stock" not in df.columns:
        df["Persentase Stock"] = np.where(
            df["Min Stock"] > 0,
            (df["Stock Cabang"] / df["Min Stock"]) * 100,
            np.where(df["Stock Cabang"] > 0, 10000, 0)
        ).round(1)

    # ── Penjelasan ─────────────────────────────────────────────────────────────
    with st.expander("📖 Cara Kerja & Arti Setiap Kolom (klik untuk membaca)", expanded=False):
        st.markdown("""
### Apa yang dihitung di sini?

Halaman ini menjawab: **"Berapa total kebutuhan tiap cabang, dan berapa total
kelebihan stok yang tersedia di seluruh cabang?"** — tanpa menentukan dulu
siapa kirim ke siapa.

---

### Rumus:

**① Siapa yang KELEBIHAN?**
Cabang dengan status **Overstock** → stoknya melebihi Max Stock.
`Qty bisa didonorkan = Stock Cabang − Max Stock`
*Contoh: Jogja stock 75, Max 46 → bisa kirim 29 unit*

Surabaya adalah pengecualian: dihitung boleh donor kalau stoknya masih **di
atas Min Stock miliknya sendiri** (tidak harus sampai Overstock dulu), karena
posisinya sebagai hub pusat.
`Sisa SBY = max(0, Stock SBY − Min Stock SBY)`

**② Siapa yang KEKURANGAN?**
Cabang dengan status **Understock** → stoknya di bawah Min Stock.
`Kebutuhan (Add Stock) = Min Stock − Stock Cabang`
*Contoh: Jakarta stock 6, Min 47 → butuh 41 unit*

Surabaya juga dihitung sebagai butuh kalau statusnya Understock.

**③ Total Pool vs Total Need (per SKU):**
- **Total Tersedia = Sisa SBY + semua Qty Bisa Donor cabang**
- **Total Butuh = jumlah semua Add Stock cabang yang Understock**
- **Selisih = Total Tersedia − Total Butuh** → positif berarti secara total
  stok di seluruh cabang cukup untuk SKU ini (belum tentu otomatis terbagi
  rata ke tiap cabang yang butuh, karena logika kirim-per-cabang tidak
  dihitung di sini).

Kategori ABC = **F** selalu dikeluarkan dari perhitungan (SKU sangat lambat,
tidak layak dipindah-pindah antar cabang).

---

### Arti Setiap Kolom:

| Kolom | Artinya |
|---|---|
| **Stock Cabang** | Stok fisik saat ini |
| **Min Stock** | Batas bawah aman berdasarkan penjualan × buffer ABC |
| **Max Stock** | Batas atas ideal |
| **Add Stock (Butuh)** | Unit yang dibutuhkan agar mencapai Min Stock |
| **% Stock** | `(Stock ÷ Min Stock) × 100` — makin kecil makin kritis |
| **Qty Bisa Donor (Tersedia)** | Unit kelebihan yang bisa didonorkan `(Stock − Max)` |
| **Total Butuh Semua Cabang** | Jumlah kebutuhan di seluruh cabang untuk SKU ini |
| **Total Tersedia Semua Cabang** | Jumlah ketersediaan di seluruh cabang untuk SKU ini |
| **Selisih (Tersedia − Butuh)** | Positif = pool cukup; Negatif = kurang |
        """)

    # ── Filter Produk ──────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("🔍 Filter Produk")
    c1, c2, c3 = st.columns(3)
    sel_kat   = c1.multiselect("Kategori Barang:", sorted(df["Kategori Barang"].dropna().unique().astype(str)), key="donor_kat")
    sel_brand = c2.multiselect("Brand:",           sorted(df["BRAND Barang"].dropna().unique().astype(str)),    key="donor_brand")
    sel_prod  = c3.multiselect("Nama Produk:",     sorted(df["Nama Barang"].dropna().unique().astype(str)),     key="donor_prod")
    c4, c5    = st.columns(2)
    sel_abc     = c4.multiselect("Kategori ABC:",  sorted(df[KAT_COL].dropna().unique().astype(str)), key="donor_abc")
    only_active = c5.checkbox("Hanya SKU dengan kebutuhan atau kelebihan stok", value=True, key="donor_active")

    if sel_kat:   df = df[df["Kategori Barang"].astype(str).isin(sel_kat)]
    if sel_brand: df = df[df["BRAND Barang"].astype(str).isin(sel_brand)]
    if sel_prod:  df = df[df["Nama Barang"].astype(str).isin(sel_prod)]
    if sel_abc:   df = df[df[KAT_COL].isin(sel_abc)]

    # ── Hitung ────────────────────────────────────────────────────────────────
    st.markdown("---")
    with st.spinner("⏳ Menghitung demand & pool..."):
        detail_df, agg_df = _run_demand_summary(df)

    if agg_df.empty:
        st.info("Tidak ada data untuk diproses.")
        st.stop()

    if only_active:
        active_skus = detail_df[
            (detail_df["Add Stock (Butuh)"] > 0) | (detail_df["Qty Bisa Donor (Tersedia)"] > 0)
        ]["No. Barang"].unique()
        detail_disp = detail_df[detail_df["No. Barang"].isin(active_skus)].copy()
        agg_disp    = agg_df[agg_df["No. Barang"].isin(active_skus)].copy()
    else:
        detail_disp = detail_df.copy()
        agg_disp    = agg_df.copy()

    # ── Metrics ───────────────────────────────────────────────────────────────
    st.subheader("📊 Ringkasan Demand vs Pool")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Butuh (semua SKU & cabang)",    f"{int(agg_disp['Total Butuh Semua Cabang'].sum()):,}")
    m2.metric("Total Tersedia (semua SKU & cabang)", f"{int(agg_disp['Total Tersedia Semua Cabang'].sum()):,}")
    m3.metric("SKU dengan pool CUKUP",                int((agg_disp["Kesimpulan"] == "✅ CUKUP").sum()))
    m4.metric("SKU dengan pool KURANG",                int((agg_disp["Kesimpulan"] == "⚠️ KURANG").sum()))

    # ── Rekap per Cabang ─────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📍 Rekap Demand & Pool per Cabang")
    st.caption("Total kebutuhan dan ketersediaan digabung dari semua SKU yang lolos filter di atas.")
    city_summary = detail_disp.groupby("City").agg(
        Total_Butuh           = ("Add Stock (Butuh)", "sum"),
        Total_Tersedia        = ("Qty Bisa Donor (Tersedia)", "sum"),
        Jumlah_SKU_Understock = ("Status Stock", lambda s: int((s == "Understock").sum())),
        Jumlah_SKU_Overstock  = ("Status Stock", lambda s: int((s == "Overstock").sum())),
    ).reset_index().sort_values("Total_Butuh", ascending=False)

    st.dataframe(city_summary, use_container_width=True, hide_index=True)
    st.bar_chart(city_summary.set_index("City")[["Total_Butuh", "Total_Tersedia"]])

    # ── Rekap per SKU ─────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📦 Rekap per SKU — Total Butuh vs Total Tersedia (Semua Cabang)")
    st.caption("1 baris = 1 SKU. Selisih positif = secara total pool cukup untuk SKU ini di seluruh cabang.")
    st.dataframe(
        agg_disp.style.map(_hl_kesimpulan, subset=["Kesimpulan"]),
        use_container_width=True,
        height=450,
    )

    # ── Detail per Cabang ────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("📋 Detail per Cabang")
    for city in sorted(detail_disp["City"].unique()):
        cdf = detail_disp[detail_disp["City"] == city].copy()
        n_act = ((cdf["Add Stock (Butuh)"] > 0) | (cdf["Qty Bisa Donor (Tersedia)"] > 0)).sum()
        with st.expander(f"📍 {city}  —  {n_act} SKU aktif dari {len(cdf)}", expanded=(n_act > 0)):
            styled = cdf.style
            for col in cdf.columns:
                if col == "Status Stock":
                    styled = styled.map(highlight_status_stock, subset=[col])
                elif col == "Kategori ABC":
                    styled = styled.map(highlight_kategori_abc_log, subset=[col])
            st.dataframe(styled, use_container_width=True)

    # ── Download ─────────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("💾 Unduh Rekap Demand & Pool")
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        agg_disp.to_excel(writer, sheet_name="Rekap per SKU", index=False)
        city_summary.to_excel(writer, sheet_name="Rekap per Cabang", index=False)
        detail_disp.to_excel(writer, sheet_name="Detail per Cabang-SKU", index=False)
    st.download_button(
        "📥 Unduh Excel — Demand & Pool",
        data=output.getvalue(),
        file_name="Demand_Pool_Semua_Cabang.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

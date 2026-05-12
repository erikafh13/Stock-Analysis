"""
pages/abc_analysis_v3.py
Analisis ABC V3 — sama dengan V2 + tambahan kolom ABC per Platform (Online / Offline).

Mapping Platform:
  ONLINE  : Nama Pelanggan mengandung "AIRPAY" atau diakhiri "- SHOPEE" → Shopee
            Nama Pelanggan diawali "TOKOPEDIA"                         → Tokopedia
            No. Faktur diawali AO, BO, DO, EO, FO, HO                 → Website/Retail
  OFFLINE : selain ONLINE
"""

import re
import numpy as np
import pandas as pd
import streamlit as st
from io import BytesIO

import matplotlib.pyplot as plt

from utils import (
    map_nama_dept,
    map_city,
    classify_abc_log_benchmark,
    highlight_kategori_abc_log,
)


# ── Konstanta ──────────────────────────────────────────────────────────────────
PLATFORM_ONLINE  = "ONLINE"
PLATFORM_OFFLINE = "OFFLINE"

_ONLINE_FAKTUR_PREFIX = re.compile(r"^(AO|BO|DO|EO|FO|HO)", re.IGNORECASE)


def _map_platform(row) -> str:
    """Tentukan platform ONLINE atau OFFLINE untuk setiap baris transaksi."""
    nama = str(row.get("Nama Pelanggan", "")).strip().upper()
    faktur = str(row.get("No. Faktur", "")).strip().upper()

    # Shopee: nama diawali AIRPAY atau mengandung '- SHOPEE' / diakhiri SHOPEE
    if nama.startswith("AIRPAY") or "- SHOPEE" in nama or nama.endswith("SHOPEE"):
        return PLATFORM_ONLINE

    # Tokopedia: nama diawali TOKOPEDIA
    if nama.startswith("TOKOPEDIA"):
        return PLATFORM_ONLINE

    # Website / Retail: No. Faktur diawali AO, BO, DO, EO, FO, HO
    if _ONLINE_FAKTUR_PREFIX.match(faktur):
        return PLATFORM_ONLINE

    return PLATFORM_OFFLINE


# ── Entry Point ────────────────────────────────────────────────────────────────
def render():
    st.title("📊 Analisis ABC V3 — Log-Benchmark + Platform (Online / Offline)")
    tab1, tab2 = st.tabs(["Hasil Tabel", "Dashboard"])

    with tab1:
        _render_table_tab()
    with tab2:
        _render_dashboard_tab()


# ── Tab Tabel ──────────────────────────────────────────────────────────────────
def _render_table_tab():
    if st.session_state.df_penjualan.empty or st.session_state.produk_ref.empty:
        st.warning(
            "⚠️ Harap muat file **Penjualan** dan **Produk Referensi** "
            "di halaman **'Input Data'** terlebih dahulu."
        )
        st.stop()

    # ── Preprocessing ──────────────────────────────────────────────────────────
    so_df      = st.session_state.df_penjualan.copy()
    produk_ref = st.session_state.produk_ref.copy()

    for df in [so_df, produk_ref]:
        if "No. Barang" in df.columns:
            df["No. Barang"] = df["No. Barang"].astype(str).str.strip()

    so_df.rename(columns={"Qty": "Kuantitas"}, inplace=True, errors="ignore")
    so_df["Nama Dept"] = so_df.apply(map_nama_dept, axis=1)
    so_df["City"]      = so_df["Nama Dept"].apply(map_city)

    # Mapping Platform (ONLINE / OFFLINE)
    so_df["Platform"] = so_df.apply(_map_platform, axis=1)

    if "Kategori Barang" in produk_ref.columns:
        produk_ref["Kategori Barang"] = (
            produk_ref["Kategori Barang"].astype(str).str.strip().str.upper()
        )
    if "City" in so_df.columns:
        so_df["City"] = so_df["City"].astype(str).str.strip().str.upper()

    so_df["Tgl Faktur"] = pd.to_datetime(
        so_df["Tgl Faktur"], dayfirst=True, errors="coerce"
    )
    so_df.dropna(subset=["Tgl Faktur"], inplace=True)

    # ── Filter Tanggal ─────────────────────────────────────────────────────────
    st.header("Filter Rentang Waktu Analisis ABC V3")
    st.info(
        "Analisis didasarkan pada data penjualan 90 hari *sebelum* "
        "**Tanggal Akhir** yang dipilih."
    )
    min_date        = so_df["Tgl Faktur"].min().date()
    max_date        = so_df["Tgl Faktur"].max().date()
    end_date_input  = st.date_input(
        "Tanggal Akhir", value=max_date, min_value=min_date, max_value=max_date
    )

    if st.button("Jalankan Analisa ABC V3 (Log-Benchmark + Platform)"):
        _run_abc_analysis_v3(so_df, produk_ref, end_date_input)

    if st.session_state.get("abc_v3_result") is None:
        return

    result_display = st.session_state.abc_v3_result.copy()
    result_display = result_display[result_display["City"] != "OTHERS"]

    # ── Filter UI ──────────────────────────────────────────────────────────────
    st.header("Filter Hasil Analisis")
    col_f1, col_f2 = st.columns(2)
    sel_kat   = col_f1.multiselect(
        "Filter Kategori:",
        sorted(produk_ref["Kategori Barang"].dropna().unique().astype(str)),
        key="abc_v3_cat_filter",
    )
    sel_brand = col_f2.multiselect(
        "Filter Brand:",
        sorted(produk_ref["BRAND Barang"].dropna().unique().astype(str)),
        key="abc_v3_brand_filter",
    )
    if sel_kat:   result_display = result_display[result_display["Kategori Barang"].astype(str).isin(sel_kat)]
    if sel_brand: result_display = result_display[result_display["BRAND Barang"].astype(str).isin(sel_brand)]

    KEYS = ["No. Barang", "Kategori Barang", "BRAND Barang", "Nama Barang"]

    # ── Tabel per Kota ─────────────────────────────────────────────────────────
    st.header("Hasil Analisis ABC V3 per Kota")
    for city in sorted(result_display["City"].unique()):
        with st.expander(f"🏙️ Lihat Hasil ABC V3 untuk Kota: {city}"):
            city_df = result_display[result_display["City"] == city]
            col_order = [
                "No. Barang", "BRAND Barang", "Nama Barang", "Kategori Barang",
                # Overall (semua transaksi)
                "AVG Mean", "AVG WMA",
                "Kategori ABC (Log-Benchmark - Mean)",
                "Kategori ABC (Log-Benchmark - WMA)",
                "Log (10) Mean", "Avg Log Mean", "Ratio Log Mean",
                "Log (10) WMA",  "Avg Log WMA",  "Ratio Log WMA",
                # Online
                "AVG Mean_ONLINE", "AVG WMA_ONLINE",
                "Kategori ABC (Log-Benchmark - Mean)_ONLINE",
                "Kategori ABC (Log-Benchmark - WMA)_ONLINE",
                # Offline
                "AVG Mean_OFFLINE", "AVG WMA_OFFLINE",
                "Kategori ABC (Log-Benchmark - Mean)_OFFLINE",
                "Kategori ABC (Log-Benchmark - WMA)_OFFLINE",
            ]
            display_cols = [c for c in col_order if c in city_df.columns]
            df_show = city_df[display_cols]

            fmt = {}
            for col in df_show.columns:
                if col in KEYS or not pd.api.types.is_numeric_dtype(df_show[col]):
                    continue
                fmt[col] = (
                    "{:.2f}"
                    if any(x in col for x in ["Ratio", "Log", "Avg Log"])
                    else "{:.0f}"
                )

            style = df_show.style.format(fmt, na_rep="-")
            for abc_col in [
                "Kategori ABC (Log-Benchmark - Mean)",
                "Kategori ABC (Log-Benchmark - WMA)",
                "Kategori ABC (Log-Benchmark - Mean)_ONLINE",
                "Kategori ABC (Log-Benchmark - WMA)_ONLINE",
                "Kategori ABC (Log-Benchmark - Mean)_OFFLINE",
                "Kategori ABC (Log-Benchmark - WMA)_OFFLINE",
            ]:
                if abc_col in df_show.columns:
                    style = style.apply(
                        lambda x: x.map(highlight_kategori_abc_log), subset=[abc_col]
                    )
            st.dataframe(style, use_container_width=True)

    # ── Tabel Pivot Gabungan ───────────────────────────────────────────────────
    st.header("📊 Tabel Gabungan Seluruh Kota (ABC V3)")
    with st.spinner("Membuat tabel pivot gabungan..."):
        _render_pivot_abc_v3(result_display, KEYS, end_date_input)


# ── Perhitungan Utama ──────────────────────────────────────────────────────────
def _compute_abc_for_subset(
    df_subset: pd.DataFrame,
    produk_ref: pd.DataFrame,
    barang_list: pd.DataFrame,
    city_list,
    end_dt,
    suffix: str = "",
) -> pd.DataFrame:
    """
    Hitung AVG Mean / WMA dan klasifikasi ABC untuk subset data (semua / online / offline).
    Mengembalikan DataFrame dengan kolom:
        City, No. Barang, AVG Mean{suffix}, AVG WMA{suffix},
        Kategori ABC (Log-Benchmark - Mean){suffix},
        Kategori ABC (Log-Benchmark - WMA){suffix},
        Log (10) Mean{suffix}, Avg Log Mean{suffix}, Ratio Log Mean{suffix},
        Log (10) WMA{suffix},  Avg Log WMA{suffix},  Ratio Log WMA{suffix},
    """
    r1_end,  r1_start = end_dt, end_dt - pd.DateOffset(days=29)
    r2_end,  r2_start = end_dt - pd.DateOffset(days=30),  end_dt - pd.DateOffset(days=59)
    r3_end,  r3_start = end_dt - pd.DateOffset(days=60),  end_dt - pd.DateOffset(days=89)

    def _sales(start, end, col):
        return (
            df_subset[df_subset["Tgl Faktur"].between(start, end)]
            .groupby(["City", "No. Barang"])["Kuantitas"]
            .sum()
            .reset_index(name=col)
        )

    s1 = _sales(r1_start, r1_end, "Penjualan Bln 1")
    s2 = _sales(r2_start, r2_end, "Penjualan Bln 2")
    s3 = _sales(r3_start, r3_end, "Penjualan Bln 3")

    kombinasi = pd.MultiIndex.from_product(
        [city_list, barang_list["No. Barang"]], names=["City", "No. Barang"]
    ).to_frame(index=False)

    grouped = pd.merge(kombinasi, barang_list, on="No. Barang", how="left")
    for sm in [s1, s2, s3]:
        grouped = pd.merge(grouped, sm, on=["City", "No. Barang"], how="left")
    grouped.fillna(
        {"Penjualan Bln 1": 0, "Penjualan Bln 2": 0, "Penjualan Bln 3": 0},
        inplace=True,
    )

    avg_mean_col = f"AVG Mean{suffix}"
    avg_wma_col  = f"AVG WMA{suffix}"

    grouped[avg_mean_col] = (
        grouped["Penjualan Bln 1"] + grouped["Penjualan Bln 2"] + grouped["Penjualan Bln 3"]
    ) / 3
    grouped[avg_wma_col] = np.ceil(
        grouped["Penjualan Bln 1"] * 0.5
        + grouped["Penjualan Bln 2"] * 0.3
        + grouped["Penjualan Bln 3"] * 0.2
    )

    # Rename sementara agar classify_abc_log_benchmark bisa pakai kolom standar
    tmp_mean = grouped.rename(columns={avg_mean_col: "AVG Mean", avg_wma_col: "AVG WMA"})
    res_mean = classify_abc_log_benchmark(tmp_mean.copy(), metric_col="AVG Mean")
    res_wma  = classify_abc_log_benchmark(tmp_mean.copy(), metric_col="AVG WMA")

    # Ambil kolom hasil klasifikasi dan rename dengan suffix
    rename_map_mean = {
        "AVG Mean":                              avg_mean_col,
        "Kategori ABC (Log-Benchmark - Mean)":   f"Kategori ABC (Log-Benchmark - Mean){suffix}",
        "Log (10) Mean":                         f"Log (10) Mean{suffix}",
        "Avg Log Mean":                          f"Avg Log Mean{suffix}",
        "Ratio Log Mean":                        f"Ratio Log Mean{suffix}",
    }
    rename_map_wma = {
        "AVG WMA":                               avg_wma_col,
        "Kategori ABC (Log-Benchmark - WMA)":    f"Kategori ABC (Log-Benchmark - WMA){suffix}",
        "Log (10) WMA":                          f"Log (10) WMA{suffix}",
        "Avg Log WMA":                           f"Avg Log WMA{suffix}",
        "Ratio Log WMA":                         f"Ratio Log WMA{suffix}",
    }

    keep_mean = ["City", "No. Barang"] + list(rename_map_mean.keys())
    keep_wma  = ["City", "No. Barang"] + list(rename_map_wma.keys())

    res_mean_slim = res_mean[[c for c in keep_mean if c in res_mean.columns]].rename(columns=rename_map_mean)
    res_wma_slim  = res_wma[[c for c in keep_wma  if c in res_wma.columns]].rename(columns=rename_map_wma)

    out = pd.merge(res_mean_slim, res_wma_slim, on=["City", "No. Barang"], how="left")

    # Bulatkan
    for col in [avg_mean_col, avg_wma_col]:
        if col in out.columns:
            out[col] = out[col].round(0).astype(int)
    for col in [
        f"Log (10) Mean{suffix}", f"Avg Log Mean{suffix}", f"Ratio Log Mean{suffix}",
        f"Log (10) WMA{suffix}",  f"Avg Log WMA{suffix}",  f"Ratio Log WMA{suffix}",
    ]:
        if col in out.columns:
            out[col] = out[col].round(2)

    return out


def _run_abc_analysis_v3(so_df, produk_ref, end_date_input):
    with st.spinner("Melakukan perhitungan analisis ABC V3 (Overall + Online + Offline)..."):
        end_dt   = pd.to_datetime(end_date_input)
        start_90 = end_dt - pd.DateOffset(days=89)
        df_90    = so_df[so_df["Tgl Faktur"].between(start_90, end_dt)]

        if df_90.empty:
            st.error("Tidak ada data penjualan pada rentang 90 hari yang dipilih.")
            st.session_state["abc_v3_result"] = None
            return

        produk_ref.rename(
            columns={
                "Keterangan Barang":    "Nama Barang",
                "Nama Kategori Barang": "Kategori Barang",
            },
            inplace=True,
            errors="ignore",
        )
        barang_list = produk_ref[
            ["No. Barang", "BRAND Barang", "Kategori Barang", "Nama Barang"]
        ].drop_duplicates()
        city_list = so_df["City"].dropna().unique()

        # 1. Overall (semua platform) — sama persis dengan V2
        res_overall = _compute_abc_for_subset(
            df_90, produk_ref, barang_list, city_list, end_dt, suffix=""
        )

        # 2. Online saja
        df_online = df_90[df_90["Platform"] == PLATFORM_ONLINE]
        res_online = _compute_abc_for_subset(
            df_online, produk_ref, barang_list, city_list, end_dt, suffix="_ONLINE"
        )

        # 3. Offline saja
        df_offline = df_90[df_90["Platform"] == PLATFORM_OFFLINE]
        res_offline = _compute_abc_for_subset(
            df_offline, produk_ref, barang_list, city_list, end_dt, suffix="_OFFLINE"
        )

        # Gabung semua
        KEYS_MERGE = ["City", "No. Barang"]
        result_final = pd.merge(res_overall, res_online,  on=KEYS_MERGE, how="left")
        result_final = pd.merge(result_final, res_offline, on=KEYS_MERGE, how="left")

        # Tambahkan info barang (kalau belum ada dari overall)
        for col in ["Kategori Barang", "BRAND Barang", "Nama Barang"]:
            if col not in result_final.columns:
                result_final = pd.merge(
                    result_final,
                    barang_list[["No. Barang", col]].drop_duplicates(),
                    on="No. Barang", how="left"
                )

        st.session_state["abc_v3_result"] = result_final.copy()
        st.success(
            "✅ Analisis ABC V3 (Overall + Online + Offline) berhasil dijalankan!"
        )

        # Tampilkan ringkasan mapping platform
        total_rows    = len(df_90)
        online_rows   = (df_90["Platform"] == PLATFORM_ONLINE).sum()
        offline_rows  = (df_90["Platform"] == PLATFORM_OFFLINE).sum()
        st.info(
            f"📦 Total transaksi 90 hari: **{total_rows:,}** — "
            f"ONLINE: **{online_rows:,}** | OFFLINE: **{offline_rows:,}**"
        )


# ── Pivot Gabungan ─────────────────────────────────────────────────────────────
def _render_pivot_abc_v3(result_display: pd.DataFrame, KEYS: list, end_date_input):
    """
    Pivot tabel:
    - Kolom per kota (Overall + Online + Offline)
    - Kolom ALL (semua kota, Overall + Online + Offline)
    """
    # Tentukan semua kolom nilai yang tersedia
    base_vals = [
        "Penjualan Bln 1", "Penjualan Bln 2", "Penjualan Bln 3",
        "AVG Mean", "AVG WMA",
        "Kategori ABC (Log-Benchmark - Mean)", "Ratio Log Mean", "Log (10) Mean", "Avg Log Mean",
        "Kategori ABC (Log-Benchmark - WMA)",  "Ratio Log WMA",  "Log (10) WMA",  "Avg Log WMA",
    ]
    platform_suffix_vals = []
    for sfx in ["_ONLINE", "_OFFLINE"]:
        platform_suffix_vals += [
            f"AVG Mean{sfx}", f"AVG WMA{sfx}",
            f"Kategori ABC (Log-Benchmark - Mean){sfx}",
            f"Kategori ABC (Log-Benchmark - WMA){sfx}",
            f"Ratio Log Mean{sfx}", f"Log (10) Mean{sfx}", f"Avg Log Mean{sfx}",
            f"Ratio Log WMA{sfx}",  f"Log (10) WMA{sfx}",  f"Avg Log WMA{sfx}",
        ]

    all_vals = base_vals + platform_suffix_vals
    existing_vals = [c for c in all_vals if c in result_display.columns]

    # Pivot per kota
    pivot = result_display.pivot_table(
        index=KEYS, columns="City", values=existing_vals, aggfunc="first"
    )
    pivot.columns = [f"{city}_{val}" for val, city in pivot.columns]
    pivot.reset_index(inplace=True)

    # ── ALL Summary ────────────────────────────────────────────────────────────
    def _build_all_summary(df_sub: pd.DataFrame, sfx: str) -> pd.DataFrame:
        """Hitung ABC ALL untuk subset tertentu (all / online / offline)."""
        avg_mean_src  = f"AVG Mean{sfx}"  if sfx else "AVG Mean"
        avg_wma_src   = f"AVG WMA{sfx}"   if sfx else "AVG WMA"
        bln1_src      = "Penjualan Bln 1"
        bln2_src      = "Penjualan Bln 2"
        bln3_src      = "Penjualan Bln 3"

        # Cek kolom tersedia
        needed = [bln1_src, bln2_src, bln3_src]
        missing = [c for c in needed if c not in df_sub.columns]
        if missing:
            return pd.DataFrame(columns=KEYS)

        total = df_sub.groupby(KEYS).agg({
            bln1_src: "sum", bln2_src: "sum", bln3_src: "sum"
        }).reset_index()
        total[avg_mean_src] = (total[bln1_src] + total[bln2_src] + total[bln3_src]) / 3
        total[avg_wma_src]  = np.ceil(
            total[bln1_src] * 0.5 + total[bln2_src] * 0.3 + total[bln3_src] * 0.2
        )
        for col in [bln1_src, bln2_src, bln3_src, avg_mean_src, avg_wma_src]:
            if col in total.columns:
                total[col] = total[col].round(0).astype(int)

        total["City"] = "ALL"

        res_mean = classify_abc_log_benchmark(total.copy(), metric_col=avg_mean_src)
        res_wma  = classify_abc_log_benchmark(total.copy(), metric_col=avg_wma_src)

        total_final = total.drop(columns=["City"], errors="ignore")
        for src_df, keywords in [
            (res_mean, [
                f"Log-Benchmark - Mean{sfx}" if sfx else "Log-Benchmark - Mean",
                f"Log (10) Mean{sfx}" if sfx else "Log (10) Mean",
                f"Avg Log Mean{sfx}" if sfx else "Avg Log Mean",
                f"Ratio Log Mean{sfx}" if sfx else "Ratio Log Mean",
            ]),
            (res_wma, [
                f"Log-Benchmark - WMA{sfx}" if sfx else "Log-Benchmark - WMA",
                f"Log (10) WMA{sfx}" if sfx else "Log (10) WMA",
                f"Avg Log WMA{sfx}" if sfx else "Avg Log WMA",
                f"Ratio Log WMA{sfx}" if sfx else "Ratio Log WMA",
            ]),
        ]:
            extra_cols = [c for c in src_df.columns if any(kw in c for kw in keywords)]
            if extra_cols:
                total_final = pd.merge(
                    total_final, src_df[KEYS + extra_cols], on=KEYS, how="left"
                )

        for col in [
            f"Log (10) Mean{sfx}", f"Avg Log Mean{sfx}", f"Ratio Log Mean{sfx}",
            f"Log (10) WMA{sfx}",  f"Avg Log WMA{sfx}",  f"Ratio Log WMA{sfx}",
        ]:
            real_col = col if col in total_final.columns else col.replace(sfx, "") if sfx else col
            if real_col in total_final.columns:
                total_final[real_col] = total_final[real_col].round(2)

        # Tambah prefix ALL_
        prefix = f"All_{sfx.lstrip('_')}_" if sfx else "All_"
        total_final.columns = [
            f"{prefix}{c}" if c not in KEYS else c for c in total_final.columns
        ]
        return total_final

    # ALL Overall
    all_overall = _build_all_summary(result_display, sfx="")

    # ALL Online — butuh data penjualan per bulan untuk online
    # Karena result_display sudah digabung, kita buat ulang penjualan per bulan dari sumber
    # yang sudah di-simpan di session_state.abc_v3_result (menggunakan AVG Mean_ONLINE sebagai proxy)
    # Untuk ALL Online/Offline, kita hitung dari kolom AVG Mean_ONLINE / AVG WMA_ONLINE
    # yang ada di result_display dengan asumsi data sudah per kota.
    # Kita re-aggregate AVG Mean_ONLINE/OFFLINE across cities:
    def _all_from_platform_cols(sfx: str) -> pd.DataFrame:
        avg_mean_col = f"AVG Mean{sfx}"
        avg_wma_col  = f"AVG WMA{sfx}"
        if avg_mean_col not in result_display.columns:
            return pd.DataFrame(columns=KEYS)

        # Kita butuh penjualan per bulan untuk re-compute WMA yang benar.
        # Karena tidak tersimpan terpisah per platform, gunakan AVG langsung
        # sebagai aproksimasi (sudah di-compute per kota dari data asli).
        # ALL = jumlah AVG per kota (proxy terbaik tanpa menyimpan ulang per bulan per platform).
        agg = result_display.groupby(KEYS).agg({
            avg_mean_col: "sum",
            avg_wma_col:  "sum",
        }).reset_index()

        agg["City"] = "ALL"
        res_mean = classify_abc_log_benchmark(agg.copy(), metric_col=avg_mean_col)
        res_wma  = classify_abc_log_benchmark(agg.copy(), metric_col=avg_wma_col)

        total_final = agg.drop(columns=["City"], errors="ignore")
        for src_df, extra_pattern in [
            (res_mean, f"Log-Benchmark - Mean{sfx}"),
            (res_wma,  f"Log-Benchmark - WMA{sfx}"),
        ]:
            extra_cols = [c for c in src_df.columns if extra_pattern in c or
                          any(x in c for x in [f"Log (10) Mean{sfx}", f"Avg Log Mean{sfx}", f"Ratio Log Mean{sfx}",
                                               f"Log (10) WMA{sfx}", f"Avg Log WMA{sfx}", f"Ratio Log WMA{sfx}"])]
            if extra_cols:
                total_final = pd.merge(
                    total_final, src_df[KEYS + extra_cols].drop_duplicates(subset=KEYS),
                    on=KEYS, how="left"
                )

        prefix = f"All_{sfx.lstrip('_')}_"
        total_final.columns = [
            f"{prefix}{c}" if c not in KEYS else c for c in total_final.columns
        ]
        return total_final

    all_online  = _all_from_platform_cols("_ONLINE")
    all_offline = _all_from_platform_cols("_OFFLINE")

    # Gabung semua ALL ke pivot
    pivot_final = pivot.copy()
    for all_df in [all_overall, all_online, all_offline]:
        if not all_df.empty and len(all_df.columns) > len(KEYS):
            pivot_final = pd.merge(pivot_final, all_df, on=KEYS, how="left")

    # ── Styling ────────────────────────────────────────────────────────────────
    df_style = pivot_final.copy()
    num_cols   = [
        c for c in df_style.columns
        if c not in KEYS
        and pd.api.types.is_numeric_dtype(df_style[c])
        and not any(x in c for x in ["Ratio", "Log", "Avg Log"])
    ]
    float_cols = [
        c for c in df_style.columns
        if c not in KEYS and any(x in c for x in ["Ratio", "Log", "Avg Log"])
    ]
    obj_cols   = [
        c for c in df_style.columns
        if c not in KEYS and c not in num_cols and c not in float_cols
    ]

    df_style[num_cols]   = df_style[num_cols].fillna(0).astype(int)
    df_style[float_cols] = df_style[float_cols].fillna(0)
    df_style[obj_cols]   = df_style[obj_cols].fillna("-")

    col_cfg = {}
    for c in num_cols:   col_cfg[c] = st.column_config.NumberColumn(format="%.0f")
    for c in float_cols: col_cfg[c] = st.column_config.NumberColumn(format="%.2f")
    st.dataframe(df_style, column_config=col_cfg, use_container_width=True)

    # ── Download Excel ─────────────────────────────────────────────────────────
    st.header("💾 Unduh Hasil Analisis ABC V3")
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_style.to_excel(writer, sheet_name="All Cities Pivot V3", index=False)
        for city in sorted(result_display["City"].unique()):
            city_data = result_display[result_display["City"] == city]
            city_data.to_excel(writer, sheet_name=city[:31], index=False)
    st.download_button(
        "📥 Unduh Hasil Analisis ABC V3 (Excel)",
        data=output.getvalue(),
        file_name=f"Hasil_Analisis_ABC_V3_{end_date_input}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Tab Dashboard ──────────────────────────────────────────────────────────────
def _render_dashboard_tab():
    if st.session_state.get("abc_v3_result") is None:
        st.info("Tidak ada data. Jalankan analisis terlebih dahulu.")
        return

    result = st.session_state.abc_v3_result.copy()

    col_p, col_m = st.columns(2)
    platform_sel = col_p.selectbox(
        "Platform:",
        ("Overall", "ONLINE", "OFFLINE"),
        key="dash_v3_platform",
    )
    metode_sel = col_m.selectbox(
        "Metode ABC:",
        ("Log-Benchmark - WMA", "Log-Benchmark - Mean"),
        key="dash_v3_metode",
    )

    sfx = "" if platform_sel == "Overall" else f"_{platform_sel}"
    metric_col = f"AVG WMA{sfx}" if "WMA" in metode_sel else f"AVG Mean{sfx}"
    kat_col    = (
        f"Kategori ABC (Log-Benchmark - WMA){sfx}"
        if "WMA" in metode_sel
        else f"Kategori ABC (Log-Benchmark - Mean){sfx}"
    )

    if kat_col not in result.columns:
        st.warning(f"Kolom '{kat_col}' belum tersedia. Jalankan analisis terlebih dahulu.")
        return

    LABELS = ["A", "B", "C", "D", "E", "F"]
    COLORS = ["#cce5ff", "#d4edda", "#fff3cd", "#f8d7da", "#e9ecef", "#6c757d"]

    summary = result.groupby(kat_col)[metric_col].agg(["count", "sum"])
    for label in LABELS:
        if label not in summary.index:
            summary.loc[label] = [0, 0]
    summary    = summary.reindex(LABELS).fillna(0)
    summary["avg_unit"] = np.where(
        summary["count"] > 0, summary["sum"] / summary["count"], 0
    )

    st.markdown("---")
    cols = st.columns(len(LABELS))
    for i, label in enumerate(LABELS):
        count = int(summary.loc[label, "count"])
        avg   = summary.loc[label, "avg_unit"]
        delta = "Tidak Terjual" if label == "F" else f"{avg:.1f} Rata-rata Penjualan"
        cols[i].metric(f"Produk Kelas {label}", f"{count} SKU", delta)

    st.markdown("---")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.subheader(f"Komposisi Produk per Kelas [{platform_sel}] (SKU Count)")
        data_pie = summary[summary["count"] > 0]
        if not data_pie.empty:
            fig, ax = plt.subplots()
            ax.pie(
                data_pie["count"],
                labels=data_pie.index,
                autopct="%1.1f%%",
                startangle=90,
                colors=[COLORS[LABELS.index(i)] for i in data_pie.index],
            )
            ax.axis("equal")
            st.pyplot(fig)
        else:
            st.info("Tidak ada data untuk pie chart.")

    with col_c2:
        st.subheader(f"Kontribusi {metric_col} per Kelas [{platform_sel}]")
        data_bar = summary[summary["sum"] > 0]
        if not data_bar.empty:
            st.bar_chart(data_bar[["sum"]].rename(columns={"sum": metric_col}))
        else:
            st.info("Tidak ada kontribusi penjualan untuk ditampilkan.")

    st.markdown("---")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.subheader(f"Top 10 Produk Terlaris [{platform_sel}]")
        top = result.groupby("Nama Barang")[metric_col].sum().nlargest(10)
        st.bar_chart(top)
    with col_t2:
        st.subheader(f"Performa Penjualan per Kota [{platform_sel}]")
        city_sales = (
            result.groupby("City")[metric_col].sum().sort_values(ascending=False)
        )
        st.bar_chart(city_sales)

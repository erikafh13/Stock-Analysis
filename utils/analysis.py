"""
utils/analysis.py
Semua fungsi kalkulasi & analisis (ABC, WMA, Min/Max Stock, dll.).
"""

import math
import numpy as np
import pandas as pd
import streamlit as st
from io import BytesIO


# ── Konstanta Multiplier ───────────────────────────────────────────────────────
# FIX: Nilai C (0.75) dan D (0.50) dikoreksi agar sinkron dengan info yang
#      ditampilkan ke user di st.info(). Sebelumnya C=0.5, D=0.3 (tidak sinkron).
DAYS_MULTIPLIER: dict[str, float] = {
    "A": 1.00,   # Buffer 30 hari
    "B": 1.00,   # Buffer 30 hari
    "C": 0.75,   # Buffer 24 hari  ← FIX (was 0.5)
    "D": 0.50,   # Buffer 15 hari  ← FIX (was 0.3)
    "E": 0.25,   # Buffer 7 hari
    "F": 0.00,   # Tidak distock
}

MAX_MULTIPLIER: dict[str, float] = {
    "A": 2.0,
    "B": 2.0,
    "C": 1.5,
    "D": 1.0,
    "E": 1.0,
    "F": 0.0,   # F → Max Stock = 1 (ditangani terpisah)
}

# ── Mapping ────────────────────────────────────────────────────────────────────
BULAN_INDONESIA = {
    1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL",
    5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS",
    9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER",
}

CITY_PREFIX_MAP = {
    "SURABAYA": ["A - ITC", "AT - TRANSIT ITC", "C", "C6", "CT - TRANSIT PUSAT",
                 "Y - SBY", "Y3 - Display Y", "YT - TRANSIT Y"],
    "JAKARTA":  ["B", "BT - TRANSIT JKT"],
    "SEMARANG": ["D - SMG", "DT - TRANSIT SMG"],
    "JOGJA":    ["E - JOG", "ET - TRANSIT JOG"],
    "MALANG":   ["F - MLG", "FT - TRANSIT MLG"],
    "BALI":     ["H - BALI", "HT - TRANSIT BALI"],
}


# ── Helper Umum ────────────────────────────────────────────────────────────────
def get_days_multiplier(kategori: str) -> float:
    return DAYS_MULTIPLIER.get(kategori, 1.0)


def map_nama_dept(row) -> str:
    dept = str(row.get("Dept.", "")).strip().upper()
    pelanggan = str(row.get("Nama Pelanggan", "")).strip().upper()
    if dept == "A":
        if pelanggan in ["A - CASH", "AIRPAY INTERNATIONAL INDONESIA", "TOKOPEDIA"]:
            return "A - ITC"
        return "A - RETAIL"
    mapping = {
        "B": "B - JKT", "C": "C - PUSAT", "D": "D - SMG",
        "E": "E - JOG", "F": "F - MLG", "G": "G - PROJECT",
        "H": "H - BALI", "X": "X",
    }
    return mapping.get(dept, "X")


def map_city(nama_dept: str) -> str:
    if nama_dept in ["A - ITC", "A - RETAIL", "C - PUSAT", "G - PROJECT"]:
        return "Surabaya"
    city_map = {
        "B - JKT": "Jakarta", "D - SMG": "Semarang",
        "E - JOG": "Jogja",   "F - MLG": "Malang", "H - BALI": "Bali",
    }
    return city_map.get(nama_dept, "Others")


@st.cache_data
def convert_df_to_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    return output.getvalue()


# ── WMA ────────────────────────────────────────────────────────────────────────
def calculate_daily_wma(group: pd.DataFrame, end_date) -> float:
    """
    Hitung Weighted Moving Average harian selama 90 hari.
    Bobot: 30 hari terakhir=0.5, 31-60=0.3, 61-90=0.2. Hasil di-ceil.
    """
    end_date = pd.to_datetime(end_date)
    r1_start = end_date - pd.DateOffset(days=29)
    r2_start = end_date - pd.DateOffset(days=59)
    r2_end   = end_date - pd.DateOffset(days=30)
    r3_start = end_date - pd.DateOffset(days=89)
    r3_end   = end_date - pd.DateOffset(days=60)

    s1 = group[group["Tgl Faktur"].between(r1_start, end_date)]["Kuantitas"].sum()
    s2 = group[group["Tgl Faktur"].between(r2_start, r2_end)]["Kuantitas"].sum()
    s3 = group[group["Tgl Faktur"].between(r3_start, r3_end)]["Kuantitas"].sum()

    wma = (s1 * 0.5) + (s2 * 0.3) + (s3 * 0.2)
    return math.ceil(wma)


# ── ABC Log-Benchmark ──────────────────────────────────────────────────────────
def classify_abc_log_benchmark(df_grouped: pd.DataFrame, metric_col: str) -> pd.DataFrame:
    """
    Klasifikasi ABC menggunakan metode Log-Benchmark.

    FIX pembagian nol: ratio sekarang menggunakan np.where agar tidak
    menghasilkan inf ketika avg_log_col == 0 (fillna(0) tidak handle inf).
    """
    df = df_grouped.copy()

    if "Kategori Barang" not in df.columns:
        st.warning("Kolom 'Kategori Barang' tidak ada untuk metode benchmark.")
        return df

    clean_metric   = metric_col.replace("SO ", "").replace("AVG ", "")
    log_col        = f"Log (10) {clean_metric}"
    avg_log_col    = f"Avg Log {clean_metric}"
    ratio_col      = f"Ratio Log {clean_metric}"
    kategori_col   = f"Kategori ABC (Log-Benchmark - {clean_metric})"

    # 1. Bulatkan metrik (sinkron dengan Excel)
    df["_rounded"] = df[metric_col].apply(np.round)

    # 2. Log individu (minimum input = 1 agar log tidak negatif)
    df["_log_input"] = np.maximum(1, df["_rounded"]).astype(float)
    df[log_col] = np.where(df[metric_col] > 0, np.log10(df["_log_input"]), np.nan)

    # 3. Rata-rata log per City × Kategori Barang (hanya rounded >= 1)
    valid = df[df["_rounded"] >= 1]
    avg_ref = (
        valid.groupby(["City", "Kategori Barang"])[log_col]
        .mean()
        .reset_index()
        .rename(columns={log_col: avg_log_col})
    )

    df.drop(columns=[avg_log_col], errors="ignore", inplace=True)
    df = pd.merge(df, avg_ref, on=["City", "Kategori Barang"], how="left")
    df[avg_log_col] = df[avg_log_col].fillna(0)

    # 4. Rasio — FIX: gunakan np.where agar avg=0 → ratio=0 (bukan inf)
    df[ratio_col] = np.where(
        df[avg_log_col] != 0,
        df[log_col] / df[avg_log_col],
        0,
    ).astype(float)

    # 5. Kategorisasi
    df[kategori_col] = df.apply(_apply_category_log, axis=1,
                                 args=(metric_col, ratio_col))

    df.drop(columns=["_rounded", "_log_input"], errors="ignore", inplace=True)
    return df


def _apply_category_log(row, metric_col: str, ratio_col: str) -> str:
    if row[metric_col] <= 0 or pd.isna(row[metric_col]):
        return "F"
    ratio = row[ratio_col]
    if ratio > 2:   return "A"
    if ratio > 1.5: return "B"
    if ratio > 1:   return "C"
    if ratio > 0.5: return "D"
    return "E"


# ── Min & Max Stock (Vectorized) ───────────────────────────────────────────────
def calculate_min_stock(df: pd.DataFrame, kategori_col: str, so_col: str) -> pd.Series:
    """
    Hitung Min Stock secara vectorized.
    FIX performa: tidak lagi menggunakan apply(axis=1) row-by-row.
    """
    mult = df[kategori_col].map(DAYS_MULTIPLIER).fillna(1.0)
    min_stock = np.where(
        (df[so_col] <= 0) | (df[kategori_col] == "F"),
        0,
        np.ceil(df[so_col] * mult),
    )
    return pd.Series(min_stock.astype(int), index=df.index)


def calculate_max_stock(df: pd.DataFrame, kategori_col: str, so_col: str) -> pd.Series:
    """
    Hitung Max Stock secara vectorized.
    FIX performa: tidak lagi menggunakan apply(axis=1) row-by-row.
    Kategori F → Max Stock = 1.
    """
    mult = df[kategori_col].map(MAX_MULTIPLIER).fillna(1.0)
    max_stock = np.where(
        df[kategori_col] == "F",
        1,
        np.ceil(df[so_col] * mult),
    )
    return pd.Series(max_stock.astype(int), index=df.index)


def calculate_add_stock(df: pd.DataFrame, kategori_col: str,
                        min_stock_col: str, stock_col: str) -> pd.Series:
    """Hitung Add Stock secara vectorized."""
    add_stock = np.where(
        df[kategori_col] == "F",
        0,
        np.maximum(0, df[min_stock_col] - df[stock_col]),
    )
    return pd.Series(add_stock.astype(int), index=df.index)


# ── Suggested PO per Cabang (Proporsional) ────────────────────────────────────
def calculate_suggested_po(df: pd.DataFrame) -> pd.Series:
    """
    Hitung Suggested PO per cabang secara vectorized dengan logika proporsional.

    Logika:
    - Surabaya adalah gudang pusat → tidak menerima kiriman dari dirinya sendiri
    - Jika Total Add Stock non-Surabaya = 0 → semua PO = 0
    - Hitung porsi tiap cabang: Add Stock Cabang / Total Add Stock non-Surabaya
    - Jika Stock Surabaya >= Total Add Stock non-Surabaya:
        → Tiap cabang dapat Add Stock penuh
    - Jika Stock Surabaya < Total Add Stock non-Surabaya:
        → Tiap cabang dapat CEIL(porsi × Stock Surabaya)
    """
    result = pd.Series(0, index=df.index, dtype=int)

    for no_barang, group in df.groupby("No. Barang"):
        # Pisahkan Surabaya (sumber) dari cabang penerima
        mask_sby    = group["City"] == "SURABAYA"
        mask_cabang = ~mask_sby

        stock_sby_val = group.loc[mask_sby, "Stock Surabaya"].sum()
        add_cabang    = group.loc[mask_cabang, "Add Stock"]
        total_add     = add_cabang.sum()

        # Jika tidak ada kebutuhan sama sekali → semua 0
        if total_add == 0:
            continue

        if stock_sby_val >= total_add:
            # Stok Surabaya cukup → tiap cabang dapat Add Stock penuh
            result.loc[add_cabang.index] = add_cabang.values
        else:
            # Stok Surabaya tidak cukup → bagi proporsional
            porsi = add_cabang / total_add
            po_proporsional = np.ceil(porsi * stock_sby_val).astype(int)
            result.loc[add_cabang.index] = po_proporsional.values

        # Surabaya sebagai cabang → selalu 0 (tidak kirim ke diri sendiri)
        result.loc[group.loc[mask_sby].index] = 0

    return result


def get_status_stock(row) -> str:
    kategori = row.get("Kategori ABC (Log-Benchmark - WMA)", "")
    if kategori == "F":
        return "Overstock F" if row["Stock Cabang"] > 2 else "Balance"
    if row["Stock Cabang"] > row["Max Stock"]:  return "Overstock"
    if row["Stock Cabang"] < row["Min Stock"]:  return "Understock"
    if row["Stock Cabang"] >= row["Min Stock"]: return "Balance"
    return "-"


# ── Stock Cabang Melt ──────────────────────────────────────────────────────────
def melt_stock_by_city(stock_df_raw: pd.DataFrame) -> pd.DataFrame:
    """Ubah kolom-kolom gudang menjadi format panjang (City, Stock)."""
    stok_cols = [c for c in stock_df_raw.columns if c not in ["No. Barang", "Keterangan Barang"]]
    result = []
    for city_name, prefixes in CITY_PREFIX_MAP.items():
        city_cols = [c for c in stok_cols if any(c.startswith(p) for p in prefixes)]
        if not city_cols:
            continue
        tmp = stock_df_raw[["No. Barang"] + city_cols].copy()
        tmp["Stock"] = tmp[city_cols].sum(axis=1)
        tmp["City"]  = city_name
        result.append(tmp[["No. Barang", "City", "Stock"]])
    return pd.concat(result, ignore_index=True) if result else pd.DataFrame()


# ── Styling ────────────────────────────────────────────────────────────────────
def highlight_kategori_abc_log(val: str) -> str:
    warna = {
        "A": "#cce5ff", "B": "#d4edda", "C": "#fff3cd",
        "D": "#f8d7da", "E": "#e9ecef", "F": "#6c757d",
    }
    color = warna.get(val, "")
    text  = "white" if val == "F" else "black"
    return f"background-color: {color}; color: {text}"


def highlight_status_stock(val: str) -> str:
    colors = {
        "Understock":  "#fff3cd",
        "Balance":     "#d4edda",
        "Overstock":   "#ffd6a5",
        "Overstock F": "#f5c6cb",
    }
    return f"background-color: {colors.get(val, '')}"


# ==============================================================================
# LOGIKA V2 — 3 Skenario Distribusi (tanpa buffer lead time)
# ==============================================================================

LEAD_TIME_CABANG   = 7   # hari (referensi saja, tidak dipakai di Add Stock)
LEAD_TIME_SUPPLIER = 30  # hari pengiriman distributor → Surabaya

# Urutan prioritas kategori ABC
_ABC_ORDER = {"A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 99}


def calculate_add_stock_v2(df: pd.DataFrame, kategori_col: str,
                            so_col: str, stock_col: str) -> pd.Series:
    """
    Hitung Add Stock V2 — TANPA buffer lead time.
    Formula sama dengan V1:
        Add Stock = Max(0, Min Stock - Stock Cabang)
        Kategori F → 0
    """
    mult      = df[kategori_col].map(DAYS_MULTIPLIER).fillna(1.0)
    min_stock = np.where(
        (df[so_col] <= 0) | (df[kategori_col] == "F"),
        0,
        np.ceil(df[so_col] * mult),
    )
    add_stock = np.where(
        df[kategori_col] == "F",
        0,
        np.maximum(0, min_stock - df[stock_col]),
    )
    return pd.Series(add_stock.astype(int), index=df.index)

def calculate_persentase_stock(df: pd.DataFrame) -> pd.Series:
    """
    Hitung persentase stock terhadap Min Stock.
    """
    persen = np.where(
        df["Min Stock"] > 0,
        (df["Stock Cabang"] / df["Min Stock"]) * 100,
        0
    )
    return pd.Series(np.round(persen, 1), index=df.index)

def calculate_suggested_po_v2(df: pd.DataFrame) -> pd.Series:
    """
    Hitung Suggested PO V2 dengan 3 skenario distribusi final.

    Definisi:
        Add_Stock_Sby  = Add Stock toko Surabaya sendiri
        Stok_Sisa      = Max(0, Stock Surabaya - Add_Stock_Sby)
        All_Add_Cabang = SUM(Add Stock semua cabang NON-Surabaya)

        TIDAK AMAN = Stock Cabang < 50% x Add Stock
                     AND Add Stock > 0
                     AND Kategori != F
        AMAN       = Stock Cabang >= 50% x Add Stock
                     AND Add Stock > 0
                     AND Kategori != F
        SKIP       = Add Stock = 0 OR Kategori = F → PO = 0

    3 Skenario:
        1. KURANG       : Stok_Sisa = 0
                          → Semua cabang PO = 0

        2. TERBATAS_LEBIH: 0 < Stok_Sisa < All_Add_Cabang
                          STEP 1: Proses TIDAK AMAN (ABC → SO WMA)
                          STEP 2: Jika masih ada sisa → proses AMAN (ABC → SO WMA)
                          Dalam tiap kategori: penuh jika cukup, proporsional jika tidak

        3. OVER         : Stok_Sisa >= All_Add_Cabang
                          → Semua cabang dapat Add Stock penuh
    """
    # Urutan prioritas kategori ABC
    ABC_ORDER = {"A": 0, "B": 1, "C": 2, "D": 3, "E": 4, "F": 99}

    result = pd.Series(0, index=df.index, dtype=int)

    for no_barang, group in df.groupby("No. Barang"):
        mask_sby = group["City"] == "SURABAYA"
        mask_cab = ~mask_sby

        # ── Data Surabaya ──────────────────────────────────────────────────────
        sby_rows      = group.loc[mask_sby]
        stock_sby     = sby_rows["Stock Cabang"].sum()
        min_stock_sby = sby_rows["Min Stock"].sum()
        stok_sisa     = max(0, stock_sby - min_stock_sby)

        # Surabaya selalu PO = 0
        result.loc[sby_rows.index] = 0

        # ── Data Cabang Non-Surabaya ───────────────────────────────────────────
        cab_rows      = group.loc[mask_cab].copy()
        all_add_cabang = cab_rows["Add Stock"].sum()

        if all_add_cabang == 0:
            continue

        # ── Skenario 1: KURANG ────────────────────────────────────────────────
        if stok_sisa <= 0:
            result.loc[cab_rows.index] = 0
            continue

        # ── Skenario 3: OVER ──────────────────────────────────────────────────
        if stok_sisa >= all_add_cabang:
            # Semua cabang dapat penuh (kecuali F dan Add Stock = 0)
            po_over = np.where(
                (cab_rows["Kategori ABC (Log-Benchmark - WMA)"] == "F") | (cab_rows["Add Stock"] == 0),
                0,
                cab_rows["Add Stock"]
            )
            result.loc[cab_rows.index] = po_over
            continue

        # ── Skenario 2: TERBATAS & LEBIH ─────────────────────────────────────
        # 0 < Stok_Sisa < All_Add_Cabang

        # Klasifikasi tiap cabang
        kat_col = "Kategori ABC (Log-Benchmark - WMA)"

        def _klasifikasi(row):
            if row["Add Stock"] == 0 or row[kat_col] == "F":
                return "SKIP"
            if row["Stock Cabang"] < 0.5 * row["Add Stock"]:
                return "TIDAK_AMAN"
            return "AMAN"

        cab_rows["_status"] = cab_rows.apply(_klasifikasi, axis=1)

        # Pisahkan grup
        tidak_aman = cab_rows[cab_rows["_status"] == "TIDAK_AMAN"].copy()
        aman       = cab_rows[cab_rows["_status"] == "AMAN"].copy()
        # SKIP → PO = 0 (default sudah 0)

        sisa = stok_sisa

        # ── Fungsi distribusi per grup ─────────────────────────────────────────
        def _distribusi(grup: pd.DataFrame, stok_tersedia: int) -> tuple:
            """
            Distribusi berbasis:
            - Urgency (dari Persentase Stock per cabang)
            - Sequential (prioritas tertinggi dulu)
            - Target: 50% Min Stock
            - Tidak harus menghabiskan stok Surabaya
            """

            if grup.empty or stok_tersedia <= 0:
                return pd.Series(0, index=grup.index, dtype=int), stok_tersedia

            po = pd.Series(0, index=grup.index, dtype=int)

            # =========================
            # 🔥 HITUNG URGENCY DINAMIS
            # =========================
            persen = pd.Series(index=grup.index, dtype=float)

            for city in grup["City"].unique():
                mask = grup["City"] == city
                col_name = f"{city}_Persentase Stock"

                if col_name in grup.columns:
                    persen.loc[mask] = grup.loc[mask, col_name]
                else:
                    persen.loc[mask] = 100  # fallback aman

            persen = persen.fillna(100)

            # urgency: makin kecil persen → makin tinggi
            urgency = 1 - (persen / 100)

            grup["_urgency"] = urgency

            # =========================
            # 🔥 URUTKAN PRIORITAS
            # =========================
            grup_sorted = grup.sort_values("_urgency", ascending=False)

            # =========================
            # 🔥 DISTRIBUSI SEQUENTIAL
            # =========================
            for idx, row in grup_sorted.iterrows():
                if stok_tersedia <= 0:
                    break

                # target minimal (50% min stock)
                target_min = 0.5 * row["Min Stock"]

                # kebutuhan aktual
                need = max(0, target_min - row["Stock Cabang"])

                if need <= 0:
                    continue

                # alokasi
                alloc = min(need, stok_tersedia)

                po[idx] = int(np.floor(alloc))
                stok_tersedia -= po[idx]

            return po, stok_tersedia

        # STEP 1: Proses TIDAK AMAN dulu
        po_tidak_aman, sisa = _distribusi(tidak_aman, sisa)
        result.loc[tidak_aman.index] = po_tidak_aman.values

        # STEP 2: Jika masih ada sisa → proses AMAN
        if sisa > 0 and not aman.empty:
            po_aman, sisa = _distribusi(aman, sisa)
            result.loc[aman.index] = po_aman.values

    return result


def calculate_all_summary_v2(df: pd.DataFrame) -> pd.DataFrame:
    """
    Hitung kolom summary ALL untuk logika V2.

    Kolom yang dihasilkan per No. Barang:
        All_Add_Stock_Cabang   : Total Add Stock cabang NON-Surabaya
        All_Add_Stock_Surabaya : Add Stock toko Surabaya sendiri
        All_Need_From_Supplier : Max(0, Kol1 + Kol2 - Stock Surabaya)
        All_Restock_1_Bulan    : "PO" jika Need_Supplier > 0, else "NO"
        Skenario_Distribusi    : Label skenario (KURANG / TERBATAS_LEBIH / OVER)
    """
    KEYS = ["No. Barang", "Kategori Barang", "BRAND Barang", "Nama Barang"]
    rows = []

    for no_barang, group in df.groupby("No. Barang"):
        mask_sby = group["City"] == "SURABAYA"
        mask_cab = ~mask_sby

        sby_rows       = group.loc[mask_sby]
        stock_sby      = sby_rows["Stock Cabang"].sum()
        min_stock_sby = sby_rows["Min Stock"].sum()
        stok_sisa     = max(0, stock_sby - min_stock_sby)
        all_add_cabang = group.loc[mask_cab, "Add Stock"].sum()
        need_supplier  = max(0, all_add_cabang - (stock_sby - min_stock_sby))

        # Tentukan skenario (3 skenario final)
        if stok_sisa <= 0:
            skenario = "1 - KURANG"
        elif stok_sisa >= all_add_cabang:
            skenario = "3 - OVER"
        else:
            skenario = "2 - TERBATAS/LEBIH"

        first = group.iloc[0]
        row   = {k: first[k] for k in KEYS if k in group.columns}
        row.update({
            "All_Add_Stock_Cabang":   int(all_add_cabang),
            "All_Need_From_Supplier": int(need_supplier),
            "All_Restock_1_Bulan":    "PO" if need_supplier > 0 else "NO",
            "Skenario_Distribusi":    skenario,
        })
        rows.append(row)

    return pd.DataFrame(rows)

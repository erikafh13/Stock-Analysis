"""
DB — database pusat (MySQL via XAMPP) untuk SEMUA data scraping
====================================================================
Ganti penyimpanan file (Excel/JSON/CSV) jadi MySQL. Dipakai bersama oleh
scraping_new.py, scrape_varian.py, kirim_ke_portal.py, bersihkan_data_lama.py.

SETUP YANG DIBUTUHKAN:
  1. Jalankan XAMPP, nyalakan modul MySQL (klik "Start" di XAMPP Control Panel).
  2. Buka http://localhost/phpmyadmin -- pastikan bisa diakses (server jalan).
  3. TIDAK PERLU bikin database manual -- script ini otomatis membuatnya
     ('shopee_scraper') beserta semua tabel, kalau belum ada.
  4. Isi .env (opsional, kalau beda dari default XAMPP):
       DB_HOST=127.0.0.1
       DB_PORT=3306
       DB_USER=root
       DB_PASS=
       DB_NAME=shopee_scraper
     Default XAMPP: user root, password KOSONG -- kalau kamu belum ubah
     password root XAMPP, biarkan .env kosong/tidak usah diisi, otomatis
     pakai default ini.

TABEL:
  produk          - hasil scraping toko (pengganti Excel Tersedia/Habis)
  progress        - checkpoint per toko per run (pengganti file JSON progress/)
  varian          - hasil scraping varian (pengganti Excel output_varian/)
  varian_memori   - kata & skornya untuk belajar deteksi varian (pengganti
                    varian_memory.json)
  varian_produk_dipelajari - histori produk yang sudah dipelajari (biar tidak
                    dihitung dobel kalau di-scrape ulang)
  captcha_log     - catatan tiap captcha (pengganti captcha_log.csv)
"""

import os
import re
import sys
from datetime import datetime, timedelta

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import mysql.connector
    from mysql.connector import pooling
except ImportError:
    print("[X] mysql-connector-python belum terpasang. Jalankan:")
    print("    pip install mysql-connector-python")
    sys.exit(1)

DB_HOST = os.environ.get("DB_HOST", "127.0.0.1")
DB_PORT = int(os.environ.get("DB_PORT", "3306"))
DB_USER = os.environ.get("DB_USER", "root")
DB_PASS = os.environ.get("DB_PASS", "")
DB_NAME = os.environ.get("DB_NAME", "shopee_scraper")

# ── SSL -- WAJIB untuk database cloud (Aiven, PlanetScale, dll) yang
# menolak koneksi tanpa enkripsi. Untuk XAMPP lokal biasanya TIDAK perlu
# (biarkan kosong di .env). Isi DB_SSL_CA dengan path file CA certificate
# yang didownload dari dashboard provider cloud-mu (mis. ca.pem dari Aiven). ──
DB_SSL_CA = os.environ.get("DB_SSL_CA", "").strip() or None

_pool = None


def _param_koneksi(**tambahan):
    """Kumpulkan parameter koneksi dasar (host/port/user/pass + SSL kalau
    diisi), supaya tidak ditulis berulang di tiap fungsi yang connect."""
    params = dict(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASS,
        connection_timeout=10,
    )
    if DB_SSL_CA:
        # HANYA ssl_ca -- JANGAN tambahkan ssl_verify_cert=True.
        # Kombinasi keduanya di beberapa versi mysql-connector-python
        # (termasuk yang dipakai di sini) memicu error
        # "SSL_CTX_set_default_verify_paths failed" -- connector mencoba
        # cari sertifikat SISTEM DEFAULT (bukan yang kita kasih), gagal,
        # padahal ssl_ca sendirian SUDAH CUKUP untuk verifikasi terhadap
        # sertifikat spesifik yang kita berikan (mis. ca.pem dari Aiven).
        params["ssl_ca"] = DB_SSL_CA
    params.update(tambahan)
    return params


def _koneksi_tanpa_db():
    """Koneksi ke server MySQL TANPA pilih database -- dipakai sekali di awal
    untuk membuat database kalau belum ada."""
    return mysql.connector.connect(**_param_koneksi())


def pastikan_database_ada():
    """Buat database DB_NAME kalau belum ada. Dipanggil sekali di awal
    (init_db()) sebelum connection pool dibuat."""
    conn = _koneksi_tanpa_db()
    cur = conn.cursor()
    cur.execute(f"CREATE DATABASE IF NOT EXISTS `{DB_NAME}` "
               f"CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
    conn.commit()
    cur.close()
    conn.close()


SKEMA_TABEL = {
    "produk": """
        CREATE TABLE IF NOT EXISTS produk (
            id INT AUTO_INCREMENT PRIMARY KEY,
            toko VARCHAR(255) NOT NULL,
            nama_produk TEXT NOT NULL,
            harga VARCHAR(100),
            penjualan_bulan INT DEFAULT 0,
            pendapatan_bulan BIGINT DEFAULT 0,
            penjualan_minggu INT DEFAULT 0,
            pendapatan_minggu BIGINT DEFAULT 0,
            rating VARCHAR(20),
            url_produk TEXT,
            status ENUM('tersedia','habis') NOT NULL DEFAULT 'tersedia',
            run_id VARCHAR(20) NOT NULL,
            tanggal DATE NOT NULL,
            dibuat_pada DATETIME DEFAULT CURRENT_TIMESTAMP,
            INDEX idx_toko_tanggal (toko, tanggal),
            INDEX idx_run_id (run_id)
        ) ENGINE=InnoDB
    """,
    "progress": """
        CREATE TABLE IF NOT EXISTS progress (
            id INT AUTO_INCREMENT PRIMARY KEY,
            run_id VARCHAR(20) NOT NULL,
            toko VARCHAR(255) NOT NULL,
            status VARCHAR(50) NOT NULL,
            halaman_terakhir INT DEFAULT 0,
            jumlah_produk INT DEFAULT 0,
            diperbarui_pada DATETIME DEFAULT CURRENT_TIMESTAMP
                ON UPDATE CURRENT_TIMESTAMP,
            UNIQUE KEY uniq_run_toko (run_id, toko)
        ) ENGINE=InnoDB
    """,
    "varian": """
        CREATE TABLE IF NOT EXISTS varian (
            id INT AUTO_INCREMENT PRIMARY KEY,
            toko VARCHAR(255),
            nama_produk TEXT NOT NULL,
            grup_varian VARCHAR(255),
            varian VARCHAR(255),
            harga VARCHAR(100),
            tipe VARCHAR(50),
            url_produk TEXT,
            tanggal DATE NOT NULL,
            dibuat_pada DATETIME DEFAULT CURRENT_TIMESTAMP,
            INDEX idx_toko_tanggal (toko, tanggal)
        ) ENGINE=InnoDB
    """,
    "varian_memori": """
        CREATE TABLE IF NOT EXISTS varian_memori (
            kata VARCHAR(100) PRIMARY KEY,
            jumlah_varian INT DEFAULT 0,
            jumlah_single INT DEFAULT 0,
            diperbarui_pada DATETIME DEFAULT CURRENT_TIMESTAMP
                ON UPDATE CURRENT_TIMESTAMP
        ) ENGINE=InnoDB
    """,
    "varian_produk_dipelajari": """
        CREATE TABLE IF NOT EXISTS varian_produk_dipelajari (
            url_produk VARCHAR(768) PRIMARY KEY,
            nama_produk TEXT,
            punya_varian BOOLEAN NOT NULL,
            ada_harga_beda BOOLEAN NOT NULL DEFAULT FALSE,
            dipelajari_pada DATETIME DEFAULT CURRENT_TIMESTAMP
        ) ENGINE=InnoDB
    """,
    "captcha_log": """
        CREATE TABLE IF NOT EXISTS captcha_log (
            id INT AUTO_INCREMENT PRIMARY KEY,
            toko VARCHAR(255),
            akun VARCHAR(100),
            jenis VARCHAR(50),
            url TEXT,
            waktu DATETIME DEFAULT CURRENT_TIMESTAMP
        ) ENGINE=InnoDB
    """,
    "pengguna": """
        CREATE TABLE IF NOT EXISTS pengguna (
            orang_ke INT PRIMARY KEY,
            nama VARCHAR(100) NOT NULL,
            nomor_wa VARCHAR(20),
            diperbarui_pada DATETIME DEFAULT CURRENT_TIMESTAMP
                ON UPDATE CURRENT_TIMESTAMP
        ) ENGINE=InnoDB
    """,
}


def init_db():
    """Panggil SEKALI di awal tiap script (scraping_new.py, scrape_varian.py,
    dll) sebelum operasi database lain. Bikin database & semua tabel kalau
    belum ada -- jadi setup-nya 'sekali colok langsung jalan', tidak perlu
    bikin tabel manual di phpMyAdmin."""
    global _pool
    try:
        pastikan_database_ada()
    except mysql.connector.Error as e:
        print(f"[X] Gagal konek ke MySQL ({DB_HOST}:{DB_PORT}): {e}")
        print("    Kalau pakai XAMPP lokal: pastikan sudah di-Start.")
        print("    Kalau pakai database cloud (mis. Aiven): cek DB_SSL_CA")
        print("    di .env sudah menunjuk ke file CA certificate yang benar.")
        raise

    conn = mysql.connector.connect(**_param_koneksi(database=DB_NAME))
    cur = conn.cursor()
    for nama_tabel, ddl in SKEMA_TABEL.items():
        cur.execute(ddl)

    # ── MIGRASI: tambah kolom baru ke tabel yang SUDAH ADA sebelumnya ──
    # CREATE TABLE IF NOT EXISTS tidak menambah kolom ke tabel lama yang
    # sudah terlanjur dibuat (mis. database Aiven yang sudah jalan lama).
    # Cek dulu apakah kolom sudah ada, kalau belum baru ALTER TABLE.
    def _pastikan_kolom(tabel, kolom, definisi):
        cur.execute(
            "SELECT COUNT(*) FROM information_schema.columns "
            "WHERE table_schema=%s AND table_name=%s AND column_name=%s",
            (DB_NAME, tabel, kolom)
        )
        if cur.fetchone()[0] == 0:
            cur.execute(f"ALTER TABLE {tabel} ADD COLUMN {kolom} {definisi}")
            print(f"[*] Migrasi: kolom '{kolom}' ditambahkan ke tabel '{tabel}'.")

    _pastikan_kolom("varian_produk_dipelajari", "ada_harga_beda",
                    "BOOLEAN NOT NULL DEFAULT FALSE")

    conn.commit()
    cur.close()
    conn.close()

    _pool = pooling.MySQLConnectionPool(
        pool_name="shopee_pool", pool_size=5,
        **_param_koneksi(database=DB_NAME)
    )
    print(f"[*] Database '{DB_NAME}' siap ({len(SKEMA_TABEL)} tabel)."
          + (" [SSL aktif]" if DB_SSL_CA else ""))


def get_conn():
    """Ambil koneksi dari pool. Panggil init_db() dulu sebelum ini."""
    if _pool is None:
        init_db()
    return _pool.get_connection()


# ══════════════════════════════════════════════════════
# PRODUK TOKO — pengganti Excel Tersedia/Habis
# ══════════════════════════════════════════════════════

def _bersihkan_angka(nilai):
    """Ubah nilai angka yang mungkin masih berformat TAMPILAN ('Rp85.410.200',
    '1.234', 'Rp 500rb', dll) jadi integer murni untuk disimpan ke database.

    KENAPA PERLU: kolom penjualan_bulan/pendapatan_bulan/dst di database
    bertipe INT/BIGINT, tapi data dari scraping_new.py sering masih dalam
    format TAMPILAN Excel (ada 'Rp', titik pemisah ribuan) -- MySQL menolak
    string seperti itu dengan error 'Incorrect integer value'.

    Aman untuk berbagai bentuk input: sudah int (dibiarkan), string berformat
    Rupiah, string angka polos, None/kosong (jadi 0), atau sampah yang sama
    sekali tidak bisa dibaca (jadi 0, bukan bikin crash)."""
    if nilai is None:
        return 0
    if isinstance(nilai, (int,)):
        return nilai
    if isinstance(nilai, float):
        return int(nilai)
    # String: buang semua karakter selain digit (buang "Rp", ".", ",", " ", dll)
    teks = str(nilai).strip()
    if not teks:
        return 0
    hanya_digit = re.sub(r"[^\d]", "", teks)
    if not hanya_digit:
        return 0
    try:
        return int(hanya_digit)
    except ValueError:
        return 0


def simpan_produk(toko, produk_list, status, run_id, tanggal=None):
    """Simpan banyak produk sekaligus (bulk insert). produk_list = list of
    dict dengan key: nama_produk, harga, penjualan_bulan, pendapatan_bulan,
    penjualan_minggu, pendapatan_minggu, rating, url_produk.

    Menghapus dulu data toko+tanggal+status yang lama (kalau run ulang di
    hari yang sama), supaya tidak menumpuk duplikat -- perilaku setara
    dengan Excel yang MENIMPA file lama."""
    if not produk_list:
        return 0
    tanggal = tanggal or datetime.now().date()
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute(
            "DELETE FROM produk WHERE toko=%s AND tanggal=%s AND status=%s",
            (toko, tanggal, status)
        )
        rows = [
            (toko, p.get("nama_produk") or p.get("Nama Produk"),
             p.get("harga") or p.get("Harga"),
             _bersihkan_angka(p.get("penjualan_bulan") or p.get("Penjualan (bln)")),
             _bersihkan_angka(p.get("pendapatan_bulan") or p.get("Pendapatan (bln)")),
             _bersihkan_angka(p.get("penjualan_minggu") or p.get("Penjualan (mg)")),
             _bersihkan_angka(p.get("pendapatan_minggu") or p.get("Pendapatan (mg)")),
             p.get("rating") or p.get("Rating"),
             p.get("url_produk") or p.get("URL Produk"),
             status, run_id, tanggal)
            for p in produk_list
        ]
        cur.executemany(
            """INSERT INTO produk
               (toko, nama_produk, harga, penjualan_bulan, pendapatan_bulan,
                penjualan_minggu, pendapatan_minggu, rating, url_produk,
                status, run_id, tanggal)
               VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
            rows
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def ambil_produk_toko(toko, tanggal=None, status=None):
    """Ambil produk toko tertentu (tanggal default = hari ini). Return
    list of dict, key sesuai nama kolom."""
    tanggal = tanggal or datetime.now().date()
    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)
        if status:
            cur.execute(
                "SELECT * FROM produk WHERE toko=%s AND tanggal=%s AND status=%s",
                (toko, tanggal, status)
            )
        else:
            cur.execute(
                "SELECT * FROM produk WHERE toko=%s AND tanggal=%s",
                (toko, tanggal)
            )
        return cur.fetchall()
    finally:
        conn.close()


def toko_yang_ada_data(tanggal=None):
    """Daftar nama toko yang SUDAH punya data di tanggal tertentu (default
    hari ini). Dipakai kirim_ke_portal.py untuk cek kelengkapan."""
    tanggal = tanggal or datetime.now().date()
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT toko FROM produk WHERE tanggal=%s", (tanggal,))
        return [r[0] for r in cur.fetchall()]
    finally:
        conn.close()


# ══════════════════════════════════════════════════════
# PROGRESS / CHECKPOINT — pengganti file JSON progress/
# ══════════════════════════════════════════════════════

def simpan_progress(run_id, toko, status, halaman_terakhir=0, jumlah_produk=0):
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO progress (run_id, toko, status, halaman_terakhir, jumlah_produk)
               VALUES (%s,%s,%s,%s,%s)
               ON DUPLICATE KEY UPDATE status=%s, halaman_terakhir=%s, jumlah_produk=%s""",
            (run_id, toko, status, halaman_terakhir, jumlah_produk,
             status, halaman_terakhir, jumlah_produk)
        )
        conn.commit()
    finally:
        conn.close()


def ambil_progress(run_id, toko):
    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM progress WHERE run_id=%s AND toko=%s", (run_id, toko))
        return cur.fetchone()
    finally:
        conn.close()


# ══════════════════════════════════════════════════════
# VARIAN — pengganti Excel output_varian/
# ══════════════════════════════════════════════════════

def simpan_varian(toko, baris_varian, tanggal=None):
    """baris_varian = list of dict dengan key: Nama Produk, Grup Varian,
    Varian, Harga, Tipe, URL (format yang sudah dipakai scrape_varian.py)."""
    if not baris_varian:
        return 0
    tanggal = tanggal or datetime.now().date()
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM varian WHERE toko=%s AND tanggal=%s", (toko, tanggal))
        rows = [
            (toko, b.get("Nama Produk"), b.get("Grup Varian"), b.get("Varian"),
             b.get("Harga"), b.get("Tipe"), b.get("URL"), tanggal)
            for b in baris_varian
        ]
        cur.executemany(
            """INSERT INTO varian
               (toko, nama_produk, grup_varian, varian, harga, tipe, url_produk, tanggal)
               VALUES (%s,%s,%s,%s,%s,%s,%s,%s)""",
            rows
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


# ══════════════════════════════════════════════════════
# MEMORI BELAJAR VARIAN — pengganti varian_memory.json
# ══════════════════════════════════════════════════════

def catat_hasil_belajar(url, nama, punya_varian, ada_harga_beda=False):
    """Catat FAKTA (dari buka card) ke database, update skor kata-katanya.
    Setara dengan catat_hasil() versi file JSON, tapi ke MySQL.

    ada_harga_beda: True kalau produk ini punya varian yang HARGANYA
    BERBEDA (storage/RAM/tipe -- yang benar-benar diklik & dicatat harga
    presisinya). False kalau produk single ATAU varian yang ada cuma
    warna/packing/harga-sama (tidak ada nilai baru didapat dari buka
    ulang). Dipakai boleh_skip_produk() supaya run berikutnya TIDAK buka
    ulang produk yang sudah pasti tidak akan menghasilkan info baru."""
    import re as _re

    conn = get_conn()
    try:
        cur = conn.cursor()

        # Cek apakah produk ini sudah pernah dipelajari (hindari hitung dobel)
        cur.execute("SELECT punya_varian FROM varian_produk_dipelajari WHERE url_produk=%s",
                    (url[:768],))
        row = cur.fetchone()
        sudah = row is not None
        sebelum = bool(row[0]) if row else None

        cur.execute(
            """INSERT INTO varian_produk_dipelajari
               (url_produk, nama_produk, punya_varian, ada_harga_beda)
               VALUES (%s,%s,%s,%s)
               ON DUPLICATE KEY UPDATE nama_produk=%s, punya_varian=%s, ada_harga_beda=%s""",
            (url[:768], nama, punya_varian, ada_harga_beda,
             nama, punya_varian, ada_harga_beda)
        )

        if sudah and sebelum == bool(punya_varian):
            conn.commit()
            return   # tidak ada perubahan, tidak perlu update skor kata

        # Pecah nama jadi kata (sama seperti versi file)
        n = nama.lower()
        tokens = _re.findall(r'[a-z]+[0-9]*[a-z]*|[0-9]+(?:gb|tb|mb|inch|mm|cm|w|watt|ml)', n)
        kata_list = list(set(t for t in tokens if len(t) >= 2 and not t.isdigit()))

        for kata in kata_list:
            if sudah and sebelum is not None:
                # koreksi kontribusi lama dulu
                if sebelum:
                    cur.execute(
                        "UPDATE varian_memori SET jumlah_varian = GREATEST(0, jumlah_varian-1) "
                        "WHERE kata=%s", (kata,))
                else:
                    cur.execute(
                        "UPDATE varian_memori SET jumlah_single = GREATEST(0, jumlah_single-1) "
                        "WHERE kata=%s", (kata,))
            kolom = "jumlah_varian" if punya_varian else "jumlah_single"
            cur.execute(
                f"""INSERT INTO varian_memori (kata, {kolom}) VALUES (%s, 1)
                    ON DUPLICATE KEY UPDATE {kolom} = {kolom} + 1""",
                (kata,)
            )
        conn.commit()
    finally:
        conn.close()


def boleh_skip_produk(url):
    """True kalau produk ini SUDAH PERNAH dipelajari dan TIDAK AKAN
    menghasilkan info baru kalau dibuka lagi -- yaitu produk single (tidak
    ada varian sama sekali), ATAU produk yang varian-nya cuma warna/packing
    dengan harga sama (bukan storage/RAM/tipe yang harganya beda).

    Produk dengan ada_harga_beda=True TETAP dibuka lagi tiap run -- karena
    harga bisa berubah dari waktu ke waktu, jadi masih ada nilai info baru."""
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT ada_harga_beda FROM varian_produk_dipelajari WHERE url_produk=%s",
            (url[:768],)
        )
        row = cur.fetchone()
        if row is None:
            return False   # belum pernah dipelajari -- WAJIB dibuka
        return not bool(row[0])   # sudah dipelajari & TIDAK ada harga beda -> skip
    finally:
        conn.close()


def url_yang_boleh_skip(daftar_url):
    """Versi BULK dari boleh_skip_produk() -- satu query untuk banyak URL
    sekaligus, jauh lebih cepat daripada query satu-satu per produk saat
    mau proses ratusan produk dalam satu toko."""
    if not daftar_url:
        return set()
    conn = get_conn()
    try:
        cur = conn.cursor()
        urls_potong = [u[:768] for u in daftar_url]
        format_strings = ",".join(["%s"] * len(urls_potong))
        cur.execute(
            f"SELECT url_produk FROM varian_produk_dipelajari "
            f"WHERE url_produk IN ({format_strings}) AND ada_harga_beda=FALSE",
            tuple(urls_potong)
        )
        return set(r[0] for r in cur.fetchall())
    finally:
        conn.close()


def skor_nama_db(nama):
    """Nilai kemungkinan nama ini punya varian, dari skor kata-kata di
    database. Return (skor -1..+1, penjelasan) -- sama seperti skor_nama()
    versi file JSON."""
    import re as _re
    n = nama.lower()
    tokens = _re.findall(r'[a-z]+[0-9]*[a-z]*|[0-9]+(?:gb|tb|mb|inch|mm|cm|w|watt|ml)', n)
    kata_list = list(set(t for t in tokens if len(t) >= 2 and not t.isdigit()))
    if not kata_list:
        return 0.0, "tidak ada kata dikenali"

    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)
        format_strings = ",".join(["%s"] * len(kata_list))
        cur.execute(
            f"SELECT * FROM varian_memori WHERE kata IN ({format_strings})",
            tuple(kata_list)
        )
        hasil = {r["kata"]: r for r in cur.fetchall()}
    finally:
        conn.close()

    total, dihitung, detail = 0.0, 0, []
    for kata in kata_list:
        slot = hasil.get(kata)
        if not slot:
            continue
        v, s = slot["jumlah_varian"], slot["jumlah_single"]
        n_total = v + s
        if n_total < 2:
            continue
        skor_kata = (v - s) / n_total
        total += skor_kata
        dihitung += 1
        if abs(skor_kata) > 0.3:
            detail.append(f"{kata}:{skor_kata:+.1f}")

    if dihitung == 0:
        return 0.0, "belum ada data kata yang dikenal"
    return total / dihitung, (", ".join(detail) if detail else "netral")


def statistik_memori():
    """Ringkasan memori: total produk dipelajari, jumlah kata dikenal, dll."""
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM varian_produk_dipelajari")
        total_produk = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM varian_produk_dipelajari WHERE punya_varian=1")
        total_varian = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM varian_memori")
        total_kata = cur.fetchone()[0]
        return {"total_dibuka": total_produk, "total_varian": total_varian,
                "total_kata": total_kata}
    finally:
        conn.close()


def kata_teratas(urutan="desc", limit=8, minimal_data=3):
    """Kata dengan skor tertinggi ('desc', penanda BERVARIAN) atau terendah
    ('asc', penanda SINGLE). Untuk laporan akhir run."""
    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute(
            "SELECT *, (jumlah_varian - jumlah_single) / "
            "(jumlah_varian + jumlah_single) AS skor, "
            "(jumlah_varian + jumlah_single) AS total "
            "FROM varian_memori HAVING total >= %s ORDER BY skor "
            + ("DESC" if urutan == "desc" else "ASC") + f" LIMIT {int(limit)}",
            (minimal_data,)
        )
        return cur.fetchall()
    finally:
        conn.close()


# ══════════════════════════════════════════════════════
# CAPTCHA LOG
# ══════════════════════════════════════════════════════

def catat_captcha(toko, akun, jenis, url=""):
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO captcha_log (toko, akun, jenis, url) VALUES (%s,%s,%s,%s)",
            (toko, akun, jenis, url[:1000] if url else "")
        )
        conn.commit()
    finally:
        conn.close()


# ══════════════════════════════════════════════════════
# ROLLING CLEANUP — pengganti bersihkan_data_lama.py versi file
# ══════════════════════════════════════════════════════

def hapus_data_lebih_dari(hari=7):
    """Hapus baris produk/varian/progress/captcha_log yang tanggalnya lebih
    tua dari 'hari' hari yang lalu. Return dict jumlah baris terhapus per
    tabel."""
    cutoff = (datetime.now() - timedelta(days=hari)).date()
    conn = get_conn()
    hasil = {}
    try:
        cur = conn.cursor()
        for tabel, kolom_tanggal in [("produk", "tanggal"), ("varian", "tanggal")]:
            cur.execute(f"DELETE FROM {tabel} WHERE {kolom_tanggal} < %s", (cutoff,))
            hasil[tabel] = cur.rowcount
        cur.execute("DELETE FROM captcha_log WHERE waktu < %s",
                    (datetime.now() - timedelta(days=hari),))
        hasil["captcha_log"] = cur.rowcount
        conn.commit()
        return hasil
    finally:
        conn.close()


# ══════════════════════════════════════════════════════
# PENGGUNA — nomor WA tiap orang (1-5), untuk notifikasi personal
# ══════════════════════════════════════════════════════

def simpan_pengguna(orang_ke, nama, nomor_wa):
    """Simpan/perbarui data satu orang (nomor WA-nya)."""
    conn = get_conn()
    try:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO pengguna (orang_ke, nama, nomor_wa) VALUES (%s,%s,%s)
               ON DUPLICATE KEY UPDATE nama=%s, nomor_wa=%s""",
            (orang_ke, nama, nomor_wa, nama, nomor_wa)
        )
        conn.commit()
    finally:
        conn.close()


def ambil_pengguna(orang_ke):
    """Ambil data satu orang. Return dict {orang_ke, nama, nomor_wa} atau
    None kalau belum terdaftar."""
    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM pengguna WHERE orang_ke=%s", (orang_ke,))
        return cur.fetchone()
    finally:
        conn.close()


def semua_pengguna():
    """Daftar semua orang terdaftar, urut nomor."""
    conn = get_conn()
    try:
        cur = conn.cursor(dictionary=True)
        cur.execute("SELECT * FROM pengguna ORDER BY orang_ke")
        return cur.fetchall()
    finally:
        conn.close()

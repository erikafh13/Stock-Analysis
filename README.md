# 📦 Analisis Stock & ABC

Aplikasi Streamlit untuk analisis stok dan klasifikasi ABC produk berdasarkan data penjualan dari Google Drive.

---

## 🗂️ Struktur Project

```
stock_app/
├── app.py                        ← Entry point utama
├── requirements.txt              ← Dependencies
├── .gitignore                    ← File yang diabaikan Git
├── .streamlit/
│   ├── config.toml               ← Konfigurasi tampilan Streamlit
│   └── secrets.toml.example      ← Template secrets (jangan commit secrets.toml asli!)
├── utils/
│   ├── __init__.py
│   ├── gdrive.py                 ← Koneksi & operasi Google Drive
│   └── analysis.py               ← Fungsi kalkulasi (ABC, WMA, Min/Max Stock)
└── pages/
    ├── input_data.py             ← Halaman Input Data
    ├── stock_analysis.py         ← Halaman Hasil Analisa Stock
    └── abc_analysis.py           ← Halaman Hasil Analisa ABC
```

---

## 🚀 Deploy ke Streamlit Cloud (Step-by-Step)

### Langkah 1 — Push ke GitHub

```bash
# 1. Inisialisasi repository (jika belum)
git init
git add .
git commit -m "Initial commit: Stock & ABC Analysis App"

# 2. Buat repo baru di github.com, lalu:
git remote add origin https://github.com/USERNAME/REPO_NAME.git
git branch -M main
git push -u origin main
```

> ⚠️ Pastikan file `.streamlit/secrets.toml` dan `credentials.json` **tidak ikut ter-push**.  
> Cek dengan: `git status` — keduanya harus ada di `.gitignore`.

---

### Langkah 2 — Setup Google Service Account

1. Buka [Google Cloud Console](https://console.cloud.google.com/)
2. Pilih atau buat project baru
3. Aktifkan **Google Drive API**:  
   *APIs & Services → Library → Google Drive API → Enable*
4. Buat Service Account:  
   *APIs & Services → Credentials → Create Credentials → Service Account*
5. Buat key JSON:  
   *Klik service account → Keys → Add Key → JSON → Download*
6. **Share folder Google Drive** ke email service account  
   (format: `nama@project-id.iam.gserviceaccount.com`) dengan akses **Editor**

---

### Langkah 3 — Deploy di Streamlit Cloud

1. Buka [share.streamlit.io](https://share.streamlit.io) → Login dengan GitHub
2. Klik **"New app"**
3. Isi form:
   - **Repository**: pilih repo GitHub Anda
   - **Branch**: `main`
   - **Main file path**: `app.py`
4. Klik **"Advanced settings"** → tab **"Secrets"**
5. Masukkan isi secrets (lihat format di bawah) → **Save**
6. Klik **"Deploy!"**

---

### Format Secrets di Streamlit Cloud

Salin isi file JSON service account ke kolom Secrets dengan format berikut:

```toml
[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "abc123..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\nMIIE...\n-----END RSA PRIVATE KEY-----\n"
client_email = "your-sa@your-project.iam.gserviceaccount.com"
client_id = "123456789"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/your-sa%40your-project.iam.gserviceaccount.com"
```

> 💡 **Tips**: Salin langsung dari file JSON yang di-download, cocokkan key-nya satu per satu.  
> Perhatikan `private_key` — newline harus ditulis sebagai `\n` (satu baris).

---

## 💻 Menjalankan Secara Lokal

```bash
# 1. Clone repo
git clone https://github.com/USERNAME/REPO_NAME.git
cd REPO_NAME

# 2. Buat virtual environment
python -m venv venv
source venv/bin/activate        # Mac/Linux
venv\Scripts\activate           # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Setup secrets lokal
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Edit secrets.toml dengan kredensial asli Anda

# 5. Jalankan
streamlit run app.py
```

---

## 🔒 Keamanan

| File | Status | Keterangan |
|------|--------|------------|
| `.streamlit/secrets.toml` | ❌ Jangan commit | Berisi private key |
| `credentials.json` | ❌ Jangan commit | Berisi private key |
| `.streamlit/secrets.toml.example` | ✅ Aman di-commit | Template kosong |
| `.gitignore` | ✅ Aman di-commit | Melindungi file sensitif |

---

## ✅ Checklist Sebelum Deploy

- [ ] `requirements.txt` sudah lengkap
- [ ] `.gitignore` sudah include `secrets.toml` dan `credentials.json`
- [ ] Folder Google Drive sudah di-share ke email service account
- [ ] Google Drive API sudah diaktifkan di Google Cloud Console
- [ ] Secrets sudah diisi di Streamlit Cloud
- [ ] `app.py` ada di root folder repo

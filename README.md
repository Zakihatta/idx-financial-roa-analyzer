# IDX Financial Statement & ROA Analyzer
### Sistem Cerdas Ekstraksi e-XBRL BEI, Analisis Auditor & Kantor Akuntan Publik (KAP), serta Kalkulator ROA

Aplikasi web analitik keuangan tingkat lanjut berbasis **Python (FastAPI)** dan **Interactive Financial Intelligence Dashboard** untuk membaca, mengekstraksi, dan menganalisis laporan keuangan terstruktur emiten Bursa Efek Indonesia (BEI / IDX) dari file Excel format e-XBRL (`.xlsx` dan `.xls`).

---

## 🌟 Fitur Utama

1. **Multi-File Upload & Intelligent Parsing**:
   - Mendukung drag & drop multi-file Excel sekaligus.
   - Parsing otomatis berbasis `pandas` dan `openpyxl`.
   - Deteksi fleksibel terhadap kode sheet standar e-XBRL BEI (`1000000`, `1110000`, `1210000`, `1311000`, dll.) maupun nama sheet deskriptif (`Informasi Umum`, `Laporan Posisi Keuangan`, `Laporan Laba Rugi`).

2. **Ekstraksi Informasi Entitas & Tata Kelola Audit (KAP)**:
   - **Nama Entitas & Ticker Saham**: Ekstraksi nama lengkap perusahaan dan kode saham 4 huruf.
   - **Auditor Tahun Berjalan (KAP)**: Identifikasi nama Kantor Akuntan Publik independen.
   - **Partner Audit (Signing Partner)**: Nama akuntan publik rekan yang menandatangani opini audit.
   - **Auditor Tahun Sebelumnya**: Nama KAP periode komparasi.
   - **Deteksi Rotasi KAP (Auditor Switching)**: Komparasi otomatis antara auditor tahun berjalan dan tahun sebelumnya untuk mendeteksi apakah terjadi rotasi auditor (Berganti / Tetap).
   - **Klasifikasi KAP Tier**: Identifikasi otomatis Big Four (PwC, EY, Deloitte, KPMG) vs Non-Big Four.

3. **Ekstraksi Data Keuangan & Normalisasi Skala**:
   - **Total Aset (Total Assets)**: Pencarian nilai aset periode berjalan pada sheet Posisi Keuangan dengan filter baris presisi.
   - **Laba Bersih (Net Income)**: Pencarian laba (rugi) periode berjalan pada sheet Laba Rugi Komprehensif.
   - **Faktor Pembulatan (Rounding Multiplier)**: Deteksi otomatis skala penyajian (Satuan Penuh, Dalam Ribuan, Dalam Jutaan, Dalam Miliaran) dan konversi ke nilai nominal rupiah riil penuh.

4. **Perhitungan & Analisis Kesehatan ROA**:
   - Menghitung rasio profitabilitas:
     $$\text{ROA (\%)} = \left( \frac{\text{Laba Bersih}}{\text{Total Aset}} \right) \times 100\%$$
   - Penilaian kesehatan berdasarkan standar OJK & analisis keuangan:
     - $\ge 10.0\%$: Sangat Unggul (High Profitability)
     - $5.0\% - 9.99\%$: Sehat / Baik (Good)
     - $2.0\% - 4.99\%$: Cukup / Standar (Moderate)
     - $0.0\% - 1.99\%$: Rendah / Kurang (Low)
     - $< 0.0\%$: Rugi / Negatif (Loss)

5. **Visualisasi Interaktif & Verifikasi (Audit Trail)**:
   - Kartu KPI Ringkasan: Total emiten, rata-rata ROA industri, akumulasi aset, pangsa Big Four, jumlah rotasi KAP.
   - Chart.js: Grafik perbandingan ROA antar emiten & grafik pangsa pasar KAP.
   - Modal Jejak Audit: Verifikasi asal sheet, nomor baris, label baris asli, dan penguraian rumus matematika langkah demi langkah.
   - Filter & Pencarian real-time (Cari berdasarkan emiten, KAP, partner, filter Big Four, Rotasi KAP, status ROA).

6. **Ekspor & Data Demo IDX**:
   - Ekspor ke Excel (`.xlsx`) dengan formula aktif dan styling sel profesional.
   - Ekspor ke CSV.
   - Fitur "Muat Demo IDX" (6 emiten BEI: BBCA, BBRI, TLKM, ASII, ICBP, GOTO).
   - Unduh sampel file Excel IDX asli langsung dari dashboard.

---

## 🚀 Cara Menjalankan Aplikasi

Pastikan Python telah terpasang di komputer Anda.

### 1. Jalankan Server
Buka terminal / PowerShell di folder proyek ini dan jalankan:
```bash
python run.py
```
atau menggunakan Uvicorn langsung:
```bash
python -m uvicorn main:app --reload --port 8000
```

### 2. Akses Aplikasi
Buka peramban (browser) Anda dan akses:
```
http://127.0.0.1:8000
```

---

## 📁 Struktur Direktori

```
app/
├── core/
│   ├── extractor.py        # Mesin ekstraksi regex & parsing data keuangan e-XBRL
│   ├── sample_data.py      # Generator file sampel Excel IDX realistis
│   └── exporter.py         # Modul ekspor ke file Excel (.xlsx berformula) & CSV
├── static/
│   ├── css/
│   │   └── style.css       # Desain antarmuka Dark Luxury Financial Design System
│   ├── js/
│   │   └── app.js          # Logika frontend, reactive state, Chart.js, modal
│   └── index.html          # Halaman dashboard web responsif
├── samples/                # Direktori file contoh laporan keuangan BEI
├── main.py                 # Backend API FastAPI
├── run.py                  # Script launcher praktis dengan auto-open browser
├── test_extractor.py       # Script pengujian akurasi parsing
└── requirements.txt        # Daftar pustaka Python yang dibutuhkan
```

---

## 📡 Dokumentasi Endpoint API

| Method | Endpoint | Deskripsi |
|---|---|---|
| `GET` | `/` | Antarmuka web dashboard interaktif |
| `POST` | `/api/upload` | Mengunggah satu atau beberapa file Excel (`.xlsx`, `.xls`) |
| `GET` | `/api/demo-data` | Mengembalikan 6 data contoh emiten BEI hasil ekstraksi real-time |
| `POST` | `/api/export/excel` | Mengunduh hasil analisis ke file Excel (`.xlsx`) berformat rapi |
| `POST` | `/api/export/csv` | Mengunduh hasil analisis ke format CSV |
| `GET` | `/api/download-sample/{ticker}` | Mengunduh file Excel contoh e-XBRL IDX asli (BBCA, BBRI, dll.) |

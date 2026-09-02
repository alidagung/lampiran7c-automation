# Automation Lampiran 7C ASPI

## Apa Itu Project Ini?

Project ini adalah sebuah **script otomatis** yang membantu Anda membuat dokumen **Lampiran 7C ASPI** (format Word/.docx) dari data **UAT Script** (format Excel/.xlsx).

Dengan script ini, Anda tidak perlu lagi copy-paste manual data URL, Headers, Request Body, dan Response dari file UAT Script ke dokumen Lampiran 7C. Cukup jalankan satu perintah, dan dokumen akan dibuat secara otomatis!

### Fitur Utama:
- Membaca file UAT Script Excel secara otomatis
- Parsing data log (URL, Headers, Request Body, Response) dari kolom Remarks
- Mendukung 2 format Remarks: format "Mitra HIT" dan "BSS YANG HIT"
- Membuat dokumen Word Lampiran 7C dengan format tabel yang benar
- Mapping otomatis skenario UAT ke section Lampiran 7C
- **Aplikasi web**: cukup unggah file, klik proses, unduh hasilnya (tanpa perlu paham teknis)

---

## 🌐 Cara Cepat: Aplikasi Web Lokal (Direkomendasikan)

Aplikasi ini berjalan **100% di komputer Anda** (lokal). File yang Anda unggah
**tidak dikirim ke internet** — aman untuk data sensitif.

### 🖱️ Cara paling mudah (Windows): cukup double-click

1. **Install Python** dulu jika belum (lihat [Langkah 1](#langkah-1-install-python)).
   Pastikan mencentang **"Add Python to PATH"** saat instalasi.

2. **Setup sekali saja:** double-click file **`setup-pertama-kali.bat`**.
   Ini akan menginstall komponen yang dibutuhkan (tunggu sampai selesai).

3. **Menjalankan aplikasi:** double-click file **`jalankan-aplikasi.bat`**.
   Browser akan terbuka otomatis di `http://localhost:8501`.

4. Di halaman web:
   - **Unggah** file `UAT Script.xlsx` Anda (klik area unggah atau seret file).
   - Klik tombol **🚀 Proses & Buat Lampiran 7C**.
   - Klik **⬇️ Unduh Lampiran 7C (.docx)**.

5. Untuk **menghentikan** aplikasi: tutup jendela hitam (Command Prompt) yang muncul,
   atau tekan `Ctrl + C` di dalamnya.

> 💡 Untuk pemakaian berikutnya, Anda **cukup double-click `jalankan-aplikasi.bat`** saja
> (tidak perlu setup lagi).

### ⌨️ Cara manual (via terminal, semua OS)

```
pip install -r requirements.txt      (cukup sekali)
streamlit run app.py
```
Lalu buka `http://localhost:8501` di browser.

---

## 🖥️ Cara Alternatif: Baris Perintah (CLI)

Jika Anda lebih suka cara lama (menaruh file di folder `input/` dan menjalankan
script), ikuti panduan lengkap di bawah ini mulai dari [Langkah 1](#langkah-1-install-python).

---

## Prasyarat

Sebelum menggunakan project ini, pastikan komputer Anda sudah terinstall:
- **Python** versi 3.8 atau lebih baru (direkomendasikan Python 3.11)
- **pip** (package manager Python, biasanya sudah terinstall bersama Python)

---

## Langkah 1: Install Python

### Untuk Windows:

1. Buka browser, kunjungi: https://www.python.org/downloads/
2. Klik tombol **"Download Python 3.x.x"** (versi terbaru)
3. Jalankan file installer yang sudah di-download
4. **PENTING:** Centang checkbox **"Add Python to PATH"** di bagian bawah installer
5. Klik **"Install Now"**
6. Tunggu sampai selesai, lalu klik **"Close"**

### Verifikasi Instalasi Python:

Buka **Command Prompt** (tekan `Win + R`, ketik `cmd`, tekan Enter), lalu ketik:

```
python --version
```

Jika berhasil, akan muncul sesuatu seperti: `Python 3.11.5`

Kemudian cek pip:

```
pip --version
```

Jika muncul versi pip, berarti instalasi berhasil.

---

## Langkah 2: Download/Clone Project

### Cara 1: Download sebagai ZIP
1. Download file project ini
2. Extract/unzip ke folder yang Anda inginkan (misalnya `D:\lampiran7c-automation\`)

### Cara 2: Menggunakan Git (jika sudah install Git)
```
git clone <repository-url>
cd lampiran7c-automation
```

---

## Langkah 3: Setup Virtual Environment (Opsional tapi Direkomendasikan)

Virtual environment berguna untuk mengisolasi package Python project ini agar tidak bentrok dengan project lain.

### Buat Virtual Environment:

Buka Command Prompt, masuk ke folder project:
```
cd D:\lampiran7c-automation
```

Buat virtual environment:
```
python -m venv venv
```

### Aktifkan Virtual Environment:

**Windows:**
```
venv\Scripts\activate
```

**Linux/Mac:**
```
source venv/bin/activate
```

Jika berhasil, akan muncul `(venv)` di awal baris command prompt Anda.

---

## Langkah 4: Install Dependencies (Library yang Dibutuhkan)

Pastikan Anda sudah berada di folder project dan virtual environment sudah aktif (jika menggunakan venv). Kemudian jalankan:

```
pip install -r requirements.txt
```

Perintah ini akan menginstall:
- **openpyxl** - untuk membaca file Excel (.xlsx)
- **python-docx** - untuk membuat file Word (.docx)
- **pytest** - untuk menjalankan unit test

Tunggu sampai proses download dan install selesai.

---

## Langkah 5: Siapkan File Input

### File yang Dibutuhkan:
Anda perlu menyiapkan file **UAT Script** dalam format Excel (.xlsx).

### Cara Menyiapkan:
1. Buka folder `input/` di dalam folder project
2. Copy file UAT Script Anda ke folder tersebut
3. **Rename** file menjadi: `UAT Script.xlsx` (perhatikan spasi dan huruf besar/kecil)

### Struktur folder setelah menyiapkan input:
```
lampiran7c-automation/
├── input/
│   └── UAT Script.xlsx    <-- file Anda taruh di sini
├── output/                <-- hasil akan muncul di sini
├── main.py
├── requirements.txt
└── README.md
```

### Syarat File UAT Script:
- Harus memiliki sheet bernama **"UAT Script"**
- Kolom-kolom yang harus ada (berurutan):
  1. Kategori Tes
  2. Nama Modul
  3. Nomor Skenario
  4. Nomor Kasus Tes
  5. Langkah Tes
  6. Hasil yang diharapkan
  7. Hasil Aktual
  8. Remarks
  9. Tanggal Pelaksanaan
  10. Jenis Script
  11. Pelaksana

---

## Langkah 6: Jalankan Script

Buka Command Prompt, masuk ke folder project, dan jalankan:

```
cd D:\lampiran7c-automation
python main.py
```

### Jika berhasil, Anda akan melihat output seperti ini:
```
============================================================
  AUTOMATION LAMPIRAN 7C ASPI
  Konversi UAT Script (Excel) -> Lampiran 7C (Word)
============================================================

  [1/3] Membaca file UAT Script: input\UAT Script.xlsx
        Ditemukan 9 section:
        - Balance Services: 11 baris data
        - Intrabank Transfer: 12 baris data
        ...

  [2/3] Memproses mapping UAT Script -> Lampiran 7C...
        - API Balance Inquiry: 11/11 baris terisi
        - Intrabank Transfer: 12/12 baris terisi
        ...

  [3/3] Membuat dokumen Lampiran 7C: output\Lampiran 7C - Hasil UAT.docx

  ========================================================
  SELESAI! Total XX baris data berhasil dipindahkan.
  File output: D:\lampiran7c-automation\output\Lampiran 7C - Hasil UAT.docx
  ========================================================
```

---

## Langkah 7: Cek Hasil Output

Setelah script selesai, buka folder `output/`. Di sana akan ada file:
- **Lampiran 7C - Hasil UAT.docx**

Buka file tersebut dengan Microsoft Word untuk melihat hasilnya.

---

## Troubleshooting (Solusi Masalah Umum)

### 1. Error: "python is not recognized as an internal or external command"
**Penyebab:** Python belum terinstall atau belum ditambahkan ke PATH.
**Solusi:** Install ulang Python dan pastikan centang "Add Python to PATH".

### 2. Error: "No module named 'openpyxl'" atau "No module named 'docx'"
**Penyebab:** Dependencies belum diinstall.
**Solusi:** Jalankan `pip install -r requirements.txt`

### 3. Error: "File UAT Script tidak ditemukan!"
**Penyebab:** File UAT Script belum di-copy ke folder input/ atau nama file salah.
**Solusi:**
- Pastikan file ada di folder `input/`
- Pastikan nama file persis: `UAT Script.xlsx` (dengan spasi, huruf besar S)

### 4. Error: "Sheet 'UAT Script' tidak ditemukan"
**Penyebab:** Nama sheet di file Excel berbeda.
**Solusi:** Buka file Excel, cek nama sheet-nya. Rename menjadi "UAT Script" jika perlu.

### 5. Hasil dokumen Word kosong (tidak ada data)
**Penyebab:** Format UAT Script mungkin berbeda dari yang diharapkan.
**Solusi:**
- Pastikan ada section header (Balance Services, Intrabank Transfer, dll.)
- Pastikan ada data di kolom Remarks
- Pastikan format Nomor Kasus Tes sesuai (misal: 1.1, 2.3, 7.15)

### 6. Error saat menjalankan di virtual environment
**Penyebab:** Virtual environment belum diaktifkan.
**Solusi:** Jalankan `venv\Scripts\activate` (Windows) sebelum menjalankan script.

---

## Struktur Project

```
lampiran7c-automation/
├── input/                  # Folder untuk file input (UAT Script.xlsx) - dipakai mode CLI
├── output/                 # Folder untuk file output (hasil Lampiran 7C) - dipakai mode CLI
├── tests/                  # Folder unit tests
│   ├── test_parser.py      # Test untuk parsing Remarks
│   └── test_mapping.py     # Test untuk mapping skenario
├── app.py                  # Aplikasi web (Streamlit) - upload & unduh via browser
├── main.py                 # Logika inti + script CLI
├── setup-pertama-kali.bat  # (Windows) install komponen - jalankan sekali di awal
├── jalankan-aplikasi.bat   # (Windows) jalankan aplikasi web - double-click
├── requirements.txt        # Daftar library yang dibutuhkan
└── README.md               # File panduan ini
```

---

## Menjalankan Tests (Untuk Developer)

Jika Anda ingin menjalankan unit test untuk memastikan script berjalan dengan benar:

```
python -m pytest tests/ -v
```

---

## Catatan Penting

- Script ini membuat dokumen Lampiran 7C **dari scratch** (bukan dari template)
- Untuk section yang sharing (Interbank Transfer dan Virtual Account), data dari skenario utama diprioritas. Skenario tambahan hanya mengisi baris yang masih kosong.
- Jika kolom "Hasil Aktual" bertuliskan "Tidak dites", maka kolom Request dan Response di Lampiran 7C akan dikosongkan.
- File UAT Script HARUS dalam format .xlsx (Excel 2007 ke atas)

---

## Kontak & Bantuan

Jika Anda mengalami masalah atau memiliki pertanyaan, silakan buat issue di repository ini atau hubungi tim development.

# Cek Bespoke — BAST Checker & CCTV Checker (versi web, UI v2)

Versi web dari `bast_checker_pdf.py` dan `cctv_checker.py`.
Logika & threshold pengecekan **sama persis** dengan script asli. Yang berubah: tidak perlu edit path `D:\...`
atau run di VS Code — cukup buka browser, upload file, klik tombol, download Excel.

## Isi folder
| File | Fungsi |
|---|---|
| `app.py` | Tampilan web (Streamlit) dengan sidebar, stepper, dan halaman hasil |
| `bast_core.py` | Logika BAST Checker (dari script asli) |
| `cctv_core.py` | Logika CCTV Checker (dari script asli) |
| `requirements.txt` | Daftar library |
| `run.bat` | Klik 2x untuk menjalankan di Windows |
| `.streamlit/config.toml` | Batas upload 1 GB, tema |

## Menjalankan di komputer sendiri (paling mudah)
**Windows:** klik 2x `run.bat`. Pertama kali akan install library (beberapa menit), lalu browser terbuka otomatis.

**Manual / lewat VS Code terminal:**
```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```
Buka http://localhost:8501

## Dipakai bersama satu tim (tanpa deploy ke internet)
Jalankan di satu PC/server kantor yang menyala:
```
streamlit run app.py --server.address 0.0.0.0
```
Rekan satu jaringan buka `http://<IP-PC-itu>:8501`. Gratis, dan dokumen tidak keluar dari jaringan kantor.

## Deploy ke cloud (opsional)
- **Streamlit Community Cloud** (gratis): push folder ini ke GitHub, lalu deploy di share.streamlit.io. RAM ±1 GB, jadi batch sangat besar bisa gagal.
- **Hugging Face Spaces** (gratis, RAM lebih besar, bisa dibuat private): pilih SDK Streamlit.
- Vercel **tidak disarankan** untuk app ini (lihat penjelasan di chat).
- Kalau pakai fitur OCR label kamera (opsional), tambah file `packages.txt` berisi `tesseract-ocr` dan aktifkan `pytesseract` di requirements.

> ⚠️ Dokumen BAST/CCTV adalah data internal. Sebelum deploy ke layanan cloud publik, pastikan dulu boleh menurut kebijakan perusahaan.

## Cara pakai
1. Pilih tab **BAST Checker** atau **CCTV Checker**.
2. Upload file (banyak sekaligus) **atau** 1 file **ZIP** berisi folder-folder (nama subfolder ikut muncul di laporan).
3. Klik **Mulai Cek**, tunggu progress selesai.
4. Download laporan Excel (BAST juga bisa download ZIP foto duplikat).

## Mengubah sensitivitas
Semua threshold ada di bagian atas `bast_core.py` (mis. `HASH_DISTANCE_THRESHOLD`, `MAX_DOC_OCCURRENCES`)
dan `cctv_core.py` (mis. `FILTERED_BLUR_THRESHOLD`, `FISHEYE_*`). Kolom/sheet "Detail Skor (Tuning)" di Excel CCTV membantu menyetelnya.

## Catatan tampilan (UI v2)
Tampilan memakai CSS khusus di bagian atas `app.py`. Kalau suatu saat Streamlit di-update dan ada bagian yang tampil aneh
(misal teks dropzone kembali berbahasa Inggris), itu hanya kosmetik; fungsi tetap jalan. Warna utama bisa diganti di `CSS` dan `.streamlit/config.toml`.

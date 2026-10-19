# Deploy ke Streamlit Community Cloud (gratis)

1. Buat akun GitHub (sebaiknya akun milik tim/atasan, bukan akun pribadi magang, supaya kepemilikan aman setelah magang selesai).
2. Buat repository baru (boleh Private). Upload SEMUA isi folder ini lewat web GitHub: Add file > Upload files.
   JANGAN upload folder `.venv`.
3. Buka https://share.streamlit.io, login dengan GitHub, klik "Create app".
4. Isi: Repository = repo tadi, Branch = main, Main file path = app.py.
5. Klik "Advanced settings":
   - Python version: 3.12 (atau 3.11)
   - Secrets: tulis satu baris:  APP_PASSWORD = "kata-sandi-pilihan-kamu"
6. Klik Deploy. Tunggu build beberapa menit. Link `https://<nama>.streamlit.app` siap dibagikan.

Kalau APP_PASSWORD tidak diisi, aplikasi bisa dibuka siapa saja yang punya link.
Untuk jalan lokal (run.bat) tidak ada kata sandi.

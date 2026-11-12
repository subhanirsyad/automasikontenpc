# PC Content Append Only

App lokal untuk menambah baris dari export TAP, SC, dan AFF ke salinan PC Content. Kalkulasi TAP dan SC memakai modul yang sama dengan paket sebelumnya.

## Pakai folder Content Automate yang sudah ada

1. Hentikan Streamlit dengan Ctrl+C. Ekstrak ZIP dan salin **app.py**, **pipeline.py**, dan **report_formulas.json** dari paket ke folder Content Automate. Timpa dua file Python lama. Tiga file ini harus berada dalam folder yang sama.
2. Gunakan file `credentials.json` dan konfigurasi lokal yang sudah ada. Untuk laptop kantor tanpa login tambahan, `.streamlit/config.toml` berisi `[server]` dan `address = "127.0.0.1"`. Isi `.streamlit/secrets.toml` untuk upload:

   ```toml
   [pc_content]
   enable_write = true
   ```

3. Jalankan dari terminal di folder Content Automate.

   ```powershell
   .\.venv\Scripts\python.exe -m streamlit run app.py
   ```

4. Buka http://127.0.0.1:8501. Masukkan link salinan tujuan, lalu klik **Cek periode di template** untuk melihat tanggal awal dan akhir TAP SV, TAP LS, SC SV, dan SC LS. Pilih periode, unggah file dan klik **Proses dan preview**. Cek tanggal yang sudah ada dalam periode upload, baris baru, dan baris yang sudah ada. Konfirmasi nama salinan, lalu klik **Kirim hasil preview**. Service account perlu akses Viewer ke List Produk dan Editor ke salinan tujuan.

Jika mulai di folder baru, salin `credentials.json` ke samping `app.py`, buat `secrets.toml` dari contoh dan pasang dependensi. Buat venv dari Python 3.12 dengan `py -3.12 -m venv .venv`, lalu jalankan `.\.venv\Scripts\python.exe -m pip install -r requirements.txt`.

## Template dan perilaku append

- Pertahankan nama tab RAW SV dan RAW LS serta header di baris 1. Area data boleh kosong mulai baris 2. Untuk RAW SV, header **Type Content** tetap tersedia. Nilai atau rumus di kolom itu dikelola oleh pengguna.
- Baris baru memperoleh rumus Date, Type Video pada SV, link video untuk TAP SV, dan Duration untuk SC LS ketika waktu mulai dan selesai tersedia. Duration TAP LS berasal dari export. Type Video mengikuti kategori di Type Content. Jika Type Content kosong, kategori jatuh ke Mandatory sesuai logika template.
- Baris RAW yang sudah ada dilewati berdasarkan Video ID + Brand untuk SV dan Livestream room ID + Brand untuk LS. ID ditulis sebagai teks. TAP diprioritaskan ketika key yang sama muncul dari TAP dan SC dalam batch. Metrik yang tidak ada di sumber dibiarkan kosong.
- Tab rekap bersifat opsional. Pada salinan dengan struktur referensi November, pertahankan tab, header, dan tanggal minggu di CONFIG. Saat upload pertama dengan baris baru, app mengisi sel rumus rekap yang kosong di CALC SV, CALC LS, RINGKAS SV, RINGKAS LS, COUNTS, dan CARA PAKAI. Rumus rekap yang sudah ada dipertahankan.
- Versi ini membuat rencana append sendiri. Mulai dari preview baru untuk proses yang tertunda dari versi sebelumnya. Baris yang sudah berhasil masuk akan dilewati berdasarkan key. Untuk proses yang tertunda dari versi append-only ini, gunakan **Lanjutkan batch**.

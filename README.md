# Training Management System

Kode sumber untuk aplikasi Training Management System berbasis **Google Apps Script + Google Sheets + Google Drive**.

## Penting: GitHub bukan server backend aplikasi ini

Aplikasi ini menggunakan `HtmlService` di `Code.gs` dan memanggil fungsi server melalui `google.script.run` di `Index.html`. Oleh karena itu:

- **GitHub Repository** dipakai untuk menyimpan kode, riwayat perubahan, dan kolaborasi.
- **Google Apps Script Web App** tetap menjadi tempat menjalankan aplikasi.
- **GitHub Pages tidak dapat menjalankan backend Google Apps Script ini secara langsung.** Supaya bisa berjalan di GitHub Pages, backend, autentikasi, penyimpanan data, dan pemanggilan `google.script.run` harus dimigrasikan ke layanan lain.

## Isi repositori

- `Code.gs` — backend Google Apps Script.
- `Index.html` — tampilan/front-end aplikasi.
- `appsscript.json` — manifest Apps Script minimal dengan runtime V8 dan zona waktu Asia/Makassar.
- `.gitignore` — mencegah file konfigurasi autentikasi lokal masuk ke Git.

> Saat menghubungkan ke proyek Apps Script yang sudah ada, simpan salinan cadangan terlebih dahulu dan pertahankan pengaturan `appsscript.json` asli jika memuat konfigurasi lain. Jangan langsung melakukan `clasp push` ke proyek produksi sebelum memastikan semua file proyek sudah tercakup.

## Upload ke GitHub melalui browser

1. Masuk ke <https://github.com/> dan login.
2. Klik tombol **+** di kanan atas, lalu pilih **New repository**.
3. Isi nama repositori, misalnya `training-management-system`.
4. Pilih **Private**. Kode memuat logika autentikasi dan operasi data; repositori publik tidak disarankan tanpa audit keamanan.
5. Buat repositori. Jika akan mengunggah file dari paket ini, lebih mudah bila repositori dibuat tanpa menambahkan README, `.gitignore`, atau lisensi otomatis agar tidak terjadi konflik.
6. Ekstrak ZIP paket ini di komputer.
7. Pada halaman repositori, klik **Add file → Upload files**.
8. Seret file dan folder yang ingin diunggah ke area upload. File tersembunyi seperti `.gitignore` kadang tidak ikut terpilih saat drag-and-drop; bila tidak muncul, buat file `.gitignore` di repositori secara manual.
9. Isi pesan commit, misalnya `Upload source awal Training Management System`.
10. Klik **Commit changes**.

Alternatif: jika sudah memasang Git, buat folder proyek lokal, salin file-file di atas, lalu gunakan `git init`, `git add .`, `git commit`, dan `git push` setelah menambahkan URL repositori GitHub sebagai remote.

## Menjalankan aplikasi

### Jika aplikasi sudah berjalan di Google Apps Script

1. Buka proyek Apps Script yang sekarang digunakan aplikasi.
2. Jangan mengganti URL Web App yang sedang dipakai pengguna hanya karena kode sudah tersimpan di GitHub.
3. Jika ada perubahan kode yang akan diterapkan, lakukan melalui salinan proyek/backup dan uji dahulu sebelum memperbarui deployment.

### Jika perlu menerbitkan Web App baru

1. Buka spreadsheet utama yang menjadi sumber data.
2. Pilih **Extensions → Apps Script** untuk membuka proyek yang terikat ke spreadsheet tersebut.
3. Pastikan `Code.gs` dan `Index.html` berada di proyek itu. Jangan membuat proyek standalone baru tanpa menyesuaikan koneksi spreadsheet, karena kode saat ini menggunakan `SpreadsheetApp.getActiveSpreadsheet()`.
4. Simpan, lalu pilih **Deploy → New deployment**.
5. Pilih jenis **Web app** dan tentukan **Execute as** serta siapa yang boleh mengakses berdasarkan kebijakan organisasi Anda. Jangan membuka akses publik kecuali memang diperlukan dan telah ditinjau.
6. Klik **Deploy**, selesaikan otorisasi yang diminta, lalu buka URL Web App untuk pengujian.

## Sinkronisasi lokal ↔ Google Apps Script dengan clasp (opsional)

`clasp` adalah alat command-line untuk mengunduh/mengunggah file proyek Apps Script. Dokumentasi resmi: <https://developers.google.com/apps-script/guides/clasp>.

### Persiapan

1. Pasang Node.js versi 22 atau versi yang didukung `clasp` saat ini.
2. Buka terminal/Command Prompt dan jalankan:

   ```bash
   npm install -g @google/clasp
   ```

3. Aktifkan **Google Apps Script API** di <https://script.google.com/home/usersettings>.
4. Login:

   ```bash
   clasp login
   ```

### Hubungkan repositori ke proyek Apps Script yang sudah ada

1. Salin repositori ke komputer menggunakan GitHub Desktop atau perintah `git clone`.
2. Buka proyek Apps Script yang benar dan salin **Script ID** dari URL editor. Pastikan ini proyek yang terikat pada spreadsheet utama aplikasi.
3. Dari folder lokal repositori, jalankan perintah berikut. Ganti `<SCRIPT_ID>` dengan ID yang benar:

   ```bash
   clasp clone-script <SCRIPT_ID> --rootDir .
   ```

   Perintah ini menarik file proyek Apps Script ke folder lokal. Periksa perubahan Git setelahnya dan pastikan `Code.gs`, `Index.html`, serta file proyek lain yang sebelumnya ada tidak hilang.
4. Setelah melakukan backup dan memastikan folder lokal berisi seluruh file proyek yang benar, kirim perubahan ke Apps Script dengan:

   ```bash
   clasp push
   ```

5. Simpan perubahan ke GitHub:

   ```bash
   git add .
   git commit -m "Update aplikasi Training Management System"
   git push
   ```

6. Bila ingin perubahan langsung digunakan oleh pengguna, buka Apps Script dan perbarui deployment Web App yang sesuai. `clasp push` mengunggah kode sumber, tetapi tidak otomatis menjamin URL deployment memakai versi produksi terbaru.

> **Peringatan:** `clasp push` mengganti konten sumber proyek online berdasarkan file lokal. Jangan menjalankannya dari folder yang hanya berisi sebagian file proyek. Lakukan backup, clone proyek online terlebih dahulu, dan periksa daftar file sebelum push.

## Pemeriksaan keamanan sebelum digunakan

- Gunakan repositori GitHub privat kecuali ada alasan kuat untuk membuka kode.
- Kode menggunakan data login dari spreadsheet; pastikan file spreadsheet memiliki akses terbatas dan jangan memasukkan password, token, atau data peserta ke repositori.
- Fungsi `_setDriveAnyoneEditor_` pada `Code.gs` mengatur resource Google Drive ke akses `Anyone` dengan izin `Editor`. Tinjau dan ubah perilaku ini bila file tidak seharusnya dapat diakses atau diedit oleh siapa pun yang memiliki tautan.
- Tinjau otorisasi deployment Web App, hak akses spreadsheet, folder Drive, file tugas, dokumen laporan, dan file hasil ujian.
- Uji dengan data dummy dan akun uji sebelum menerapkan perubahan ke aplikasi aktif.

## Referensi

- clasp: <https://developers.google.com/apps-script/guides/clasp>
- Dokumentasi GitHub: <https://docs.github.com/en/repositories/working-with-files/managing-files/adding-a-file-to-a-repository>

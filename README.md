# Pencatat Temuan Patroli Tower

Aplikasi Streamlit sederhana untuk menempel teks laporan patroli (format grup WA)
dan otomatis mengubahnya menjadi baris-baris data rapi di spreadsheet Excel.

## Cara Menjalankan

1. Pastikan Python sudah terpasang (3.9+).
2. Install dependensi:
   ```
   pip install -r requirements.txt
   ```
3. Jalankan aplikasi:
   ```
   streamlit run app.py
   ```
4. Browser akan otomatis terbuka ke `http://localhost:8501`.

## Cara Pakai

1. Buka tab **Input Laporan Baru**.
2. Tempel teks laporan, contohnya:
   ```
   Temuan patroli tower Clifford sbb :
   1). Termonitor Tembok Retak Di PTD 2 Lt 11
   2.) Termonitor Tembok Terkupas Di Lt 01 PTD 1
   ```
3. Pilih tanggal patroli, lalu klik **Parse Laporan**.
4. Aplikasi otomatis memecah teks menjadi baris per temuan (Tower, PTD, Lantai,
   Temuan, Kategori, PIC, Status). Kamu bisa edit langsung di tabel sebelum
   disimpan (misalnya ubah Status jadi "Selesai", tambah/hapus baris, dsb).
5. Klik **Simpan ke Spreadsheet** — data akan ditambahkan (append) ke file
   `data_temuan.xlsx` di folder yang sama.
6. Buka tab **Riwayat Temuan** untuk melihat, memfilter (per tower/status/kata
   kunci), dan mengunduh seluruh data yang sudah tersimpan.

## Format Teks yang Didukung

Parser mengenali heading:
```
Temuan patroli tower <NamaTower> sbb :
```
(tidak case-sensitive, spasi/tanda baca fleksibel), lalu baris bernomor
seperti `1).`, `1.)`, `2.)`, dst sebagai temuan. Baris `@Nama Orang` di akhir
blok otomatis dianggap sebagai PIC (penanggung jawab).

## Catatan

- Data tersimpan lokal di `data_temuan.xlsx` — kalau mau deploy ke server/cloud
  agar bisa diakses tim, sebaiknya file ini dipindah ke penyimpanan bersama
  (misal Google Sheets/database) — beri tahu saya kalau butuh versi itu.
- Kolom **Kategori** ditebak otomatis dari kata kunci (retak, gompal, jamur,
  panel, dst) tapi tetap bisa diedit manual sebelum disimpan.

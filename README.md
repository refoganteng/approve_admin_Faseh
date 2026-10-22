# approve_admin_Faseh

Script otomasi menggunakan **Playwright** untuk melakukan proses **Approve Admin** secara otomatis pada survei FASIH Web (BPS).

---

## 📋 Alur Kerja Otomasi

1. **Inisialisasi Browser**: Membuka Google Chrome dengan mode remote debugging (port 9222) dan profil pengguna tersimpan (`~/chrome_fasih_profile`), sehingga session login tetap tersimpan.
2. **Login Manual**: Jika belum login, script memberi jeda kepada pengguna untuk login manual di browser Chrome. Setelah login berhasil, pengguna cukup menekan `ENTER` di terminal.
3. **Navigasi & Eksekusi Tiap Assignment**:
   - Membuka URL `link_fasih` assignment ([01_bukalink.html](file:///Users/user/Documents/Python/approve_admin_faseh/01_bukalink.html)).
   - Menunggu indikator loading FormGear & FASIH overlay selesai.
   - Mengklik **Tombol Approve** (icon centang hijau di floating action bar, sesuai [02_btnapprove.html](file:///Users/user/Documents/Python/approve_admin_faseh/02_btnapprove.html)).
   - Menunggu popup modal **"Approve Assignment?"** muncul ([03_modalkonfirmasi.html](file:///Users/user/Documents/Python/approve_admin_faseh/03_modalkonfirmasi.html)).
   - Mengklik tombol **"Konfirmasi"** pada modal dialog.
   - Menunggu proses penyimpanan server selesai.
4. **Pembaruan Data & Riwayat**:
   - Status di Excel diperbarui menjadi `BERHASIL` (atau `SUDAH_APPROVED`) beserta timestamp dan `bot_runner` (APPROVE_ADMIN).
   - Dituliskan ke log ringkas (`log_simple_approve.txt`) dan log progres real-time (`log_terakhir_approve.txt`).
   - Mendukung fitur **Resume otomatis** (jika proses dihentikan di tengah jalan, baris yang sudah sukses tidak akan diulang).

---

## 🚀 Persyaratan Sistem

- Python 3.10+
- Google Chrome terinstal di komputer
- Dependensi:
  ```bash
  pip install -r requirements.txt
  playwright install chromium
  ```

---

## 💻 Cara Menjalankan

### 1. Menjalankan Script (Pilihan File Excel Interaktif)
Jalankan script langsung:
```bash
python auto_approve_admin.py
```
Di awal terminal, script akan otomatis mendeteksi semua file Excel (`.xlsx` / `.xls`) di folder ini dan menampilkan menu pilihan interaktif:
```text
📂 PILIH FILE EXCEL DATA ASSIGNMENT:
=================================================================
  [1] data_edited.xlsx (7.3 KB)
  [2] data_survei_batch2.xlsx (15.2 KB)
  [3] ✍️  Ketik nama / path file lain secara manual...
-----------------------------------------------------------------
Pilih nomor file Excel [1-3] (Tekan ENTER untuk [1]): 
```
- Anda cukup menekan **ENTER** untuk memilih file nomor `[1]`, atau ketik nomor yang diinginkan.
- Nama file bisa berbeda-beda tanpa batasan.

### 2. Menjalankan Langsung dengan File Tertentu (Tanpa Menu Interaktif)
Jika ingin melewati menu pilihan dan langsung mengeksekusi file tertentu:
```bash
python auto_approve_admin.py --excel nama_file_anda.xlsx
```

### 3. Mode Simulasi / Uji Coba (`--dry-run`)
Mode ini akan membuka halaman dan mengklik tombol Approve, tetapi **tidak akan mengklik Konfirmasi akhir** (akan membatalkan modal), cocok untuk memastikan selector berjalan dengan baik:
```bash
python auto_approve_admin.py --dry-run
```

### 4. Membatasi Jumlah Assignment (`--limit`)
Contoh hanya memproses 5 assignment pertama:
```bash
python auto_approve_admin.py --limit 5
```

### 5. Mulai dari Baris Tertentu (`--start-row`)
```bash
python auto_approve_admin.py --start-row 10
```

### 6. Mengatur Jeda Antar Assignment (`--delay`)
Default jeda adalah **0.7 detik**. Anda dapat mengubahnya jika ingin jeda lebih lama, contoh jeda 2 detik antar assignment:
```bash
python auto_approve_admin.py --delay 2.0
```
   
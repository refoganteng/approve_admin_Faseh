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

### 1. Menjalankan Default (menggunakan `data_edited.xlsx`)
```bash
python auto_approve_admin.py
```

### 2. Menjalankan dengan Data dari Folder `input_fasih_ntb`
```bash
python auto_approve_admin.py --excel input_fasih_ntb/data_edited.xlsx
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

---

## 📂 Struktur File Terkait

- [auto_approve_admin.py](file:///Users/user/Documents/Python/approve_admin_faseh/auto_approve_admin.py) : Script otomasi utama Approve Admin.
- [data_edited.xlsx](file:///Users/user/Documents/Python/approve_admin_faseh/data_edited.xlsx) : File Excel data assignment (default root).
- [input_fasih_ntb/data_edited.xlsx](file:///Users/user/Documents/Python/approve_admin_faseh/input_fasih_ntb/data_edited.xlsx) : File Excel data assignment lengkap dari survei NTB.
- [01_bukalink.html](file:///Users/user/Documents/Python/approve_admin_faseh/01_bukalink.html) : Referensi struktur halaman saat membuka assignment.
- [02_btnapprove.html](file:///Users/user/Documents/Python/approve_admin_faseh/02_btnapprove.html) : Referensi elemen tombol Approve (`f:bg-success`, icon `tabler-icon-check`).
- [03_modalkonfirmasi.html](file:///Users/user/Documents/Python/approve_admin_faseh/03_modalkonfirmasi.html) : Referensi modal dialog konfirmasi Approve.
- [04_revoke.html](file:///Users/user/Documents/Python/approve_admin_faseh/04_revoke.html) : Referensi elemen tombol Revoke (rotasi putar). Jika tombol ini muncul dan tombol Approve tidak ada, berarti assignment sudah di-approve admin dan otomatis di-skip.
- [log_approve.txt](file:///Users/user/Documents/Python/approve_admin_faseh/log_approve.txt) : Log detail eksekusi.
- [log_simple_approve.txt](file:///Users/user/Documents/Python/approve_admin_faseh/log_simple_approve.txt) : Log ringkas per assignment untuk resume.
- [log_terakhir_approve.txt](file:///Users/user/Documents/Python/approve_admin_faseh/log_terakhir_approve.txt) : Pemantau status progres secara real-time.

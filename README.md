# Laporan Harian Pekerjaan - PT. Berjaya Makmur Indonesia

Web app laporan harian pekerjaan berbasis Google Sheets + Google Apps Script.

## Fitur
- Login admin, kepala divisi, dan karyawan (akun dikelola admin, tersimpan di Google Sheets)
- Manajemen karyawan (divisi, jabatan, status akun, foto profil)
- Laporan harian dengan banyak pekerjaan: status (Selesai/Progres/Tunda), kendala, solusi, rencana besok
- Upload foto/file pekerjaan (maks. 5 per pekerjaan) ke Google Drive
- Dashboard admin: KPI, donut, bar, pie, tren, progress per divisi
- Akses Kepala Divisi: memantau laporan anggota divisinya
- Rekap bulanan per karyawan / per divisi
- Cetak / simpan PDF dengan kop surat (laporan berfoto dan rekap)

## File
| File | Fungsi |
|---|---|
| `setup.gs` | Inisialisasi database (Safe Migrate, tidak menghapus data) |
| `code.gs` | Backend: auth, CRUD, upload, dashboard, rekap |
| `index.html` | Frontend lengkap (responsif HP dan laptop) |

## Cara deploy ke Google Apps Script
1. Buat Google Sheet baru, buka **Extensions > Apps Script**.
2. Buat 3 file bernama persis `setup.gs`, `code.gs`, `index.html` lalu salin isinya dari repo ini.
3. Jalankan fungsi `setupDatabase` sekali (izinkan akses Sheets dan Drive). Jalankan ulang setelah setiap pembaruan skema.
4. **Deploy > New deployment > Web app** (Execute as: Me, Who has access: Anyone).
5. Login awal: `admin` / `admin123` (wajib diganti lewat menu Ubah Password).

Catatan: GitHub hanya menyimpan kode. Aplikasi berjalan lewat deployment Apps Script, bukan GitHub Pages.

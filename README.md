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

## Sinkron otomatis GitHub ke Apps Script (clasp + GitHub Actions)
Setiap push ke `main` yang mengubah `setup.gs`, `code.gs`, `index.html`, atau `appsscript.json` otomatis dikirim ke proyek Apps Script lewat workflow `.github/workflows/deploy-apps-script.yml`.

Persiapan satu kali:
1. Aktifkan **Google Apps Script API** di https://script.google.com/home/usersettings
2. Di komputermu: `npm install -g @google/clasp@2.4.2`, lalu `clasp login`. Ini membuat file `~/.clasprc.json`.
3. Di GitHub: **Settings > Secrets and variables > Actions > New repository secret**, buat:
   - `CLASPRC_JSON`: seluruh isi file `~/.clasprc.json`
   - `SCRIPT_ID`: Script ID proyek (Apps Script: Project Settings > IDs)
   - `DEPLOYMENT_ID` (opsional tapi disarankan): ID deployment web app (Deploy > Manage deployments), agar URL web app tetap sama dan versinya ikut diperbarui. Tanpa ini, kode ter-push tetapi kamu perlu membuat versi deployment baru manual.
4. Push ke `main` atau jalankan manual lewat tab **Actions > Deploy ke Google Apps Script > Run workflow**.

Catatan:
- Sinkron satu arah (GitHub ke Apps Script). Perubahan yang dibuat langsung di editor Apps Script akan tertimpa saat push berikutnya.
- Setelah perubahan skema database, jalankan `setupDatabase` sekali dari editor Apps Script.
- `CLASPRC_JSON` berisi token akses akunmu. Simpan hanya sebagai GitHub Secret, jangan pernah di-commit.

// =====================================================================
// setup.gs - Inisialisasi Database (SAFE MIGRATE, non-destruktif)
// PT. BERJAYA MAKMUR INDONESIA - Laporan Harian Pekerjaan
// Cara pakai: pilih fungsi setupDatabase lalu klik Run (sekali saja).
// Aman dijalankan ulang: tidak pernah memakai clear() dan tidak menghapus data.
// =====================================================================

// Skema header setiap sheet
var SCHEMA = {
  Users: ['ID_User', 'Username', 'Password_Hash', 'Role', 'Nama_Karyawan', 'Divisi', 'Jabatan', 'Status_Akun', 'Dibuat_Pada', 'Diubah_Pada', 'Foto_Profil'],
  Laporan: ['ID_Laporan', 'Tanggal', 'ID_User', 'Nama_Karyawan', 'Divisi', 'Jabatan', 'Rencana_Besok', 'Dibuat_Pada', 'Diubah_Pada', 'Diubah_Oleh'],
  Detail_Pekerjaan: ['ID_Detail', 'ID_Laporan', 'Tanggal', 'Deskripsi_Pekerjaan', 'Status_Pekerjaan', 'Kendala', 'Solusi', 'Link_File', 'Nama_File'],
  Sesi: ['Token', 'ID_User', 'Role', 'Berlaku_Hingga'],
  Config: ['key', 'value']
};

// Data dummy akun (hanya dimasukkan jika sheet Users masih kosong)
var DUMMY_USERS = [
  ['admin', 'admin123', 'Admin', 'Administrator', '-', '-'],
  ['budi.mech', 'budi123', 'Karyawan', 'Budi Santoso', 'Mechanical', 'Kepala Divisi'],
  ['andi.mech', 'andi123', 'Karyawan', 'Andi Pratama', 'Mechanical', 'Anggota Divisi'],
  ['sari.elec', 'sari123', 'Karyawan', 'Sari Wulandari', 'Electrical', 'Kepala Divisi'],
  ['dedi.elec', 'dedi123', 'Karyawan', 'Dedi Kurniawan', 'Electrical', 'Anggota Divisi'],
  ['rina.wh', 'rina123', 'Karyawan', 'Rina Marlina', 'Warehouse', 'Kepala Divisi'],
  ['joko.wh', 'joko123', 'Karyawan', 'Joko Susilo', 'Warehouse', 'Anggota Divisi'],
  ['tika.eng', 'tika123', 'Karyawan', 'Tika Permata', 'Engineering', 'Kepala Divisi'],
  ['fajar.eng', 'fajar123', 'Karyawan', 'Fajar Ramadhan', 'Engineering', 'Anggota Divisi'],
  ['maya.fin', 'maya123', 'Karyawan', 'Maya Anggraini', 'Finance', 'Kepala Divisi'],
  ['hendra.fin', 'hendra123', 'Karyawan', 'Hendra Gunawan', 'Finance', 'Anggota Divisi']
];

// Contoh pekerjaan dummy per divisi
var DUMMY_JOBS = {
  Mechanical: ['Perawatan rutin pompa hydrant', 'Penggantian bearing conveyor', 'Pengecekan oli genset'],
  Electrical: ['Inspeksi panel MDP', 'Penggantian lampu area produksi', 'Pengukuran beban trafo'],
  Warehouse: ['Penerimaan barang masuk', 'Stock opname gudang A', 'Pengiriman material ke proyek'],
  Engineering: ['Revisi gambar kerja layout mesin', 'Survey lokasi instalasi baru', 'Review spesifikasi material'],
  Finance: ['Rekonsiliasi kas harian', 'Input invoice vendor', 'Pembuatan laporan pengeluaran mingguan']
};

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = Object.keys(SCHEMA);

  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var header = SCHEMA[name];
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
    }
    // Pastikan jumlah kolom cukup
    if (sh.getMaxColumns() < header.length) {
      sh.insertColumnsAfter(sh.getMaxColumns(), header.length - sh.getMaxColumns());
    }
    // Semua kolom berformat teks agar tanggal/jam tidak berubah jadi objek Date
    sh.getRange(1, 1, sh.getMaxRows(), header.length).setNumberFormat('@');
    // Safe Migrate: hanya update baris header, data pengguna tidak disentuh
    sh.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#2F7D5B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }

  // Hapus sheet bawaan hanya jika benar-benar kosong
  var defaults = ['Sheet1', 'Lembar1'];
  for (var d = 0; d < defaults.length; d++) {
    var def = ss.getSheetByName(defaults[d]);
    if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(def);
    }
  }

  seedConfig_(ss);
  var usersSeeded = seedUsers_(ss);
  if (usersSeeded) {
    seedLaporan_(ss);
  }
  getUploadFolder_(); // buat folder Drive untuk lampiran jika belum ada
  SpreadsheetApp.flush();

  Logger.log('Setup selesai. Akun awal: admin / admin123 (WAJIB diganti setelah login pertama).');
}

// Tambah key Config yang belum ada (tidak menimpa nilai yang sudah ada)
function seedConfig_(ss) {
  var sh = ss.getSheetByName('Config');
  var existing = {};
  if (sh.getLastRow() > 1) {
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues();
    for (var i = 0; i < vals.length; i++) { existing[vals[i][0]] = true; }
  }
  var defaults = [
    ['nama_perusahaan', 'PT. BERJAYA MAKMUR INDONESIA'],
    ['alamat', 'Komplek DDN No.Kav 23, Harjamukti, Kec. Cimanggis, Kota Depok, Jawa Barat 16454'],
    ['telepon', '021-38852384'],
    ['website', 'https://www.berjaya-makmur-indonesia.co.id/'],
    ['hari_kerja', '5'],
    ['folder_id', '']
  ];
  for (var j = 0; j < defaults.length; j++) {
    if (!existing[defaults[j][0]]) {
      sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setValues([defaults[j]]);
    }
  }
}

// Isi akun dummy hanya jika belum ada data. Return true jika di-seed.
function seedUsers_(ss) {
  var sh = ss.getSheetByName('Users');
  if (sh.getLastRow() > 1) { return false; }
  var now = nowStr_();
  var rows = [];
  for (var i = 0; i < DUMMY_USERS.length; i++) {
    var u = DUMMY_USERS[i];
    rows.push(['USR-' + pad4_(i + 1), u[0], hashPassword_(u[1]), u[2], u[3], u[4], u[5], 'Aktif', now, now, '']);
  }
  sh.getRange(2, 1, rows.length, SCHEMA.Users.length).setValues(rows);
  return true;
}

// Laporan dummy 7 hari terakhir. ID reset harian: LAP-0001, DTL-0001
function seedLaporan_(ss) {
  var lap = ss.getSheetByName('Laporan');
  var det = ss.getSheetByName('Detail_Pekerjaan');
  if (lap.getLastRow() > 1 || det.getLastRow() > 1) { return; }

  var kary = [];
  for (var i = 0; i < DUMMY_USERS.length; i++) {
    if (DUMMY_USERS[i][2] === 'Karyawan') {
      kary.push({ id: 'USR-' + pad4_(i + 1), nama: DUMMY_USERS[i][3], divisi: DUMMY_USERS[i][4], jabatan: DUMMY_USERS[i][5] });
    }
  }
  var cycle = ['Selesai', 'Selesai', 'Progres', 'Tunda', 'Selesai', 'Progres'];
  var now = nowStr_();
  var base = parseD_(todayStr_());
  var lapRows = [];
  var detRows = [];

  for (var d = 6; d >= 0; d--) {
    var tgl = fmtD_(base - d * 86400000);
    var lc = 0;
    var dc = 0;
    for (var k = 0; k < kary.length; k++) {
      if ((k + d) % 5 === 4) { continue; }      // sebagian karyawan tidak lapor
      if (d === 0 && k % 3 === 2) { continue; } // hari ini sebagian belum lapor
      lc++;
      var lid = 'LAP-' + pad4_(lc);
      var u = kary[k];
      var jobs = DUMMY_JOBS[u.divisi];
      var n = 1 + ((k + d) % 3);
      lapRows.push([lid, tgl, u.id, u.nama, u.divisi, u.jabatan, 'Melanjutkan pekerjaan prioritas dan koordinasi dengan tim', now, now, u.nama]);
      for (var j = 0; j < n; j++) {
        dc++;
        var st = cycle[(k + d + j) % cycle.length];
        detRows.push([
          'DTL-' + pad4_(dc), lid, tgl, jobs[j % jobs.length], st,
          st === 'Selesai' ? '' : 'Material atau persetujuan belum lengkap',
          st === 'Selesai' ? '' : 'Koordinasi dengan pihak terkait dan jadwalkan ulang',
          '', ''
        ]);
      }
    }
  }
  if (lapRows.length) { lap.getRange(2, 1, lapRows.length, SCHEMA.Laporan.length).setValues(lapRows); }
  if (detRows.length) { det.getRange(2, 1, detRows.length, SCHEMA.Detail_Pekerjaan.length).setValues(detRows); }
}

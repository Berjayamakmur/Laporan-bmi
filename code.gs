// =====================================================================
// code.gs - Backend Laporan Harian Pekerjaan
// PT. BERJAYA MAKMUR INDONESIA
// Semua pemanggilan frontend lewat satu pintu: api(action, token, payload)
// =====================================================================

var TZ = 'Asia/Jakarta';
var DIVISI_LIST = ['Mechanical', 'Electrical', 'Warehouse', 'Engineering', 'Finance'];
var JABATAN_LIST = ['Kepala Divisi', 'Anggota Divisi'];
var STATUS_LIST = ['Selesai', 'Progres', 'Tunda'];
var SESSION_MS = 12 * 60 * 60 * 1000; // sesi login 12 jam
var MAX_FILE_B64 = 7500000;           // batas base64 (sekitar 5 MB file)
var ALLOWED_MIME = [
  'image/jpeg', 'image/png', 'image/webp', 'application/pdf',
  'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
];

// ---------------------------------------------------------------------
// Entry point web app
// ---------------------------------------------------------------------
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('Laporan Harian - PT. Berjaya Makmur Indonesia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ---------------------------------------------------------------------
// Dispatcher API
// ---------------------------------------------------------------------
function api(action, token, p) {
  try {
    p = p || {};
    switch (action) {
      case 'login': return apiLogin_(p);
      case 'logout': return apiLogout_(token);
      case 'me': return apiMe_(token);
      case 'changePassword': return apiChangePassword_(auth_(token), p);
      case 'dashboard': return apiDashboard_(auth_(token, 'monitor'), p);
      case 'rekap': return apiRekap_(auth_(token, 'monitor'), p);
      case 'profilePhoto': return apiProfilePhoto_(auth_(token), p);
      case 'fotoData': return apiFotoData_(auth_(token), p);
      case 'myStats': return apiMyStats_(auth_(token));
      case 'reports': return apiReportList_(auth_(token), p);
      case 'reportSave': return apiReportSave_(auth_(token), p);
      case 'reportDelete': return apiReportDelete_(auth_(token), p);
      case 'upload': return apiUpload_(auth_(token), p);
      case 'users': return apiUsersList_(auth_(token, true), p);
      case 'userSave': return apiUserSave_(auth_(token, true), p);
      case 'userStatus': return apiUserStatus_(auth_(token, true), p);
      case 'karyawanOpts': return apiKaryawanOpts_(auth_(token, true));
      default: throw new Error('Aksi tidak dikenal.');
    }
  } catch (e) {
    var m = String(e && e.message ? e.message : e);
    if (m === 'AUTH') { return { ok: false, auth: true, msg: 'Sesi berakhir, silakan login ulang.' }; }
    return { ok: false, msg: m };
  }
}

// ---------------------------------------------------------------------
// Helper umum
// ---------------------------------------------------------------------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function pad2_(n) { return (n < 10 ? '0' : '') + n; }
function pad4_(n) { return ('0000' + n).slice(-4); }
function nowStr_() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'); }
function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy'); }

// dd/MM/yyyy -> angka yyyyMMdd (untuk perbandingan)
function dkey_(s) {
  var p = String(s || '').split('/');
  if (p.length !== 3) { return 0; }
  return Number(p[2]) * 10000 + Number(p[1]) * 100 + Number(p[0]);
}
// dd/MM/yyyy -> epoch UTC (aman dari beda timezone)
function parseD_(s) {
  var p = String(s).split('/');
  return Date.UTC(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
}
function fmtD_(t) {
  var d = new Date(t);
  return pad2_(d.getUTCDate()) + '/' + pad2_(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear();
}

// Baca seluruh isi sheet (tanpa header) sebagai string murni
function readSheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) { throw new Error('Sheet ' + name + ' belum ada. Jalankan setupDatabase() dulu.'); }
  var lr = sh.getLastRow();
  var rows = [];
  if (lr > 1) {
    rows = sh.getRange(2, 1, lr - 1, SCHEMA[name].length).getDisplayValues();
  }
  return { sh: sh, rows: rows };
}

function obj_(name, row) {
  var h = SCHEMA[name];
  var o = {};
  for (var i = 0; i < h.length; i++) { o[h[i]] = row[i]; }
  return o;
}

function hashPassword_(pw) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'BMI|' + String(pw), Utilities.Charset.UTF_8);
  var s = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] < 0 ? bytes[i] + 256 : bytes[i]).toString(16);
    if (b.length < 2) { b = '0' + b; }
    s += b;
  }
  return s;
}

function getConfigMap_() {
  var r = readSheet_('Config');
  var m = {};
  for (var i = 0; i < r.rows.length; i++) { m[r.rows[i][0]] = r.rows[i][1]; }
  return m;
}

function getUploadFolder_() {
  var cfg = readSheet_('Config');
  var id = '';
  var rowIdx = -1;
  for (var i = 0; i < cfg.rows.length; i++) {
    if (cfg.rows[i][0] === 'folder_id') { id = cfg.rows[i][1]; rowIdx = i + 2; }
  }
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* folder hilang, buat baru */ }
  }
  var f = DriveApp.createFolder('Laporan Harian BMI - Lampiran');
  if (rowIdx > 0) {
    cfg.sh.getRange(rowIdx, 2).setValue(f.getId());
  } else {
    cfg.sh.getRange(cfg.sh.getLastRow() + 1, 1, 1, 2).setValues([['folder_id', f.getId()]]);
  }
  return f;
}

function publicConfig_() {
  var m = getConfigMap_();
  return {
    perusahaan: m.nama_perusahaan || 'PT. BERJAYA MAKMUR INDONESIA',
    alamat: m.alamat || 'Komplek DDN No.Kav 23, Harjamukti, Kec. Cimanggis, Kota Depok, Jawa Barat 16454',
    telepon: m.telepon || '021-38852384',
    website: m.website || 'https://www.berjaya-makmur-indonesia.co.id/',
    divisi: DIVISI_LIST, jabatan: JABATAN_LIST, status: STATUS_LIST
  };
}

// Kepala Divisi = karyawan dengan jabatan Kepala Divisi (otomatis dapat akses monitoring divisinya)
function isKepala_(o) {
  var role = o.Role !== undefined ? o.Role : o.role;
  var jab = o.Jabatan !== undefined ? o.Jabatan : o.jabatan;
  return role === 'Karyawan' && jab === 'Kepala Divisi';
}

function publicUser_(o) {
  return { id: o.ID_User, username: o.Username, role: o.Role, nama: o.Nama_Karyawan, divisi: o.Divisi, jabatan: o.Jabatan, kepala: isKepala_(o), foto: o.Foto_Profil || '' };
}

function userOut_(row) {
  return { id: row[0], username: row[1], role: row[3], nama: row[4], divisi: row[5], jabatan: row[6], status: row[7], foto: row[10] || '' };
}

function findUserById_(id) {
  var r = readSheet_('Users');
  for (var i = 0; i < r.rows.length; i++) {
    if (r.rows[i][0] === id) { return obj_('Users', r.rows[i]); }
  }
  return null;
}

// ---------------------------------------------------------------------
// Sesi & Auth
// ---------------------------------------------------------------------
function createSession_(user) {
  var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  var exp = Date.now() + SESSION_MS;
  var sh = ss_().getSheetByName('Sesi');
  if (sh.getLastRow() > 150) { purgeSessions_(); }
  sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues([[token, user.ID_User, user.Role, String(exp)]]);
  CacheService.getScriptCache().put('S_' + token, JSON.stringify({ id: user.ID_User, exp: exp }), 600);
  return token;
}

function purgeSessions_() {
  var r = readSheet_('Sesi');
  var now = Date.now();
  for (var i = r.rows.length - 1; i >= 0; i--) {
    if (Number(r.rows[i][3]) < now) { r.sh.deleteRow(i + 2); }
  }
}

function auth_(token, needAdmin) {
  if (!token) { throw new Error('AUTH'); }
  var cache = CacheService.getScriptCache();
  var uid = null;
  var exp = 0;
  var c = cache.get('S_' + token);
  if (c) {
    var o = JSON.parse(c);
    uid = o.id; exp = o.exp;
  } else {
    var r = readSheet_('Sesi');
    for (var i = 0; i < r.rows.length; i++) {
      if (r.rows[i][0] === token) { uid = r.rows[i][1]; exp = Number(r.rows[i][3]); break; }
    }
    if (uid) { cache.put('S_' + token, JSON.stringify({ id: uid, exp: exp }), 600); }
  }
  if (!uid || exp < Date.now()) { throw new Error('AUTH'); }
  var user = findUserById_(uid);
  if (!user || user.Status_Akun !== 'Aktif') { throw new Error('AUTH'); }
  if (needAdmin === 'monitor') {
    if (user.Role !== 'Admin' && !isKepala_(user)) { throw new Error('Akses ditolak: khusus Admin atau Kepala Divisi.'); }
  } else if (needAdmin && user.Role !== 'Admin') {
    throw new Error('Akses ditolak: khusus Admin.');
  }
  return user;
}

function apiLogin_(p) {
  var username = String(p.username || '').trim().toLowerCase();
  var password = String(p.password || '');
  if (!username || !password) { throw new Error('ID dan password wajib diisi.'); }
  var r = readSheet_('Users');
  var hash = hashPassword_(password);
  for (var i = 0; i < r.rows.length; i++) {
    var row = r.rows[i];
    if (String(row[1]).toLowerCase() === username && row[2] === hash) {
      if (row[7] !== 'Aktif') { throw new Error('Akun dinonaktifkan. Hubungi admin.'); }
      var u = obj_('Users', row);
      return { ok: true, token: createSession_(u), user: publicUser_(u), cfg: publicConfig_() };
    }
  }
  throw new Error('ID atau password salah.');
}

function apiLogout_(token) {
  if (token) {
    CacheService.getScriptCache().remove('S_' + token);
    var r = readSheet_('Sesi');
    for (var i = r.rows.length - 1; i >= 0; i--) {
      if (r.rows[i][0] === token) { r.sh.deleteRow(i + 2); break; }
    }
  }
  return { ok: true };
}

function apiMe_(token) {
  var u = auth_(token);
  return { ok: true, user: publicUser_(u), cfg: publicConfig_() };
}

function apiChangePassword_(user, p) {
  var oldPw = String(p.oldPassword || '');
  var newPw = String(p.newPassword || '');
  if (hashPassword_(oldPw) !== user.Password_Hash) { throw new Error('Password lama salah.'); }
  if (newPw.length < 6) { throw new Error('Password baru minimal 6 karakter.'); }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var r = readSheet_('Users');
    for (var i = 0; i < r.rows.length; i++) {
      if (r.rows[i][0] === user.ID_User) {
        r.sh.getRange(i + 2, 3).setValue(hashPassword_(newPw));
        r.sh.getRange(i + 2, 10).setValue(nowStr_());
        break;
      }
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return { ok: true };
}

// ---------------------------------------------------------------------
// Manajemen karyawan (Admin)
// ---------------------------------------------------------------------
function apiUsersList_(admin, p) {
  var r = readSheet_('Users');
  var q = String(p.search || '').toLowerCase();
  var out = [];
  for (var i = 0; i < r.rows.length; i++) {
    var row = r.rows[i];
    if (p.divisi && row[5] !== p.divisi) { continue; }
    if (p.status && row[7] !== p.status) { continue; }
    if (q && (row[4] + ' ' + row[1]).toLowerCase().indexOf(q) < 0) { continue; }
    out.push(row);
  }
  out.sort(function (a, b) { return a[4].toLowerCase() < b[4].toLowerCase() ? -1 : 1; });
  var offset = Math.max(0, Number(p.offset) || 0);
  var limit = Math.min(100, Math.max(1, Number(p.limit) || 15));
  var page = out.slice(offset, offset + limit).map(userOut_);
  return { ok: true, rows: page, total: out.length, hasMore: offset + limit < out.length };
}

function apiKaryawanOpts_(admin) {
  var r = readSheet_('Users');
  var out = [];
  for (var i = 0; i < r.rows.length; i++) {
    var row = r.rows[i];
    if (row[3] === 'Karyawan' && row[7] === 'Aktif') {
      out.push({ id: row[0], nama: row[4], divisi: row[5], jabatan: row[6] });
    }
  }
  out.sort(function (a, b) { return a.nama.toLowerCase() < b.nama.toLowerCase() ? -1 : 1; });
  return { ok: true, rows: out };
}

function apiUserSave_(admin, p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var nama = String(p.nama || '').trim();
    var username = String(p.username || '').trim();
    var role = String(p.role || 'Karyawan');
    var status = String(p.status || 'Aktif');
    var divisi = String(p.divisi || '');
    var jabatan = String(p.jabatan || '');
    var pw = String(p.password || '');
    var id = String(p.id || '');

    if (!nama) { throw new Error('Nama karyawan wajib diisi.'); }
    if (!/^[A-Za-z0-9._-]{3,30}$/.test(username)) { throw new Error('ID login 3-30 karakter (huruf, angka, titik, strip, underscore).'); }
    if (['Admin', 'Karyawan'].indexOf(role) < 0) { throw new Error('Role tidak valid.'); }
    if (['Aktif', 'Nonaktif'].indexOf(status) < 0) { throw new Error('Status akun tidak valid.'); }
    if (role === 'Karyawan') {
      if (DIVISI_LIST.indexOf(divisi) < 0) { throw new Error('Divisi tidak valid.'); }
      if (JABATAN_LIST.indexOf(jabatan) < 0) { throw new Error('Jabatan tidak valid.'); }
    } else {
      divisi = '-'; jabatan = '-';
    }

    var r = readSheet_('Users');
    var idx = -1;
    var maxN = 0;
    for (var i = 0; i < r.rows.length; i++) {
      var rw = r.rows[i];
      var n = parseInt(String(rw[0]).split('-')[1], 10) || 0;
      if (n > maxN) { maxN = n; }
      if (String(rw[1]).toLowerCase() === username.toLowerCase() && rw[0] !== id) { throw new Error('ID login sudah dipakai.'); }
      if (id && rw[0] === id) { idx = i; }
    }

    var now = nowStr_();
    var row;
    if (id) {
      if (idx < 0) { throw new Error('Data karyawan tidak ditemukan.'); }
      var old = r.rows[idx];
      var hash = old[2];
      if (pw) {
        if (pw.length < 6) { throw new Error('Password minimal 6 karakter.'); }
        hash = hashPassword_(pw);
      }
      if (id === admin.ID_User) { role = 'Admin'; status = 'Aktif'; divisi = old[5]; jabatan = old[6]; }
      row = [id, username, hash, role, nama, divisi, jabatan, status, old[8], now, old[10] || ''];
      r.sh.getRange(idx + 2, 1, 1, 11).setValues([row]);
      CacheService.getScriptCache().remove('U_' + id);
    } else {
      if (pw.length < 6) { throw new Error('Password wajib diisi minimal 6 karakter.'); }
      row = ['USR-' + pad4_(maxN + 1), username, hashPassword_(pw), role, nama, divisi, jabatan, status, now, now, ''];
      r.sh.getRange(r.sh.getLastRow() + 1, 1, 1, 11).setValues([row]);
    }
    SpreadsheetApp.flush();
    return { ok: true, user: userOut_(row) };
  } finally { lock.releaseLock(); }
}

function apiUserStatus_(admin, p) {
  var id = String(p.id || '');
  var status = String(p.status || '');
  if (['Aktif', 'Nonaktif'].indexOf(status) < 0) { throw new Error('Status tidak valid.'); }
  if (id === admin.ID_User) { throw new Error('Akun admin yang sedang login tidak bisa dinonaktifkan.'); }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var r = readSheet_('Users');
    for (var i = 0; i < r.rows.length; i++) {
      if (r.rows[i][0] === id) {
        r.sh.getRange(i + 2, 8).setValue(status);
        r.sh.getRange(i + 2, 10).setValue(nowStr_());
        SpreadsheetApp.flush();
        return { ok: true };
      }
    }
  } finally { lock.releaseLock(); }
  throw new Error('Data karyawan tidak ditemukan.');
}

// ---------------------------------------------------------------------
// Laporan
// ---------------------------------------------------------------------
function reportOut_(r, its) {
  var items = [];
  var c = { selesai: 0, progres: 0, tunda: 0, total: 0 };
  for (var i = 0; i < its.length; i++) {
    var x = its[i];
    items.push({ id: x[0], deskripsi: x[3], status: x[4], kendala: x[5], solusi: x[6], files: filesOut_(x[7], x[8]) });
    if (x[4] === 'Selesai') { c.selesai++; } else if (x[4] === 'Progres') { c.progres++; } else if (x[4] === 'Tunda') { c.tunda++; }
  }
  c.total = items.length;
  return {
    id: r[0], tanggal: r[1], userId: r[2], nama: r[3], divisi: r[4], jabatan: r[5],
    rencana: r[6], dibuat: r[7], diubah: r[8], diubahOleh: r[9], items: items, counts: c
  };
}

// Filter + gabung laporan & detail, seluruh proses di backend
function queryReports_(user, p) {
  var lap = readSheet_('Laporan').rows;
  var det = readSheet_('Detail_Pekerjaan').rows;
  var map = {};
  for (var i = 0; i < det.length; i++) {
    var k = det[i][1] + '|' + det[i][2];
    if (!map[k]) { map[k] = []; }
    map[k].push(det[i]);
  }
  var from = dkey_(p.dateFrom) || 0;
  var to = dkey_(p.dateTo) || 99999999;
  var q = String(p.search || '').toLowerCase();
  var out = [];
  for (var j = 0; j < lap.length; j++) {
    var r = lap[j];
    var dk = dkey_(r[1]);
    if (dk < from || dk > to) { continue; }
    if (user.Role !== 'Admin') {
      if (isKepala_(user)) { if (r[4] !== user.Divisi) { continue; } }   // Kepala Divisi: hanya divisinya
      else if (r[2] !== user.ID_User) { continue; }                      // Karyawan: hanya miliknya
    }
    if (p.userId && r[2] !== p.userId) { continue; }
    if (p.divisi && r[4] !== p.divisi) { continue; }
    var its = map[r[0] + '|' + r[1]] || [];
    if (p.status) {
      var has = false;
      for (var a = 0; a < its.length; a++) { if (its[a][4] === p.status) { has = true; break; } }
      if (!has) { continue; }
    }
    if (q) {
      var hay = r[0] + ' ' + r[3] + ' ' + r[4];
      for (var b = 0; b < its.length; b++) { hay += ' ' + its[b][3]; }
      if (hay.toLowerCase().indexOf(q) < 0) { continue; }
    }
    out.push({ r: r, its: its, dk: dk });
  }
  out.sort(function (x, y) {
    if (y.dk !== x.dk) { return y.dk - x.dk; }
    return y.r[0] < x.r[0] ? -1 : (y.r[0] > x.r[0] ? 1 : 0);
  });
  return out;
}

function apiReportList_(user, p) {
  var all = queryReports_(user, p);
  var offset = Math.max(0, Number(p.offset) || 0);
  var limit = Math.min(200, Math.max(1, Number(p.limit) || 15));
  var rows = [];
  var slice = all.slice(offset, offset + limit);
  for (var i = 0; i < slice.length; i++) { rows.push(reportOut_(slice[i].r, slice[i].its)); }
  return { ok: true, rows: rows, total: all.length, hasMore: offset + limit < all.length };
}

function nextSeq_(rows, idxDate, idxId, tanggal) {
  var max = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][idxDate] === tanggal) {
      var n = parseInt(String(rows[i][idxId]).split('-')[1], 10) || 0;
      if (n > max) { max = n; }
    }
  }
  return max + 1;
}

function cleanItems_(items) {
  if (!items || !items.length) { throw new Error('Minimal satu pekerjaan harus diisi.'); }
  if (items.length > 30) { throw new Error('Maksimal 30 pekerjaan per laporan.'); }
  var out = [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i] || {};
    var d = String(it.deskripsi || '').trim();
    var st = String(it.status || '');
    if (!d) { throw new Error('Deskripsi pekerjaan #' + (i + 1) + ' wajib diisi.'); }
    if (STATUS_LIST.indexOf(st) < 0) { throw new Error('Status pekerjaan #' + (i + 1) + ' tidak valid.'); }
    // Lampiran: maksimal 5 file per pekerjaan, disimpan dipisah tanda |
    var urls = [];
    var names = [];
    var fl = it.files || [];
    for (var f = 0; f < fl.length && urls.length < 5; f++) {
      var fu = String((fl[f] && fl[f].url) || '');
      if (fu.indexOf('https://') !== 0 || fu.indexOf('|') >= 0) { continue; }
      urls.push(fu);
      names.push(String((fl[f] && fl[f].name) || 'Lampiran').replace(/[|]/g, '_').slice(0, 100));
    }
    out.push({
      deskripsi: d.slice(0, 1000), status: st,
      kendala: String(it.kendala || '').slice(0, 1000),
      solusi: String(it.solusi || '').slice(0, 1000),
      link: urls.join('|'), namaFile: names.join('|')
    });
  }
  return out;
}

function apiReportSave_(user, p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var tgl = String(p.tanggal || '');
    if (!/^\d{2}\/\d{2}\/\d{4}$/.test(tgl)) { throw new Error('Format tanggal tidak valid.'); }
    if (dkey_(tgl) > dkey_(todayStr_())) { throw new Error('Tanggal laporan tidak boleh melebihi hari ini.'); }
    var items = cleanItems_(p.items);
    var rencana = String(p.rencana || '').slice(0, 2000);
    var id = String(p.id || '');
    var now = nowStr_();
    var lap = readSheet_('Laporan');
    var lapRow;

    if (id) {
      // ---------- EDIT ----------
      var idx = -1;
      for (var i = 0; i < lap.rows.length; i++) {
        if (lap.rows[i][0] === id && lap.rows[i][1] === tgl) { idx = i; break; }
      }
      if (idx < 0) { throw new Error('Laporan tidak ditemukan.'); }
      var old = lap.rows[idx];
      if (user.Role !== 'Admin' && old[2] !== user.ID_User) { throw new Error('Anda tidak berhak mengubah laporan ini.'); }
      lapRow = [old[0], old[1], old[2], old[3], old[4], old[5], rencana, old[7], now, user.Nama_Karyawan];
      lap.sh.getRange(idx + 2, 1, 1, 10).setValues([lapRow]);
      deleteDetails_(id, tgl);
    } else {
      // ---------- BARU ----------
      var owner = user;
      if (user.Role === 'Admin') {
        owner = findUserById_(String(p.userId || ''));
        if (!owner || owner.Role !== 'Karyawan' || owner.Status_Akun !== 'Aktif') { throw new Error('Pilih karyawan yang valid.'); }
      }
      for (var k = 0; k < lap.rows.length; k++) {
        if (lap.rows[k][1] === tgl && lap.rows[k][2] === owner.ID_User) {
          throw new Error('Laporan tanggal ' + tgl + ' untuk ' + owner.Nama_Karyawan + ' sudah ada (' + lap.rows[k][0] + '). Silakan edit laporan tersebut.');
        }
      }
      id = 'LAP-' + pad4_(nextSeq_(lap.rows, 1, 0, tgl));
      lapRow = [id, tgl, owner.ID_User, owner.Nama_Karyawan, owner.Divisi, owner.Jabatan, rencana, now, now, user.Nama_Karyawan];
      lap.sh.getRange(lap.sh.getLastRow() + 1, 1, 1, 10).setValues([lapRow]);
    }

    // Tulis detail pekerjaan (batch)
    var det = readSheet_('Detail_Pekerjaan');
    var seq = nextSeq_(det.rows, 2, 0, tgl);
    var detRows = [];
    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      detRows.push(['DTL-' + pad4_(seq + j), id, tgl, it.deskripsi, it.status, it.kendala, it.solusi, it.link, it.namaFile]);
    }
    det.sh.getRange(det.sh.getLastRow() + 1, 1, detRows.length, 9).setValues(detRows);
    SpreadsheetApp.flush();
    return { ok: true, report: reportOut_(lapRow, detRows) };
  } finally { lock.releaseLock(); }
}

function deleteDetails_(id, tgl) {
  var det = readSheet_('Detail_Pekerjaan');
  for (var i = det.rows.length - 1; i >= 0; i--) {
    if (det.rows[i][1] === id && det.rows[i][2] === tgl) { det.sh.deleteRow(i + 2); }
  }
}

function apiReportDelete_(user, p) {
  var id = String(p.id || '');
  var tgl = String(p.tanggal || '');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var lap = readSheet_('Laporan');
    for (var i = 0; i < lap.rows.length; i++) {
      if (lap.rows[i][0] === id && lap.rows[i][1] === tgl) {
        if (user.Role !== 'Admin' && lap.rows[i][2] !== user.ID_User) { throw new Error('Anda tidak berhak menghapus laporan ini.'); }
        deleteDetails_(id, tgl);
        lap.sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { ok: true };
      }
    }
  } finally { lock.releaseLock(); }
  throw new Error('Laporan tidak ditemukan.');
}

// ---------------------------------------------------------------------
// Upload file bukti pekerjaan ke Google Drive
// ---------------------------------------------------------------------
function apiUpload_(user, p) {
  var name = String(p.name || 'file').replace(/[^A-Za-z0-9._ -]/g, '_').slice(0, 100);
  var mime = String(p.mime || '');
  var data = String(p.data || '');
  if (ALLOWED_MIME.indexOf(mime) < 0) { throw new Error('Tipe file tidak diizinkan (gunakan gambar, PDF, Word, atau Excel).'); }
  if (!data || data.length > MAX_FILE_B64) { throw new Error('Ukuran file terlalu besar (maksimal sekitar 5 MB).'); }
  var blob = Utilities.newBlob(Utilities.base64Decode(data), mime, user.ID_User + '_' + Date.now() + '_' + name);
  var file = getUploadFolder_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* kebijakan domain bisa menolak, abaikan */ }
  return { ok: true, url: file.getUrl(), name: name };
}

// ---------------------------------------------------------------------
// Dashboard
// ---------------------------------------------------------------------
function summarize_(rows) {
  var s = { laporan: rows.length, pekerjaan: 0, selesai: 0, progres: 0, tunda: 0 };
  for (var i = 0; i < rows.length; i++) {
    var its = rows[i].its;
    for (var j = 0; j < its.length; j++) {
      s.pekerjaan++;
      if (its[j][4] === 'Selesai') { s.selesai++; } else if (its[j][4] === 'Progres') { s.progres++; } else if (its[j][4] === 'Tunda') { s.tunda++; }
    }
  }
  s.pctSelesai = s.pekerjaan ? Math.round(s.selesai * 100 / s.pekerjaan) : 0;
  return s;
}

function apiMyStats_(user) {
  var today = todayStr_();
  var from = fmtD_(parseD_(today) - 29 * 86400000);
  var rows = queryReports_(user, { dateFrom: from, dateTo: today, userId: user.ID_User });
  var s = summarize_(rows);
  var sudah = false;
  for (var i = 0; i < rows.length; i++) { if (rows[i].r[1] === today) { sudah = true; break; } }
  s.sudahLaporHariIni = sudah;
  return { ok: true, kpi: s };
}

function apiDashboard_(user, p) {
  var divF = user.Role === 'Admin' ? String(p.divisi || '') : user.Divisi;
  var today = todayStr_();
  var to = /^\d{2}\/\d{2}\/\d{4}$/.test(String(p.dateTo || '')) ? p.dateTo : today;
  var from = /^\d{2}\/\d{2}\/\d{4}$/.test(String(p.dateFrom || '')) ? p.dateFrom : fmtD_(parseD_(to) - 6 * 86400000);
  if (parseD_(from) > parseD_(to)) { var tmp = from; from = to; to = tmp; }
  if (parseD_(to) - parseD_(from) > 59 * 86400000) { from = fmtD_(parseD_(to) - 59 * 86400000); }

  var rows = queryReports_(user, { dateFrom: from, dateTo: to, divisi: divF });
  var kpi = summarize_(rows);

  // Per divisi
  var divMap = {};
  for (var d = 0; d < DIVISI_LIST.length; d++) {
    divMap[DIVISI_LIST[d]] = { divisi: DIVISI_LIST[d], laporan: 0, pekerjaan: 0, selesai: 0, progres: 0, tunda: 0, pct: 0 };
  }
  // Per hari
  var hari = [];
  var hariMap = {};
  var t = parseD_(from);
  var tEnd = parseD_(to);
  while (t <= tEnd && hari.length < 60) {
    var dd = new Date(t);
    var o = { key: dkey_(fmtD_(t)), label: pad2_(dd.getUTCDate()) + '/' + pad2_(dd.getUTCMonth() + 1), laporan: 0, pekerjaan: 0, selesai: 0, pct: 0 };
    hari.push(o);
    hariMap[o.key] = o;
    t += 86400000;
  }
  var kMap = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i].r;
    var its = rows[i].its;
    var dv = divMap[r[4]];
    var hh = hariMap[rows[i].dk];
    var km = kMap[r[2]];
    if (!km) { km = kMap[r[2]] = { nama: r[3], laporan: 0, pekerjaan: 0, selesai: 0, pct: 0 }; }
    km.laporan++;
    if (dv) { dv.laporan++; }
    if (hh) { hh.laporan++; }
    for (var j = 0; j < its.length; j++) {
      var st = its[j][4];
      if (dv) { dv.pekerjaan++; if (st === 'Selesai') { dv.selesai++; } else if (st === 'Progres') { dv.progres++; } else if (st === 'Tunda') { dv.tunda++; } }
      if (hh) { hh.pekerjaan++; if (st === 'Selesai') { hh.selesai++; } }
      km.pekerjaan++; if (st === 'Selesai') { km.selesai++; }
    }
  }
  var perDivisi = [];
  for (var d2 = 0; d2 < DIVISI_LIST.length; d2++) {
    var x = divMap[DIVISI_LIST[d2]];
    x.pct = x.pekerjaan ? Math.round(x.selesai * 100 / x.pekerjaan) : 0;
    perDivisi.push(x);
  }
  if (divF) { perDivisi = perDivisi.filter(function (z) { return z.divisi === divF; }); }
  for (var h = 0; h < hari.length; h++) {
    hari[h].pct = hari[h].pekerjaan ? Math.round(hari[h].selesai * 100 / hari[h].pekerjaan) : 0;
    delete hari[h].key;
  }

  // Karyawan yang belum lapor hari ini
  var users = readSheet_('Users').rows;
  var lapToday = {};
  var lapAll = readSheet_('Laporan').rows;
  for (var a = 0; a < lapAll.length; a++) { if (lapAll[a][1] === today) { lapToday[lapAll[a][2]] = true; } }
  var belum = [];
  var aktif = 0;
  var sudah = 0;
  for (var u = 0; u < users.length; u++) {
    var us = users[u];
    if (us[3] !== 'Karyawan' || us[7] !== 'Aktif') { continue; }
    if (divF && us[5] !== divF) { continue; }
    if (divF && !kMap[us[0]]) { kMap[us[0]] = { nama: us[4], laporan: 0, pekerjaan: 0, selesai: 0, pct: 0 }; }
    aktif++;
    if (lapToday[us[0]]) { sudah++; } else { belum.push({ nama: us[4], divisi: us[5], jabatan: us[6] }); }
  }
  kpi.karyawanAktif = aktif;
  kpi.sudahLapor = sudah;
  kpi.belumLapor = belum.length;

  var perAnggota = [];
  if (divF) {
    Object.keys(kMap).forEach(function (key) {
      var m = kMap[key];
      m.pct = m.pekerjaan ? Math.round(m.selesai * 100 / m.pekerjaan) : 0;
      perAnggota.push(m);
    });
    perAnggota.sort(function (a, b) { return b.pct - a.pct || (a.nama < b.nama ? -1 : 1); });
  }

  var recent = [];
  for (var k = 0; k < rows.length && k < 8; k++) { recent.push(reportOut_(rows[k].r, rows[k].its)); }

  return { ok: true, range: { from: from, to: to }, kpi: kpi, perDivisi: perDivisi, perAnggota: perAnggota, perHari: hari, belum: belum.slice(0, 50), recent: recent };
}

// ---------------------------------------------------------------------
// Rekap bulanan (Admin: semua divisi, Kepala Divisi: hanya divisinya)
// ---------------------------------------------------------------------
function newK_(id, nama, divisi, jabatan) {
  return { id: id, nama: nama, divisi: divisi, jabatan: jabatan, hariLapor: 0, pekerjaan: 0, selesai: 0, progres: 0, tunda: 0, kendala: 0, pctSelesai: 0, pctLapor: 0 };
}

function apiRekap_(user, p) {
  var bulan = String(p.bulan || '');
  if (!/^\d{2}\/\d{4}$/.test(bulan)) { throw new Error('Format bulan tidak valid.'); }
  var mm = Number(bulan.substr(0, 2));
  var yy = Number(bulan.substr(3));
  if (mm < 1 || mm > 12) { throw new Error('Bulan tidak valid.'); }
  var lastDay = new Date(Date.UTC(yy, mm, 0)).getUTCDate();
  var from = '01/' + bulan;
  var to = pad2_(lastDay) + '/' + bulan;
  var divF = user.Role === 'Admin' ? String(p.divisi || '') : user.Divisi;
  var userF = String(p.userId || '');

  // Hari kerja: Senin-Jumat (atau Senin-Sabtu jika Config hari_kerja = 6), dihitung sampai hari ini bila bulan berjalan
  var maxDow = Number(getConfigMap_().hari_kerja) === 6 ? 6 : 5;
  var endKey = Math.min(dkey_(to), dkey_(todayStr_()));
  var hariKerja = 0;
  for (var d = 1; d <= lastDay; d++) {
    if (yy * 10000 + mm * 100 + d > endKey) { break; }
    var dow = new Date(Date.UTC(yy, mm - 1, d)).getUTCDay();
    if (dow >= 1 && dow <= maxDow) { hariKerja++; }
  }

  var rows = queryReports_(user, { dateFrom: from, dateTo: to, divisi: divF, userId: userF });

  // Daftar karyawan aktif (agar yang belum pernah lapor tetap muncul)
  var kmap = {};
  var users = readSheet_('Users').rows;
  for (var u = 0; u < users.length; u++) {
    var us = users[u];
    if (us[3] !== 'Karyawan' || us[7] !== 'Aktif') { continue; }
    if (divF && us[5] !== divF) { continue; }
    if (userF && us[0] !== userF) { continue; }
    kmap[us[0]] = newK_(us[0], us[4], us[5], us[6]);
  }

  var sum = { laporan: rows.length, pekerjaan: 0, selesai: 0, progres: 0, tunda: 0 };
  var kendala = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i].r;
    var its = rows[i].its;
    var k = kmap[r[2]];
    if (!k) { k = kmap[r[2]] = newK_(r[2], r[3], r[4], r[5]); }
    k.hariLapor++;
    for (var j = 0; j < its.length; j++) {
      var st = its[j][4];
      k.pekerjaan++; sum.pekerjaan++;
      if (st === 'Selesai') { k.selesai++; sum.selesai++; } else if (st === 'Progres') { k.progres++; sum.progres++; } else if (st === 'Tunda') { k.tunda++; sum.tunda++; }
      if (its[j][5]) {
        k.kendala++;
        if (kendala.length < 30) {
          kendala.push({ tanggal: r[1], nama: r[3], divisi: r[4], deskripsi: its[j][3], status: st, kendala: its[j][5], solusi: its[j][6] });
        }
      }
    }
  }

  var karyawan = [];
  var totLapor = 0;
  Object.keys(kmap).forEach(function (key) {
    var x = kmap[key];
    x.pctSelesai = x.pekerjaan ? Math.round(x.selesai * 100 / x.pekerjaan) : 0;
    x.pctLapor = hariKerja ? Math.min(100, Math.round(x.hariLapor * 100 / hariKerja)) : 0;
    totLapor += x.hariLapor;
    karyawan.push(x);
  });
  karyawan.sort(function (a, b) {
    if (a.divisi !== b.divisi) { return a.divisi < b.divisi ? -1 : 1; }
    return a.nama.toLowerCase() < b.nama.toLowerCase() ? -1 : 1;
  });

  // Agregasi per divisi
  var perDivisi = [];
  for (var dI = 0; dI < DIVISI_LIST.length; dI++) {
    var dv = { divisi: DIVISI_LIST[dI], karyawan: 0, hariLapor: 0, pekerjaan: 0, selesai: 0, progres: 0, tunda: 0, kendala: 0, pctSelesai: 0, pctLapor: 0 };
    for (var q = 0; q < karyawan.length; q++) {
      if (karyawan[q].divisi !== dv.divisi) { continue; }
      dv.karyawan++; dv.hariLapor += karyawan[q].hariLapor; dv.pekerjaan += karyawan[q].pekerjaan;
      dv.selesai += karyawan[q].selesai; dv.progres += karyawan[q].progres; dv.tunda += karyawan[q].tunda; dv.kendala += karyawan[q].kendala;
    }
    if (!dv.karyawan) { continue; }
    dv.pctSelesai = dv.pekerjaan ? Math.round(dv.selesai * 100 / dv.pekerjaan) : 0;
    dv.pctLapor = hariKerja ? Math.min(100, Math.round(dv.hariLapor * 100 / (hariKerja * dv.karyawan))) : 0;
    perDivisi.push(dv);
  }

  sum.karyawan = karyawan.length;
  sum.pctSelesai = sum.pekerjaan ? Math.round(sum.selesai * 100 / sum.pekerjaan) : 0;
  sum.pctLapor = (hariKerja && karyawan.length) ? Math.min(100, Math.round(totLapor * 100 / (hariKerja * karyawan.length))) : 0;

  var out = { ok: true, bulan: bulan, from: from, to: to, hariKerja: hariKerja, divisi: divF, summary: sum, perDivisi: perDivisi, karyawan: karyawan, kendala: kendala };
  if (userF) {
    out.laporan = [];
    for (var z = 0; z < rows.length && z < 62; z++) { out.laporan.push(reportOut_(rows[z].r, rows[z].its)); }
  }
  return out;
}

// ---------------------------------------------------------------------
// Lampiran (banyak file per pekerjaan) & foto
// ---------------------------------------------------------------------
function filesOut_(link, nama) {
  var urls = String(link || '').split('|');
  var names = String(nama || '').split('|');
  var out = [];
  for (var i = 0; i < urls.length; i++) {
    var u = urls[i].trim();
    if (u.indexOf('https://') !== 0) { continue; }
    out.push({ url: u, name: String(names[i] || ('Lampiran ' + (out.length + 1))).trim() });
  }
  return out;
}

function driveId_(url) {
  var m = String(url).match(/\/d\/([A-Za-z0-9_-]+)/) || String(url).match(/[?&]id=([A-Za-z0-9_-]+)/);
  return m ? m[1] : '';
}

// Simpan / hapus foto profil (data URI JPEG kecil disimpan di kolom Foto_Profil)
function apiProfilePhoto_(user, p) {
  var targetId = user.ID_User;
  if (p.userId && String(p.userId) !== user.ID_User) {
    if (user.Role !== 'Admin') { throw new Error('Anda hanya dapat mengubah foto profil sendiri.'); }
    targetId = String(p.userId);
  }
  var data = '';
  if (!p.remove) {
    data = String(p.data || '');
    if (!/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(data) || data.length > 45000) {
      throw new Error('Foto tidak valid atau terlalu besar.');
    }
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var r = readSheet_('Users');
    for (var i = 0; i < r.rows.length; i++) {
      if (r.rows[i][0] === targetId) {
        r.sh.getRange(i + 2, 11).setValue(data);
        SpreadsheetApp.flush();
        return { ok: true, foto: data };
      }
    }
  } finally { lock.releaseLock(); }
  throw new Error('Pengguna tidak ditemukan.');
}

// Ambil foto pekerjaan (dan foto profil pemilik) sebagai data URI untuk dicetak
function apiFotoData_(user, p) {
  var id = String(p.id || '');
  var tgl = String(p.tanggal || '');
  var all = queryReports_(user, { dateFrom: tgl, dateTo: tgl });
  var found = null;
  for (var i = 0; i < all.length; i++) {
    if (all[i].r[0] === id && all[i].r[1] === tgl) { found = all[i]; break; }
  }
  if (!found) { throw new Error('Laporan tidak ditemukan.'); }
  var fotos = {};
  var count = 0;
  for (var j = 0; j < found.its.length; j++) {
    var files = filesOut_(found.its[j][7], found.its[j][8]);
    for (var k = 0; k < files.length; k++) {
      if (count >= 15) { break; }
      if (!/\.(jpe?g|png|webp|gif)$/i.test(files[k].name)) { continue; }
      var fid = driveId_(files[k].url);
      if (!fid) { continue; }
      try {
        var blob = DriveApp.getFileById(fid).getBlob();
        var ct = blob.getContentType();
        var bytes = blob.getBytes();
        if (String(ct).indexOf('image/') !== 0 || bytes.length > 2000000) { continue; }
        fotos[files[k].url] = 'data:' + ct + ';base64,' + Utilities.base64Encode(bytes);
        count++;
      } catch (e) { /* file hilang atau tidak bisa diakses, lewati */ }
    }
  }
  var owner = findUserById_(found.r[2]);
  return { ok: true, fotos: fotos, profil: owner ? (owner.Foto_Profil || '') : '' };
}

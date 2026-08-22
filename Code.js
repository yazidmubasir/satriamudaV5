/**
 * ============================================================
 *  WEBAPP BUILDER SEKOLAH - GOOGLE APPS SCRIPT
 * ============================================================
 * Cara pakai singkat:
 * 1. Buat 1 Google Spreadsheet baru -> jadikan "Spreadsheet Master".
 * 2. Extensions > Apps Script. Hapus isi default, paste file ini sebagai Code.gs
 *    dan file Index.html sebagai Index.html (HTML file baru bernama "Index").
 * 3. Jalankan fungsi bootstrapSystem sekali dari editor (pilih fungsi
 *    "runBootstrapFromEditor" lalu klik Run) UNTUK MEMBUAT SHEET-SHEET AWAL
 *    + akun owner default (username: owner / password: owner123).
 *    Bisa juga dilakukan lewat tombol "Bootstrap" di halaman login pertama kali.
 * 4. Deploy > New deployment > Web app.
 *    - Execute as: Me
 *    - Who has access: Anyone  (sesuai poin 11: "me" dan "anyone")
 * 5. Login sebagai owner, segera ganti password lewat menu Ganti Password.
 * ============================================================
 */

// ================= KONFIGURASI DASAR =================
const SHEET_NAMES = {
  USERS: 'master_users',
  GURU: 'master_guru',
  KARYAWAN: 'master_karyawan',
  KELAS: 'master_kelas',
  MAPEL: 'master_mapel',
  MENU: 'master_menu',
  MODUL: 'master_modul',
  MODUL_KELAS: 'master_modul_kelas'
};

const HEADERS = {
  master_users: ['username', 'password_hash', 'role', 'nama_lengkap', 'spreadsheet_id', 'status', 'created_at'],
  master_guru: ['nip', 'nama', 'mapel', 'no_hp', 'alamat', 'status'],
  master_karyawan: ['nip', 'nama', 'jabatan', 'no_hp', 'alamat', 'status'],
  master_kelas: ['kode_kelas', 'nama_kelas', 'wali_kelas', 'spreadsheet_id', 'admin_kelas_username', 'status'],
  master_mapel: ['kode_mapel', 'nama_mapel', 'guru_pengampu', 'status'],
  master_menu: ['menu_id', 'nama_menu', 'icon', 'urutan', 'role_akses', 'target_type', 'sheet_name', 'status'],
  master_modul: ['modul_id', 'nama_modul', 'keterangan', 'template_json', 'status'],
  master_modul_kelas: ['kode_kelas', 'modul_id', 'status_aktif', 'tanggal_aktif']
};

// Primary key per sheet master (dipakai untuk update/delete generik)
const PK = {
  master_users: 'username',
  master_guru: 'nip',
  master_karyawan: 'nip',
  master_kelas: 'kode_kelas',
  master_mapel: 'kode_mapel',
  master_menu: 'menu_id',
  master_modul: 'modul_id'
};

const SISWA_HEADERS = ['nis', 'nama', 'username', 'password_hash', 'jenis_kelamin', 'status'];
const ROLES = ['owner', 'guru', 'karyawan', 'admin_kelas'];

// ================= ENTRY POINT WEB APP =================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistem Sekolah - Web App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getMasterSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ================= UTIL: HASH PASSWORD (SHA-256) =================
function sha256(text) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return raw.map(function (byte) {
    var v = (byte < 0) ? byte + 256 : byte;
    var hex = v.toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

// ================= UTIL: SHEET HELPERS =================
function ensureSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function sheetToObjects(sh) {
  var data = sh.getDataRange().getValues();
  if (data.length < 1) return [];
  var headers = data[0];
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = { __row: i + 1 };
    var empty = true;
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
      if (row[j] !== '' && row[j] !== null && row[j] !== undefined) empty = false;
    }
    if (!empty) out.push(obj);
  }
  return out;
}

// google.script.run tidak dapat mengirim objek Date langsung ke browser.
// Ubah nilai tanggal menjadi teks sebelum data dikembalikan ke frontend.
function makeClientSafeRows(rows) {
  return rows.map(function (row) {
    var safe = {};
    Object.keys(row).forEach(function (key) {
      var value = row[key];
      safe[key] = value instanceof Date ? value.toISOString() : value;
    });
    return safe;
  });
}

// ================= BOOTSTRAP (poin 2) =================
function runBootstrapFromEditor() {
  // Jalankan manual dari editor Apps Script (tanpa token) untuk setup pertama kali.
  return bootstrapSystem('');
}

function bootstrapSystem(token) {
  var ss = getMasterSS();
  var usersSh = ss.getSheetByName(SHEET_NAMES.USERS);
  var firstRun = !usersSh || usersSh.getLastRow() <= 1;
  if (!firstRun) {
    requireOwner(token); // setelah pertama kali, hanya owner yg boleh re-run bootstrap
  }

  Object.keys(HEADERS).forEach(function (name) {
    ensureSheet(ss, name, HEADERS[name]);
  });

  var uSh = ss.getSheetByName(SHEET_NAMES.USERS);
  if (uSh.getLastRow() <= 1) {
    uSh.appendRow(['owner', sha256('owner123'), 'owner', 'Pemilik Sistem', '', 'aktif', new Date()]);
  }

  seedDefaultModules();

  return { message: 'Bootstrap selesai. Sheet master siap. Login default: owner / owner123 (segera ganti password setelah login).' };
}

function seedDefaultModules() {
  var ss = getMasterSS();
  var sh = ensureSheet(ss, SHEET_NAMES.MODUL, HEADERS.master_modul);
  if (sh.getLastRow() > 1) return;

  var modules = [
    {
      modul_id: 'MODUL_7KAIH',
      nama_modul: 'Modul 7 Kebiasaan Anak Indonesia Hebat',
      keterangan: 'Mencatat pelaksanaan 7 kebiasaan harian siswa (bangun pagi, beribadah, olahraga, dst).',
      template_json: JSON.stringify([{
        sheet_name: 'TRX_7KAIH',
        columns: [
          { name: 'tanggal', note: 'Format: YYYY-MM-DD, wajib diisi' },
          { name: 'nis', note: 'Nomor Induk Siswa, harus sama dengan sheet siswa' },
          { name: 'nama_siswa', note: 'Nama siswa (opsional, memudahkan pembacaan)' },
          { name: 'bangun_pagi', note: 'Isi: Ya / Tidak' },
          { name: 'beribadah', note: 'Isi: Ya / Tidak' },
          { name: 'berolahraga', note: 'Isi: Ya / Tidak' },
          { name: 'makan_sehat', note: 'Isi: Ya / Tidak' },
          { name: 'gemar_belajar', note: 'Isi: Ya / Tidak' },
          { name: 'bermasyarakat', note: 'Isi: Ya / Tidak' },
          { name: 'tidur_cepat', note: 'Isi: Ya / Tidak' },
          { name: 'keterangan', note: 'Catatan tambahan, boleh dikosongkan' }
        ]
      }])
      ,
      status: 'aktif'
    },
    {
      modul_id: 'MODUL_BELAJAR_MANDIRI',
      nama_modul: 'Modul Belajar Mandiri',
      keterangan: 'Mencatat aktivitas belajar mandiri siswa di luar jam sekolah.',
      template_json: JSON.stringify([{
        sheet_name: 'TRX_BELAJAR_MANDIRI',
        columns: [
          { name: 'tanggal', note: 'Format: YYYY-MM-DD' },
          { name: 'nis', note: 'Nomor Induk Siswa' },
          { name: 'mapel', note: 'Nama mata pelajaran yang dipelajari' },
          { name: 'durasi_menit', note: 'Angka, lama belajar dalam menit' },
          { name: 'ringkasan', note: 'Ringkasan singkat materi yang dipelajari' },
          { name: 'paraf_ortu', note: 'Isi: Sudah / Belum' }
        ]
      }]),
      status: 'aktif'
    },
    {
      modul_id: 'MODUL_LITERASI',
      nama_modul: 'Modul Kegiatan Literasi',
      keterangan: 'Mencatat kegiatan membaca / literasi harian siswa.',
      template_json: JSON.stringify([{
        sheet_name: 'TRX_LITERASI',
        columns: [
          { name: 'tanggal', note: 'Format: YYYY-MM-DD' },
          { name: 'nis', note: 'Nomor Induk Siswa' },
          { name: 'judul_buku', note: 'Judul buku / bahan bacaan' },
          { name: 'halaman_dibaca', note: 'Angka, jumlah halaman yang dibaca' },
          { name: 'resume', note: 'Resume singkat isi bacaan' }
        ]
      }]),
      status: 'aktif'
    }
  ];

  modules.forEach(function (m) {
    sh.appendRow([m.modul_id, m.nama_modul, m.keterangan, m.template_json, m.status]);
  });
}

// ================= AUTH & SESSION (poin 9, 11) =================
function login(username, password) {
  username = (username || '').trim();
  var hash = sha256(password);
  var ss = getMasterSS();
  var usersSh = ensureSheet(ss, SHEET_NAMES.USERS, HEADERS.master_users);
  var users = sheetToObjects(usersSh);

  var found = users.find(function (u) {
    return u.username === username && u.password_hash === hash && u.status !== 'nonaktif';
  });

  if (found) {
    var kodeKelas = '';
    if (found.role === 'admin_kelas' && found.spreadsheet_id) {
      var kelasList = sheetToObjects(ensureSheet(ss, SHEET_NAMES.KELAS, HEADERS.master_kelas));
      var kl = kelasList.find(function (k) { return k.spreadsheet_id === found.spreadsheet_id; });
      if (kl) kodeKelas = kl.kode_kelas;
    }
    return createSession(found.username, found.role, found.spreadsheet_id || '', found.nama_lengkap || '', kodeKelas);
  }

  // Siswa: kredensial ada di sheet 'siswa' pada spreadsheet kelas masing-masing (poin 10)
  var kelasSh = ensureSheet(ss, SHEET_NAMES.KELAS, HEADERS.master_kelas);
  var kelasList2 = sheetToObjects(kelasSh);
  for (var i = 0; i < kelasList2.length; i++) {
    var k = kelasList2[i];
    if (!k.spreadsheet_id) continue;
    try {
      var kss = SpreadsheetApp.openById(k.spreadsheet_id);
      var siswaSh = kss.getSheetByName('siswa');
      if (!siswaSh) continue;
      var siswaList = sheetToObjects(siswaSh);
      var s = siswaList.find(function (x) { return x.username === username && x.password_hash === hash && x.status !== 'nonaktif'; });
      if (s) {
        return createSession(s.username, 'siswa', k.spreadsheet_id, s.nama || '', k.kode_kelas);
      }
    } catch (err) {
      continue;
    }
  }
  throw new Error('Username atau password salah');
}

function createSession(username, role, spreadsheetId, nama, kodeKelas) {
  var token = Utilities.getUuid();
  var payload = {
    username: username,
    role: role,
    spreadsheet_id: spreadsheetId || '',
    nama: nama || '',
    kode_kelas: kodeKelas || ''
  };
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(payload), 21600); // 6 jam
  return { token: token, user: payload };
}

function getSession(token) {
  if (!token) throw new Error('Sesi tidak ditemukan, silakan login ulang');
  var raw = CacheService.getScriptCache().get('sess_' + token);
  if (!raw) throw new Error('Sesi habis, silakan login ulang');
  return JSON.parse(raw);
}

function whoAmI(token) {
  return getSession(token);
}

function logout(token) {
  CacheService.getScriptCache().remove('sess_' + token);
  return true;
}

function requireAuth(token) {
  return getSession(token);
}

function requireOwner(token) {
  var u = getSession(token);
  if (u.role !== 'owner') throw new Error('Akses ditolak: khusus owner');
  return u;
}

function requireRole(token, roles) {
  var u = getSession(token);
  if (roles.indexOf(u.role) === -1) throw new Error('Akses ditolak');
  return u;
}

function changePassword(token, oldPassword, newPassword) {
  var u = requireAuth(token);
  var oldHash = sha256(oldPassword);
  var newHash = sha256(newPassword);

  if (u.role === 'siswa') {
    var kss = SpreadsheetApp.openById(u.spreadsheet_id);
    var sh = kss.getSheetByName('siswa');
    if (!sh) throw new Error('Sheet siswa tidak ditemukan');
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('username')] === u.username) {
        if (data[i][headers.indexOf('password_hash')] !== oldHash) throw new Error('Password lama salah');
        sh.getRange(i + 1, headers.indexOf('password_hash') + 1).setValue(newHash);
        return { message: 'Password berhasil diubah' };
      }
    }
    throw new Error('User tidak ditemukan');
  }

  var ss = getMasterSS();
  var msh = ensureSheet(ss, SHEET_NAMES.USERS, HEADERS.master_users);
  var mheaders = HEADERS.master_users;
  var mdata = msh.getDataRange().getValues();
  for (var j = 1; j < mdata.length; j++) {
    if (mdata[j][mheaders.indexOf('username')] === u.username) {
      if (mdata[j][mheaders.indexOf('password_hash')] !== oldHash) throw new Error('Password lama salah');
      msh.getRange(j + 1, mheaders.indexOf('password_hash') + 1).setValue(newHash);
      return { message: 'Password berhasil diubah' };
    }
  }
  throw new Error('User tidak ditemukan');
}

// ================= MASTER DATA CRUD (poin 1, 5) - HANYA OWNER YANG EDIT =================
function getMasterList(token, sheetName) {
  requireAuth(token); // semua role login boleh baca (untuk kebutuhan menu), edit tetap dibatasi
  var ss = getMasterSS();
  var sh = ensureSheet(ss, sheetName, HEADERS[sheetName]);
  var list = makeClientSafeRows(sheetToObjects(sh));
  if (sheetName === SHEET_NAMES.USERS) {
    list.forEach(function (u) { delete u.password_hash; });
  }
  return list;
}

function addMasterRow(token, sheetName, obj) {
  requireOwner(token);
  var ss = getMasterSS();
  var sh = ensureSheet(ss, sheetName, HEADERS[sheetName]);
  var headers = HEADERS[sheetName];
  var pk = PK[sheetName];

  if (pk) {
    var existing = sheetToObjects(sh);
    if (existing.some(function (r) { return String(r[pk]) === String(obj[pk]); })) {
      throw new Error('Data dengan ' + pk + ' tersebut sudah ada');
    }
  }
  if (sheetName === SHEET_NAMES.USERS) {
    obj.password_hash = sha256(obj.password || '123456');
    obj.created_at = new Date();
  }
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sh.appendRow(row);
  return { message: 'Data ditambahkan' };
}

function updateMasterRow(token, sheetName, pkValue, obj) {
  requireOwner(token);
  return updateMasterRowInternal(sheetName, pkValue, obj);
}

function updateMasterRowInternal(sheetName, pkValue, obj) {
  var ss = getMasterSS();
  var sh = ensureSheet(ss, sheetName, HEADERS[sheetName]);
  var headers = HEADERS[sheetName];
  var pk = PK[sheetName];
  var data = sh.getDataRange().getValues();
  var pkIdx = headers.indexOf(pk);

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][pkIdx]) === String(pkValue)) {
      if (sheetName === SHEET_NAMES.USERS) {
        var phIdx = headers.indexOf('password_hash');
        if (obj.password) { obj.password_hash = sha256(obj.password); }
        else { obj.password_hash = data[i][phIdx]; }
      }
      var row = headers.map(function (h, idx) { return obj[h] !== undefined ? obj[h] : data[i][idx]; });
      sh.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      return { message: 'Data diperbarui' };
    }
  }
  throw new Error('Data tidak ditemukan');
}

function deleteMasterRow(token, sheetName, pkValue) {
  requireOwner(token);
  var ss = getMasterSS();
  var sh = ensureSheet(ss, sheetName, HEADERS[sheetName]);
  var headers = HEADERS[sheetName];
  var pk = PK[sheetName];
  var data = sh.getDataRange().getValues();
  var pkIdx = headers.indexOf(pk);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][pkIdx]) === String(pkValue)) {
      sh.deleteRow(i + 1);
      return { message: 'Data dihapus' };
    }
  }
  throw new Error('Data tidak ditemukan');
}

// ================= ASSIGN ADMIN KELAS (poin 7, 8, 18, 19) - HANYA OWNER =================
function assignAdminKelas(token, kodeKelas, username, spreadsheetId) {
  requireOwner(token);
  updateMasterRowInternal(SHEET_NAMES.KELAS, kodeKelas, { spreadsheet_id: spreadsheetId, admin_kelas_username: username });
  updateMasterRowInternal(SHEET_NAMES.USERS, username, { role: 'admin_kelas', spreadsheet_id: spreadsheetId });
  return { message: 'Admin kelas berhasil diatur untuk kelas ' + kodeKelas };
}

function getUsersForAssign(token) {
  requireOwner(token);
  var list = getMasterList(token, SHEET_NAMES.USERS);
  return list.filter(function (u) { return u.role !== 'owner'; });
}

// ================= SHEET & KOLOM BUILDER GENERIK (poin 2, 14, 15, 17, 18) =================
function checkSpreadsheetAccess(u, spreadsheetId) {
  if (u.role === 'owner') return true;
  if ((u.role === 'admin_kelas' || u.role === 'siswa') && u.spreadsheet_id === spreadsheetId) return true;
  throw new Error('Akses ditolak ke spreadsheet ini');
}

function listSpreadsheetSheets(token, spreadsheetId) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  var ss = SpreadsheetApp.openById(spreadsheetId);
  return ss.getSheets().map(function (s) { return s.getName(); });
}

function createSheetWithHeaders(token, spreadsheetId, sheetName, columns) {
  // columns: array of {name, note} -- note hanya dipakai di frontend sbg petunjuk
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  if (u.role === 'siswa') throw new Error('Siswa tidak berwenang membuat sheet');
  var ss = SpreadsheetApp.openById(spreadsheetId);
  if (ss.getSheetByName(sheetName)) throw new Error('Sheet "' + sheetName + '" sudah ada');
  var sh = ss.insertSheet(sheetName);
  var colNames = columns.map(function (c) { return c.name; });
  sh.appendRow(colNames);
  sh.setFrozenRows(1);
  return { message: 'Sheet ' + sheetName + ' berhasil dibuat' };
}

function getSheetHeadersFn(token, spreadsheetId, sheetName) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return [];
  var lastCol = sh.getLastColumn();
  if (lastCol === 0) return [];
  return sh.getRange(1, 1, 1, lastCol).getValues()[0];
}

function addColumnToSheetFn(token, spreadsheetId, sheetName, columnName) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  if (u.role === 'siswa') throw new Error('Siswa tidak berwenang');
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan');
  var lastCol = sh.getLastColumn();
  var headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (headers.indexOf(columnName) > -1) throw new Error('Kolom sudah ada');
  sh.getRange(1, lastCol + 1).setValue(columnName);
  return { message: 'Kolom ' + columnName + ' ditambahkan' };
}

function deleteColumnFromSheetFn(token, spreadsheetId, sheetName, columnName) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  if (u.role === 'siswa') throw new Error('Siswa tidak berwenang');
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan');
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headers.indexOf(columnName);
  if (idx === -1) throw new Error('Kolom tidak ditemukan');
  sh.deleteColumn(idx + 1);
  return { message: 'Kolom ' + columnName + ' dihapus' };
}

// ================= DATA BARIS GENERIK UNTUK SHEET KEGIATAN KELAS =================
function canWriteSheet(u, spreadsheetId, sheetName) {
  if (u.role === 'owner') return true;
  if (u.role === 'admin_kelas' && u.spreadsheet_id === spreadsheetId) return true;
  // guru / karyawan: diizinkan hanya jika ada menu aktif yang mengarah ke sheet ini untuk role mereka
  var menus = sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.MENU, HEADERS.master_menu));
  return menus.some(function (m) {
    return m.sheet_name === sheetName && m.status === 'aktif' &&
      (m.role_akses || '').split(',').map(function (s) { return s.trim(); }).indexOf(u.role) > -1;
  });
}

function getSheetDataFn(token, spreadsheetId, sheetName) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan');
  var list = sheetToObjects(sh);
  if (sheetName === 'siswa') list.forEach(function (r) { delete r.password_hash; });
  return list;
}

function addRowToSheet(token, spreadsheetId, sheetName, obj) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  if (!canWriteSheet(u, spreadsheetId, sheetName)) throw new Error('Tidak berwenang menulis ke sheet ini');
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet tidak ditemukan');
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sh.appendRow(row);
  return { message: 'Baris ditambahkan' };
}

function updateRowInSheet(token, spreadsheetId, sheetName, rowNumber, obj) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  if (!canWriteSheet(u, spreadsheetId, sheetName)) throw new Error('Tidak berwenang');
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(sheetName);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var current = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  var row = headers.map(function (h, idx) { return obj[h] !== undefined ? obj[h] : current[idx]; });
  sh.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
  return { message: 'Baris diperbarui' };
}

function deleteRowFromSheet(token, spreadsheetId, sheetName, rowNumber) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  if (!canWriteSheet(u, spreadsheetId, sheetName)) throw new Error('Tidak berwenang');
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(sheetName);
  sh.deleteRow(rowNumber);
  return { message: 'Baris dihapus' };
}

// ================= SISWA CRUD OLEH ADMIN_KELAS (poin 10) =================
function getSiswaList(token, spreadsheetId) {
  var u = requireRole(token, ['admin_kelas', 'owner']);
  var sid = (u.role === 'admin_kelas') ? u.spreadsheet_id : spreadsheetId;
  if (!sid) throw new Error('spreadsheet_id diperlukan');
  checkSpreadsheetAccess(u, sid);
  var ss = SpreadsheetApp.openById(sid);
  var sh = ensureSheet(ss, 'siswa', SISWA_HEADERS);
  var list = sheetToObjects(sh);
  list.forEach(function (r) { delete r.password_hash; });
  return list;
}

function addSiswa(token, spreadsheetId, obj) {
  var u = requireRole(token, ['admin_kelas', 'owner']);
  var sid = (u.role === 'admin_kelas') ? u.spreadsheet_id : spreadsheetId;
  checkSpreadsheetAccess(u, sid);
  var ss = SpreadsheetApp.openById(sid);
  var sh = ensureSheet(ss, 'siswa', SISWA_HEADERS);
  var existing = sheetToObjects(sh);
  if (existing.some(function (r) { return r.username === obj.username; })) throw new Error('Username sudah dipakai');
  obj.password_hash = sha256(obj.password || '123456');
  var row = SISWA_HEADERS.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sh.appendRow(row);
  return { message: 'Siswa ditambahkan' };
}

function updateSiswa(token, spreadsheetId, rowNumber, obj) {
  var u = requireRole(token, ['admin_kelas', 'owner']);
  var sid = (u.role === 'admin_kelas') ? u.spreadsheet_id : spreadsheetId;
  checkSpreadsheetAccess(u, sid);
  var ss = SpreadsheetApp.openById(sid);
  var sh = ss.getSheetByName('siswa');
  if (!sh) throw new Error('Sheet siswa tidak ditemukan');
  var current = sh.getRange(rowNumber, 1, 1, SISWA_HEADERS.length).getValues()[0];
  if (obj.password) { obj.password_hash = sha256(obj.password); }
  var row = SISWA_HEADERS.map(function (h, idx) { return obj[h] !== undefined ? obj[h] : current[idx]; });
  sh.getRange(rowNumber, 1, 1, SISWA_HEADERS.length).setValues([row]);
  return { message: 'Data siswa diperbarui' };
}

function deleteSiswa(token, spreadsheetId, rowNumber) {
  var u = requireRole(token, ['admin_kelas', 'owner']);
  var sid = (u.role === 'admin_kelas') ? u.spreadsheet_id : spreadsheetId;
  checkSpreadsheetAccess(u, sid);
  var ss = SpreadsheetApp.openById(sid);
  var sh = ss.getSheetByName('siswa');
  sh.deleteRow(rowNumber);
  return { message: 'Siswa dihapus' };
}

// ================= MODUL KEGIATAN (poin 19, 20, 21, 22) =================
function listModulesAll(token) {
  requireOwner(token);
  return sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.MODUL, HEADERS.master_modul));
}

function listModulesForActivation(token) {
  requireRole(token, ['admin_kelas', 'owner']);
  var list = sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.MODUL, HEADERS.master_modul));
  return list.filter(function (m) { return m.status === 'aktif'; });
}

function saveModul(token, modulId, namaModul, keterangan, templateArr, isNew) {
  requireOwner(token);
  var templateJson = JSON.stringify(templateArr);
  if (isNew) {
    return addMasterRow(token, SHEET_NAMES.MODUL, {
      modul_id: modulId, nama_modul: namaModul, keterangan: keterangan,
      template_json: templateJson, status: 'aktif'
    });
  } else {
    return updateMasterRow(token, SHEET_NAMES.MODUL, modulId, {
      nama_modul: namaModul, keterangan: keterangan, template_json: templateJson
    });
  }
}

function getActiveModulesForKelas(token, kodeKelas) {
  requireAuth(token);
  var list = sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.MODUL_KELAS, HEADERS.master_modul_kelas));
  return list.filter(function (r) { return r.kode_kelas === kodeKelas; });
}

function activateModulForKelas(token, kodeKelas, modulId) {
  var u = requireRole(token, ['admin_kelas', 'owner']);
  var ss = getMasterSS();
  var kelasList = sheetToObjects(ensureSheet(ss, SHEET_NAMES.KELAS, HEADERS.master_kelas));
  var kelas = kelasList.find(function (k) { return k.kode_kelas === kodeKelas; });
  if (!kelas) throw new Error('Kelas tidak ditemukan');
  if (u.role === 'admin_kelas' && u.spreadsheet_id !== kelas.spreadsheet_id) {
    throw new Error('Anda tidak berwenang untuk kelas ini');
  }
  if (!kelas.spreadsheet_id) throw new Error('Kelas ini belum memiliki spreadsheet_id, hubungi owner');

  var modulList = sheetToObjects(ensureSheet(ss, SHEET_NAMES.MODUL, HEADERS.master_modul));
  var modul = modulList.find(function (m) { return m.modul_id === modulId; });
  if (!modul) throw new Error('Modul tidak ditemukan');

  var template = JSON.parse(modul.template_json);
  var kss = SpreadsheetApp.openById(kelas.spreadsheet_id);
  template.forEach(function (t) {
    var sh = kss.getSheetByName(t.sheet_name);
    var colNames = t.columns.map(function (c) { return c.name; });
    if (!sh) {
      sh = kss.insertSheet(t.sheet_name);
      sh.appendRow(colNames);
      sh.setFrozenRows(1);
    } else {
      var lastCol = sh.getLastColumn();
      var existingHeaders = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
      colNames.forEach(function (c) {
        if (existingHeaders.indexOf(c) === -1) {
          sh.getRange(1, sh.getLastColumn() + 1).setValue(c);
        }
      });
    }
  });

  var mkSh = ensureSheet(ss, SHEET_NAMES.MODUL_KELAS, HEADERS.master_modul_kelas);
  var mkList = sheetToObjects(mkSh);
  var idx = -1;
  for (var i = 0; i < mkList.length; i++) {
    if (mkList[i].kode_kelas === kodeKelas && mkList[i].modul_id === modulId) { idx = i; break; }
  }
  if (idx === -1) {
    mkSh.appendRow([kodeKelas, modulId, 'aktif', new Date()]);
  } else {
    mkSh.getRange(mkList[idx].__row, 3, 1, 2).setValues([['aktif', new Date()]]);
  }
  return { message: 'Modul "' + modul.nama_modul + '" berhasil diaktifkan untuk kelas ' + kodeKelas };
}

// ================= MENU CRUD (poin 12, 13, 16) - HANYA OWNER YANG MENAMBAH =================
function getMenusForUser(token) {
  var u = requireAuth(token);
  var menus = sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.MENU, HEADERS.master_menu));
  return menus.filter(function (m) {
    if (m.status !== 'aktif') return false;
    var roles = (m.role_akses || '').split(',').map(function (s) { return s.trim(); });
    return roles.indexOf(u.role) > -1 || roles.indexOf('*') > -1;
  }).sort(function (a, b) { return (Number(a.urutan) || 0) - (Number(b.urutan) || 0); });
}

// ================= AGREGASI DATA SEMUA KELAS (poin 23) =================
function listAllKelasSpreadsheets(token) {
  requireOwner(token);
  return sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.KELAS, HEADERS.master_kelas));
}

function aggregateSheetAcrossKelas(token, sheetName) {
  requireOwner(token);
  var ss = getMasterSS();
  var kelasList = sheetToObjects(ensureSheet(ss, SHEET_NAMES.KELAS, HEADERS.master_kelas));
  var result = [];
  kelasList.forEach(function (k) {
    if (!k.spreadsheet_id) return;
    try {
      var kss = SpreadsheetApp.openById(k.spreadsheet_id);
      var sh = kss.getSheetByName(sheetName);
      if (!sh) return;
      var rows = sheetToObjects(sh);
      rows.forEach(function (r) { r.__kode_kelas = k.kode_kelas; r.__nama_kelas = k.nama_kelas; });
      result = result.concat(rows);
    } catch (e) { /* spreadsheet tidak bisa diakses, lewati */ }
  });
  return result;
}

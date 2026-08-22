/**
 * ============================================================
 * MASTER CONTROLLER BACKEND - CODE.GS
 * ============================================================
 */

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

// ================= ENTRY POINT =================
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

// ================= UTILITAS DATA & SERIALISASI =================
function sha256(text) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return raw.map(function (byte) {
    var v = (byte < 0) ? byte + 256 : byte;
    var hex = v.toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

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
  var tz = Session.getScriptTimeZone();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = { __row: i + 1 };
    var empty = true;
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      // FIX CRITICAL: Format Date ke String agar aman dikirim via google.script.run
      if (val instanceof Date) {
        val = Utilities.formatDate(val, tz, "yyyy-MM-dd HH:mm:ss");
      }
      obj[headers[j]] = (val !== null && val !== undefined) ? val : '';
      if (row[j] !== '' && row[j] !== null && row[j] !== undefined) empty = false;
    }
    if (!empty) out.push(obj);
  }
  return out;
}

// ================= BOOTSTRAP =================
function runBootstrapFromEditor() {
  return bootstrapSystem('');
}

function bootstrapSystem(token) {
  var ss = getMasterSS();
  var usersSh = ss.getSheetByName(SHEET_NAMES.USERS);
  var firstRun = !usersSh || usersSh.getLastRow() <= 1;
  if (!firstRun) {
    requireOwner(token);
  }

  Object.keys(HEADERS).forEach(function (name) {
    ensureSheet(ss, name, HEADERS[name]);
  });

  var uSh = ss.getSheetByName(SHEET_NAMES.USERS);
  if (uSh.getLastRow() <= 1) {
    uSh.appendRow(['owner', sha256('owner123'), 'owner', 'Pemilik Sistem', '', 'aktif', new Date()]);
  }

  seedDefaultModules();
  return { message: 'Bootstrap selesai. Sheet master siap. Login default: owner / owner123' };
}

function seedDefaultModules() {
  var ss = getMasterSS();
  var sh = ensureSheet(ss, SHEET_NAMES.MODUL, HEADERS.master_modul);
  if (sh.getLastRow() > 1) return;

  var modules = [
    {
      modul_id: 'MODUL_7KAIH',
      nama_modul: 'Modul 7 Kebiasaan Anak Indonesia Hebat',
      keterangan: 'Mencatat pelaksanaan 7 kebiasaan harian siswa.',
      template_json: JSON.stringify([{
        sheet_name: 'TRX_7KAIH',
        columns: [
          { name: 'tanggal', note: 'Format: YYYY-MM-DD' },
          { name: 'nis', note: 'Nomor Induk Siswa' },
          { name: 'nama_siswa', note: 'Nama lengkap siswa' },
          { name: 'bangun_pagi', note: 'Ya / Tidak' },
          { name: 'beribadah', note: 'Ya / Tidak' },
          { name: 'berolahraga', note: 'Ya / Tidak' },
          { name: 'makan_sehat', note: 'Ya / Tidak' },
          { name: 'gemar_belajar', note: 'Ya / Tidak' },
          { name: 'bermasyarakat', note: 'Ya / Tidak' },
          { name: 'tidur_cepat', note: 'Ya / Tidak' },
          { name: 'keterangan', note: 'Catatan harian' }
        ]
      }]),
      status: 'aktif'
    },
    {
      modul_id: 'MODUL_BELAJAR_MANDIRI',
      nama_modul: 'Modul Belajar Mandiri',
      keterangan: 'Mencatat aktivitas belajar mandiri di luar jam sekolah.',
      template_json: JSON.stringify([{
        sheet_name: 'TRX_BELAJAR_MANDIRI',
        columns: [
          { name: 'tanggal', note: 'Format: YYYY-MM-DD' },
          { name: 'nis', note: 'Nomor Induk Siswa' },
          { name: 'nama_siswa', note: 'Nama Siswa' },
          { name: 'mapel', note: 'Mata Pelajaran' },
          { name: 'durasi_menit', note: 'Durasi (Menit)' },
          { name: 'ringkasan', note: 'Ringkasan Materi' },
          { name: 'paraf_ortu', note: 'Ya / Tidak' }
        ]
      }]),
      status: 'aktif'
    }
  ];

  modules.forEach(function (m) {
    sh.appendRow([m.modul_id, m.nama_modul, m.keterangan, m.template_json, m.status]);
  });
}

// ================= AUTH & AUTENTIKASI =================
function login(username, password) {
  username = (username || '').trim();
  var hash = sha256(password);
  var ss = getMasterSS();
  var usersSh = ensureSheet(ss, SHEET_NAMES.USERS, HEADERS.master_users);
  var users = sheetToObjects(usersSh);

  var found = users.find(function (u) {
    return String(u.username) === username && u.password_hash === hash && u.status !== 'nonaktif';
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

  // Login Siswa dari Spreadsheet Kelas
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
      var s = siswaList.find(function (x) { return String(x.username) === username && x.password_hash === hash && x.status !== 'nonaktif'; });
      if (s) {
        return createSession(s.username, 'siswa', k.spreadsheet_id, s.nama || '', k.kode_kelas);
      }
    } catch (err) { continue; }
  }
  throw new Error('Username atau password salah');
}

function createSession(username, role, spreadsheetId, nama, kodeKelas) {
  var token = Utilities.getUuid();
  var payload = {
    username: String(username),
    role: role,
    spreadsheet_id: spreadsheetId || '',
    nama: nama || '',
    kode_kelas: kodeKelas || ''
  };
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(payload), 21600);
  return { token: token, user: payload };
}

function getSession(token) {
  if (!token) throw new Error('Sesi tidak ditemukan, silakan login ulang');
  var raw = CacheService.getScriptCache().get('sess_' + token);
  if (!raw) throw new Error('Sesi habis, silakan login ulang');
  return JSON.parse(raw);
}

function logout(token) {
  CacheService.getScriptCache().remove('sess_' + token);
  return true;
}

function requireAuth(token) { return getSession(token); }
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
      if (String(data[i][headers.indexOf('username')]) === String(u.username)) {
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
    if (String(mdata[j][mheaders.indexOf('username')]) === String(u.username)) {
      if (mdata[j][mheaders.indexOf('password_hash')] !== oldHash) throw new Error('Password lama salah');
      msh.getRange(j + 1, mheaders.indexOf('password_hash') + 1).setValue(newHash);
      return { message: 'Password berhasil diubah' };
    }
  }
  throw new Error('User tidak ditemukan');
}

// ================= MASTER CRUD DATA =================
function getMasterList(token, sheetName) {
  requireAuth(token);
  var ss = getMasterSS();
  var sh = ensureSheet(ss, sheetName, HEADERS[sheetName]);
  var list = sheetToObjects(sh);
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
    obj.created_at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  }
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sh.appendRow(row);
  return { message: 'Data berhasil ditambahkan' };
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
      return { message: 'Data berhasil diperbarui' };
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
      return { message: 'Data berhasil dihapus' };
    }
  }
  throw new Error('Data tidak ditemukan');
}

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
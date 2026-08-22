/**
 * ============================================================
 * MODULE & CLASS SPREADSHEET ENGINE BACKEND - MODULE.GS
 * ============================================================
 */

function checkSpreadsheetAccess(u, spreadsheetId) {
  spreadsheetId = String(spreadsheetId || '');
  if (!spreadsheetId) throw new Error('Spreadsheet tidak ditentukan');
  if (u.role === 'owner') return true;
  if (u.role !== 'admin_kelas' && u.role !== 'siswa') {
    throw new Error('Akses ditolak ke spreadsheet ini');
  }

  var kelas = getKelasForUser(u);
  if (!kelas || String(kelas.spreadsheet_id) !== spreadsheetId) {
    throw new Error('Akses ditolak: Anda hanya dapat mengakses spreadsheet kelas Anda');
  }
  // Selalu ambil identitas terbaru dari roster kelas, bukan data yang mungkin
  // sudah lama tersimpan di sesi browser.
  if (u.role === 'siswa') {
    var student = getStudentForSession(u);
    u.nis = student.nis || '';
    u.nama = student.nama || '';
  }
  return true;
}

function getKelasForUser(u) {
  if (!u || !u.kode_kelas || !u.spreadsheet_id) return null;
  var kelasList = sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.KELAS, HEADERS.master_kelas));
  return kelasList.find(function (k) {
    return String(k.kode_kelas) === String(u.kode_kelas) &&
      String(k.spreadsheet_id) === String(u.spreadsheet_id) &&
      (u.role !== 'admin_kelas' || String(k.admin_kelas_username || '') === String(u.username || '')) &&
      String(k.status || 'aktif') !== 'nonaktif';
  }) || null;
}

function requireAuthorizedKelas(u, kodeKelas) {
  var kelasList = sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.KELAS, HEADERS.master_kelas));
  var kelas = kelasList.find(function (k) { return String(k.kode_kelas) === String(kodeKelas); });
  if (!kelas) throw new Error('Kelas tidak ditemukan');
  if (String(kelas.status || 'aktif') === 'nonaktif') throw new Error('Kelas sedang nonaktif');
  if (u.role !== 'owner' && (!u.kode_kelas || String(u.kode_kelas) !== String(kodeKelas))) {
    throw new Error('Akses ditolak: data kelas lain tidak dapat diakses');
  }
  return kelas;
}

function getStudentRecord(spreadsheetId, identifier, fieldName) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName('siswa');
  if (!sh) return null;
  var list = sheetToObjects(sh);
  return list.find(function (s) {
    return String(s[fieldName] || '') === String(identifier || '') && String(s.status || 'aktif') !== 'nonaktif';
  }) || null;
}

function getStudentForSession(u) {
  if (u.role !== 'siswa') return null;
  var student = getStudentRecord(u.spreadsheet_id, u.username, 'username');
  if (!student) throw new Error('Data siswa tidak ditemukan pada kelas Anda');
  return student;
}

function applyStudentIdentity(u, spreadsheetId, headers, obj) {
  var nisIdx = headers.indexOf('nis');
  var nameIdx = headers.indexOf('nama_siswa');
  if (nisIdx === -1 && nameIdx === -1) return obj;

  var student;
  if (u.role === 'siswa') {
    student = getStudentForSession(u);
  } else if (obj.nis) {
    student = getStudentRecord(spreadsheetId, obj.nis, 'nis');
    if (!student) throw new Error('NIS tidak terdaftar pada kelas ini');
  } else if (obj.username) {
    student = getStudentRecord(spreadsheetId, obj.username, 'username');
  }

  if (student) {
    if (nisIdx > -1) obj.nis = student.nis;
    if (nameIdx > -1) obj.nama_siswa = student.nama;
  }
  return obj;
}

function isStudentRowOwned(u, row) {
  if (u.role !== 'siswa') return true;
  var identity = String(u.nis || u.username || '');
  return String(row.nis || '') === identity || String(row.username || '') === String(u.username || '');
}

function canWriteSheet(u, spreadsheetId, sheetName) {
  if (u.role === 'owner') return true;
  if (u.role === 'admin_kelas' && u.spreadsheet_id === spreadsheetId) return true;
  // Siswa boleh mengisi modul di kelasnya, tetapi tidak boleh mengubah roster siswa.
  if (u.role === 'siswa' && u.spreadsheet_id === spreadsheetId && sheetName !== 'siswa') return true;

  var menus = sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.MENU, HEADERS.master_menu));
  return menus.some(function (m) {
    return m.sheet_name === sheetName && m.status === 'aktif' &&
      (m.role_akses || '').split(',').map(function (s) { return s.trim(); }).indexOf(u.role) > -1;
  });
}

function listSpreadsheetSheets(token, spreadsheetId) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  var ss = SpreadsheetApp.openById(spreadsheetId);
  return ss.getSheets().map(function (s) { return s.getName(); });
}

function createSheetWithHeaders(token, spreadsheetId, sheetName, columns) {
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

// ================= CRUD BARIS MODUL / LEMBAR KEGIATAN =================
function getSheetDataFn(token, spreadsheetId, sheetName) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet ' + sheetName + ' tidak ditemukan');
  var list = sheetToObjects(sh);
  if (u.role === 'siswa') {
    list = list.filter(function (r) { return isStudentRowOwned(u, r); });
  }
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
  obj = applyStudentIdentity(u, spreadsheetId, headers, obj || {});
  
  var row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sh.appendRow(row);
  return { message: 'Data berhasil disimpan' };
}

function updateRowInSheet(token, spreadsheetId, sheetName, rowNumber, obj) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  if (!canWriteSheet(u, spreadsheetId, sheetName)) throw new Error('Tidak berwenang');
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(sheetName);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var current = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  var currentObj = {};
  headers.forEach(function (h, i) { currentObj[h] = current[i]; });
  if (!isStudentRowOwned(u, currentObj)) throw new Error('Siswa hanya dapat mengubah data miliknya sendiri');
  obj = applyStudentIdentity(u, spreadsheetId, headers, obj || {});
  var row = headers.map(function (h, idx) { return obj[h] !== undefined ? obj[h] : current[idx]; });
  sh.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
  return { message: 'Data berhasil diperbarui' };
}

function deleteRowFromSheet(token, spreadsheetId, sheetName, rowNumber) {
  var u = requireAuth(token);
  checkSpreadsheetAccess(u, spreadsheetId);
  if (!canWriteSheet(u, spreadsheetId, sheetName)) throw new Error('Tidak berwenang');
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheetByName(sheetName);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var current = sh.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  var currentObj = {};
  headers.forEach(function (h, i) { currentObj[h] = current[i]; });
  if (!isStudentRowOwned(u, currentObj)) throw new Error('Siswa hanya dapat menghapus data miliknya sendiri');
  sh.deleteRow(rowNumber);
  return { message: 'Data berhasil dihapus' };
}

// ================= KELOLA SISWA (OLEH ADMIN KELAS) =================
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
  if (existing.some(function (r) { return String(r.username) === String(obj.username); })) throw new Error('Username sudah dipakai');
  obj.password_hash = sha256(obj.password || '123456');
  var row = SISWA_HEADERS.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sh.appendRow(row);
  return { message: 'Siswa berhasil ditambahkan' };
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

// ================= MODUL & AKTIVASI =================
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
  var u = requireAuth(token);
  requireAuthorizedKelas(u, kodeKelas);
  var list = sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.MODUL_KELAS, HEADERS.master_modul_kelas));
  return list.filter(function (r) { return r.kode_kelas === kodeKelas; });
}

function activateModulForKelas(token, kodeKelas, modulId) {
  var u = requireRole(token, ['admin_kelas', 'owner']);
  var ss = getMasterSS();
  var kelas = requireAuthorizedKelas(u, kodeKelas);
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
    mkSh.appendRow([kodeKelas, modulId, 'aktif', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")]);
  } else {
    mkSh.getRange(mkList[idx].__row, 3, 1, 2).setValues([['aktif', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")]]);
  }
  return { message: 'Modul "' + modul.nama_modul + '" berhasil diaktifkan untuk kelas ' + kodeKelas };
}

function getMenusForUser(token) {
  var u = requireAuth(token);
  var menus = sheetToObjects(ensureSheet(getMasterSS(), SHEET_NAMES.MENU, HEADERS.master_menu));
  return menus.filter(function (m) {
    if (m.status !== 'aktif') return false;
    var roles = (m.role_akses || '').split(',').map(function (s) { return s.trim(); });
    return roles.indexOf(u.role) > -1 || roles.indexOf('*') > -1;
  }).sort(function (a, b) { return (Number(a.urutan) || 0) - (Number(b.urutan) || 0); });
}

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
    } catch (e) {}
  });
  return result;
}

/**
 * ============================================
 * APLIKASI KELULUSAN MADRASAH
 * Backend: Google Apps Script
 * Database: Google Spreadsheet
 * ============================================
 */

var CONFIG = {
  SHEET_SETTINGS: 'Settings',
  SHEET_STUDENTS: 'Students'
};

// ==========================================
// HANDLER UTAMA
// ==========================================

function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Cek Kelulusan Madrasah')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var result = routeApi(body.path, body.data || {});
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: 'Server error: ' + error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// LAYER KOMUNIKASI
// ==========================================

function callApi(path, data) { return routeApi(path, data || {}); }

function routeApi(path, data) {
  try {
    switch (path) {
      case '/api/health': return apiHealth();
      case '/api/login': return apiLogin(data);
      case '/api/settings': return apiGetSettings();
      case '/api/settings/update': return apiUpdateSettings(data);
      case '/api/students': return apiGetStudents();
      case '/api/students/create': return apiCreateStudent(data);
      case '/api/students/update': return apiUpdateStudent(data);
      case '/api/students/delete': return apiDeleteStudent(data);
      case '/api/students/batch-create': return apiBatchCreateStudents(data);
      case '/api/logo/upload': return apiUploadLogo(data);
      case '/api/public/settings': return apiGetPublicSettings();
      case '/api/students/search': return apiSearchStudent(data);
      default: return { success: false, message: 'Endpoint tidak ditemukan: ' + path };
    }
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ==========================================
// API ENDPOINTS — UMUM
// ==========================================

function apiHealth() {
  return { success: true, message: 'Koneksi ke backend berhasil', timestamp: new Date().toISOString() };
}

// ==========================================
// API ENDPOINTS — AUTH & SETTINGS
// ==========================================

function apiLogin(data) {
  var storedPassword = getSettingValue('admin_password');
  if (!storedPassword) return { success: true, message: 'Login berhasil (password belum diatur)', firstTime: true };
  if (data.password === storedPassword) return { success: true, message: 'Login berhasil' };
  return { success: false, message: 'Password salah' };
}

function apiGetSettings() {
  var sheet = getSheet(CONFIG.SHEET_SETTINGS);
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    if (key) settings[key] = data[i][1] || '';
  }
  return { success: true, data: settings };
}

function apiUpdateSettings(data) {
  var sheet = getSheet(CONFIG.SHEET_SETTINGS);
  var keys = Object.keys(data);
  for (var i = 0; i < keys.length; i++) setSettingValue(sheet, keys[i], data[keys[i]]);
  return { success: true, message: 'Pengaturan berhasil disimpan' };
}

function apiGetPublicSettings() {
  var sheet = getSheet(CONFIG.SHEET_SETTINGS);
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    if (key && key !== 'admin_password') settings[key] = data[i][1] || '';
  }
  return { success: true, data: settings };
}

// ==========================================
// API ENDPOINTS — STUDENTS
// ==========================================

/**
 * Kolom Students: A=id, B=nisn, C=nama, D=kelas, E=status, F=created_at, G=nilai
 */
function apiGetStudents() {
  var sheet = getSheet(CONFIG.SHEET_STUDENTS);
  var data = sheet.getDataRange().getValues();
  var students = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    students.push({
      id: String(data[i][0]),
      nisn: String(data[i][1]),
      nama: String(data[i][2]),
      kelas: String(data[i][3]),
      status: String(data[i][4]),
      created_at: String(data[i][5]),
      nilai: data[i][6] !== undefined && data[i][6] !== '' ? String(data[i][6]) : ''
    });
  }
  return { success: true, data: students };
}

function apiCreateStudent(data) {
  if (!data.nisn || !data.nama || !data.kelas || !data.status) {
    return { success: false, message: 'NISN, Nama, Kelas, dan Status wajib diisi' };
  }
  var sheet = getSheet(CONFIG.SHEET_STUDENTS);
  if (findRowByNisn(sheet, data.nisn)) {
    return { success: false, message: 'NISN ' + data.nisn + ' sudah terdaftar' };
  }
  var id = String(Date.now()) + String(Math.floor(Math.random() * 1000));
  sheet.appendRow([id, data.nisn, data.nama, data.kelas, data.status, new Date().toISOString(), data.nilai || '']);
  return { success: true, message: 'Siswa berhasil ditambahkan' };
}

function apiUpdateStudent(data) {
  if (!data.id) return { success: false, message: 'ID siswa diperlukan' };
  var sheet = getSheet(CONFIG.SHEET_STUDENTS);
  var allData = sheet.getDataRange().getValues();
  var row = findRowById(allData, data.id);
  if (!row) return { success: false, message: 'Data siswa tidak ditemukan' };
  var nisnRow = findRowByNisn(sheet, data.nisn);
  if (nisnRow && nisnRow !== row) return { success: false, message: 'NISN ' + data.nisn + ' sudah digunakan siswa lain' };
  sheet.getRange(row, 2, 1, 4).setValues([[data.nisn, data.nama, data.kelas, data.status]]);
  sheet.getRange(row, 7, 1, 1).setValue(data.nilai || '');
  return { success: true, message: 'Data siswa berhasil diupdate' };
}

function apiDeleteStudent(data) {
  if (!data.id) return { success: false, message: 'ID siswa diperlukan' };
  var sheet = getSheet(CONFIG.SHEET_STUDENTS);
  var allData = sheet.getDataRange().getValues();
  var row = findRowById(allData, data.id);
  if (!row) return { success: false, message: 'Data siswa tidak ditemukan' };
  sheet.deleteRow(row);
  return { success: true, message: 'Data siswa berhasil dihapus' };
}

/**
 * POST /api/students/batch-create
 * Import banyak siswa sekaligus
 * @param {object} data - { students: [{ nisn, nama, kelas, status, nilai }, ...] }
 */
function apiBatchCreateStudents(data) {
  var students = data.students;
  if (!students || !Array.isArray(students) || students.length === 0) {
    return { success: false, message: 'Tidak ada data untuk diimport' };
  }

  var sheet = getSheet(CONFIG.SHEET_STUDENTS);

  // Ambil semua NISN yang sudah ada (untuk cek duplikat)
  var existingNisn = getExistingNisn(sheet);

  var validRows = [];
  var errors = [];

  for (var i = 0; i < students.length; i++) {
    var s = students[i];

    // Validasi wajib isi
    if (!s.nisn || !s.nama || !s.kelas || !s.status) {
      errors.push('Baris ' + (i + 1) + ': NISN, Nama, Kelas, Status wajib diisi');
      continue;
    }

    // Cek duplikat dengan data di spreadsheet
    if (existingNisn.indexOf(String(s.nisn)) !== -1) {
      errors.push('Baris ' + (i + 1) + ': NISN ' + s.nisn + ' sudah terdaftar');
      continue;
    }

    // Tambah ke existingNisn (cek duplikat antar baris yang diimport)
    existingNisn.push(String(s.nisn));

    validRows.push([
      String(Date.now()) + String(Math.floor(Math.random() * 1000)) + String(i),
      String(s.nisn),
      String(s.nama),
      String(s.kelas),
      String(s.status),
      new Date().toISOString(),
      s.nilai ? String(s.nilai) : ''
    ]);
  }

  if (validRows.length === 0) {
    return { success: false, message: 'Tidak ada data valid', errors: errors };
  }

  // Batch insert — lebih cepat dari appendRow berulang kali
  var lastRow = sheet.getLastRow();
  // Jika sheet benar-benar kosong (hanya header atau sama sekali kosong)
  var startRow = lastRow < 1 ? 1 : lastRow + 1;
  sheet.getRange(startRow, 1, validRows.length, 7).setValues(validRows);

  return {
    success: true,
    message: 'Import selesai: ' + validRows.length + ' data berhasil, ' + errors.length + ' dilewati',
    inserted: validRows.length,
    skipped: errors.length,
    errors: errors
  };
}

/**
 * POST /api/students/search
 */
function apiSearchStudent(data) {
  if (!data.nisn) return { success: false, message: 'NISN wajib diisi' };
  var sheet = getSheet(CONFIG.SHEET_STUDENTS);
  var allData = sheet.getDataRange().getValues();
  var foundStudent = null;
  var lulusWithValue = [];

  for (var i = 1; i < allData.length; i++) {
    if (!allData[i][1]) continue;
    if (String(allData[i][1]) === String(data.nisn)) {
      foundStudent = {
        nisn: String(allData[i][1]), nama: String(allData[i][2]),
        kelas: String(allData[i][3]), status: String(allData[i][4]),
        nilai: (allData[i][6] !== undefined && allData[i][6] !== '') ? String(allData[i][6]) : ''
      };
    }
    if (String(allData[i][4]) === 'Lulus' && allData[i][6] !== undefined && allData[i][6] !== '') {
      lulusWithValue.push({ nama: String(allData[i][2]), nilai: Number(allData[i][6]) });
    }
  }

  if (!foundStudent) return { success: false, message: 'NISN tidak ditemukan' };
  lulusWithValue.sort(function (a, b) { return b.nilai - a.nilai; });
  return { success: true, data: { student: foundStudent, topThree: lulusWithValue.slice(0, 3) } };
}

// ==========================================
// HELPER FUNCTIONS
// ==========================================

function getSpreadsheet() { return SpreadsheetApp.getActiveSpreadsheet(); }
function getSheet(name) { return getSpreadsheet().getSheetByName(name); }

function getSettingValue(key) {
  var data = getSheet(CONFIG.SHEET_SETTINGS).getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1] || '';
  }
  return '';
}

function setSettingValue(sheet, key, value) {
  var data = sheet.getDataRange().getValues();
  var lastRow = sheet.getLastRow();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.getRange(lastRow + 1, 1).setValue(key);
  sheet.getRange(lastRow + 1, 2).setValue(value);
}

function findRowById(allData, id) {
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(id)) return i + 1;
  }
  return null;
}

function findRowByNisn(sheet, nisn) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(nisn)) return i + 1;
  }
  return null;
}

/**
 * Mengambil semua NISN yang sudah terdaftar di sheet Students
 * @param {Sheet} sheet
 * @returns {string[]} Array of NISN strings
 */
function getExistingNisn(sheet) {
  var data = sheet.getDataRange().getValues();
  var nisn = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) nisn.push(String(data[i][1]));
  }
  return nisn;
}

/**
 * POST /api/logo/upload
 * Upload logo PNG ke Google Drive, simpan URL-nya ke settings
 */
function apiUploadLogo(data) {
  try {
    // Jika base64 kosong, ini adalah permintaan hapus
    if (!data.base64) {
      var oldUrl = getSettingValue('logo_url');
      if (oldUrl) {
        var idMatch = oldUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (idMatch) {
          try { DriveApp.getFileById(idMatch[1]).setTrashed(true); } catch (e) {}
        }
      }
      return { success: true, message: 'Logo dihapus' };
    }

    var base64Str = data.base64;
    var commaIdx = base64Str.indexOf(',');
    if (commaIdx !== -1) {
      base64Str = base64Str.substring(commaIdx + 1);
    }

    // Decode base64 → blob
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Str),
      'image/png',
      'logo.png'
    );

    // Hapus logo lama jika ada
    var oldUrl = getSettingValue('logo_url');
    if (oldUrl) {
      var idMatch = oldUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (idMatch) {
        try { DriveApp.getFileById(idMatch[1]).setTrashed(true); } catch (e) {}
      }
    }

    // Simpan file baru ke Drive (otomatis ke root "Drive Saya")
    var fileName = 'logo_' + getSpreadsheet().getId() + '.png';
    var file = DriveApp.createFile(blob);
    file.setName(fileName);

    // Set agar bisa diakses publik tanpa login
    file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

    // Buat URL langsung (pendek, ~60 karakter)
    var url = 'https://lh3.googleusercontent.com/d/' + file.getId();

    // Simpan URL ke spreadsheet
    var sheet = getSheet(CONFIG.SHEET_SETTINGS);
    setSettingValue(sheet, 'logo_url', url);

    return {
      success: true,
      message: 'Logo berhasil diupload',
      data: { url: url }
    };

  } catch (error) {
    return { success: false, message: 'Gagal upload logo: ' + error.message };
  }
}
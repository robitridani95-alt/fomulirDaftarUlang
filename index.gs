// ============================================================
//  SMKN RAKIT KULIM — Sistem Daftar Ulang
//  Google Apps Script Backend
//  Versi: 1.0 | 2025
// ============================================================
//
//  CARA DEPLOY:
//  1. Buka script.google.com → Buat project baru
//  2. Salin seluruh kode ini
//  3. Ganti SPREADSHEET_ID di bawah dengan ID spreadsheet Anda
//  4. Klik Deploy → New Deployment → Web App
//     - Execute as: Me
//     - Who has access: Anyone
//  5. Copy URL deployment → tempel ke variabel SCRIPT_URL di formulir HTML
// ============================================================

const SPREADSHEET_ID = '1XXBK7GoCxsyGFAlDNT1INhVaAv1Kh1YvRevAAEjUFSg';
const SHEET_SISWA    = 'DaftarUlang';
const SHEET_LOG      = 'Log';
const SHEET_STATS    = 'Statistik';

// ─────────────────────────────────────────────
//  HEADER HTTP — izinkan CORS dari semua origin
// ─────────────────────────────────────────────
function setCORSHeaders(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────
//  ENTRY POINT — POST
// ─────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;

    if (action === 'addRow')    return setCORSHeaders(ContentService.createTextOutput(JSON.stringify(addRow(payload.data))));
    if (action === 'checkNISN') return setCORSHeaders(ContentService.createTextOutput(JSON.stringify(checkNISN(payload.nisn))));
    if (action === 'getStats')  return setCORSHeaders(ContentService.createTextOutput(JSON.stringify(getStats())));

    return setCORSHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Action tidak dikenal' })));
  } catch (err) {
    return setCORSHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() })));
  }
}

// ─────────────────────────────────────────────
//  ENTRY POINT — GET (untuk cek status / ping)
// ─────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  if (action === 'ping') {
    return setCORSHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'ok', server: 'SMKN Rakit Kulim Daftar Ulang v1.0' })));
  }
  if (action === 'getStats') {
    return setCORSHeaders(ContentService.createTextOutput(JSON.stringify(getStats())));
  }
  return setCORSHeaders(ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'Server aktif' })));
}

// ─────────────────────────────────────────────
//  FUNGSI: Tambah Data Siswa ke Sheet
// ─────────────────────────────────────────────
function addRow(data) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_SISWA);

  // Inisialisasi header jika sheet masih kosong
  if (sheet.getLastRow() === 0) {
    initHeaders(sheet);
  }

  // Tambah data
  sheet.appendRow(data);

  // Format baris terakhir
  const lastRow = sheet.getLastRow();
  formatNewRow(sheet, lastRow);

  // Catat ke log
  writeLog(ss, data[0], data[1], data[2], 'SUBMIT', 'Berhasil');

  // Update statistik
  updateStats(ss);

  // Kirim notifikasi email ke admin
  sendNotifEmail(data);

  return { status: 'ok', regNo: data[1], row: lastRow };
}

// ─────────────────────────────────────────────
//  FUNGSI: Cek duplikasi NISN
// ─────────────────────────────────────────────
function checkNISN(nisn) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_SISWA);

  if (sheet.getLastRow() <= 1) return { exists: false };

  const data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  // Kolom NISN = indeks 2 (kolom C, 0-based)
  const found = data.find(row => String(row[2]).trim() === String(nisn).trim());

  return found
    ? { exists: true,  name: found[2], regNo: found[1] }
    : { exists: false };
}

// ─────────────────────────────────────────────
//  FUNGSI: Ambil statistik ringkasan
// ─────────────────────────────────────────────
function getStats() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_SISWA);

  if (sheet.getLastRow() <= 1) {
    return { total: 0, byKelas: {}, byJurusan: {}, today: 0 };
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const today = new Date().toLocaleDateString('id-ID');

  let byKelas   = {};
  let byJurusan = {};
  let todayCount = 0;

  data.forEach(row => {
    const kelas   = row[20] || 'Tidak diketahui';   // Kolom U
    const jurusan = row[21] || 'Tidak diketahui';   // Kolom V
    const tgl     = new Date(row[0]).toLocaleDateString('id-ID');

    byKelas[kelas]     = (byKelas[kelas]     || 0) + 1;
    byJurusan[jurusan] = (byJurusan[jurusan] || 0) + 1;
    if (tgl === today) todayCount++;
  });

  return {
    total    : data.length,
    byKelas,
    byJurusan,
    today    : todayCount,
    lastUpdate: new Date().toLocaleString('id-ID')
  };
}

// ─────────────────────────────────────────────
//  HELPER: Inisialisasi header kolom
// ─────────────────────────────────────────────
function initHeaders(sheet) {
  const headers = [
    'Timestamp', 'No. Registrasi',
    // Identitas Siswa
    'Nama Lengkap', 'NISN', 'NIK',
    'Tempat Lahir', 'Tanggal Lahir', 'Jenis Kelamin', 'Agama',
    'Alamat Lengkap', 'Kode Pos',
    'No. HP Siswa', 'Email Siswa',
    // Keluarga
    'Nama Ayah', 'Pekerjaan Ayah', 'Pendidikan Ayah', 'Penghasilan Ayah', 'No. HP Ayah',
    'Nama Ibu', 'Pekerjaan Ibu', 'Pendidikan Ibu', 'No. HP Ibu',
    'Nama Wali', 'Hub. Wali', 'No. HP Wali',
    // Akademik
    'Tahun Ajaran', 'Kelas', 'Kompetensi Keahlian', 'Rombel',
    'Asal Sekolah', 'Tahun Lulus SMP', 'No. Ijazah SMP',
    // Kesehatan
    'Gol. Darah', 'Tinggi Badan', 'Berat Badan',
    'Jarak Rumah', 'Transportasi',
    // Pembayaran & Dokumen
    'Status Bayar', 'No. Bukti Bayar', 'PIP/KIP', 'No. KIP',
    // Tambahan
    'Prestasi / Minat', 'Catatan', 'Status Proses'
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // Style header
  headerRange
    .setBackground('#1a4a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  // Lebar kolom
  sheet.setColumnWidth(1, 160);   // Timestamp
  sheet.setColumnWidth(2, 140);   // No. Reg
  sheet.setColumnWidth(3, 200);   // Nama
  sheet.setColumnWidth(4, 120);   // NISN
  sheet.setColumnWidth(5, 140);   // NIK

  // Freeze & filter
  sheet.getRange(1, 1, 1, headers.length).createFilter();
}

// ─────────────────────────────────────────────
//  HELPER: Format baris baru
// ─────────────────────────────────────────────
function formatNewRow(sheet, row) {
  const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  range
    .setBackground(row % 2 === 0 ? '#f0f7f2' : '#ffffff')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, '#d0e8d8', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(row, 30);

  // Kolom "Status Proses" = kolom terakhir → default Menunggu Verifikasi
  const lastCol = sheet.getLastColumn();
  const statusCell = sheet.getRange(row, lastCol);
  statusCell.setValue('Menunggu Verifikasi');
  statusCell
    .setBackground('#fff8e1')
    .setFontColor('#e65100')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

// ─────────────────────────────────────────────
//  HELPER: Sheet log aktivitas
// ─────────────────────────────────────────────
function writeLog(ss, timestamp, regNo, nama, action, keterangan) {
  const logSheet = getOrCreateSheet(ss, SHEET_LOG);
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['Timestamp', 'No. Reg', 'Nama', 'Action', 'Keterangan', 'IP/Agent']);
    logSheet.getRange(1,1,1,6).setBackground('#1a4a2e').setFontColor('#fff').setFontWeight('bold');
  }
  logSheet.appendRow([timestamp, regNo, nama, action, keterangan, 'Web Form']);
}

// ─────────────────────────────────────────────
//  HELPER: Update sheet statistik otomatis
// ─────────────────────────────────────────────
function updateStats(ss) {
  const statsSheet = getOrCreateSheet(ss, SHEET_STATS);
  statsSheet.clearContents();

  const stats = getStats();

  statsSheet.getRange('A1').setValue('📊 STATISTIK DAFTAR ULANG — SMKN RAKIT KULIM');
  statsSheet.getRange('A1').setFontSize(14).setFontWeight('bold').setFontColor('#1a4a2e');

  statsSheet.getRange('A3').setValue('Total Pendaftar');
  statsSheet.getRange('B3').setValue(stats.total);

  statsSheet.getRange('A4').setValue('Pendaftar Hari Ini');
  statsSheet.getRange('B4').setValue(stats.today);

  statsSheet.getRange('A5').setValue('Terakhir Update');
  statsSheet.getRange('B5').setValue(stats.lastUpdate);

  statsSheet.getRange('A7').setValue('Per Kelas');
  statsSheet.getRange('A7').setFontWeight('bold');
  let r = 8;
  Object.entries(stats.byKelas).forEach(([k, v]) => {
    statsSheet.getRange(r, 1).setValue(k);
    statsSheet.getRange(r, 2).setValue(v);
    r++;
  });

  r += 1;
  statsSheet.getRange(r, 1).setValue('Per Jurusan').setFontWeight('bold');
  r++;
  Object.entries(stats.byJurusan).forEach(([k, v]) => {
    statsSheet.getRange(r, 1).setValue(k);
    statsSheet.getRange(r, 2).setValue(v);
    r++;
  });
}

// ─────────────────────────────────────────────
//  HELPER: Buat sheet jika belum ada
// ─────────────────────────────────────────────
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ─────────────────────────────────────────────
//  HELPER: Notifikasi email ke admin
// ─────────────────────────────────────────────
function sendNotifEmail(data) {
  try {
    const adminEmail = Session.getActiveUser().getEmail(); // email pemilik script
    const subject    = `[SMKN Rakit Kulim] Daftar Ulang Baru: ${data[2]} — ${data[1]}`;
    const body = `
Pendaftaran baru telah masuk ke sistem Daftar Ulang SMKN Rakit Kulim.

━━━━━━━━━━━━━━━━━━━━━━
  No. Registrasi : ${data[1]}
  Nama Siswa     : ${data[2]}
  NISN           : ${data[3]}
  Kelas          : ${data[26] || '-'}
  Jurusan        : ${data[27] || '-'}
  Waktu Submit   : ${data[0]}
━━━━━━━━━━━━━━━━━━━━━━

Silakan periksa Google Spreadsheet untuk detail lengkap.

— Sistem Otomatis SMKN Rakit Kulim
    `;
    GmailApp.sendEmail(adminEmail, subject, body);
  } catch (e) {
    // Email opsional — jangan hentikan proses jika gagal
    Logger.log('Email notif gagal: ' + e.toString());
  }
}

// ─────────────────────────────────────────────
//  FUNGSI: Ekspor data ke CSV (dipanggil manual)
// ─────────────────────────────────────────────
function exportToCSV() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SISWA);
  const data  = sheet.getDataRange().getValues();

  const csv = data.map(row =>
    row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
  ).join('\n');

  const file = DriveApp.createFile('DaftarUlang_Export_' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmm') + '.csv', csv, MimeType.CSV);
  Logger.log('CSV tersimpan: ' + file.getUrl());
  return file.getUrl();
}

// ─────────────────────────────────────────────
//  FUNGSI: Setup awal (jalankan 1x manual)
// ─────────────────────────────────────────────
function setupSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Rename spreadsheet
  ss.rename('SMKN Rakit Kulim — Sistem Daftar Ulang 2025/2026');

  // Buat semua sheet
  [SHEET_SISWA, SHEET_LOG, SHEET_STATS].forEach(name => getOrCreateSheet(ss, name));

  // Init header sheet utama
  const sheet = ss.getSheetByName(SHEET_SISWA);
  if (sheet.getLastRow() === 0) initHeaders(sheet);

  Logger.log('Setup selesai! Spreadsheet: ' + ss.getUrl());
  return 'Setup berhasil';
}

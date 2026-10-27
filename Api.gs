// ============================================================================
// Api.gs — Jembatan HTTP agar frontend bisa jalan di luar Apps Script
// ============================================================================

// Tanpa login
const API_PUBLIC = ['login', 'logoutUser', 'getPengaturanTampilanAwal', 'getDeploymentUrl'];

// Hanya admin
const API_ADMIN = [
  'bersihkanDuplikatAbsensiSholat', 'archiveAndResetYear', 'processGradePromotion',
  'processIndividualPromotion', 'importSiswaBulk', 'importGuruBulk', 'importHariLiburBulk',
  'addKelas', 'updateKelas', 'deleteKelas', 'getKelasMaster', 'saveAppConfig', 'setModeAbsen',
  'addHariLibur', 'addHariLiburRange', 'deleteHariLibur', 'updateHariLibur'
];

// Admin / guru / wakel
const API_STAFF = [
  'scanAbsensi', 'lookupSiswaForScan', 'batchScanAbsensi', 'getAbsensiList', 'getMonitoringRealtime',
  'getSiswaList', 'getSiswaByKelas', 'getSiswaByNisn', 'getKelasList', 'getMonthlyReportData',
  'generateExcel', 'getTemplateExcel', 'submitAbsensiSholat', 'getAbsensiSholatByKelas',
  'getRekapSholat', 'getRekapBulananSholat', 'generateExcelRekapSholatBulanan',
  'getDetailHarianSholatSiswa', 'generateExcelDetailSholatSiswa', 'updateAbsensiSholatSiswa',
  'hapusAbsensiSholatSiswa', 'getLaporanSholatDetailBulanan', 'getRekapSholatPerKelas',
  'generateExcelLaporanSholatDetail', 'getLaporanSholatPrintHTML',
  'getDaftarHadirByKelasAndTanggal', 'getDaftarHadirDetailByKelasAndTanggal',
  'submitPulangSatu', 'submitPulangGuru', 'submitDaftarHadirGuru'
];

// Semua yang punya sesi valid (fungsi-fungsi ini juga mengecek role sendiri lewat token)
const API_ANY = [
  'getProfilSiswa', 'updateProfilSiswa', 'getPengaturanWaktuSholat', 'simpanPengaturanWaktuSholat',
  'izinkanEditSiswaSholat', 'submitAbsensiSholatSiswa', 'getAbsensiSholatSiswaHariIni',
  'getAbsensiToday', 'getHariLibur', 'getAppConfig', 'getModeAbsen', 'getDaftarKelas', 'getRankingSholat',
  'addSiswa', 'updateSiswa', 'deleteSiswa', 'getGuruList', 'addGuru', 'updateGuru', 'deleteGuru',
  'updateAbsensiStatus', 'saveTampilanAplikasi', 'resetTampilanAplikasi',
  'nilaiPengaturanGet', 'nilaiPengaturanSimpan', 'kuisBuat', 'kuisListGuru', 'kuisSetStatus',
  'kuisHapus', 'kuisHasilList', 'kuisListSiswa', 'kuisAmbilSoal', 'kuisKirimJawaban',
  'nilaiManualSimpanBatch', 'getRekapNilai', 'getDetailNilaiSiswa', 'generateExcelRekapNilai',
  'tautanBuat', 'tautanUbah', 'tautanList', 'tautanSinkron', 'tautanSetStatus', 'tautanHapus',
  'tautanAnalisisSoal', 'tautanListSiswa', 'tautanSinkronSemua',
  'canvaImportAnalisis', 'canvaImportSimpan'
];

function api_json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function api_handle_(req) {
  const fn = String(req.fn || '');
  const args = Array.isArray(req.args) ? req.args : [];

  const dikenal = API_PUBLIC.concat(API_ADMIN, API_STAFF, API_ANY);
  if (dikenal.indexOf(fn) === -1) throw new Error('Fungsi tidak tersedia: ' + fn);

  if (API_PUBLIC.indexOf(fn) === -1) {
    const s = getSessionInfo(req.token);   // melempar error bila token salah / kedaluwarsa
    if (API_ADMIN.indexOf(fn) !== -1 && s.role !== 'admin') {
      throw new Error('Akses ditolak: hanya admin.');
    }
    if (API_STAFF.indexOf(fn) !== -1 && ['admin', 'guru', 'wakel'].indexOf(s.role) === -1) {
      throw new Error('Akses ditolak: hanya guru/admin.');
    }
  }

  const f = globalThis[fn];
  if (typeof f !== 'function') throw new Error('Fungsi belum ada di server: ' + fn);
  return f.apply(null, args);
}

// Satu-satunya doPost. Permintaan API punya field "fn"; selain itu dianggap webhook nilai lama.
function doPost(e) {
  let body = null;
  try { body = JSON.parse(e.postData.contents); } catch (x) {}
  if (body && body.fn) {
    try {
      return api_json_({ ok: true, result: api_handle_(body) });
    } catch (err) {
      return api_json_({ ok: false, error: String(err.message || err) });
    }
  }
  return tautan_doPostWebhook_(e);
}

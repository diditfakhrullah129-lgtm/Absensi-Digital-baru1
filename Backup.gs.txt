// ============================================================================
// CADANGAN DATA OTOMATIS — Absensi Siswa Digital
// Taruh sebagai file baru bernama "Backup" (Backup.gs) di project Apps Script,
// satu project dengan Code.gs, Nilai.gs, dan Tautan.gs.
// Tidak perlu mengubah file lain: memakai SPREADSHEET_ID & PARENT_FOLDER_ID
// yang sudah ada di Code.gs.
//
// Cara kerja:
//   - Tiap hari (default pukul 01:00) seluruh spreadsheet disalin ke folder
//     "Cadangan_Harian_Absensi" di dalam folder Drive utama aplikasi.
//   - Hanya N salinan terbaru yang disimpan (default 30). Salinan yang lebih
//     lama dipindah ke Trash otomatis, jadi Drive tidak penuh.
//   - Jika backup gagal, email pemberitahuan dikirim ke pemilik skrip.
//
// LANGKAH PASANG (sekali saja):
//   1. Simpan file ini, lalu pilih fungsi  pasangPemicuBackup  -> klik Run.
//   2. Setujui izin yang diminta Google (Drive, Trigger, Email).
//   3. (Opsional) Jalankan  backupSekarang  untuk mencoba langsung.
// ============================================================================

const BACKUP_CFG = {
  namaFolder: 'Cadangan_Harian_Absensi',   // subfolder di PARENT_FOLDER_ID
  prefix: 'Backup Absensi ',               // awalan nama file salinan
  simpanTerbaru: 30,                       // jumlah salinan yang dipertahankan
  jam: 1                                   // jam jalan harian (0-23, zona waktu skrip)
};

// ---------------------------------------------------------------------------
// HELPER
// ---------------------------------------------------------------------------
function backup_folder_() {
  const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const it = parent.getFoldersByName(BACKUP_CFG.namaFolder);
  return it.hasNext() ? it.next() : parent.createFolder(BACKUP_CFG.namaFolder);
}

// Pindahkan salinan lama ke Trash, sisakan N terbaru
function backup_bersihkan_(folder) {
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(BACKUP_CFG.prefix) === 0) files.push(f);
  }
  files.sort((a, b) => b.getDateCreated().getTime() - a.getDateCreated().getTime());
  let dihapus = 0;
  for (let i = BACKUP_CFG.simpanTerbaru; i < files.length; i++) {
    files[i].setTrashed(true);
    dihapus++;
  }
  return { total: Math.min(files.length, BACKUP_CFG.simpanTerbaru), dihapus: dihapus };
}

function backup_simpanStatus_(obj) {
  PropertiesService.getScriptProperties().setProperty('backup_terakhir', JSON.stringify(obj));
}

function backup_kabarGagal_(pesan) {
  try {
    MailApp.sendEmail(
      Session.getEffectiveUser().getEmail(),
      'GAGAL: Cadangan otomatis Absensi Siswa Digital',
      'Cadangan harian gagal dibuat pada ' + formatWaktuSheet(new Date(), 'yyyy-MM-dd HH:mm:ss') +
      '.\n\nPenyebab: ' + pesan +
      '\n\nSilakan buka Apps Script, jalankan fungsi backupSekarang, dan periksa log eksekusi.'
    );
  } catch (e) { /* email opsional, abaikan bila tidak diizinkan */ }
}

// ---------------------------------------------------------------------------
// INTI BACKUP
// ---------------------------------------------------------------------------
function backup_jalankan_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Backup lain sedang berjalan, coba lagi sebentar.');
  try {
    const folder = backup_folder_();
    const nama = BACKUP_CFG.prefix + formatWaktuSheet(new Date(), 'yyyy-MM-dd HH-mm');
    const salinan = DriveApp.getFileById(SPREADSHEET_ID).makeCopy(nama, folder);
    const sisa = backup_bersihkan_(folder);
    const hasil = {
      waktu: formatWaktuSheet(new Date(), 'yyyy-MM-dd HH:mm:ss'),
      status: 'Berhasil',
      nama: nama,
      url: salinan.getUrl(),
      pesan: 'Tersimpan ' + sisa.total + ' salinan' + (sisa.dihapus ? ', ' + sisa.dihapus + ' salinan lama dibuang ke Trash' : '') + '.'
    };
    backup_simpanStatus_(hasil);
    return hasil;
  } finally {
    lock.releaseLock();
  }
}

// Dipanggil otomatis oleh trigger harian
function cadanganHarian() {
  try {
    backup_jalankan_();
  } catch (e) {
    backup_simpanStatus_({
      waktu: formatWaktuSheet(new Date(), 'yyyy-MM-dd HH:mm:ss'),
      status: 'Gagal', nama: '', url: '', pesan: e.message
    });
    backup_kabarGagal_(e.message);
    throw e; // supaya tercatat juga di log eksekusi Apps Script
  }
}

// Jalankan manual dari editor untuk mencoba
function backupSekarang() {
  const h = backup_jalankan_();
  Logger.log('Backup selesai: ' + h.nama + ' | ' + h.url + ' | ' + h.pesan);
  return h;
}

// ---------------------------------------------------------------------------
// PASANG / LEPAS TRIGGER
// ---------------------------------------------------------------------------
function pasangPemicuBackup() {
  hapusPemicuBackup();
  ScriptApp.newTrigger('cadanganHarian')
    .timeBased()
    .everyDays(1)
    .atHour(BACKUP_CFG.jam)
    .create();
  Logger.log('Pemicu cadangan harian terpasang: tiap hari sekitar pukul ' + BACKUP_CFG.jam + ':00.');
}

function hapusPemicuBackup() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'cadanganHarian') ScriptApp.deleteTrigger(t);
  });
}

// ---------------------------------------------------------------------------
// OPSIONAL: dipanggil dari tampilan (hanya admin) bila nanti ingin dibuatkan tombol
//   google.script.run.backupSekarangAdmin(currentUser.token)
//   google.script.run.backupStatusAdmin(currentUser.token)
// ---------------------------------------------------------------------------
function backupSekarangAdmin(token) {
  try {
    verifyUser(token, 'admin');
    const h = backup_jalankan_();
    return { success: true, message: 'Cadangan dibuat. ' + h.pesan, data: h };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function backupStatusAdmin(token) {
  try {
    verifyUser(token, 'admin');
    const raw = PropertiesService.getScriptProperties().getProperty('backup_terakhir');
    const adaPemicu = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'cadanganHarian');
    return { success: true, pemicuAktif: adaPemicu, terakhir: raw ? JSON.parse(raw) : null,
             simpanTerbaru: BACKUP_CFG.simpanTerbaru };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

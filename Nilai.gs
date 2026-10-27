// ============================================================================
// REKAP NILAI DIGITAL & KUIS ONLINE  —  Absensi Siswa Digital
// Taruh sebagai file "Nilai.gs" di project Apps Script (ganti file lama).
//
// PERUBAHAN VERSI INI:
//  1. Nilai Pemahaman & Pengetahuan (kunci internal "Harian") dan Nilai Analisa
//     sekarang DIHITUNG TERPISAH. Dulu keduanya memakai data yang sama.
//  2. Nilai Akhir memakai bobot: Pemahaman 45%, Analisa 45%, Disiplin 10%.
//  3. Kolom Score ikut ditampilkan sesuai spreadsheet (tidak masuk Nilai Akhir).
//
// Sheet yang dibuat OTOMATIS saat pertama dipakai:
//   kuis          : daftar kuis + soal + kunci jawaban
//   hasil_kuis    : hasil pengerjaan siswa (score, benar, salah)
//   nilai_manual  : nilai yang diinput guru
// ============================================================================

const JENIS_NILAI = ['Harian', 'Analisa', 'Refleksi'];
const NL_KOLOM = ['Harian', 'Analisa', 'Refleksi', 'Disiplin'];

// ---------------------------------------------------------------------------
// PENGATURAN NILAI AKHIR  (ubah di sini kalau bobot ingin diganti)
// Kunci "Harian" = Nilai Pemahaman & Pengetahuan. Total bobot sebaiknya = 1.
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// PENGATURAN NILAI (bobot, KKM, mode) — disimpan di sheet 'konfigurasi'
// Nilai di bawah ini hanya DEFAULT bila guru belum pernah mengatur.
// ---------------------------------------------------------------------------
const NL_BOBOT_DEFAULT = { Harian: 0.45, Analisa: 0.45, Disiplin: 0.1 };
const NL_KKM_DEFAULT = 75;
const NL_MODE_DEFAULT = 'lengkap';
const NL_MODE_VALID = ['lengkap', 'nol', 'proporsional'];
var NL_CFG_CACHE_ = null;

function nl_pengaturan_() {
  if (NL_CFG_CACHE_) return NL_CFG_CACHE_;
  const cfg = { bobot: Object.assign({}, NL_BOBOT_DEFAULT), kkm: NL_KKM_DEFAULT, mode: NL_MODE_DEFAULT };
  try {
    const sh = getSpreadsheet().getSheetByName('konfigurasi');
    if (sh) {
      const d = sh.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        if (d[i][0] !== 'nilai_pengaturan') continue;
        const p = JSON.parse(String(d[i][1]));
        if (p.bobot) {
          const b = {}; let sum = 0, ok = true;
          Object.keys(NL_BOBOT_DEFAULT).forEach(j => {
            const v = Number(p.bobot[j]);
            if (isNaN(v) || v < 0 || v > 1) ok = false;
            b[j] = v; sum += v;
          });
          if (ok && Math.abs(sum - 1) < 0.001) cfg.bobot = b;
        }
        const k = Number(p.kkm);
        if (!isNaN(k) && k >= 0 && k <= 100) cfg.kkm = k;
        if (NL_MODE_VALID.indexOf(p.mode) !== -1) cfg.mode = p.mode;
        break;
      }
    }
  } catch (e) {}
  NL_CFG_CACHE_ = cfg;
  return cfg;
}

// Semua yang sudah login boleh membaca (siswa perlu KKM untuk tanda merah)
function nilaiPengaturanGet(token) {
  try {
    nl_session_(token, ['admin', 'guru', 'wakel', 'siswa']);
    const c = nl_pengaturan_();
    return { success: true, bobot: c.bobot, kkm: c.kkm, mode: c.mode };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Hanya guru & admin yang boleh mengubah. (Pengaturan berlaku untuk SEMUA kelas.
// Kalau ingin hanya admin: ganti ['guru','admin'] menjadi ['admin'].)
function nilaiPengaturanSimpan(token, data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    nl_session_(token, ['guru', 'admin']);

    const pct = { Harian: Number(data.harian), Analisa: Number(data.analisa), Disiplin: Number(data.disiplin) };
    let sum = 0;
    Object.keys(pct).forEach(j => {
      if (isNaN(pct[j]) || pct[j] < 0 || pct[j] > 100) throw new Error('Bobot harus berupa angka 0–100.');
      sum += pct[j];
    });
    if (Math.abs(sum - 100) > 0.01) throw new Error('Total bobot harus tepat 100% (sekarang ' + (Math.round(sum * 100) / 100) + '%).');

    const kkm = Number(data.kkm);
    if (isNaN(kkm) || kkm < 0 || kkm > 100) throw new Error('KKM harus berupa angka 0–100.');
    const mode = String(data.mode || '');
    if (NL_MODE_VALID.indexOf(mode) === -1) throw new Error('Mode perhitungan tidak valid.');

    const ss = getSpreadsheet();
    let sh = ss.getSheetByName('konfigurasi');
    if (!sh) { sh = ss.insertSheet('konfigurasi'); sh.appendRow(['key', 'value', 'keterangan']); }

    const simpan = {
      bobot: { Harian: pct.Harian / 100, Analisa: pct.Analisa / 100, Disiplin: pct.Disiplin / 100 },
      kkm: kkm, mode: mode
    };
    _setKonfigurasiValue(sh, 'nilai_pengaturan', JSON.stringify(simpan));
    NL_CFG_CACHE_ = null;
    const c = nl_pengaturan_();
    return { success: true, message: 'Pengaturan nilai disimpan. Rekap langsung memakai bobot & KKM baru.', bobot: c.bobot, kkm: c.kkm, mode: c.mode };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Nama pendek komponen (untuk keterangan "Belum: ...")
const NL_PENDEK = { Harian: 'Pemahaman', Analisa: 'Analisa', Disiplin: 'Disiplin' };

const NL_HDR = {
  kuis:         ['ID', 'Judul', 'Mapel', 'Kelas', 'Jenis', 'Durasi (menit)', 'Status', 'Dibuat Oleh', 'Tanggal Dibuat', 'Soal (JSON)'],
  hasil_kuis:   ['Waktu', 'ID Kuis', 'Judul', 'Jenis', 'NISN', 'Nama', 'Kelas', 'Score', 'Benar', 'Salah', 'Total Soal', 'Jawaban (JSON)'],
  nilai_manual: ['Waktu', 'NISN', 'Nama', 'Kelas', 'Jenis', 'Nilai', 'Keterangan', 'Diinput Oleh'],
  detail_soal:  ['Waktu', 'ID Tautan', 'Judul', 'NISN', 'Nama', 'Kelas', 'Benar', 'Salah', 'Sebagian', 'Total Soal', 'No Benar', 'No Salah', 'Detail (JSON)']
};

// ---------------------------------------------------------------------------
// HELPER
// ---------------------------------------------------------------------------
function nl_sheet_(name) {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    const headers = NL_HDR[name];
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1B5E20').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Opsional: jalankan sekali dari editor untuk membuat semua sheet sekaligus
function setupNilaiSheets() {
  Object.keys(NL_HDR).forEach(n => nl_sheet_(n));
  Logger.log('Sheet kuis, hasil_kuis, nilai_manual siap.');
}

function nl_norm_(k) { return String(k || '').trim().toUpperCase().replace(/\s+/g, ' '); }
function nl_nisn_(v) { return String(v == null ? '' : v).replace(/^'/, '').trim(); }
function nl_round_(n) { return Math.round(n * 10) / 10; }
function nl_fmt_(v) {
  if (v instanceof Date) return formatWaktuSheet(v, 'yyyy-MM-dd HH:mm:ss');
  return String(v || '');
}

// Validasi token + role. Melempar error jika tidak valid / tidak berhak.
function nl_session_(token, allowedRoles) {
  const sh = getSpreadsheet().getSheetByName('sessions');
  if (!sh) throw new Error('Sesi tidak ditemukan. Silakan login ulang.');
  const data = sh.getDataRange().getValues();
  const now = new Date();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === token) {
      if (!(data[i][3] instanceof Date) || data[i][3] <= now) throw new Error('Sesi berakhir. Silakan login ulang.');
      const s = { role: String(data[i][2]).trim(), identifier: nl_nisn_(data[i][1]) };
      if (allowedRoles && allowedRoles.indexOf(s.role) === -1) throw new Error('Akses ditolak: Anda tidak memiliki izin.');
      return s;
    }
  }
  throw new Error('Token tidak valid. Silakan login ulang.');
}

// Daftar kelas yang boleh diakses guru. null = akses semua kelas.
function nl_kelasGuru_(session) {
  if (session.role !== 'guru') return null;
  const data = getSpreadsheet().getSheetByName('users').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (nl_nisn_(data[i][0]) === session.identifier) {
      const list = String(data[i][3] || '').split(',').map(k => k.trim()).filter(Boolean);
      return list.length ? list : null;
    }
  }
  return null;
}

function nl_kelasBoleh_(session, kelas) {
  const akses = nl_kelasGuru_(session);
  if (!akses) return true;
  return akses.map(nl_norm_).indexOf(nl_norm_(kelas)) !== -1;
}

// Data siswa (nama & kelas SAMA dengan sheet 'siswa' yang dipakai absensi)
function nl_loadSiswa_() {
  const data = getSpreadsheet().getSheetByName('siswa').getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const n = nl_nisn_(data[i][1]);
    map[n] = { nisn: n, nama: data[i][0], kelas: String(data[i][8]).trim() };
  }
  return map;
}

// Kumpulkan semua nilai (kuis + manual) per NISN
function nl_recordsByNisn_() {
  const out = {};
  const add = (nisn, rec) => { (out[nisn] = out[nisn] || []).push(rec); };

  const hk = nl_sheet_('hasil_kuis').getDataRange().getValues();
  for (let i = 1; i < hk.length; i++) {
    if (!hk[i][4]) continue;

    // Ambil teks refleksi dari kolom JSON (kolom ke-12)
    let refl = '';
    try { refl = String((JSON.parse(hk[i][11] || '{}') || {}).refleksi || '').trim(); } catch (e) {}

    add(nl_nisn_(hk[i][4]), {
      waktu: nl_fmt_(hk[i][0]), jenis: String(hk[i][3]), judul: String(hk[i][2]),
      score: Number(hk[i][7]) || 0, benar: Number(hk[i][8]) || 0, salah: Number(hk[i][9]) || 0,
      total: Number(hk[i][10]) || 0, sumber: 'Kuis', keterangan: '',
      refleksi: refl
    });
  }

  const nm = nl_sheet_('nilai_manual').getDataRange().getValues();
  for (let i = 1; i < nm.length; i++) {
    if (!nm[i][1]) continue;
    add(nl_nisn_(nm[i][1]), {
      waktu: nl_fmt_(nm[i][0]), jenis: String(nm[i][4]), judul: String(nm[i][6] || 'Input guru'),
      score: Number(nm[i][5]) || 0, benar: 0, salah: 0, total: 0, sumber: 'Manual',
      keterangan: String(nm[i][6] || ''),
      refleksi: ''
    });
  }
  return out;
}

function nl_tgl_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) ? String(s) : ''; }

function nl_filterRange_(recs, mulai, akhir) {
  if (!mulai && !akhir) return recs;
  return recs.filter(r => {
    const t = String(r.waktu).substring(0, 10);
    return (!mulai || t >= mulai) && (!akhir || t <= akhir);
  });
}

// ---------------------------------------------------------------------------
// HITUNG KOMPONEN + NILAI AKHIR BERBOBOT
//  - Harian  (= Pemahaman & Pengetahuan) : dari record berjenis "Harian"
//  - Analisa                              : dari record berjenis "Analisa"
//  - Refleksi                             : hanya ditampilkan, tidak masuk bobot
//  - Disiplin                             : dari kehadiran
// ---------------------------------------------------------------------------
function nl_komponen_(recs, dis, mulai, akhir) {
  recs = nl_filterRange_(recs, mulai, akhir);
  const harian = [], analisa = [], refl = [], skor = [];
  let waktuKuis = '';
  recs.forEach(r => {
    const j = (r.jenis === 'Evaluasi') ? 'Refleksi' : r.jenis;
    if (j === 'Harian' || j === 'Analisa') {
      (j === 'Harian' ? harian : analisa).push(r);
      if (String(r.waktu) > waktuKuis) waktuKuis = String(r.waktu);
    } else if (j === 'Score') {
      skor.push(r);
      if (String(r.waktu) > waktuKuis) waktuKuis = String(r.waktu);
    } else if (j === 'Refleksi') {
      refl.push(r);
    }
  });

  const hitung = arr => {
    if (!arr.length) return { ada: false, nilai: 0, ket: '' };
    let sum = 0, b = 0, s = 0;
    arr.forEach(r => { sum += r.score; b += r.benar || 0; s += r.salah || 0; });
    return { ada: true, nilai: nl_round_(sum / arr.length),
             ket: (b + s) > 0 ? 'Benar ' + b + ' • Salah ' + s : arr.length + ' nilai' };
  };

  const kH = hitung(harian);
  const kA = hitung(analisa);

  // Score mengikuti kolom Score di spreadsheet (angka apa adanya, tidak dibobot).
  // Bila ada beberapa, dipakai yang terbaru.
  let kS = { ada: false, nilai: 0, ket: '' };
  if (skor.length) {
    const terbaru = skor.slice().sort((a, b) => String(b.waktu).localeCompare(String(a.waktu)))[0];
    kS = { ada: true, nilai: terbaru.score, ket: terbaru.total > 0 ? 'dari ' + terbaru.total : '' };
  }

  // Refleksi teks: ambil yang terbaru dari semua record
  let teksR = '', waktuR = '';
  recs.forEach(r => {
    if (r.refleksi && String(r.waktu) >= waktuR) { teksR = r.refleksi; waktuR = String(r.waktu); }
  });
  let reflKomp = hitung(refl);
  if (teksR) {
    if (reflKomp.ada) reflKomp.teks = teksR;
    else reflKomp = { ada: true, nilai: 0, ket: '', teks: teksR, hanyaTeks: true };
  }

  const komponen = {
    Score: kS,
    Harian: kH,
    Analisa: kA,
    Refleksi: reflKomp,
    Disiplin: (dis && dis.ada)
      ? { ada: true, nilai: dis.nilai, ket: '', hadir: dis.hadir, hari: dis.hari }
      : { ada: false, nilai: 0, ket: '' }
  };

    // ---- Nilai Akhir berbobot (bobot, mode & KKM dari pengaturan guru) ----
  const cfg = nl_pengaturan_();
  const BOBOT = cfg.bobot;
  let total = 0, jumlahBobot = 0, sumBobot = 0, lengkap = true;
  const kurang = [];
  Object.keys(BOBOT).forEach(j => {
    komponen[j].bobot = BOBOT[j];
    if (komponen[j].ada) total += komponen[j].nilai;
    if (!(BOBOT[j] > 0)) return;            // bobot 0% = komponen tidak diwajibkan
    if (komponen[j].ada) {
      jumlahBobot += komponen[j].nilai * BOBOT[j];
      sumBobot += BOBOT[j];
    } else {
      lengkap = false;
      kurang.push(NL_PENDEK[j] || j);
    }
  });

  let nilaiAkhir = 0, akhirAda = false;
  if (cfg.mode === 'proporsional') {
    if (sumBobot > 0) { nilaiAkhir = jumlahBobot / sumBobot; akhirAda = true; }
  } else if (cfg.mode === 'nol') {
    if (sumBobot > 0) { nilaiAkhir = jumlahBobot; akhirAda = true; }
  } else {
    if (lengkap && sumBobot > 0) { nilaiAkhir = jumlahBobot; akhirAda = true; }
  }

  return {
    komponen: komponen,
    total: nl_round_(total),
    akhir: nl_round_(nilaiAkhir),
    akhirAda: akhirAda,
    lengkap: lengkap,
    kurang: kurang,
    bobot: BOBOT,
    mode: cfg.mode,
    kkm: cfg.kkm,
    waktuKuis: waktuKuis
  };
}

// Nilai disiplin dari sheet 'absensi' (hari libur di 'hari_libur' tidak dihitung)
function nl_disiplinMap_(mulai, akhir) {
  const ss = getSpreadsheet();
  const today = formatWaktuSheet(new Date(), 'yyyy-MM-dd');

  const libur = {};
  const ls = ss.getSheetByName('hari_libur');
  if (ls) {
    ls.getDataRange().getValues().slice(1).forEach(r => {
      if (r[0]) libur[formatWaktuSheet(new Date(r[0]), 'yyyy-MM-dd')] = true;
    });
  }
  const cfg = getAppConfig();
  const jadwal = (cfg.success && cfg.data.jadwal_harian) ? cfg.data.jadwal_harian : null;
  const liburRutin = tgl => {
    const key = formatWaktuSheet(new Date(tgl + 'T12:00:00'), 'u');
    if (!jadwal) return key === '7';
    const sc = jadwal[key];
    return !sc || sc.libur === true;
  };

  const hari = {}, hadir = {};
  const as = ss.getSheetByName('absensi');
  if (as) {
    const d = as.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (!d[i][0]) continue;
      const tgl = formatWaktuSheet(new Date(d[i][0]), 'yyyy-MM-dd');
      if (tgl > today) continue;
      if (mulai && tgl < mulai) continue;
      if (akhir && tgl > akhir) continue;
      if (libur[tgl] || liburRutin(tgl)) continue;
      hari[tgl] = true;
      if (String(d[i][7]).trim() === 'Hadir') {
        const n = nl_nisn_(d[i][1]);
        (hadir[n] = hadir[n] || {})[tgl] = true;
      }
    }
  }
  const jumlahHari = Object.keys(hari).length;
  return function (nisn) {
    if (!jumlahHari) return { ada: false };
    const h = hadir[nisn] ? Object.keys(hadir[nisn]).length : 0;
    return { ada: true, hadir: h, hari: jumlahHari, nilai: Math.min(100, Math.round(h / jumlahHari * 100)) };
  };
}

// Sinkron form terbaru (dibatasi sekali per 60 detik)
function nl_autoSinkron_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('autosync')) return;
  cache.put('autosync', '1', 60);
  try { sinkronOtomatis(); } catch (e) {}
}

// ---------------------------------------------------------------------------
// KUIS — SISI GURU / ADMIN
// ---------------------------------------------------------------------------
function kuisBuat(token, data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const s = nl_session_(token, ['guru', 'admin']);

    const judul = String(data.judul || '').trim();
    const mapel = String(data.mapel || '').trim();
    const kelas = String(data.kelas || '').trim();
    const jenis = String(data.jenis || '').trim();
    const durasi = Math.max(0, parseInt(data.durasi, 10) || 0);

    if (!judul) throw new Error('Judul kuis wajib diisi.');
    if (!kelas) throw new Error('Pilih kelas tujuan kuis.');
    if (JENIS_NILAI.indexOf(jenis) === -1) throw new Error('Jenis nilai tidak valid.');
    if (kelas.toUpperCase() !== 'SEMUA' && !nl_kelasBoleh_(s, kelas)) throw new Error('Anda tidak punya akses ke kelas ' + kelas + '.');
    if (kelas.toUpperCase() === 'SEMUA' && s.role === 'guru' && nl_kelasGuru_(s)) throw new Error('Pilih salah satu kelas Anda, bukan "Semua Kelas".');

    const soalIn = data.soal || [];
    if (!soalIn.length) throw new Error('Tambahkan minimal 1 soal.');
    if (soalIn.length > 100) throw new Error('Maksimal 100 soal per kuis.');

    const soal = soalIn.map((q, i) => {
      const opsi = (q.opsi || []).map(o => String(o || '').trim());
      const kunci = parseInt(q.kunci, 10);
      const terisi = opsi.filter(Boolean).length;
      if (!String(q.q || '').trim()) throw new Error('Soal ' + (i + 1) + ' belum ada pertanyaannya.');
      if (terisi < 2) throw new Error('Soal ' + (i + 1) + ' minimal punya 2 pilihan jawaban.');
      if (isNaN(kunci) || !opsi[kunci]) throw new Error('Soal ' + (i + 1) + ': kunci jawaban belum dipilih / pilihannya kosong.');
      return { q: String(q.q).trim(), opsi: opsi, kunci: kunci, penjelasan: String(q.penjelasan || '').trim() };
    });

    const json = JSON.stringify(soal);
    if (json.length > 45000) throw new Error('Soal terlalu panjang untuk satu sel. Kurangi jumlah/teks soal.');

    const id = 'K' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
    nl_sheet_('kuis').appendRow([
      id, judul, mapel, kelas, jenis, durasi, 'Aktif', s.identifier,
      formatWaktuSheet(new Date(), 'yyyy-MM-dd HH:mm:ss'), json
    ]);
    return { success: true, message: 'Kuis "' + judul + '" berhasil dibuat dan langsung aktif untuk siswa.', id: id };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function kuisListGuru(token) {
  try {
    const s = nl_session_(token, ['guru', 'admin']);
    const kuis = nl_sheet_('kuis').getDataRange().getValues();
    const hk = nl_sheet_('hasil_kuis').getDataRange().getValues();
    const count = {};
    for (let i = 1; i < hk.length; i++) { count[hk[i][1]] = (count[hk[i][1]] || 0) + 1; }

    const list = [];
    for (let i = 1; i < kuis.length; i++) {
      if (!kuis[i][0]) continue;
      if (s.role === 'guru' && nl_nisn_(kuis[i][7]) !== s.identifier) continue;
      let jml = 0;
      try { jml = JSON.parse(kuis[i][9]).length; } catch (e) {}
      list.push({
        id: kuis[i][0], judul: kuis[i][1], mapel: kuis[i][2], kelas: kuis[i][3], jenis: kuis[i][4],
        durasi: kuis[i][5], status: kuis[i][6], dibuatOleh: kuis[i][7], tanggal: nl_fmt_(kuis[i][8]),
        jumlahSoal: jml, dikerjakan: count[kuis[i][0]] || 0
      });
    }
    list.reverse(); // terbaru di atas
    return { success: true, data: list };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Cari baris kuis & pastikan pemiliknya (atau admin)
function nl_cariKuisMilik_(s, id) {
  const sh = nl_sheet_('kuis');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      if (s.role === 'guru' && nl_nisn_(data[i][7]) !== s.identifier) throw new Error('Kuis ini bukan milik Anda.');
      return { sh: sh, row: i + 1, data: data[i] };
    }
  }
  throw new Error('Kuis tidak ditemukan.');
}

function kuisSetStatus(token, id, status) {
  try {
    const s = nl_session_(token, ['guru', 'admin']);
    if (['Aktif', 'Tutup'].indexOf(status) === -1) throw new Error('Status tidak valid.');
    const k = nl_cariKuisMilik_(s, id);
    k.sh.getRange(k.row, 7).setValue(status);
    return { success: true, message: 'Kuis sekarang ' + (status === 'Aktif' ? 'dibuka' : 'ditutup') + '.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Hapus kuis. Hasil pengerjaan siswa yang sudah masuk TIDAK ikut terhapus
// (nilai di rekap tetap aman).
function kuisHapus(token, id) {
  try {
    const s = nl_session_(token, ['guru', 'admin']);
    const k = nl_cariKuisMilik_(s, id);
    k.sh.deleteRow(k.row);
    return { success: true, message: 'Kuis dihapus. Nilai siswa yang sudah masuk tetap tersimpan di rekap.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function kuisHasilList(token, id) {
  try {
    const s = nl_session_(token, ['guru', 'admin']);
    const k = nl_cariKuisMilik_(s, id);
    const hk = nl_sheet_('hasil_kuis').getDataRange().getValues();
    const list = [];
    for (let i = 1; i < hk.length; i++) {
      if (hk[i][1] !== id) continue;
      list.push({
        waktu: nl_fmt_(hk[i][0]), nisn: nl_nisn_(hk[i][4]), nama: hk[i][5], kelas: hk[i][6],
        score: hk[i][7], benar: hk[i][8], salah: hk[i][9], total: hk[i][10]
      });
    }
    list.sort((a, b) => String(a.kelas).localeCompare(String(b.kelas)) || String(a.nama).localeCompare(String(b.nama)));
    return { success: true, judul: k.data[1], data: list };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ---------------------------------------------------------------------------
// KUIS — SISI SISWA
// ---------------------------------------------------------------------------
function kuisListSiswa(token) {
  try {
    const s = nl_session_(token, ['siswa']);
    const st = nl_loadSiswa_()[s.identifier];
    if (!st) throw new Error('Data siswa tidak ditemukan.');

    const hk = nl_sheet_('hasil_kuis').getDataRange().getValues();
    const sudah = {};
    for (let i = 1; i < hk.length; i++) {
      if (nl_nisn_(hk[i][4]) === s.identifier) {
        sudah[hk[i][1]] = { score: hk[i][7], benar: hk[i][8], salah: hk[i][9], total: hk[i][10] };
      }
    }

    const kuis = nl_sheet_('kuis').getDataRange().getValues();
    const list = [];
    for (let i = 1; i < kuis.length; i++) {
      if (!kuis[i][0]) continue;
      const kelasKuis = String(kuis[i][3]).trim();
      const cocok = kelasKuis.toUpperCase() === 'SEMUA' ||
        kelasKuis.split(',').map(nl_norm_).indexOf(nl_norm_(st.kelas)) !== -1;
      if (!cocok) continue;
      const hasil = sudah[kuis[i][0]] || null;
      // Kuis yang sudah ditutup hanya tampil kalau siswa sudah mengerjakannya
      if (kuis[i][6] !== 'Aktif' && !hasil) continue;
      let jml = 0;
      try { jml = JSON.parse(kuis[i][9]).length; } catch (e) {}
      list.push({
        id: kuis[i][0], judul: kuis[i][1], mapel: kuis[i][2], jenis: kuis[i][4],
        durasi: kuis[i][5], status: kuis[i][6], jumlahSoal: jml, hasil: hasil
      });
    }
    list.reverse();
    return { success: true, data: list };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Ambil soal TANPA kunci jawaban
function kuisAmbilSoal(token, id) {
  try {
    const s = nl_session_(token, ['siswa']);
    const st = nl_loadSiswa_()[s.identifier];
    if (!st) throw new Error('Data siswa tidak ditemukan.');

    const kuis = nl_sheet_('kuis').getDataRange().getValues();
    let row = null;
    for (let i = 1; i < kuis.length; i++) { if (kuis[i][0] === id) { row = kuis[i]; break; } }
    if (!row) throw new Error('Kuis tidak ditemukan.');
    if (row[6] !== 'Aktif') throw new Error('Kuis ini sudah ditutup oleh guru.');

    const kelasKuis = String(row[3]).trim();
    const cocok = kelasKuis.toUpperCase() === 'SEMUA' ||
      kelasKuis.split(',').map(nl_norm_).indexOf(nl_norm_(st.kelas)) !== -1;
    if (!cocok) throw new Error('Kuis ini bukan untuk kelas Anda.');

    const hk = nl_sheet_('hasil_kuis').getDataRange().getValues();
    for (let i = 1; i < hk.length; i++) {
      if (hk[i][1] === id && nl_nisn_(hk[i][4]) === s.identifier) throw new Error('Anda sudah mengerjakan kuis ini.');
    }

    // Catat waktu mulai (untuk batas durasi)
    const cache = CacheService.getScriptCache();
    const key = 'mulai_' + id + '_' + s.identifier;
    if (!cache.get(key)) cache.put(key, String(Date.now()), 21600);

    const soal = JSON.parse(row[9]).map(q => ({ q: q.q, opsi: q.opsi }));
    return { success: true, id: id, judul: row[1], mapel: row[2], jenis: row[4], durasi: row[5], soal: soal };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Koreksi di server, simpan, lalu langsung tampil di Rekap Nilai
function kuisKirimJawaban(token, id, jawaban) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const s = nl_session_(token, ['siswa']);
    const st = nl_loadSiswa_()[s.identifier];
    if (!st) throw new Error('Data siswa tidak ditemukan.');

    const kuis = nl_sheet_('kuis').getDataRange().getValues();
    let row = null;
    for (let i = 1; i < kuis.length; i++) { if (kuis[i][0] === id) { row = kuis[i]; break; } }
    if (!row) throw new Error('Kuis tidak ditemukan.');

    const hkSheet = nl_sheet_('hasil_kuis');
    const hk = hkSheet.getDataRange().getValues();
    for (let i = 1; i < hk.length; i++) {
      if (hk[i][1] === id && nl_nisn_(hk[i][4]) === s.identifier) throw new Error('Jawaban Anda sudah tercatat sebelumnya.');
    }

    // Batas durasi (toleransi 60 detik)
    const durasi = Number(row[5]) || 0;
    if (durasi > 0) {
      const mulai = CacheService.getScriptCache().get('mulai_' + id + '_' + s.identifier);
      if (mulai && (Date.now() - Number(mulai)) > (durasi * 60 + 60) * 1000) {
        throw new Error('Waktu pengerjaan sudah habis.');
      }
    }

    const soal = JSON.parse(row[9]);
    let benar = 0;
    const review = soal.map((q, i) => {
      const j = (jawaban && jawaban[i] !== undefined && jawaban[i] !== null) ? parseInt(jawaban[i], 10) : -1;
      const ok = j === q.kunci;
      if (ok) benar++;
      return { q: q.q, opsi: q.opsi, jawaban: j, kunci: q.kunci, benar: ok, penjelasan: q.penjelasan || '' };
    });
    const total = soal.length;
    const salah = total - benar;
    const score = total ? nl_round_(benar / total * 100) : 0;

    hkSheet.appendRow([
      formatWaktuSheet(new Date(), 'yyyy-MM-dd HH:mm:ss'), id, row[1], row[4],
      "'" + st.nisn, st.nama, st.kelas, score, benar, salah, total,
      JSON.stringify((jawaban || []).map(x => (x === null || x === undefined) ? -1 : x))
    ]);

    return { success: true, judul: row[1], jenis: row[4], score: score, benar: benar, salah: salah, total: total, review: review };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// NILAI MANUAL — input guru
// list: [{ nisn, nilai }]  — hanya yang terisi
// ---------------------------------------------------------------------------
function nilaiManualSimpanBatch(token, jenis, keterangan, list) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const s = nl_session_(token, ['guru', 'admin']);
    if (JENIS_NILAI.indexOf(jenis) === -1) throw new Error('Jenis nilai tidak valid.');
    if (!list || !list.length) throw new Error('Belum ada nilai yang diisi.');

    const siswa = nl_loadSiswa_();
    const waktu = formatWaktuSheet(new Date(), 'yyyy-MM-dd HH:mm:ss');
    const rows = [];
    list.forEach(it => {
      const n = nl_nisn_(it.nisn);
      const v = Number(it.nilai);
      const st = siswa[n];
      if (!st || isNaN(v) || v < 0 || v > 100) return;
      if (!nl_kelasBoleh_(s, st.kelas)) return;
      rows.push([waktu, "'" + st.nisn, st.nama, st.kelas, jenis, v, String(keterangan || '').trim(), s.identifier]);
    });
    if (!rows.length) throw new Error('Tidak ada nilai valid (harus angka 0–100 dan siswa dari kelas Anda).');

    const sh = nl_sheet_('nilai_manual');
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { success: true, message: rows.length + ' nilai ' + jenis + ' berhasil disimpan.' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// REKAP NILAI
// ---------------------------------------------------------------------------
function getRekapNilai(token, kelas, mulai, akhir) {
  try {
    const s = nl_session_(token, ['admin', 'guru', 'wakel']);

    if (kelas && !nl_kelasBoleh_(s, kelas)) {
      throw new Error('Anda tidak punya akses ke kelas ' + kelas + '.');
    }

    mulai = nl_tgl_(mulai);
    akhir = nl_tgl_(akhir);

    nl_autoSinkron_();

    const akses = nl_kelasGuru_(s);
    const filter = kelas
      ? [nl_norm_(kelas)]
      : (akses ? akses.map(nl_norm_) : []);

    const siswa = nl_loadSiswa_();
    const recMap = nl_recordsByNisn_();
    const dis = nl_disiplinMap_(mulai, akhir);
    const soalMap = nl_soalMap_(mulai, akhir);

    const rows = [];

    Object.keys(siswa).forEach(n => {
      const st = siswa[n];

      if (
        filter.length &&
        filter.indexOf(nl_norm_(st.kelas)) === -1
      ) {
        return;
      }

      const r = nl_komponen_(
        recMap[n] || [],
        dis(n),
        mulai,
        akhir
      );

      rows.push({
        nisn: n,
        nama: st.nama,
        kelas: st.kelas,
        komponen: r.komponen,
        total: r.total,
        akhir: r.akhir,
        akhirAda: r.akhirAda,
        lengkap: r.lengkap,
        kurang: r.kurang,
        waktu: r.waktuKuis,
        soal: soalMap[n] || null
      });
    });

    rows.sort(
      (a, b) =>
        a.kelas.localeCompare(b.kelas) ||
        a.nama.localeCompare(b.nama)
    );

    // Gunakan konfigurasi nilai yang aktif.
    const cfg = nl_pengaturan_();

    return {
      success: true,
      data: rows,
      bobot: cfg.bobot,
      mode: cfg.mode,
      kkm: cfg.kkm
    };

  } catch (e) {
    return {
      success: false,
      message: e.message
    };
  }
}

// Rincian benar/salah per soal milik satu siswa (dari sheet 'detail_soal').
// Kunci jawaban hanya dikirim ke guru/admin, tidak ke siswa.
function nl_rincianSoal_(nisn, denganKunci, mulai, akhir) {
  const out = [];
  const data = nl_sheet_('detail_soal').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (nl_nisn_(data[i][3]) !== nisn) continue;
    const waktu = nl_fmt_(data[i][0]);
    const t = waktu.substring(0, 10);
    if ((mulai && t < mulai) || (akhir && t > akhir)) continue;
    let det = [];
    try { det = JSON.parse(data[i][12] || '[]'); } catch (e) {}
    if (!denganKunci) det = det.map(x => ({ n: x.n, s: x.s, j: x.j, p: x.p, m: x.m, t: x.t }));
    out.push({ judul: String(data[i][2]), waktu: waktu, benar: Number(data[i][6]) || 0, salah: Number(data[i][7]) || 0,
      sebagian: Number(data[i][8]) || 0, total: Number(data[i][9]) || 0,
      noBenar: String(data[i][10] || ''), noSalah: String(data[i][11] || ''), detail: det });
  }
  out.sort((a, b) => String(b.waktu).localeCompare(String(a.waktu)));
  return out;
}

// Ringkasan jumlah soal benar/salah per siswa (dari 'detail_soal') pada periode yang dipilih
function nl_soalMap_(mulai, akhir) {
  const map = {};
  const data = nl_sheet_('detail_soal').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const t = nl_fmt_(data[i][0]).substring(0, 10);
    if ((mulai && t < mulai) || (akhir && t > akhir)) continue;
    const n = nl_nisn_(data[i][3]);
    const m = map[n] = map[n] || { benar: 0, salah: 0, sebagian: 0, total: 0 };
    m.benar += Number(data[i][6]) || 0;
    m.salah += Number(data[i][7]) || 0;
    m.sebagian += Number(data[i][8]) || 0;
    m.total += Number(data[i][9]) || 0;
  }
  return map;
}

// Detail 1 siswa: ringkasan + riwayat + rincian soal. Siswa hanya bisa melihat dirinya sendiri.
function getDetailNilaiSiswa(token, nisn, mulai, akhir) {
  try {
    const s = nl_session_(token, ['admin', 'guru', 'wakel', 'siswa']);
    if (s.role === 'siswa') nisn = s.identifier;
    nisn = nl_nisn_(nisn);
    mulai = nl_tgl_(mulai); akhir = nl_tgl_(akhir);

    const st = nl_loadSiswa_()[nisn];
    if (!st) throw new Error('Data siswa tidak ditemukan.');
    if (s.role === 'guru' && !nl_kelasBoleh_(s, st.kelas)) throw new Error('Siswa ini bukan dari kelas Anda.');

    const recs = nl_filterRange_(nl_recordsByNisn_()[nisn] || [], mulai, akhir);
    const ringkasan = nl_komponen_(recs, nl_disiplinMap_(mulai, akhir)(nisn), mulai, akhir);
    const riwayat = recs.slice().sort((a, b) => String(b.waktu).localeCompare(String(a.waktu)));
    const rincian = nl_rincianSoal_(nisn, s.role !== 'siswa', mulai, akhir);
    return { success: true, siswa: st, ringkasan: ringkasan, riwayat: riwayat, rincian: rincian };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ---------------------------------------------------------------------------
// EXPORT EXCEL REKAP NILAI
// ---------------------------------------------------------------------------
function generateExcelRekapNilai(token, kelas, mulai, akhir) {
  try {
    const res = getRekapNilai(token, kelas, mulai, akhir);
    if (!res.success) return res;

    const fileName = 'Rekap Nilai' + (kelas ? ' - ' + kelas : '') + (mulai ? ' ' + mulai + ' s.d ' + akhir : '') + ' - ' + formatWaktuSheet(new Date(), 'ddMMyyyy_HHmm');
    const ss = SpreadsheetApp.create(fileName);
    const sheet = ss.getSheets()[0];
    sheet.setName('Rekap Nilai');

    const cfgX = nl_pengaturan_();
    const pct = k => Math.round(cfgX.bobot[k] * 100);
    const header = ['No', 'Nama', 'NISN', 'Kelas',
                    'Score',
                    'Nilai Pemahaman & Pengetahuan (' + pct('Harian') + '%)',
                    'Nilai Analisa (' + pct('Analisa') + '%)',
                    'Refleksi',
                    'Nilai Disiplin (' + pct('Disiplin') + '%)',
                    'Waktu Mengerjakan Kuis', 'Total', 'Nilai Akhir', 'Status'];
    sheet.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#1B5E20').setFontColor('#FFFFFF')
      .setHorizontalAlignment('center').setWrap(true);

    const v = x => x.ada ? (x.teks || x.nilai) : '';
    const rows = res.data.map((d, i) => [
      i + 1, d.nama, "'" + d.nisn, d.kelas,
      v(d.komponen.Score),
      v(d.komponen.Harian), v(d.komponen.Analisa), v(d.komponen.Refleksi),
      d.komponen.Disiplin.ada ? d.komponen.Disiplin.nilai + '%' : '',
      d.waktu || '', d.total > 0 ? d.total : '', d.akhirAda ? d.akhir : '',
      d.lengkap ? 'Lengkap' : (d.kurang.length === 3 ? 'Belum ada nilai' : 'Belum: ' + d.kurang.join(', '))
    ]);
      if (rows.length) {
      const kolomCek = [6, 7, 12]; // Pemahaman, Analisa, Nilai Akhir
      rows.forEach((r, i) => {
        kolomCek.forEach(c => {
          const val = r[c - 1];
          if (typeof val === 'number' && val < cfgX.kkm) {
            sheet.getRange(i + 2, c).setFontColor('#B91C1C').setBackground('#FEE2E2').setFontWeight('bold');
          }
        });
      });
      sheet.getRange(rows.length + 3, 1).setValue('KKM = ' + cfgX.kkm + '. Nilai di bawah KKM diberi warna merah.')
        .setFontStyle('italic').setFontColor('#B91C1C');
    }
      sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
      sheet.getRange(1, 1, rows.length + 1, header.length)
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_THIN);
    sheet.autoResizeColumns(1, header.length);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);

    const fileId = ss.getId();
    try { DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (x) {}
    return { success: true, url: 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx' };
  } catch (e) {
    return { success: false, message: 'Gagal membuat Excel: ' + e.message };
  }
}

// ---------------------------------------------------------------------------
// Dipakai index.html untuk memuat file HTML terpisah: <?!= include('nilai') ?>
// ---------------------------------------------------------------------------
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

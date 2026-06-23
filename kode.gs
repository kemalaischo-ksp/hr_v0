/*************************************************************************************************
 *  AISCHO — SISTEM DASHBOARD & HISTORY SET GAJI 2026
 *  AL-WILDAN ISLAMIC SCHOOL HOLDING
 * -----------------------------------------------------------------------------------------------
 *  Fitur:
 *   1) DASHBOARD STATUS .............. ringkasan hidup (LANJUT/HOLD/BATAL, per batch, per cabang,
 *                                      total THP bersih) yang bisa dilaporkan ke ATASAN.
 *   2) LOG UPDATE STATUS ............. tiap update = 1 ROW BARU ke bawah (riwayat progres).
 *   3) HISTORY STANDAR SET SALARY .... snapshot salary master + (opsi) tiap cabang ke tab riwayat
 *                                      berstandar, untuk dilaporkan ke OWNER.
 *   4) LAPORAN EMAIL ................. kirim ringkasan ke atasan & owner sekali klik.
 *
 *  Cara pakai:
 *   - Buka spreadsheet > Extensions > Apps Script > tempel seluruh file ini > Save.
 *   - Isi bagian CONFIG di bawah (email atasan/owner, daftar cabang).
 *   - Reload spreadsheet > muncul menu "🏫 AISCHO Salary" > klik "⚙️ Setup".
 *************************************************************************************************/

/* =========================== CONFIG — SILAKAN SESUAIKAN =========================== */
const CFG = {
  // ID file data (ambil dari URL spreadsheet: .../d/<ID>/edit).
  // Ini membuat script SELALU membaca file yang benar, di mana pun script dipasang.
  // Kosongkan ('') jika script dipasang langsung di dalam file datanya.
  SPREADSHEET_ID : '1PTGYrrr1YlR6zxVhs2ocRqIOFrx_PFxmQ1u8gQPPKa8',
  // Sumber khusus Database Pengajuan. Kosongkan ('') = pakai spreadsheet utama di atas.
  // Isi ID spreadsheet lain di sini bila data pengajuan ada di file berbeda
  // (akun yang menjalankan web app harus punya akses minimal Viewer/Editor ke file itu).
  PENGAJUAN_SPREADSHEET_ID : '',

  MASTER_SHEET   : 'SET GAJI 2026',          // sumber data utama
  PENGAJUAN_SHEETS: ['Pengajuan Salary', 'Database', 'Pengajuan', 'Form Responses 1', 'Tanggapan Formulir 1'], // tab data pengajuan (foto + Merged Doc URL)
  // Sumber Data Kesehatan (spreadsheet terpisah). Kosongkan ('') untuk menonaktifkan tab Kesehatan.
  KESEHATAN_SPREADSHEET_ID : '1QWofbSTUriQg5gtud6ODWw8xYcWXv6d58lhP_Ya-A8Y',
  KESEHATAN_SHEETS : ['Form Responses 1', 'Tanggapan Formulir 1', 'Respons', 'Form Responses', 'Kesehatan'],
  // Draft email aktivasi SDM
  KONTAK_SPREADSHEET_ID : '1PTV3JznGGYyPSjJbeq184VNyotmrE_8GGJRPNQ054wA', // DBASE EMAIL (FIN & MGT per cabang)
  EMAIL_AKTIVASI_TETAP : ['kemal.aischo@gmail.com'],
  TTD_NAMA  : 'Kemal P',
  TTD_EMAIL : 'Kemal.aischo@gmail.com',
  TTD_HP    : '0857 2945 2911',
  DASHBOARD_SHEET: 'DASHBOARD STATUS',       // dibuat otomatis
  LOG_SHEET      : 'LOG UPDATE STATUS',      // dibuat otomatis (append ke bawah)
  HISTORY_SHEET  : 'HISTORY STANDAR SALARY', // dibuat otomatis (tab baru untuk owner)
  DETAIL_SHEET   : 'DETAIL AKTIF PER CABANG', // dibuat otomatis (klik +/- buka-tutup)

  // Daftar tab cabang (untuk konsolidasi history). Sesuaikan bila ada yang berubah.
  BRANCH_SHEETS  : ['AW1','AW3','AW4','AW5','AW7','AW8','AW10','AW12','AW13','AW14',
                    'AW15','AW16','AW20','AW22','AW23','AW24','AW29','AW32'],

  // Email tujuan laporan (pisahkan dengan koma bila lebih dari satu).
  EMAIL_ATASAN   : 'ganti-email-atasan@alwildan.sch.id',
  EMAIL_OWNER    : 'ganti-email-owner@alwildan.sch.id',

  // Status yang dihitung pada dashboard (urutan menentukan tampilan).
  STATUS_LIST    : ['LANJUT','HOLD','BATAL'],

  // true = nama yang muncul ganda dalam satu sheet hanya dihitung sekali
  // (mengatasi entri duplikat seperti pada blok header kedua).
  DEDUP_BY_NAMA  : true,

  TZ             : 'Asia/Jakarta',
  ORG            : 'AL-WILDAN ISLAMIC SCHOOL HOLDING (AISCHO)'
};
/* ================================================================================= */


/* --------------------------- SPREADSHEET REFERENCE -------------------------------- */
// Selalu mengarah ke file data yang benar (lewat ID bila diisi).
function ss_() {
  return CFG.SPREADSHEET_ID ? SpreadsheetApp.openById(CFG.SPREADSHEET_ID) : SpreadsheetApp.getActive();
}
// Sumber khusus data pengajuan (boleh spreadsheet berbeda dari dashboard).
function ssPengajuan_() {
  return CFG.PENGAJUAN_SPREADSHEET_ID ? SpreadsheetApp.openById(CFG.PENGAJUAN_SPREADSHEET_ID) : ss_();
}
// Sumber khusus data kesehatan.
function ssKesehatan_() {
  return CFG.KESEHATAN_SPREADSHEET_ID ? SpreadsheetApp.openById(CFG.KESEHATAN_SPREADSHEET_ID) : ss_();
}


/* ----------------------------------- MENU ---------------------------------------- */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🏫 AISCHO Salary')
      .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
      .addItem('🔍 Detail Aktif per Cabang (klik buka-tutup)', 'buildBranchDetail')
      .addItem('➕ Catat Update Status (Append Log)', 'appendStatusLog')
      .addSeparator()
      .addItem('🗂️ Snapshot → History Standar Salary', 'snapshotMasterToHistory')
      .addItem('🏬 Snapshot Termasuk Tiap Cabang', 'snapshotAllToHistory')
      .addSeparator()
      .addItem('📨 Laporkan Dashboard ke Atasan', 'reportToAtasan')
      .addItem('👑 Laporkan History ke Owner', 'reportToOwner')
      .addItem('⚡ Update + Lapor ke Atasan (1 klik)', 'updateAndReport')
      .addSeparator()
      .addItem('⚙️ Setup (buat sheet + trigger harian)', 'setup')
      .addSeparator()
      .addItem('🌐 Lihat Link Web App (real-time)', 'showWebAppUrl')
      .addToUi();
  } catch (e) {
    // getUi() hanya tersedia saat spreadsheet dibuka langsung; abaikan di konteks lain.
  }
}


/* =============================== SETUP & TRIGGER ================================== */
function setup() {
  getOrCreateSheet_(CFG.DASHBOARD_SHEET);
  ensureLogHeader_();
  ensureHistoryHeader_();
  refreshDashboard();

  // Pasang trigger harian (07:00) bila belum ada.
  const exists = ScriptApp.getProjectTriggers()
    .some(function(t){ return t.getHandlerFunction() === 'dailyAutoUpdate'; });
  if (!exists) {
    ScriptApp.newTrigger('dailyAutoUpdate')
      .timeBased().atHour(7).everyDays(1)
      .inTimezone(CFG.TZ).create();
  }
  ss_().toast('Setup selesai. Menu siap dipakai.', 'AISCHO', 5);
}

// Dipanggil otomatis tiap hari: refresh + catat log (tanpa kirim email).
function dailyAutoUpdate() {
  refreshDashboard();
  appendStatusLog('AUTO — trigger harian');
}


/* ============================ BACA DATA (GENERIC) ================================ */
/**
 * Membaca sheet apa pun yang memiliki baris header berisi NAMA & CABANG.
 * Header boleh berulang (batch) — baris pengulangan & baris kosong otomatis dilewati.
 * @return {{records: Object[], headerRow:number}}
 */
function readSalarySheet_(sheetName) {
  const sh = findSheet_(sheetName);
  if (!sh) return { records: [], headerRow: -1, reason: 'SHEET_NOT_FOUND' };

  const values = sh.getDataRange().getValues();
  let hRow = -1, map = null;

  for (let r = 0; r < Math.min(values.length, 60); r++) {
    const m = buildHeaderMap_(values[r]);
    if (m && m.nama != null && m.cabang != null) { hRow = r; map = m; break; }
  }
  if (hRow === -1) return { records: [], headerRow: -1, reason: 'HEADER_NOT_FOUND', sheetName: sh.getName() };

  const records = [];
  const seen = {};
  for (let r = hRow + 1; r < values.length; r++) {
    const row = values[r];
    const nama = (map.nama != null ? row[map.nama] : '').toString().trim();
    if (!nama) continue;                       // lewati baris kosong
    if (nama.toUpperCase() === 'NAMA') continue; // lewati pengulangan header

    if (CFG.DEDUP_BY_NAMA) {
      const key = nama.toUpperCase().replace(/\s+/g, ' ');
      if (seen[key]) continue;                 // nama ganda dalam sheet ini -> 1x saja
      seen[key] = true;
    }

    records.push({
      sumber   : sheetName,
      email    : pick_(row, map.email),
      wa       : pick_(row, map.wa),
      cv       : pick_(row, map.cv),
      kesehatan: pick_(row, map.kesehatan),
      batch    : pick_(row, map.batch),
      status   : pick_(row, map.status).toString().toUpperCase().trim(),
      no       : pick_(row, map.no),
      nama     : nama,
      cabang   : pick_(row, map.cabang),
      s1       : pick_(row, map.s1),
      s2       : pick_(row, map.s2),
      mapel    : pick_(row, map.mapel),
      pengajuan: pick_(row, map.pengajuan),
      rekom    : pick_(row, map.rekom),
      thpKotor : parseRupiah_(pick_(row, map.thpKotor)),
      konfirm  : pick_(row, map.konfirm),
      thpBersih: parseRupiah_(pick_(row, map.thpBersih)),
      total    : parseRupiah_(pick_(row, map.total)),
      tk       : parseRupiah_(pick_(row, map.tk)),
      thr      : parseRupiah_(pick_(row, map.thr)),
      nego     : pick_(row, map.nego)
    });
  }
  return { records: records, headerRow: hRow, reason: 'OK', sheetName: sh.getName() };
}

function buildHeaderMap_(rowArr) {
  const map = {};
  let hit = 0;
  rowArr.forEach(function(cell, i){
    const h = (cell == null ? '' : cell.toString()).toUpperCase().trim();
    if (!h) return;
    if (h === 'EMAIL') { map.email = i; hit++; }
    else if (h === 'NO WA' || h === 'NOWA' || h === 'NO. WA') { map.wa = i; }
    else if (h === 'CV') { map.cv = i; }
    else if (h === 'KESEHATAN') { map.kesehatan = i; }
    else if (h === 'BATCH') { map.batch = i; }
    else if (h === 'STATUS') { map.status = i; }
    else if (h === 'NO') { map.no = i; }
    else if (h === 'NAMA') { map.nama = i; hit++; }
    else if (h === 'CABANG') { map.cabang = i; hit++; }
    else if (h === 'S1') { map.s1 = i; }
    else if (h === 'S2') { map.s2 = i; }
    else if (h.indexOf('MAPEL') > -1 || h.indexOf('SUBJEK') > -1) { map.mapel = i; }
    else if (h === 'PENGAJUAN') { map.pengajuan = i; }
    else if (h.indexOf('ACC') > -1 || h.indexOf('REKOMENDASI') > -1 || h.indexOf('PRESDIR') > -1) { map.rekom = i; }
    else if (h.indexOf('THP KOTOR') > -1) { map.thpKotor = i; }
    else if (h === 'KONFIRMASI') { map.konfirm = i; }
    else if (h.indexOf('THP BERSIH') > -1) { map.thpBersih = i; }
    else if (h.indexOf('TOTAL') > -1) { map.total = i; }
    else if (h === 'TK') { map.tk = i; }
    else if (h.indexOf('THR') > -1) { map.thr = i; }
    else if (h.indexOf('NEGO') > -1) { map.nego = i; }
  });
  return hit >= 2 ? map : null;
}

function pick_(row, idx){ return (idx == null) ? '' : (row[idx] == null ? '' : row[idx]); }

// "Rp5.884.800" / "-Rp384.800" / 5884800  -> number
function parseRupiah_(v) {
  if (typeof v === 'number') return Math.round(v);
  if (v == null) return 0;
  let s = String(v).trim();
  if (!s) return 0;
  const neg = /^-/.test(s) || /\(.*\)/.test(s);
  const digits = s.replace(/[^0-9]/g, '');
  if (!digits) return 0;
  const n = parseInt(digits, 10);
  return neg ? -n : n;
}


/* =============================== RINGKASAN ====================================== */
function computeSummary_(records) {
  const sum = {
    total: records.length,
    byStatus: {}, byBatch: {}, byCabang: {},
    thpLanjut: 0, nLanjut: 0
  };
  CFG.STATUS_LIST.forEach(function(s){ sum.byStatus[s] = 0; });
  sum.byStatus['LAINNYA'] = 0;

  records.forEach(function(r){
    if (sum.byStatus[r.status] != null) sum.byStatus[r.status]++;
    else sum.byStatus['LAINNYA']++;

    const b = (r.batch || '-').toString().trim() || '-';
    sum.byBatch[b] = (sum.byBatch[b] || 0) + 1;

    const c = normCabang_(r.cabang);
    sum.byCabang[c] = sum.byCabang[c] || { lanjut:0, hold:0, batal:0, total:0, thp:0 };
    sum.byCabang[c].total++;
    if (r.status === 'LANJUT') { sum.byCabang[c].lanjut++; sum.byCabang[c].thp += r.thpBersih; sum.thpLanjut += r.thpBersih; sum.nLanjut++; }
    else if (r.status === 'HOLD') sum.byCabang[c].hold++;
    else if (r.status === 'BATAL') sum.byCabang[c].batal++;
  });
  sum.avgLanjut = sum.nLanjut ? Math.round(sum.thpLanjut / sum.nLanjut) : 0;
  return sum;
}

// Ambil kode cabang utama: "AW16 (FULL...)" -> "AW16", "AW12,13,15" -> "AW12,13,15"
function normCabang_(c) {
  let s = (c == null ? '' : c.toString()).trim();
  if (!s) return '(kosong)';
  s = s.replace(/\s*\(.*?\)\s*/g, '').trim(); // buang keterangan dalam kurung
  return s || '(kosong)';
}


/* =============================== DASHBOARD ====================================== */
function refreshDashboard() {
  const data = readSalarySheet_(CFG.MASTER_SHEET);
  if (data.headerRow === -1) { SpreadsheetApp.getUi().alert('Sheet "'+CFG.MASTER_SHEET+'" / header tidak ditemukan.'); return; }
  const s = computeSummary_(data.records);
  const sh = getOrCreateSheet_(CFG.DASHBOARD_SHEET);
  sh.clear();

  const now = Utilities.formatDate(new Date(), CFG.TZ, 'EEEE, dd MMM yyyy — HH:mm');
  const rows = [];
  rows.push(['DASHBOARD STATUS — SET GAJI 2026']);
  rows.push([CFG.ORG]);
  rows.push(['Diperbarui: ' + now]);
  rows.push(['']);

  rows.push(['RINGKASAN STATUS','Jumlah','% dari Total']);
  CFG.STATUS_LIST.concat(['LAINNYA']).forEach(function(st){
    const n = s.byStatus[st] || 0;
    if (st === 'LAINNYA' && n === 0) return;
    rows.push([st, n, s.total ? (n/s.total) : 0]);
  });
  rows.push(['TOTAL KANDIDAT', s.total, 1]);
  rows.push(['']);

  rows.push(['FINANSIAL (status LANJUT)','Nilai']);
  rows.push(['Jumlah pegawai LANJUT', s.nLanjut]);
  rows.push(['Total THP Bersih / bulan', s.thpLanjut]);
  rows.push(['Rata-rata THP Bersih', s.avgLanjut]);
  rows.push(['Estimasi beban / tahun (×12)', s.thpLanjut * 12]);
  rows.push(['']);

  rows.push(['PER BATCH','Jumlah']);
  Object.keys(s.byBatch).sort().forEach(function(b){ rows.push([b, s.byBatch[b]]); });
  rows.push(['']);

  rows.push(['PER CABANG','LANJUT','HOLD','BATAL','TOTAL','THP Bersih LANJUT']);
  Object.keys(s.byCabang).sort().forEach(function(c){
    const x = s.byCabang[c];
    rows.push([c, x.lanjut, x.hold, x.batal, x.total, x.thp]);
  });

  // tulis
  const maxCols = rows.reduce(function(m,r){ return Math.max(m, r.length); }, 1);
  rows.forEach(function(r){ while (r.length < maxCols) r.push(''); });
  sh.getRange(1,1,rows.length,maxCols).setValues(rows);

  // format ringan
  sh.getRange('A1').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2:A3').setFontColor('#555555');
  styleHeaderRow_(sh, rows, 'RINGKASAN STATUS', maxCols);
  styleHeaderRow_(sh, rows, 'FINANSIAL (status LANJUT)', maxCols);
  styleHeaderRow_(sh, rows, 'PER BATCH', maxCols);
  styleHeaderRow_(sh, rows, 'PER CABANG', maxCols);
  formatRupiahCells_(sh, rows);
  formatPercentCells_(sh, rows);
  sh.setColumnWidth(1, 230);
  for (let c = 2; c <= maxCols; c++) sh.setColumnWidth(c, 140);
  sh.setFrozenRows(3);

  ss_().toast('Dashboard diperbarui.', 'AISCHO', 4);
  return s;
}

function styleHeaderRow_(sh, rows, label, maxCols) {
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === label) {
      sh.getRange(i+1, 1, 1, maxCols)
        .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f6f43');
      return;
    }
  }
}
function formatRupiahCells_(sh, rows) {
  for (let i = 0; i < rows.length; i++) {
    const label = (rows[i][0] || '').toString();
    if (/THP|beban|THR|TK|Total THP/i.test(label) && typeof rows[i][1] === 'number') {
      sh.getRange(i+1, 2).setNumberFormat('"Rp"#,##0');
    }
    if (rows[i][0] === 'PER CABANG') {
      // kolom THP bersih di tabel cabang
      for (let j = i+1; j < rows.length; j++){
        if (!rows[j][0]) break;
        sh.getRange(j+1, 6).setNumberFormat('"Rp"#,##0');
      }
    }
  }
}
function formatPercentCells_(sh, rows) {
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === 'RINGKASAN STATUS') {
      for (let j = i+1; j < rows.length; j++){
        if (!rows[j][0]) break;
        sh.getRange(j+1, 3).setNumberFormat('0.0%');
      }
    }
  }
}


/* ============================= LOG (APPEND KE BAWAH) ============================= */
function ensureLogHeader_() {
  const sh = getOrCreateSheet_(CFG.LOG_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp','Total Kandidat','LANJUT','HOLD','BATAL',
                  'Total THP Bersih (LANJUT)','Rata2 THP','Jml Cabang Aktif','Catatan','Oleh']);
    sh.getRange(1,1,1,10).setFontWeight('bold').setBackground('#1f6f43').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Tiap dipanggil = 1 ROW BARU di bawah (riwayat progres status).
function appendStatusLog(catatan) {
  const data = readSalarySheet_(CFG.MASTER_SHEET);
  const s = computeSummary_(data.records);
  const sh = ensureLogHeader_();
  const cabangAktif = Object.keys(s.byCabang).filter(function(c){ return s.byCabang[c].lanjut > 0; }).length;
  const ts = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm:ss');

  sh.appendRow([ ts, s.total, s.byStatus['LANJUT']||0, s.byStatus['HOLD']||0, s.byStatus['BATAL']||0,
                 s.thpLanjut, s.avgLanjut, cabangAktif,
                 (typeof catatan === 'string' ? catatan : 'Update manual'),
                 Session.getActiveUser().getEmail() ]);
  const last = sh.getLastRow();
  sh.getRange(last, 6).setNumberFormat('"Rp"#,##0');
  sh.getRange(last, 7).setNumberFormat('"Rp"#,##0');
  ss_().toast('Log status dicatat (row '+last+').', 'AISCHO', 4);
}


/* ===================== HISTORY STANDAR SALARY (untuk OWNER) ====================== */
function ensureHistoryHeader_() {
  const sh = getOrCreateSheet_(CFG.HISTORY_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Snapshot ID','Tgl Snapshot','Sumber','Batch','Cabang','Nama','Mapel/Subjek',
                  'S1','S2','THP Kotor','THP Bersih','TK','THR/bln','Status','Konfirmasi']);
    sh.getRange(1,1,1,15).setFontWeight('bold').setBackground('#0b3d2e').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Snapshot dari MASTER saja.
function snapshotMasterToHistory() { return doSnapshot_([CFG.MASTER_SHEET]); }

// Snapshot MASTER + semua tab cabang (konsolidasi salary tiap cabang).
function snapshotAllToHistory() {
  return doSnapshot_([CFG.MASTER_SHEET].concat(CFG.BRANCH_SHEETS));
}

function doSnapshot_(sheetNames) {
  const sh = ensureHistoryHeader_();
  const snapId = 'SNAP-' + Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMMdd-HHmm');
  const tgl = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm');
  const out = [];

  sheetNames.forEach(function(name){
    const data = readSalarySheet_(name);
    data.records.forEach(function(r){
      out.push([ snapId, tgl, r.sumber, r.batch, normCabang_(r.cabang), r.nama, r.mapel,
                 r.s1, r.s2, r.thpKotor, r.thpBersih, r.tk, r.thr, r.status, r.konfirm ]);
    });
  });

  if (!out.length) { SpreadsheetApp.getUi().alert('Tidak ada data untuk di-snapshot.'); return null; }
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, out.length, 15).setValues(out);
  sh.getRange(start, 10, out.length, 4).setNumberFormat('"Rp"#,##0');
  ss_().toast(out.length+' baris di-snapshot ('+snapId+').', 'AISCHO', 5);
  return { snapId: snapId, count: out.length };
}


/* ================================= EMAIL ======================================== */
function reportToAtasan() {
  const s = refreshDashboard();
  if (!s) { SpreadsheetApp.getUi().alert('Dashboard gagal dibuat — cek nama sheet master.'); return; }
  const url = ss_().getUrl() + '#gid=' + getSheetGid_(CFG.DASHBOARD_SHEET);
  const subject = '[AISCHO] Update Status Set Gaji 2026 — ' +
                  Utilities.formatDate(new Date(), CFG.TZ, 'dd MMM yyyy HH:mm');
  MailApp.sendEmail({
    to: CFG.EMAIL_ATASAN, subject: subject,
    htmlBody: buildDashboardEmail_(s, url)
  });
  appendStatusLog('Dilaporkan ke atasan via email');
  ss_().toast('Laporan terkirim ke atasan.', 'AISCHO', 4);
}

function reportToOwner() {
  const snap = snapshotMasterToHistory();
  const data = readSalarySheet_(CFG.MASTER_SHEET);
  const s = computeSummary_(data.records);
  const url = ss_().getUrl() + '#gid=' + getSheetGid_(CFG.HISTORY_SHEET);
  const subject = '[AISCHO] History Standar Set Salary — ' + (snap ? snap.snapId : '');
  MailApp.sendEmail({
    to: CFG.EMAIL_OWNER, subject: subject,
    htmlBody: buildOwnerEmail_(s, snap, url)
  });
  ss_().toast('History terkirim ke owner.', 'AISCHO', 4);
}

// reportToAtasan() sudah melakukan refresh + catat log + kirim email.
function updateAndReport() { reportToAtasan(); }

function buildDashboardEmail_(s, url) {
  const rp = function(n){ return 'Rp' + Number(n||0).toLocaleString('id-ID'); };
  let cab = '';
  Object.keys(s.byCabang).sort().forEach(function(c){
    const x = s.byCabang[c];
    cab += '<tr><td>'+c+'</td><td style="text-align:center">'+x.lanjut+'</td><td style="text-align:center">'+
           x.hold+'</td><td style="text-align:center">'+x.batal+'</td><td style="text-align:right">'+rp(x.thp)+'</td></tr>';
  });
  return ''+
   '<div style="font-family:Arial,sans-serif;color:#222;max-width:680px">'+
   '<h2 style="color:#1f6f43;margin-bottom:0">Update Status — Set Gaji 2026</h2>'+
   '<p style="color:#666;margin-top:4px">'+CFG.ORG+'<br>'+
     Utilities.formatDate(new Date(), CFG.TZ,'EEEE, dd MMM yyyy HH:mm')+' WIB</p>'+
   '<table style="border-collapse:collapse;margin:10px 0"><tr>'+
     kpi_('LANJUT', s.byStatus['LANJUT']||0,'#1f6f43')+
     kpi_('HOLD', s.byStatus['HOLD']||0,'#b8860b')+
     kpi_('BATAL', s.byStatus['BATAL']||0,'#a83232')+
     kpi_('TOTAL', s.total,'#333')+'</tr></table>'+
   '<p><b>Total THP Bersih (LANJUT):</b> '+rp(s.thpLanjut)+' / bulan &nbsp;·&nbsp; '+
     '<b>Rata-rata:</b> '+rp(s.avgLanjut)+' &nbsp;·&nbsp; <b>Estimasi/tahun:</b> '+rp(s.thpLanjut*12)+'</p>'+
   '<h3 style="color:#1f6f43">Rekap Per Cabang</h3>'+
   '<table style="border-collapse:collapse;width:100%;font-size:13px" border="1" cellpadding="6">'+
     '<tr style="background:#1f6f43;color:#fff"><th align="left">Cabang</th><th>LANJUT</th><th>HOLD</th><th>BATAL</th><th>THP Bersih</th></tr>'+
     cab+'</table>'+
   '<p style="margin-top:16px"><a href="'+url+'" style="background:#1f6f43;color:#fff;padding:10px 16px;text-decoration:none;border-radius:6px">Buka Dashboard</a></p>'+
   '<p style="color:#999;font-size:12px">Email otomatis dari sistem dashboard AISCHO.</p></div>';
}

function buildOwnerEmail_(s, snap, url) {
  const rp = function(n){ return 'Rp' + Number(n||0).toLocaleString('id-ID'); };
  return ''+
   '<div style="font-family:Arial,sans-serif;color:#222;max-width:680px">'+
   '<h2 style="color:#0b3d2e">History Standar Set Salary</h2>'+
   '<p style="color:#666">'+CFG.ORG+'<br>Snapshot: <b>'+(snap?snap.snapId:'-')+'</b> · '+
     (snap?snap.count:0)+' baris tersimpan</p>'+
   '<p><b>Pegawai aktif (LANJUT):</b> '+s.nLanjut+'<br>'+
     '<b>Total THP Bersih:</b> '+rp(s.thpLanjut)+' / bulan<br>'+
     '<b>Estimasi beban tahunan:</b> '+rp(s.thpLanjut*12)+'</p>'+
   '<p><a href="'+url+'" style="background:#0b3d2e;color:#fff;padding:10px 16px;text-decoration:none;border-radius:6px">Buka Tab History</a></p>'+
   '<p style="color:#999;font-size:12px">Email otomatis dari sistem AISCHO.</p></div>';
}

function kpi_(label, val, color) {
  return '<td style="padding:0 8px"><div style="border:1px solid #ddd;border-radius:8px;padding:10px 18px;text-align:center">'+
         '<div style="font-size:22px;font-weight:bold;color:'+color+'">'+val+'</div>'+
         '<div style="font-size:11px;color:#666">'+label+'</div></div></td>';
}


/* ============== DETAIL AKTIF PER CABANG (klik +/- buka-tutup di Sheet) =========== */
function buildBranchDetail() {
  const data = readSalarySheet_(CFG.MASTER_SHEET);
  if (data.headerRow === -1) { SpreadsheetApp.getUi().alert('Sheet master tidak ditemukan.'); return; }

  // Kelompokkan pegawai LANJUT per cabang.
  const groups = {};
  data.records.forEach(function(r){
    if (r.status !== 'LANJUT') return;
    const c = normCabang_(r.cabang);
    (groups[c] = groups[c] || []).push(r);
  });

  const sh = getFreshSheet_(CFG.DETAIL_SHEET);
  const now = Utilities.formatDate(new Date(), CFG.TZ, 'EEEE, dd MMM yyyy — HH:mm');

  const rows = [];
  rows.push(['DETAIL PEGAWAI AKTIF (LANJUT) PER CABANG','','','']);
  rows.push([CFG.ORG,'','','']);
  rows.push(['Diperbarui: ' + now,'','','']);
  rows.push(['','','','']);
  rows.push(['CABANG / NAMA','MAPEL/SUBJEK','THP BERSIH','BATCH']);

  const branchRows = [];          // baris header tiap cabang (1-based)
  const memberRanges = [];        // {start,count} baris anggota utk di-group
  Object.keys(groups).sort(sortCabang_).forEach(function(c){
    const list = groups[c];
    const totalThp = list.reduce(function(a,r){ return a + r.thpBersih; }, 0);
    rows.push([c + '   (' + list.length + ' aktif)', '', totalThp, '']);
    branchRows.push(rows.length);
    const start = rows.length + 1;
    list.forEach(function(r){ rows.push(['   • ' + r.nama, r.mapel, r.thpBersih, r.batch]); });
    memberRanges.push({ start: start, count: list.length });
  });

  sh.getRange(1,1,rows.length,4).setValues(rows);

  // Format
  sh.getRange('A1').setFontSize(13).setFontWeight('bold');
  sh.getRange('A2:A3').setFontColor('#666666');
  sh.getRange(5,1,1,4).setFontWeight('bold').setBackground('#0b3d2e').setFontColor('#ffffff');
  branchRows.forEach(function(rIdx){
    sh.getRange(rIdx,1,1,4).setFontWeight('bold').setBackground('#e1f0e9');
    sh.getRange(rIdx,3).setNumberFormat('"Rp"#,##0');
  });
  memberRanges.forEach(function(g){
    if (g.count) sh.getRange(g.start,3,g.count,1).setNumberFormat('"Rp"#,##0');
  });
  sh.setColumnWidth(1,290); sh.setColumnWidth(2,210); sh.setColumnWidth(3,140); sh.setColumnWidth(4,90);
  sh.setFrozenRows(5);

  // Buat row group (tombol +/- per cabang) lalu lipat semua.
  memberRanges.forEach(function(g){
    if (g.count) sh.getRange(g.start,1,g.count,1).shiftRowGroupDepth(1);
  });
  try { sh.collapseAllRowGroups(); } catch(e) {}

  ss_().toast('Detail per cabang dibuat. Klik +/- untuk buka-tutup.', 'AISCHO', 5);
}

// Urutkan cabang berdasar nomor (AW1, AW4, AW5, ... AW10, AW16) bukan teks.
function sortCabang_(a, b) {
  function num(s){ const m = String(s).match(/(\d+)/); return m ? parseInt(m[1],10) : 9999; }
  const na = num(a), nb = num(b);
  return na === nb ? String(a).localeCompare(String(b)) : na - nb;
}


/* ================================ UTILITIES ===================================== */
function getOrCreateSheet_(name) {
  const ss = ss_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
// Ambil tren dari LOG UPDATE STATUS (timestamp + jumlah LANJUT) untuk grafik tren.
function getTrend_() {
  const sh = findSheet_(CFG.LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const ts = vals[r][0];
    if (ts === '' || ts == null) continue;
    let label;
    if (ts instanceof Date) label = Utilities.formatDate(ts, CFG.TZ, 'dd/MM HH:mm');
    else { const d = new Date(ts); label = isNaN(d.getTime()) ? String(ts) : Utilities.formatDate(d, CFG.TZ, 'dd/MM HH:mm'); }
    out.push({ date: label, lanjut: Number(vals[r][2]) || 0, total: Number(vals[r][1]) || 0 });
  }
  return out.slice(-30);
}

// Cari sheet: cocok persis dulu, lalu toleran (abaikan beda spasi & huruf besar/kecil).
function findSheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  const want = String(name).toLowerCase().replace(/\s+/g, ' ').trim();
  const all = ss.getSheets();
  for (let i = 0; i < all.length; i++) {
    if (all[i].getName().toLowerCase().replace(/\s+/g, ' ').trim() === want) return all[i];
  }
  return null;
}
function getFreshSheet_(name) {
  const ss = ss_();
  const ex = ss.getSheetByName(name);
  if (ex) ss.deleteSheet(ex);
  return ss.insertSheet(name);
}


/* ============================ WEB APP (real-time) =============================== */
/**
 * Halaman dashboard yang bisa dibuka lewat link.
 * Deploy: Deploy > New deployment > Web app.
 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : '';
  const file = (page === 'database' || page === 'pengajuan') ? 'Database' : 'Index';
  const title = (file === 'Database') ? 'Database Pengajuan — AL-WILDAN' : 'Dashboard Aktivasi Pegawai — AISCHO';
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Dipanggil dari halaman (google.script.run) untuk data terbaru.
function getDashboardData() {
  const data = readSalarySheet_(CFG.MASTER_SHEET);

  if (data.reason && data.reason !== 'OK') {
    const names = ss_().getSheets().map(function(s){ return s.getName(); });
    return {
      error: (data.reason === 'SHEET_NOT_FOUND')
        ? ('Tab "' + CFG.MASTER_SHEET + '" tidak ditemukan.')
        : ('Header (kolom NAMA & CABANG) tidak ditemukan di tab "' + (data.sheetName || CFG.MASTER_SHEET) + '".'),
      reason: data.reason,
      masterSheet: CFG.MASTER_SHEET,
      availableSheets: names,
      updated: Utilities.formatDate(new Date(), CFG.TZ, 'EEEE, dd MMM yyyy HH:mm:ss'),
      org: CFG.ORG
    };
  }

  const s = computeSummary_(data.records);
  const inList = function(st){ return CFG.STATUS_LIST.indexOf(st) > -1; };

  const people = data.records.map(function(r){
    return {
      nama     : r.nama,
      cabang   : normCabang_(r.cabang),
      mapel    : String(r.mapel || ''),
      thp      : r.thpBersih,
      batch    : String(r.batch || ''),
      status   : inList(r.status) ? r.status : 'LAINNYA',
      konfirm  : String(r.konfirm || ''),
      s1       : String(r.s1 || ''),
      s2       : String(r.s2 || ''),
      wa       : String(r.wa || ''),
      email    : String(r.email || ''),
      cv       : String(r.cv || ''),
      kesehatan: String(r.kesehatan || ''),
      thpKotor : r.thpKotor,
      thpBersih: r.thpBersih,
      total    : r.total,
      tk       : r.tk,
      thr      : r.thr
    };
  });

  const cabangStats = Object.keys(s.byCabang).sort(sortCabang_).map(function(c){
    const x = s.byCabang[c];
    return { cabang:c, lanjut:x.lanjut, hold:x.hold, batal:x.batal, total:x.total,
             thp:x.thp, avg: x.lanjut ? Math.round(x.thp / x.lanjut) : 0 };
  });

  return {
    updated  : Utilities.formatDate(new Date(), CFG.TZ, 'EEEE, dd MMM yyyy HH:mm:ss'),
    org      : CFG.ORG,
    total    : s.total,
    lanjut   : s.byStatus['LANJUT'] || 0,
    hold     : s.byStatus['HOLD'] || 0,
    batal    : s.byStatus['BATAL'] || 0,
    lainnya  : s.byStatus['LAINNYA'] || 0,
    nLanjut  : s.nLanjut,
    thpLanjut: s.thpLanjut,
    avgLanjut: s.avgLanjut,
    byBatch  : s.byBatch,
    cabangStats: cabangStats,
    trend    : getTrend_(),
    people   : people
  };
}

// ---- Halaman PENGAJUAN (foto + Merged Doc URL) ----
// Cari tab pengajuan: utamakan yang punya kolom "Merged", lalu kandidat pertama yang ada.
function findSubmissionSheet_(ss, cands) {
  const sheets = ss.getSheets();
  function headerLower(sh) {
    if (sh.getLastColumn() < 1 || sh.getLastRow() < 1) return [];
    return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(x){ return String(x).toLowerCase(); });
  }
  function hasCol(hdr, kw) { return hdr.some(function(h){ return h.indexOf(kw) > -1; }); }
  function isConfig(name) {
    const n = String(name).toLowerCase();
    return n.indexOf('autocrat') > -1 || n.indexOf('job settings') > -1 || n.indexOf('do not delete') > -1 ||
           n.indexOf('scriptproperties') > -1 || n.indexOf('nvscript') > -1 || n.indexOf('settings') > -1;
  }
  function match(name) {
    const want = String(name).toLowerCase().replace(/\s+/g, ' ').trim();
    for (let i = 0; i < sheets.length; i++) {
      if (String(sheets[i].getName()).toLowerCase().replace(/\s+/g, ' ').trim() === want) return sheets[i];
    }
    return null;
  }
  for (let i = 0; i < cands.length; i++) {
    const sh = match(cands[i]);
    if (sh && hasCol(headerLower(sh), 'merged')) return sh;
  }
  for (let i = 0; i < sheets.length; i++) {
    if (isConfig(sheets[i].getName())) continue;
    const h = headerLower(sheets[i]);
    if (hasCol(h, 'merged doc') || hasCol(h, 'merged')) return sheets[i];
  }
  for (let j = 0; j < cands.length; j++) { const s = match(cands[j]); if (s) return s; }
  return null;
}
function findPengajuanSheet_() { return findSubmissionSheet_(ssPengajuan_(), CFG.PENGAJUAN_SHEETS); }

// === DIAGNOSTIK: jalankan fungsi ini dari editor, lalu lihat Execution log ===
function cekPengajuan() {
  const all = ssPengajuan_().getSheets().map(function(s){ return s.getName(); });
  Logger.log('Spreadsheet pengajuan: ' + ssPengajuan_().getName());
  Logger.log('Semua tab: ' + all.join(' | '));
  const sh = findPengajuanSheet_();
  if (!sh) { Logger.log('>> TAB PENGAJUAN TIDAK DITEMUKAN (dicari: ' + CFG.PENGAJUAN_SHEETS.join(', ') + ')'); return; }
  Logger.log('Tab dipakai: "' + sh.getName() + '" | baris=' + sh.getLastRow() + ' kolom=' + sh.getLastColumn());
  const hdr = sh.getRange(1, 1, 1, Math.min(sh.getLastColumn(), 50)).getValues()[0];
  Logger.log('Header baris 1: ' + hdr.join(' | '));
  const d = getPengajuanData();
  Logger.log('getPengajuanData: ' + (d.error ? ('ERROR -> ' + d.error) : (d.count + ' data')));
  if (d.rows && d.rows[0]) Logger.log('Contoh: nama="' + d.rows[0].nama + '" | jml field=' + (d.rows[0].fields || []).length + ' | foto.id=' + (d.rows[0].photo ? d.rows[0].photo.id : '-'));
}

function getPengajuanData() {
  const sh = findPengajuanSheet_();
  if (!sh) {
    const names = ssPengajuan_().getSheets().map(function(s){ return s.getName(); });
    return { error: 'Tab data pengajuan tidak ditemukan (dicari: ' + CFG.PENGAJUAN_SHEETS.join(', ') + ').',
             availableSheets: names,
             updated: Utilities.formatDate(new Date(), CFG.TZ, 'EEEE, dd MMM yyyy HH:mm:ss') };
  }
  return readSubmissionSheet_(sh);
}

function getKesehatanData() {
  if (!CFG.KESEHATAN_SPREADSHEET_ID) return { rows: [], count: 0, updated: '', error: 'Sumber Data Kesehatan belum diisi.' };
  const ss = ssKesehatan_();
  const sh = findSubmissionSheet_(ss, CFG.KESEHATAN_SHEETS);
  if (!sh) {
    const names = ss.getSheets().map(function(s){ return s.getName(); });
    return { error: 'Tab data kesehatan tidak ditemukan (dicari: ' + CFG.KESEHATAN_SHEETS.join(', ') + ').',
             availableSheets: names,
             updated: Utilities.formatDate(new Date(), CFG.TZ, 'EEEE, dd MMM yyyy HH:mm:ss') };
  }
  return readSubmissionSheet_(sh);
}

// Pembaca generik tab "form responses + merge" → {count, rows:[{nama,cabang,photo,merged,fields}]}
function readSubmissionSheet_(sh) {
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return { rows: [], count: 0, updated: Utilities.formatDate(new Date(), CFG.TZ, 'EEEE, dd MMM yyyy HH:mm:ss') };

  const headers = vals[0].map(function(h){ return String(h == null ? '' : h).trim(); });
  let idxPhotoUrl = findCol_(headers, ['photo url']);
  let idxUpload   = findCol_(headers, ['lampirkan photo', 'foto terbaru', 'upload foto', 'photo']);
  let idxMerged   = findCol_(headers, ['merged doc url', 'merged']);
  let idxNama     = findCol_(headers, ['nama lengkap', 'nama']);
  let idxCabang   = findCol_(headers, ['cabang']);
  if (idxUpload   === -1) idxUpload   = 1;
  if (idxPhotoUrl === -1) idxPhotoUrl = 2;

  const rows = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (row.join('').trim() === '') continue;
    const fields = [];
    for (let c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      if (c === idxPhotoUrl || c === idxUpload || c === idxMerged) continue;
      const v = row[c];
      if (v === '' || v == null) continue;
      fields.push({ label: headers[c], value: String(v) });
    }
    rows.push({
      nama   : (idxNama   > -1 ? String(row[idxNama]   || '') : ''),
      cabang : (idxCabang > -1 ? String(row[idxCabang] || '') : ''),
      photo  : buildPhoto_(row, idxPhotoUrl, idxUpload),
      merged : (idxMerged > -1 ? buildMerged_(row[idxMerged]) : { open:'', preview:'', thumb:'', id:'' }),
      fields : fields
    });
  }
  return { count: rows.length, org: CFG.ORG, rows: rows,
           updated: Utilities.formatDate(new Date(), CFG.TZ, 'EEEE, dd MMM yyyy HH:mm:ss') };
}

function findCol_(headers, keys) {
  for (let k = 0; k < keys.length; k++) {
    for (let i = 0; i < headers.length; i++) {
      if (String(headers[i] || '').toLowerCase().indexOf(keys[k]) > -1) return i;
    }
  }
  return -1;
}
function extractDriveId_(url) {
  url = String(url || '');
  let m = url.match(/[?&]id=([-\w]{20,})/); if (m) return m[1];
  m = url.match(/\/d\/([-\w]{20,})/);        if (m) return m[1];
  m = url.match(/([-\w]{25,})/);             if (m) return m[1];
  return '';
}
function buildPhoto_(row, idxUrl, idxUpload) {
  let raw = '';
  if (idxUrl > -1 && row[idxUrl]) raw = String(row[idxUrl]);
  if (!raw && idxUpload > -1 && row[idxUpload]) raw = String(row[idxUpload]);
  raw = raw.split(',')[0].trim();
  if (!raw) return { thumb: '', open: '', id: '', direct: false };
  if (/^https?:\/\/.*\.(jpe?g|png|gif|webp)(\?|$)/i.test(raw)) return { thumb: raw, open: raw, id: '', direct: true };
  const id = extractDriveId_(raw);
  if (id) return { thumb: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400',
                   open: 'https://drive.google.com/file/d/' + id + '/view', id: id, direct: false };
  return { thumb: '', open: raw, id: '', direct: false };
}

// Ambil gambar dari Drive (script jalan sebagai pemilik) lalu kirim sebagai data URI.
function getPhotoData(id) {
  try {
    if (!id) return '';
    const blob = DriveApp.getFileById(id).getBlob();
    const bytes = blob.getBytes();
    if (bytes.length > 4 * 1024 * 1024) return '';   // lewati file > 4MB
    const type = blob.getContentType() || 'image/jpeg';
    return 'data:' + type + ';base64,' + Utilities.base64Encode(bytes);
  } catch (e) { return ''; }
}
function buildMerged_(raw) {
  raw = String(raw || '').trim();
  if (!raw) return { open: '', preview: '', thumb: '', id: '' };
  const id = extractDriveId_(raw);
  let preview = '';
  if (/document\/d\//.test(raw))          preview = 'https://docs.google.com/document/d/' + id + '/preview';
  else if (/presentation\/d\//.test(raw)) preview = 'https://docs.google.com/presentation/d/' + id + '/preview';
  else if (/spreadsheets\/d\//.test(raw)) preview = 'https://docs.google.com/spreadsheets/d/' + id + '/preview';
  else if (id)                            preview = 'https://drive.google.com/file/d/' + id + '/preview';
  const thumb = id ? ('https://drive.google.com/thumbnail?id=' + id + '&sz=w640') : '';
  return { open: raw, id: id, preview: preview, thumb: thumb };
}

// ===== DRAFT EMAIL AKTIVASI SDM =====
function formatRp_(n) {
  n = Number(n) || 0;
  const neg = n < 0; n = Math.abs(Math.round(n));
  return (neg ? '-' : '') + 'Rp' + n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}
function escH_(s) { return String(s == null ? '' : s).replace(/[&<>]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; }); }
function emailsFrom_(s) {
  return String(s || '').split(/[\s,;]+/).map(function(e){ return e.trim(); }).filter(function(e){ return e.indexOf('@') > -1; });
}
function uniq_(arr) { const o = {}; const out = []; arr.forEach(function(x){ x = String(x).trim(); if (x && !o[x.toLowerCase()]) { o[x.toLowerCase()] = 1; out.push(x); } }); return out; }

// Cari kontak FIN & MGT untuk satu cabang (mis. "AW20" → "AISCHO 20 MATARAM").
function getKontakCabang(cabang) {
  try {
    if (!CFG.KONTAK_SPREADSHEET_ID) return { ok: false, fin: [], mgt: [] };
    const ss = SpreadsheetApp.openById(CFG.KONTAK_SPREADSHEET_ID);
    const sheets = ss.getSheets();
    let sh = null;
    for (let i = 0; i < sheets.length; i++) {
      if (sheets[i].getLastColumn() < 2) continue;
      const h = sheets[i].getRange(1, 1, 1, sheets[i].getLastColumn()).getValues()[0].map(function(x){ return String(x).toLowerCase(); });
      if (h.some(function(c){ return c.indexOf('cabang') > -1; }) && h.some(function(c){ return c.indexOf('finance') > -1 || c.indexOf('fin') > -1; })) { sh = sheets[i]; break; }
    }
    if (!sh) sh = sheets[0];
    const vals = sh.getDataRange().getValues();
    const headers = vals[0].map(function(x){ return String(x).toLowerCase(); });
    const iCab = findCol_(headers, ['cabang']);
    const iFin = findCol_(headers, ['email finance', 'finance', 'fin']);
    const iMgt = findCol_(headers, ['email mgt', 'mgt', 'manaj', 'management']);
    const num = (String(cabang).match(/(\d+)/) || [])[1];
    const cl = String(cabang).toLowerCase();
    for (let r = 1; r < vals.length; r++) {
      const cabCell = String(vals[r][iCab] || '');
      if (!cabCell) continue;
      const cnum = (cabCell.match(/(\d+)/) || [])[1];
      const matched = (num && cnum && cnum === num) || (cabCell.toLowerCase().indexOf(cl) > -1) || (cl.indexOf(cabCell.toLowerCase()) > -1);
      if (matched) {
        return { ok: true, cabangFull: cabCell,
                 fin: iFin > -1 ? emailsFrom_(vals[r][iFin]) : [],
                 mgt: iMgt > -1 ? emailsFrom_(vals[r][iMgt]) : [] };
      }
    }
    return { ok: false, cabangFull: '', fin: [], mgt: [] };
  } catch (e) { return { ok: false, error: String(e), fin: [], mgt: [] }; }
}

function buildAktivasiHtml_(p) {
  function waLink(w) {
    w = String(w || ''); if (!w) return '—';
    if (/^https?:\/\//i.test(w)) return '<a href="' + escH_(w) + '">' + escH_(w) + '</a>';
    const d = w.replace(/[^\d]/g, ''); return d ? '<a href="https://wa.me/' + d + '">' + escH_(w) + '</a>' : escH_(w);
  }
  function lnk(u, t) { u = String(u || ''); return /^https?:\/\//i.test(u) ? '<a href="' + escH_(u) + '">' + t + '</a>' : '—'; }
  const mgtHtml = (p.mgtEmails || []).map(function(e){ return escH_(e); }).join('<br>');
  const hi = 'padding:2px 4px;';
  const td = 'border:1px solid #000;padding:6px 9px;font-size:13px;vertical-align:top;';
  const th = td + 'font-weight:bold;text-align:center;';
  return '' +
    '<div style="font-family:Arial,sans-serif;color:#000;font-size:13px;line-height:1.5">' +
      '<p><b>HR SUBJECT :</b><br>' + escH_(p.hrSubject) + '</p>' +
      '<p><b>EMAIL FIN:</b> ' + ((p.finEmails || []).map(escH_).join(', ') || '—') + '</p>' +
      '<p><b>EMAIL MGT:</b><br>' + (mgtHtml || '—') + '</p>' +
      '<div style="border:1px solid #000;padding:14px">' +
        '<p style="text-align:justify">السَّلَامُ عَلَيْكُمْ وَرَحْمَةُ اللهِ وَبَرَكَاتُهُ</p>' +
        '<p><i>Assalamu\'alaikum Warahmatullahi Wabarakaatuh,</i><br><i>Bismillah,</i></p>' +
        '<table style="border-collapse:collapse;width:100%;margin:10px 0">' +
          '<tr>' +
            '<td style="' + td + 'width:38%"><b>' + escH_(p.hrSubject) + '</b></td>' +
            '<td style="' + td + '"><span style="background:#b6ff8c;' + hi + '"><b>Aktif : ' + escH_(p.tanggalAktif || '-') + '</b></span><br><br>' +
              '<span style="background:#86e3ce;' + hi + '"><b>THP BERSIH ' + escH_(p.mode || '') + '</b></span></td>' +
          '</tr>' +
          '<tr>' +
            '<td style="' + td + '"><span style="background:#fff36b;' + hi + '"><b>Action</b></span></td>' +
            '<td style="' + td + '"><ol style="margin:0;padding-left:18px">' +
              '<li><b>AKAD ' + escH_(p.akad || '') + '</b></li>' +
              '<li><b>JOB DESC ' + escH_(p.jobdesc || '') + '</b></li>' +
              '<li>' + escH_(p.unit || '') + '</li>' +
              (p.lintasCabang ? '<li>' + escH_(p.lintasCabang) + '</li>' : '') +
            '</ol></td>' +
          '</tr>' +
        '</table>' +
        '<table style="border-collapse:collapse;width:100%;margin:10px 0;text-align:center">' +
          '<tr>' +
            '<td style="' + th + 'background:#7dff7d">THP KOTOR</td>' +
            '<td style="' + th + 'background:#9fe6a0">KONFIRMASI</td>' +
            '<td style="' + th + 'background:#fff36b">THP BERSIH</td>' +
            '<td style="' + th + 'background:#f3c6c6">TOTAL (TK-THR)</td>' +
            '<td style="' + th + 'background:#f3c6c6">TK</td>' +
            '<td style="' + th + 'background:#f3c6c6">THR /bulan</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="' + td + 'text-align:center"><b>' + formatRp_(p.thpKotor) + '</b></td>' +
            '<td style="' + td + 'text-align:center"><b>' + escH_(p.konfirmasi || p.tanggalAktif || '-') + '</b></td>' +
            '<td style="' + td + 'text-align:center"><b>' + formatRp_(p.thpBersih) + '</b></td>' +
            '<td style="' + td + 'text-align:center"><b>' + formatRp_(p.total) + '</b></td>' +
            '<td style="' + td + 'text-align:center"><b>' + formatRp_(p.tk) + '</b></td>' +
            '<td style="' + td + 'text-align:center"><b>' + formatRp_(p.thr) + '</b></td>' +
          '</tr>' +
        '</table>' +
        '<table style="border-collapse:collapse;width:100%;margin:10px 0">' +
          '<tr>' +
            '<td style="' + th + '">EMAIL</td><td style="' + th + '">NO WA</td><td style="' + th + '">CV</td><td style="' + th + '">KESEHATAN</td>' +
          '</tr>' +
          '<tr>' +
            '<td style="' + td + '">' + (p.email ? '<a href="mailto:' + escH_(p.email) + '">' + escH_(p.email) + '</a>' : '—') + '</td>' +
            '<td style="' + td + '">' + waLink(p.wa) + '</td>' +
            '<td style="' + td + '">' + lnk(p.cv, 'Buka CV') + '</td>' +
            '<td style="' + td + '">' + lnk(p.kesehatan, 'Buka berkas') + '</td>' +
          '</tr>' +
        '</table>' +
        '<p><i>Syukron Jazakumullahu Khairan</i></p>' +
        '<p style="color:#d2691e"><b>' + escH_(CFG.TTD_NAMA) + '</b><br>' +
          '<span style="color:#000"><b>Email 1:</b> <a href="mailto:' + escH_(CFG.TTD_EMAIL) + '">' + escH_(CFG.TTD_EMAIL) + '</a><br>' +
          '<b>Mobile :</b> ' + escH_(CFG.TTD_HP) + '</span></p>' +
      '</div>' +
    '</div>';
}

// Dipanggil dari web app: buat draft Gmail aktivasi. Mengembalikan ringkasan penerima.
function createAktivasiDraft(p) {
  try {
    const to  = uniq_(CFG.EMAIL_AKTIVASI_TETAP.concat(p.toExtra || []));
    const cc  = uniq_((p.finEmails || []).concat(p.mgtEmails || []).concat(p.extraEmails || []))
                  .filter(function(e){ return to.indexOf(e) === -1; });
    const subject = p.hrSubject || ('Aktivasi SDM - ' + (p.nama || ''));
    const html = buildAktivasiHtml_(p);
    const plain = 'Draft aktivasi SDM untuk ' + (p.nama || '') + ' (' + (p.cabang || '') + '). Buka di Gmail (HTML) untuk tampilan tabel.';
    GmailApp.createDraft(to.join(','), subject, plain, { htmlBody: html, cc: cc.join(',') });
    return { ok: true, to: to, cc: cc, subject: subject };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

// Tampilkan link web app dari menu.
function showWebAppUrl() {
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  SpreadsheetApp.getUi().alert(
    url ? ('Link dashboard (real-time):\n\n' + url +
           '\n\nBagikan ke yang berhak. Atur akses saat deploy.')
        : 'Web App belum di-deploy.\n\nApps Script: Deploy → New deployment → Web app.');
}
function getSheetGid_(name) {
  const sh = ss_().getSheetByName(name);
  return sh ? sh.getSheetId() : 0;
}

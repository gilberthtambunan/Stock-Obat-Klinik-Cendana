/**
 * API Kartu Stok Klinik Cendana.
 * Web app: Execute as "Me", akses "Anyone". Pengamanan lewat TOKEN (Script Properties).
 * Aturan: hanya menulis ke kolom input (A,B,C,D,F,G,L). Kolom rumus (E,H,I,J,K) tidak disentuh.
 */
const SS_ID = '1mevzH1wkEa-wSJaPhTQqFB56lRB-nOsaWHbt_-mvIwU';
const SKIP = ['TEMPLATE CONTOH', 'TEMPLATE READY', 'REKAP PEKERJAAN'];
const FIRST = 19, LAST = 118;       // baris transaksi pada sheet obat
const TZ = 'Asia/Jakarta';

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
    if (!token || req.token !== token) return out({ ok: false, error: 'Kunci akses salah.' });
    let res;
    if (req.action === 'data') res = getData();
    else {
      const lock = LockService.getScriptLock();
      lock.waitLock(15000);
      try {
        if (req.action === 'add') res = addTx(req);
        else if (req.action === 'edit') res = editTx(req);
        else if (req.action === 'del') res = delTx(req);
        else throw new Error('Aksi tidak dikenal.');
      } finally { lock.releaseLock(); }
    }
    return out(Object.assign({ ok: true }, res));
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) });
  }
}
function doGet() { return out({ ok: true, info: 'API aktif' }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Dijalankan sekali dari editor untuk membuat kunci akses.
function buatToken() {
  const t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  PropertiesService.getScriptProperties().setProperty('TOKEN', t);
  Logger.log('KUNCI AKSES: ' + t);
}

/* ---------- Baca semua data (1 panggilan API) ---------- */
function getData() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const names = ss.getSheets().map(s => s.getName()).filter(n => SKIP.indexOf(n) < 0);
  const ranges = names.map(n => "'" + n.replace(/'/g, "''") + "'!A8:L" + LAST);
  const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + SS_ID + '/values:batchGet?majorDimension=ROWS' +
    '&valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER&' +
    ranges.map(r => 'ranges=' + encodeURIComponent(r)).join('&');
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Gagal membaca spreadsheet: ' + res.getContentText().slice(0, 150));
  const vr = JSON.parse(res.getContentText()).valueRanges;
  const obat = vr.map((v, i) => {
    const r = v.values || [];
    const g = (ri, ci) => (r[ri] && r[ri][ci] !== undefined) ? r[ri][ci] : '';
    const batches = [];
    for (let k = 0; k < 4; k++) {            // baris 11-14 -> indeks 3-6
      const no = g(3 + k, 2);
      batches.push({ slot: k, no: no, ed: g(3 + k, 4), sisa: no === '' ? 0 : Number(g(3 + k, 5)) || 0 });
    }
    const rows = [];
    for (let k = FIRST - 8; k < r.length; k++) {
      const x = r[k]; if (!x) continue;
      const masuk = Number(x[2]) || 0, keluar = Number(x[5]) || 0;
      if (!masuk && !keluar) continue;
      rows.push({ row: k + 8, tgl: x[0], nor: x[1] || '', masuk: masuk, bm: x[3] || '', keluar: keluar, bk: x[6] || '', ket: x[11] || '' });
    }
    return { nama: names[i], min: Number(g(0, 5)) || 0, total: Number(g(7, 5)) || 0, batches: batches, rows: rows };
  });
  return { obat: obat, now: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') };
}

/* ---------- Util ---------- */
function sheetOf(nama) {
  if (SKIP.indexOf(nama) >= 0) throw new Error('Sheet tidak valid.');
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(nama);
  if (!sh) throw new Error('Obat "' + nama + '" tidak ditemukan.');
  return sh;
}
function posInt(v) {
  const n = Number(v);
  if (!isFinite(n) || n <= 0 || Math.floor(n) !== n) throw new Error('Jumlah harus bilangan bulat lebih dari 0.');
  return n;
}
function sisaBatch(sh, batch, exceptRow) {
  const d = sh.getRange(FIRST, 3, LAST - FIRST + 1, 5).getValues(); // C..G
  let s = 0;
  d.forEach((x, i) => {
    if (exceptRow && FIRST + i === exceptRow) return;
    if (x[1] === batch) s += Number(x[0]) || 0;
    if (x[4] === batch) s -= Number(x[3]) || 0;
  });
  return s;
}
function ensureBatch(sh, batch, ed) {
  const c = sh.getRange('C11:C14').getValues().map(r => String(r[0]).trim());
  if (c.indexOf(batch) >= 0) return;
  const slot = c.indexOf('');
  if (slot < 0) throw new Error('Slot batch penuh (maks. 4). Hapus batch lama di sheet dulu.');
  if (!ed) throw new Error('ED wajib diisi untuk batch baru.');
  sh.getRange(11 + slot, 3).setValue(batch);
  const p = String(ed).split('-');
  const cell = sh.getRange(11 + slot, 5);
  cell.setValue(new Date(Number(p[0]), Number(p[1]) - 1, 1)); cell.setNumberFormat('MM/yy');
}
function rekap(nama, jenis, jumlah, batch, tgl) {
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName('REKAP PEKERJAAN');
  if (!sh) return;
  const a = sh.getRange('A1:A' + Math.max(sh.getLastRow(), 6)).getValues();
  let last = 5;
  for (let i = 5; i < a.length; i++) if (a[i][0] !== '') last = i + 1;
  const r = sh.getRange(last + 1, 1, 1, 5);
  r.setValues([[tgl, nama, jenis, jumlah, batch]]);
  sh.getRange(last + 1, 1).setNumberFormat('dd/MM/yyyy');
}

/* ---------- Tambah ---------- */
function addTx(q) {
  const sh = sheetOf(q.obat);
  const jenis = q.jenis === 'KELUAR' ? 'KELUAR' : 'MASUK';
  const jml = posInt(q.jumlah);
  const batch = String(q.batch || '').trim();
  if (!batch) throw new Error('Pilih no. batch.');
  if (jenis === 'MASUK') ensureBatch(sh, batch, q.ed);
  else {
    const ada = sh.getRange('C11:C14').getValues().map(r => String(r[0]).trim()).indexOf(batch) >= 0;
    if (!ada) throw new Error('Batch tidak ada di kartu stok.');
    const s = sisaBatch(sh, batch);
    if (jml > s) throw new Error('Stok batch ' + batch + ' hanya ' + s + '.');
  }
  const d = sh.getRange(FIRST, 3, LAST - FIRST + 1, 5).getValues();
  let last = FIRST - 1;
  d.forEach((x, i) => { if (x[0] !== '' || x[1] !== '' || x[3] !== '' || x[4] !== '') last = FIRST + i; });
  const row = last + 1;
  if (row > LAST) throw new Error('Baris kartu penuh (maks. ' + (LAST - FIRST + 1) + '). Perpanjang rumus di sheet.');
  const tgl = new Date(Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd'));
  const ta = sh.getRange(row, 1); ta.setValue(tgl); ta.setNumberFormat('dd/MM/yyyy');
  if (q.nor) sh.getRange(row, 2).setValue(String(q.nor));
  if (jenis === 'MASUK') sh.getRange(row, 3, 1, 2).setValues([[jml, batch]]);
  else sh.getRange(row, 6, 1, 2).setValues([[jml, batch]]);
  if (q.ket) sh.getRange(row, 12).setValue(String(q.ket));
  rekap(q.obat, jenis, jml, batch, tgl);
  return { row: row };
}

/* ---------- Ubah (jumlah, batch, no. R/, keterangan) ---------- */
function editTx(q) {
  const sh = sheetOf(q.obat), row = Number(q.row);
  if (row < FIRST || row > LAST) throw new Error('Baris tidak valid.');
  const v = sh.getRange(row, 1, 1, 7).getValues()[0];
  const masuk = Number(v[2]) > 0, keluar = Number(v[5]) > 0;
  if (!masuk && !keluar) throw new Error('Baris sudah kosong.');
  const jml = posInt(q.jumlah), batch = String(q.batch || '').trim();
  if (!batch) throw new Error('Pilih no. batch.');
  if (masuk) {
    ensureBatch(sh, batch, q.ed);
    sh.getRange(row, 3, 1, 2).setValues([[jml, batch]]);
  } else {
    const s = sisaBatch(sh, batch, row);
    if (jml > s) throw new Error('Stok batch ' + batch + ' hanya ' + s + '.');
    sh.getRange(row, 6, 1, 2).setValues([[jml, batch]]);
  }
  sh.getRange(row, 2).setValue(String(q.nor || ''));
  sh.getRange(row, 12).setValue(String(q.ket || ''));
  return {};
}

/* ---------- Hapus: kosongkan sel input saja, baris & rumus tetap ---------- */
function delTx(q) {
  const sh = sheetOf(q.obat), row = Number(q.row);
  if (row < FIRST || row > LAST) throw new Error('Baris tidak valid.');
  sh.getRangeList(['A' + row + ':D' + row, 'F' + row + ':G' + row, 'L' + row]).clearContent();
  return {};
}

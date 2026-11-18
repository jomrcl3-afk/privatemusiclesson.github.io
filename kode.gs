// ===== Presensi Les Pribadi - Kode.gs (biru muda · 1 pengajar · tanpa admin, tanpa honor, tanpa logo) =====

// ---------- PENGATURAN (wajib dibaca) ----------
// 1) WAJIB: ganti dengan sandi Anda sendiri. Hanya Anda yang memakai sandi ini.
const SANDI = 'GANTI-SANDI-INI';
// 2) Nama les yang tampil di atas aplikasi, di invoice, dan di pesan WhatsApp.
const NAMA_LES = 'Les Musik Privat';
// 3) Opsional: isi dengan URL web app (berakhiran /exec) kalau link ortu yang dibuat aplikasi tidak benar.
const BASE_URL = '';
// 4) Opsional: isi ID spreadsheet kalau proyek ini berdiri sendiri (tidak dibuat lewat Ekstensi > Apps Script).
const SHEET_ID = '';
// -----------------------------------------------

// Nama folder di Google Drive (dibuat terpisah dari folder tempat Anda mengajar).
const FOLDER_FOTO = 'Foto Murid Les Pribadi';
const FOLDER_DOK = 'Dokumen Les Pribadi';

const SCHEMA = {
  Siswa: ['ID','Nama','Instrumen','Level','Ortu','Telepon','TglMulai','Tarif','Jadwal','Durasi','JmlPertemuan','Foto','LinkDokumentasi','Rekening','Token','Status'],
  Presensi: ['Tanggal','ID','Nama','Periode','Pertemuan','Status','Materi','Catatan'],
  Penilaian: ['Tanggal','ID','Nama','Materi','Nilai','Catatan','NoNilai'],
  Invoice: ['NoInvoice','ID','Nama','Periode','Pertemuan','Total','Rekening','Status','Terbit'],
  Rekening: ['Kode','Bank','NoRek','AtasNama']
};
const WARNA_HEADER = '#cfe8fb';

// ---------- Dasar ----------
function ss_() {
  const s = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
  if (!s) throw new Error('Proyek ini tidak terhubung ke spreadsheet. Isi SHEET_ID di bagian atas Kode.gs.');
  return s;
}

function sh_(n) {
  const ss = ss_();
  let sh = ss.getSheetByName(n);
  if (!sh) {
    sh = ss.insertSheet(n);
    const h = SCHEMA[n];
    sh.getRange(1, 1, sh.getMaxRows(), h.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground(WARNA_HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Membaca header; menambah kolom yang belum ada.
function hdr_(n) {
  const sh = sh_(n);
  const lc = sh.getLastColumn();
  let h = lc ? sh.getRange(1, 1, 1, lc).getValues()[0].map(x => String(x).trim()) : [];
  while (h.length && h[h.length - 1] === '') h.pop();
  const miss = SCHEMA[n].filter(k => h.indexOf(k) < 0);
  if (miss.length) {
    sh.getRange(1, h.length + 1, 1, miss.length).setNumberFormat('@').setValues([miss]).setFontWeight('bold').setBackground(WARNA_HEADER);
    h = h.concat(miss);
  }
  return h;
}

function txt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Jakarta', 'yyyy-MM-dd');
  return String(v === null || v === undefined ? '' : v);
}

function read_(n) {
  const h = hdr_(n), sh = sh_(n), lr = sh.getLastRow();
  if (lr < 2) return [];
  const key = ({Presensi: 'ID', Penilaian: 'ID'})[n] || SCHEMA[n][0];
  return sh.getRange(2, 1, lr - 1, h.length).getValues().map((r, i) => {
    const o = {_row: i + 2};
    h.forEach((k, j) => { if (k) o[k] = txt_(r[j]); });
    return o;
  }).filter(o => o[key] !== undefined && o[key] !== '');
}

function val_(obj, k) { return obj[k] === undefined || obj[k] === null ? '' : String(obj[k]); }

// Menulis berdasarkan NAMA kolom. Kolom yang tidak dikirim tidak diubah.
function append_(n, obj) {
  const h = hdr_(n), sh = sh_(n);
  const row = h.map(k => SCHEMA[n].indexOf(k) >= 0 ? val_(obj, k) : '');
  sh.getRange(sh.getLastRow() + 1, 1, 1, h.length).setNumberFormat('@').setValues([row]);
}

function write_(n, rowNum, obj) {
  const h = hdr_(n), sh = sh_(n), rg = sh.getRange(rowNum, 1, 1, h.length), cur = rg.getValues()[0];
  const row = h.map((k, j) =>
    (SCHEMA[n].indexOf(k) >= 0 && Object.prototype.hasOwnProperty.call(obj, k)) ? val_(obj, k) : txt_(cur[j]));
  rg.setNumberFormat('@').setValues([row]);
}

function delBy_(n, id) {
  read_(n).filter(x => x.ID === id).map(x => x._row).sort((a, b) => b - a).forEach(r => sh_(n).deleteRow(r));
}

function newToken_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 24); }

function tokens_() {
  const sh = sh_('Siswa'), col = hdr_('Siswa').indexOf('Token') + 1;
  read_('Siswa').filter(x => !x.Token).forEach(x => sh.getRange(x._row, col).setNumberFormat('@').setValue(newToken_()));
}

function folder_(nm) {
  const it = DriveApp.getFoldersByName(nm);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nm);
}

// ---------- Keamanan (satu sandi) ----------
const DEF_PW_ = 'GANTI-SANDI-INI';
function auth_(pw) {
  if (SANDI === DEF_PW_) throw new Error('Ganti dulu SANDI di bagian atas Kode.gs, lalu deploy versi baru.');
  if (String(pw || '').trim() !== SANDI) { Utilities.sleep(1000); throw new Error('Sandi salah'); }
  return true;
}
function siswa_() { tokens_(); return read_('Siswa'); }
function ada_(id) {
  const s = read_('Siswa').find(x => x.ID === id);
  if (!s) throw new Error('Murid tidak ditemukan');
  return s;
}
function baseUrl_() { return BASE_URL || ScriptApp.getService().getUrl(); }

// ---------- Setup & tes ----------
// Jalankan sekali (dan izinkan akses Drive). Aman dijalankan ulang.
function setup() {
  Object.keys(SCHEMA).forEach(n => hdr_(n));
  tokens_();
  folder_(FOLDER_FOTO);
  folder_(FOLDER_DOK);
  Logger.log('Setup selesai. Sheet: ' + ss_().getSheets().map(s => s.getName()).join(', '));
}

function cekKoneksi() {
  Logger.log('Terhubung ke: ' + ss_().getName());
  Logger.log('Jumlah siswa: ' + read_('Siswa').length);
  Logger.log('Sandi sudah diganti: ' + (SANDI !== DEF_PW_));
}

// ---------- Web app ----------
function doGet(e) {
  const p = String((e && e.parameter && e.parameter.p) || '').replace(/[^\w-]/g, '');
  const t = HtmlService.createTemplateFromFile('Index');
  t.boot = JSON.stringify({mode: p ? 'ortu' : 'guru', token: p, nama: NAMA_LES});
  t.nama = NAMA_LES;
  return t.evaluate().setTitle(NAMA_LES)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- Tampilan orang tua (hanya baca, hanya satu anak) ----------
function getProfilOrtu(token) {
  token = String(token || '').replace(/[^\w-]/g, '');
  if (token.length < 10) throw new Error('Link tidak valid');
  const s = read_('Siswa').find(x => x.Token === token);
  if (!s) throw new Error('Link tidak valid');
  const pk = (o, ks) => { const r = {}; ks.forEach(k => r[k] = o[k]); return r; };
  return {
    s: pk(s, ['Nama','Instrumen','Level','Jadwal','Durasi','JmlPertemuan','Foto','LinkDokumentasi']),
    pres: read_('Presensi').filter(x => x.ID === s.ID).map(x => pk(x, ['Tanggal','Periode','Pertemuan','Status','Materi','Catatan'])),
    nilai: nilaiList_(s.ID).map(x => pk(x, ['Tanggal','Materi','Nilai','Catatan']))
  };
}

// ---------- Awal & login ----------
function getAwal(pw) {
  auth_(pw);
  return {siswa: siswa_(), rek: read_('Rekening'), base: baseUrl_()};
}
function getSiswa(pw) { auth_(pw); return siswa_(); }

function saveSiswa(pw, o) {
  auth_(pw);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    o.Nama = String(o.Nama || '').trim();
    delete o.Token;
    const all = read_('Siswa');
    const f = o.ID ? all.find(x => x.ID === o.ID) : null;
    if (f) {
      write_('Siswa', f._row, o);
    } else {
      const max = all.reduce((m, x) => Math.max(m, Number(String(x.ID).slice(1)) || 0), 0);
      o.ID = 'S' + String(max + 1).padStart(3, '0');
      o.Status = o.Status || 'Aktif';
      o.Token = newToken_();
      append_('Siswa', o);
    }
    return {id: o.ID, list: siswa_()};
  } finally {
    lock.releaseLock();
  }
}

// Hapus permanen: murid + presensi, penilaian, dan invoicenya.
function hapusSiswa(pw, id) {
  auth_(pw);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ['Presensi', 'Penilaian', 'Invoice', 'Siswa'].forEach(n => delBy_(n, id));
    return siswa_();
  } finally {
    lock.releaseLock();
  }
}

// Membuat link ortu baru (link lama otomatis tidak berlaku).
function resetToken(pw, id) {
  auth_(pw);
  const f = ada_(id);
  const col = hdr_('Siswa').indexOf('Token') + 1;
  sh_('Siswa').getRange(f._row, col).setNumberFormat('@').setValue(newToken_());
  return siswa_();
}

// ---------- Foto & dokumen (Google Drive) ----------
function uploadFoto(pw, b64) {
  auth_(pw);
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'foto_' + Date.now() + '.jpg');
  const f = folder_(FOLDER_FOTO).createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?id=' + f.getId();
}

// Mengambil foto dari Drive sebagai data gambar (agar foto ikut tercetak di PDF / gambar).
function fotoData(pw, url) {
  auth_(pw);
  const m = String(url).match(/[?&]id=([\w-]+)|\/d\/([\w-]+)/);
  if (!m) throw new Error('Foto tidak valid');
  const fid = m[1] || m[2];
  if (!read_('Siswa').some(s => String(s.Foto).indexOf(fid) >= 0)) throw new Error('Foto tidak valid');
  const b = DriveApp.getFileById(fid).getBlob();
  return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
}

function simpanDokumen(pw, name, mime, b64) {
  auth_(pw);
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, name);
  const f = folder_(FOLDER_DOK).createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getUrl();
}

// ---------- Presensi (per pertemuan) ----------
function getPresensi(pw, id) { auth_(pw); ada_(id); return read_('Presensi').filter(x => x.ID === id); }

function savePertemuan(pw, o) {
  auth_(pw); ada_(o.ID);
  const f = read_('Presensi').find(x => x.ID === o.ID && x.Periode === String(o.Periode) && x.Pertemuan === String(o.Pertemuan));
  if (f) write_('Presensi', f._row, o);
  else append_('Presensi', o);
  return true;
}

function hapusPertemuan(pw, id, periode, ke) {
  auth_(pw); ada_(id);
  const f = read_('Presensi').find(x => x.ID === id && x.Periode === String(periode) && x.Pertemuan === String(ke));
  if (f) sh_('Presensi').deleteRow(f._row);
  return true;
}

// ---------- Penilaian ----------
function skor_(v) { const n = Number(v); if (!(n >= 1 && n <= 5 && n === Math.floor(n))) throw new Error('Nilai harus 1 sampai 5'); return String(n); }
// Daftar penilaian satu murid (terbaru di atas).
function nilaiList_(id) { return read_('Penilaian').filter(x => x.ID === id).reverse(); }
function getPenilaian(pw, id) { auth_(pw); ada_(id); return nilaiList_(id); }
function savePenilaian(pw, o) {
  auth_(pw); ada_(o.ID);
  o.Nilai = skor_(o.Nilai);
  o.NoNilai = 'N' + newToken_().slice(0, 10);
  append_('Penilaian', o); return true;
}
function updatePenilaian(pw, o) {
  auth_(pw);
  const r = read_('Penilaian').find(x => x.NoNilai === o.NoNilai);
  if (!r) throw new Error('Penilaian tidak ditemukan');
  o.Nilai = skor_(o.Nilai);
  delete o.ID; delete o.Nama; // murid tidak berubah
  write_('Penilaian', r._row, o); return true;
}
function hapusPenilaian(pw, no) {
  auth_(pw);
  const r = read_('Penilaian').find(x => x.NoNilai === no);
  if (!r) throw new Error('Penilaian tidak ditemukan');
  sh_('Penilaian').deleteRow(r._row); return true;
}

// ---------- Rekening ----------
function simpanRekening(pw, o) {
  auth_(pw);
  const max = read_('Rekening').reduce((m, x) => Math.max(m, Number(String(x.Kode).slice(1)) || 0), 0);
  append_('Rekening', {Kode: 'R' + String(max + 1).padStart(3, '0'), Bank: String(o.Bank || '').trim(),
    NoRek: String(o.NoRek || '').trim(), AtasNama: String(o.AtasNama || '').trim()});
  return read_('Rekening');
}
function hapusRekening(pw, kode) {
  auth_(pw);
  read_('Rekening').filter(x => x.Kode === kode).map(x => x._row).sort((a, b) => b - a).forEach(r => sh_('Rekening').deleteRow(r));
  return read_('Rekening');
}

// ---------- Invoice (tagihan orang tua) ----------
function getInvoiceSiswa(pw, id) { auth_(pw); ada_(id); return read_('Invoice').filter(x => x.ID === id); }

// Terbitkan invoice untuk satu periode. Nomor, nominal, dan rekening dikunci saat diterbitkan.
function buatInvoice(pw, id, periode, rekening) {
  auth_(pw);
  const s = ada_(id);
  const sudah = read_('Invoice').find(e => e.ID === id && String(e.Periode) === String(periode));
  if (!sudah) {
    append_('Invoice', {
      NoInvoice: 'INV-' + id + '-' + periode, ID: id, Nama: s.Nama,
      Periode: periode, Pertemuan: Number(s.JmlPertemuan) || 4,
      Total: Number(s.Tarif) || 0, Rekening: rekening || '', Status: 'Belum Bayar',
      Terbit: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd')
    });
  }
  return read_('Invoice').filter(x => x.ID === id);
}

// Batalkan invoice (supaya bisa diterbitkan ulang, mis. untuk ganti rekening).
function hapusInvoice(pw, no) {
  auth_(pw);
  const f = read_('Invoice').find(x => x.NoInvoice === no);
  if (!f) return [];
  sh_('Invoice').deleteRow(f._row);
  return read_('Invoice').filter(x => x.ID === f.ID);
}

function setStatusInvoice(pw, no, status) {
  auth_(pw);
  const f = read_('Invoice').find(x => x.NoInvoice === no);
  if (f) {
    const col = hdr_('Invoice').indexOf('Status') + 1;
    sh_('Invoice').getRange(f._row, col).setNumberFormat('@').setValue(status);
  }
  return true;
}

// === AKHIR FILE ===

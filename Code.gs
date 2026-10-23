// ============================================================
// 🛒 WISHLIST — Apps Script (Code.gs)
// Web app (Index.html) + auto gambar + rapikan sheet
//
// Data mulai baris 6. Susunan kolom & merge mengikuti header baris 5
// (No, Kategori, Barang, Image, Spesifikasi, Harga, Prioritas,
// Status, Link, Keterangan).
//
// ⚠️ PAKAI FILE INI SAJA. Hapus file .gs lama di project, karena
// nama yang sama (START_ROW, formatWishlistRow, dll) akan bentrok.
//
// Fitur lama yang sudah tergabung di sini:
// AUTO IMAGE + AUTO CAPSLOCK + AUTO HEIGHT + WRAP + MERGE
// ============================================================

const SHEET_ID = '1a0lW7cSRhGtafUjUE6ww2SRuQVtthlbV2iDFR1VpjXI';
const SHEET_NAME = 'WISH LIST';   // nama tab; kalau tidak ketemu, pakai tab pertama
const FOLDER_NAME = 'Wishlist Images'; // folder Drive untuk gambar upload
const START_ROW = 6;
const TOTAL_COLS = 14;
const IMAGE_SIZE = 120;
const MIN_ROW_HEIGHT = 120;

// Posisi kolom & merge DIBACA OTOMATIS dari header baris 5,
// jadi ikut berubah kalau susunan kolom di sheet diubah.
// Angka di bawah hanya cadangan kalau header tidak terbaca.
const HEADER_ROW = START_ROW - 1;
const LAYOUT = detectLayout_();
const COL = LAYOUT.col;

function detectLayout_() {
  const col = { NO: 1, KATEGORI: 2, BARANG: 3, IMAGE: 4, SPEK: 5, HARGA: 7,
                PRIORITAS: 8, STATUS: 9, LINK: 11, KET: 13 };
  let merges = [[5, 2], [11, 2], [13, 2]];
  try {
    const sheet = getSheet_();
    const range = sheet.getRange(HEADER_ROW, 1, 1, Math.min(TOTAL_COLS, sheet.getMaxColumns()));
    const heads = range.getValues()[0].map(function (h) { return String(h).toLowerCase(); });
    const keys = { KATEGORI: 'kategori', BARANG: 'barang', IMAGE: 'image', SPEK: 'spesifikasi',
                   HARGA: 'harga', PRIORITAS: 'prioritas', STATUS: 'status',
                   LINK: 'link', KET: 'keterangan' };
    Object.keys(keys).forEach(function (k) {
      const i = heads.findIndex(function (h) { return h.indexOf(keys[k]) > -1; });
      if (i > -1) col[k] = i + 1;
    });
    const found = range.getMergedRanges()
      .filter(function (r) { return r.getRow() === HEADER_ROW && r.getNumRows() === 1; })
      .map(function (r) { return [r.getColumn(), r.getNumColumns()]; });
    if (found.length) merges = found;
  } catch (err) {
    console.log(err);
  }
  return { col: col, merges: merges };
}

// Jalankan manual untuk mengecek hasil deteksi (lihat di Execution log)
function cekLayout() {
  console.log(JSON.stringify(LAYOUT));
}

// Teks harus SAMA PERSIS dengan isi sheet (dipakai rumus hitung di baris 3-4)
const OPTIONS = {
  kategori: ['💻🛵 Aksesoris', '👕 Fashion', '👟 Sepatu', '🔧 Peralatan',
             '🍜 Makanan & Minuman', '📱 Elektronik', '🏠 Furniture',
             '📚 Buku & Edukasi', '🧴🧼 Perawatan'],
  prioritas: ['🔴 Penting', '🟡 Ga Terlalu', '🟢 Pingin'],
  status: ['🛒 Mau Beli', '🙏 Belum', '✅ Sudah Dibeli', '❌ Batal']
};


// ============================================================
// 1. WEB APP
// ============================================================

// Password untuk versi GitHub (dan untuk membuka web app dengan ?k=PASSWORD).
// GANTI sebelum deploy dengan akses "Anyone"!
const API_TOKEN = '0317';

// Fungsi yang boleh dipanggil dari luar (versi GitHub)
const API_FUNCS = {
  getInitData: getInitData,
  addWish: addWish,
  updateWish: updateWish,
  deleteWish: deleteWish,
  getImageData: getImageData
};

function doGet(e) {
  const k = e && e.parameter && e.parameter.k;
  // Halaman bawaan Apps Script hanya dibuka kalau password belum diatur
  // (pemakaian pribadi lama) atau link memakai ?k=PASSWORD
  if (API_TOKEN === 'GANTI_PASSWORD' || k === API_TOKEN) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('My Wishlist ⭐')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return ContentService.createTextOutput('Wishlist API aktif.');
}

// Dipanggil oleh versi GitHub (fetch POST, isi berupa teks JSON)
function doPost(e) {
  let out;
  try {
    if (API_TOKEN === 'GANTI_PASSWORD') throw new Error('API_TOKEN di Code.gs belum diganti.');
    const req = JSON.parse(e.postData.contents);
    if (req.token !== API_TOKEN) throw new Error('Password salah');
    const fn = API_FUNCS[req.fn];
    if (!fn) throw new Error('Fungsi tidak dikenal.');
    out = { result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function getInitData() {
  return { options: OPTIONS, wishes: getWishes_() };
}

// Dipanggil dari HTML saat tombol "Simpan" ditekan
function addWish(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const barang = String(data.barang || '').trim();
    if (!barang) throw new Error('Nama barang wajib diisi.');

    const sheet = getSheet_();
    const row = Math.max(lastDataRow_(sheet) + 1, START_ROW);
    if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 5);

    // Salin tampilan & dropdown dari baris di atasnya
    if (row > START_ROW) {
      const src = sheet.getRange(row - 1, 1, 1, TOTAL_COLS);
      const dst = sheet.getRange(row, 1, 1, TOTAL_COLS);
      src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }

    mergeWishlistRow(sheet, row);

    // Kolom No diisi otomatis oleh ARRAYFORMULA di A6, JANGAN ditulis dari sini
    // (menulis angka di bawah rumus itu membuatnya jadi #REF!).
    sheet.getRange(row, COL.KATEGORI).setValue(data.kategori || '');
    sheet.getRange(row, COL.BARANG).setValue(barang.toUpperCase());
    sheet.getRange(row, COL.SPEK).setValue(String(data.spesifikasi || '').toUpperCase());
    sheet.getRange(row, COL.HARGA)
      .setValue(Number(data.harga) || '')
      .setNumberFormat('"Rp"#,##0.00');
    sheet.getRange(row, COL.PRIORITAS).setValue(data.prioritas || '');
    sheet.getRange(row, COL.STATUS).setValue(data.status || '');
    sheet.getRange(row, COL.LINK).setValue(data.link || '');
    sheet.getRange(row, COL.KET).setValue(String(data.keterangan || '').toUpperCase());

    // Prioritas gambar: upload > URL gambar > otomatis dari link
    let imageUrl = '';
    if (data.imageBase64 || /^data:image/i.test(data.imageUrl || '')) imageUrl = saveImage_(data.imageBase64 || data.imageUrl, barang);
    else if (data.imageUrl) imageUrl = String(data.imageUrl).trim();
    else if (data.link) imageUrl = fetchImageFromLink_(String(data.link).trim());

    if (imageUrl) setImageFormula_(sheet.getRange(row, COL.IMAGE), imageUrl);

    formatWishlistRow(sheet, row);
    SpreadsheetApp.flush();

    return { ok: true, row: row, gambar: !!imageUrl, wishes: getWishes_() };
  } finally {
    lock.releaseLock();
  }
}


// ============================================================
// 2. BACA DATA
// ============================================================

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

// Baris terakhir yang kolom C (Barang)-nya terisi
function lastDataRow_(sheet) {
  const max = sheet.getMaxRows();
  if (max < START_ROW) return START_ROW - 1;
  const vals = sheet.getRange(START_ROW, COL.BARANG, max - START_ROW + 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return START_ROW + i;
  }
  return START_ROW - 1;
}

function getWishes_() {
  const sheet = getSheet_();
  const last = lastDataRow_(sheet);
  if (last < START_ROW) return [];
  const n = last - START_ROW + 1;
  const vals = sheet.getRange(START_ROW, 1, n, Math.max(COL.STATUS, COL.LINK, COL.KET)).getValues();
  const forms = sheet.getRange(START_ROW, COL.IMAGE, n, 1).getFormulas();
  const out = [];
  for (let i = 0; i < n; i++) {
    const v = vals[i];
    if (!String(v[COL.BARANG - 1]).trim()) continue;
    const m = String(forms[i][0]).match(/IMAGE\("([^"]+)"/i);
    let img = m ? m[1] : '';
    let hasImg = !!img;
    const cv = v[COL.IMAGE - 1]; // gambar yang disisipkan langsung ke sel (bukan rumus)
    if (cv && typeof cv === 'object' && cv.getContentUrl) {
      hasImg = true;
      if (!img) { try { img = cv.getUrl() || ''; } catch (err) {} }
    }
    out.push({
      row: START_ROW + i,
      hasImg: hasImg,
      spek: String(v[COL.SPEK - 1] || '').slice(0, 3000),
      link: String(v[COL.LINK - 1] || '').trim(),
      ket: String(v[COL.KET - 1] || '').slice(0, 3000),
      no: v[0],
      kategori: v[COL.KATEGORI - 1],
      barang: v[COL.BARANG - 1],
      harga: Number(v[COL.HARGA - 1]) || 0,
      prioritas: v[COL.PRIORITAS - 1],
      status: v[COL.STATUS - 1],
      img: img
    });
  }
  return out.reverse(); // terbaru di atas
}


// ============================================================
// 3. GAMBAR
// ============================================================

// Simpan gambar upload ke Drive, balikin link yang bisa dipakai =IMAGE()
function saveImage_(base64, name) {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
  const safe = name.replace(/[^\w\- ]/g, '').slice(0, 40) || 'wish';
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64.split(',').pop()), 'image/jpeg', safe + '.jpg');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}

// Ambil gambar dari link produk (og:image → twitter:image → Shopee)
function fetchImageFromLink_(link) {
  if (!link) return '';
  if (/\.(jpg|jpeg|png|gif|webp)(\?.*)?$/i.test(link)) return link;
  try {
    const html = UrlFetchApp.fetch(link, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36' }
    }).getContentText();

    const patterns = [
      /<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']+)["']/i,
      /<meta[^>]*content=["']([^"']+)["'][^>]*property=["']og:image["']/i,
      /<meta[^>]*name=["']twitter:image["'][^>]*content=["']([^"']+)["']/i,
      /<meta[^>]*content=["']([^"']+)["'][^>]*name=["']twitter:image["']/i,
      /(https?:\/\/[^"'\\ ]*img\.susercontent\.com[^"'\\ ]+)/i
    ];
    for (const p of patterns) {
      const m = html.match(p);
      if (m) {
        return m[1]
          .replace(/\\u002F/g, '/').replace(/\\\//g, '/')
          .replace(/&amp;/g, '&').replace(/\\u0026/g, '&')
          .replace(/\\u003D/g, '=').replace(/\\u003F/g, '?');
      }
    }
  } catch (err) {
    console.log(err);
  }
  return '';
}

// setFormula SELALU pakai koma (bukan titik koma), apa pun bahasa spreadsheet
function setImageFormula_(cell, url) {
  cell.setFormula('=IMAGE("' + url.replace(/"/g, '""') + '",4,' + IMAGE_SIZE + ',' + IMAGE_SIZE + ')');
}


// ============================================================
// 4. FORMAT & MERGE
// ============================================================

function formatWishlistRow(sheet, row) {
  const r = sheet.getRange(row, 1, 1, Math.max(sheet.getLastColumn(), TOTAL_COLS));
  r.setWrap(true).setVerticalAlignment('middle');
  sheet.autoResizeRows(row, 1);
  if (sheet.getRowHeight(row) < MIN_ROW_HEIGHT) sheet.setRowHeight(row, MIN_ROW_HEIGHT);
}

function mergeWishlistRow(sheet, row) {
  LAYOUT.merges.forEach(function (c) {
    const range = sheet.getRange(row, c[0], 1, c[1]);
    if (!range.isPartOfMerge()) range.merge();
  });
}

// Jalankan manual kalau sheet sudah banyak data lama
function rapikanSemuaWishlist() {
  const sheet = getSheet_();
  const end = Math.max(lastDataRow_(sheet), START_ROW);
  for (let row = START_ROW; row <= end; row++) {
    [COL.BARANG, COL.SPEK, COL.KET].forEach(function (c) {
      const cell = sheet.getRange(row, c);
      const v = cell.getValue();
      if (typeof v === 'string' && v.trim() !== '') cell.setValue(v.toUpperCase());
    });
    mergeWishlistRow(sheet, row);
    formatWishlistRow(sheet, row);
  }
  SpreadsheetApp.openById(SHEET_ID).toast('Wishlist berhasil dirapikan! ✨', 'Selesai', 3);
}


// ============================================================
// 5. EDIT MANUAL DI SHEET (installable trigger)
//
// Menggantikan onEdit lama. Fitur: AUTO CAPSLOCK nama barang,
// AUTO IMAGE dari link, MERGE, WRAP, AUTO HEIGHT.
//
// PENTING: onEdit biasa (simple trigger) tidak boleh memakai
// UrlFetchApp, jadi auto gambar gagal. Jalankan setupTrigger()
// SATU KALI dari editor Apps Script.
// ============================================================

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'onEditWishlist' || fn === 'onEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditWishlist')
    .forSpreadsheet(SHEET_ID)
    .onEdit()
    .create();
  SpreadsheetApp.openById(SHEET_ID).toast('Trigger terpasang ✅', 'Selesai', 3);
}

function onEditWishlist(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== getSheet_().getName()) return;

  const r1 = e.range.getRow();
  const rEnd = Math.min(r1 + e.range.getNumRows() - 1, r1 + 29); // maks 30 baris sekali edit
  const c1 = e.range.getColumn();
  const c2 = c1 + e.range.getNumColumns() - 1;

  for (let row = Math.max(r1, START_ROW); row <= rEnd; row++) {

    // A. CAPSLOCK: nama barang, spesifikasi, keterangan (link tidak)
    [COL.BARANG, COL.SPEK, COL.KET].forEach(function (c) {
      if (c < c1 || c > c2) return;
      const cell = sheet.getRange(row, c);
      const v = cell.getValue();
      if (typeof v === 'string' && v.trim() !== '' && v !== v.toUpperCase()) {
        cell.setValue(v.toUpperCase());
      }
    });

    // B. Gambar otomatis dari link
    if (COL.LINK >= c1 && COL.LINK <= c2) updateImageFromLink_(sheet, row);

    // C. Merge, wrap, tinggi otomatis
    mergeWishlistRow(sheet, row);
    formatWishlistRow(sheet, row);
  }
}

function updateImageFromLink_(sheet, row) {
  const link = String(sheet.getRange(row, COL.LINK).getValue()).trim();
  const imgCell = sheet.getRange(row, COL.IMAGE);
  if (!link) { imgCell.clearContent(); return; }
  const url = fetchImageFromLink_(link);
  if (url) setImageFormula_(imgCell, url);
  else imgCell.setValue('❌ Gambar tidak ditemukan');
}

// Nama lama, tetap ada supaya tombol/menu lama tidak error
function mergeWishlist() { rapikanSemuaWishlist(); }


// ============================================================
// 6. GAMBAR UNTUK DAFTAR DI WEB
//
// Dipakai kalau gambar tidak bisa dimuat langsung oleh browser
// (diblokir toko online, atau disisipkan langsung ke sel).
// Server mengambil gambarnya lalu mengirim sebagai data.
// ============================================================

function getImageData(row) {
  const sheet = getSheet_();
  const cell = sheet.getRange(row, COL.IMAGE);
  const opts = {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36' }
  };
  let url = '';
  const m = cell.getFormula().match(/IMAGE\("([^"]+)"/i);
  if (m) url = m[1].replace(/&amp;/g, '&');

  const cv = cell.getValue();
  if (!url && cv && typeof cv === 'object' && cv.getContentUrl) {
    try { url = cv.getUrl() || ''; } catch (err) {}
    if (!url) {
      url = cv.getContentUrl();
      opts.headers.Authorization = 'Bearer ' + ScriptApp.getOAuthToken();
    }
  }
  if (!url) return '';

  const res = UrlFetchApp.fetch(url, opts);
  if (res.getResponseCode() !== 200) return '';
  const blob = res.getBlob();
  const bytes = blob.getBytes();
  if (bytes.length > 1500000) return '';
  return 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(bytes);
}

// Untuk pelacakan: lihat isi kolom Image 5 baris pertama (Execution log)
function cekGambar() {
  const sheet = getSheet_();
  for (let r = START_ROW; r < START_ROW + 5; r++) {
    const c = sheet.getRange(r, COL.IMAGE);
    console.log(r, 'rumus:', c.getFormula().slice(0, 120), '| tipe:', typeof c.getValue());
  }
}


// ============================================================
// 7. UBAH & HAPUS DARI WEB
//
// Row dicek ulang dengan nama barang, supaya tidak salah
// baris kalau sheet berubah sejak daftar dimuat.
// ============================================================

function checkRow_(sheet, row, barang) {
  row = Number(row);
  if (!(row >= START_ROW) || row > lastDataRow_(sheet)) {
    throw new Error('Barang tidak ditemukan. Muat ulang halaman.');
  }
  const cur = String(sheet.getRange(row, COL.BARANG).getValue()).trim().toUpperCase();
  if (cur !== String(barang).trim().toUpperCase()) {
    throw new Error('Daftar sudah berubah. Muat ulang halaman.');
  }
  return row;
}

function updateWish(row, barang, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    row = checkRow_(sheet, row, barang);
    if (data.prioritas) sheet.getRange(row, COL.PRIORITAS).setValue(data.prioritas);
    if (data.status) sheet.getRange(row, COL.STATUS).setValue(data.status);
    sheet.getRange(row, COL.HARGA)
      .setValue(Number(data.harga) || '')
      .setNumberFormat('"Rp"#,##0.00');
    SpreadsheetApp.flush();
    return { ok: true, wishes: getWishes_() };
  } finally {
    lock.releaseLock();
  }
}

function deleteWish(row, barang) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    row = checkRow_(sheet, row, barang);
    // Rumus nomor (ARRAYFORMULA) ada di baris pertama; kalau barisnya dihapus, pasang lagi
    const noFormula = row === START_ROW ? sheet.getRange(row, COL.NO).getFormula() : '';
    sheet.deleteRow(row);
    if (noFormula) sheet.getRange(START_ROW, COL.NO).setFormula(noFormula);
    SpreadsheetApp.flush();
    return { ok: true, wishes: getWishes_() };
  } finally {
    lock.releaseLock();
  }
}

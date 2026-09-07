/**
 * ============================================================
 * DAILY GOLD TRADING MANAGEMENT SYSTEM
 * Divisi Bullion PT Pegadaian
 * Code.gs — Backend Core (Fase 2)
 * ============================================================
 * CARA PAKAI:
 * 1. Buka project Apps Script yang sama tempat Setup.gs berada
 *    (yang sudah Anda jalankan setupDatabase()-nya).
 * 2. Klik ikon "+" di samping "Files" > Script > beri nama "Code".
 * 3. Paste seluruh isi file ini ke Code.gs.
 * 4. Simpan (Ctrl+S).
 * Fase ini BELUM bisa dibuka sebagai web app (belum ada
 * frontend/index.html) — itu Fase 3. Untuk sekarang, semua
 * fungsi di sini bisa dites langsung dari editor (pilih nama
 * fungsi di dropdown toolbar > Run), hasilnya cek di Logs.
 * ============================================================
 */

// ------------------------------------------------------------
// KONFIGURASI
// ------------------------------------------------------------
const SHEET_NAMES = {
  SUPPLIER: 'Supplier',
  BUYER: 'Buyer',
  BENCHMARK: 'BenchmarkPrice',
  SUPPLIER_PRICE: 'SupplierPrice',
  OUTSTANDING: 'Outstanding',
  PRICE_SETTING: 'PriceSetting',
  INVOICE: 'Invoice',
  PEMBELIAN: 'Pembelian',
  PENJUALAN_FULL: 'PenjualanTransaksi',
  PEMBELIAN_FULL: 'PembelianTransaksi',
  USERS: 'Users',
  AUDIT: 'AuditTrail',
  SETTINGS: 'Settings',
  MITRA_PROFILE: 'MITRA_PROFILE',
  MITRA_PKS: 'MITRA_PKS',
  MITRA_PKS_DOCUMENTS: 'MITRA_PKS_DOCUMENTS',
  MITRA_COMPANY_VISIT: 'MITRA_COMPANY_VISIT',
  BOBOT: 'BobotSupplier',
  PENETAPAN_BOBOT: 'PenetapanHargaBobot'
};

const ROLE_LEVEL = { Viewer: 1, Supervisor: 2, Admin: 3 };

// ------------------------------------------------------------
// KONEKSI DATABASE
// ------------------------------------------------------------
/**
 * Mengambil object Spreadsheet database.
 * ID diambil otomatis dari Script Properties (disimpan oleh
 * Setup.gs). Kalau untuk suatu alasan Script Properties-nya
 * kosong (misalnya Code.gs dijalankan di project terpisah),
 * ganti FALLBACK_ID di bawah dengan ID spreadsheet Anda.
 */
function getDb() {
  const FALLBACK_ID = ''; // isi manual di sini kalau perlu
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('DB_SPREADSHEET_ID');
  if (!id) id = FALLBACK_ID;
  if (!id) {
    throw new Error('DB_SPREADSHEET_ID belum diset. Jalankan setupDatabase() di Setup.gs terlebih dahulu, atau isi FALLBACK_ID di getDb().');
  }
  return SpreadsheetApp.openById(id);
}

function getSheet(name) {
  const sh = getDb().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" tidak ditemukan.');
  return sh;
}

// ------------------------------------------------------------
// doGet — entry point web app
// ------------------------------------------------------------
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gold Trading Management System')
    .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/favicon3.ico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper untuk include file HTML lain (Styles.html, JavaScript.html)
 * ke dalam Index.html, dipanggil lewat <?!= include('NamaFile'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ------------------------------------------------------------
// HELPER UMUM
// ------------------------------------------------------------
/**
 * Mengubah seluruh isi sheet menjadi array of object,
 * key-nya diambil dari baris header (baris 1).
 */
function sheetToObjects(sheetName) {
  const sh = getSheet(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map((row, idx) => {
      const obj = { _row: idx + 2 }; // nomor baris asli di sheet, untuk update/delete
      headers.forEach((h, i) => {
        let val = row[i];
        // Ubah Date jadi string ISO — mencegah google.script.run
        // mengembalikan null akibat masalah serialisasi objek Date.
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
        }
        obj[h] = val;
      });
      return obj;
    });
}

/**
 * Generate ID berurutan baru, format PREFIX-001, PREFIX-002, dst.
 */
function getNextId(sheetName, idColumnName, prefix) {
  const data = sheetToObjects(sheetName);
  if (data.length === 0) return prefix + '-001';
  const numbers = data
    .map(row => {
      const val = String(row[idColumnName] || '');
      const match = val.match(/(\d+)$/);
      return match ? parseInt(match[1], 10) : 0;
    })
    .filter(n => !isNaN(n));
  const max = numbers.length ? Math.max.apply(null, numbers) : 0;
  return prefix + '-' + String(max + 1).padStart(3, '0');
}

/**
 * Mencari baris berdasarkan nilai kolom ID, lalu update
 * beberapa kolom sekaligus (updates = {NamaKolom: nilaiBaru}).
 */
function updateRowById(sheetName, idColumnName, idValue, updates) {
  const sh = getSheet(sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idColIdx = headers.indexOf(idColumnName);
  if (idColIdx === -1) throw new Error('Kolom ID "' + idColumnName + '" tidak ditemukan di sheet ' + sheetName);

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idColIdx]) === String(idValue)) {
      Object.keys(updates).forEach(key => {
        const colIdx = headers.indexOf(key);
        if (colIdx !== -1) sh.getRange(r + 1, colIdx + 1).setValue(updates[key]);
      });
      return true;
    }
  }
  throw new Error('Data dengan ID "' + idValue + '" tidak ditemukan di sheet ' + sheetName);
}

/**
 * Menghapus baris berdasarkan nilai kolom ID.
 */
function deleteRowById(sheetName, idColumnName, idValue) {
  const sh = getSheet(sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idColIdx = headers.indexOf(idColumnName);
  if (idColIdx === -1) throw new Error('Kolom ID "' + idColumnName + '" tidak ditemukan di sheet ' + sheetName);

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idColIdx]) === String(idValue)) {
      sh.deleteRow(r + 1);
      return true;
    }
  }
  throw new Error('Data dengan ID "' + idValue + '" tidak ditemukan di sheet ' + sheetName);
}

// ------------------------------------------------------------
// AUTENTIKASI & ROLE
// ------------------------------------------------------------
/**
 * Mengambil data user yang sedang login (berdasarkan email
 * akun Google aktif) dari sheet Users.
 */
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const users = sheetToObjects(SHEET_NAMES.USERS);
  const found = users.find(u => u.Email === email);
  if (!found) {
    return { Email: email, Nama: email, Role: 'Viewer', Status: 'Tidak Terdaftar' };
  }
  return found;
}

/**
 * Memastikan user yang login punya role minimal tertentu.
 * Lempar error kalau tidak cukup — dipakai di awal fungsi
 * CRUD yang butuh proteksi (misalnya delete hanya Admin).
 */
function requireRole(minRole) {
  const user = getCurrentUser();
  const userLevel = ROLE_LEVEL[user.Role] || 0;
  const requiredLevel = ROLE_LEVEL[minRole] || 99;
  if (userLevel < requiredLevel) {
    throw new Error('Akses ditolak. Aksi ini butuh role minimal "' + minRole + '", role Anda: "' + user.Role + '".');
  }
  return user;
}

// ------------------------------------------------------------
// AUDIT TRAIL
// ------------------------------------------------------------
function logAudit(action, module, recordId, detail) {
  const sh = getSheet(SHEET_NAMES.AUDIT);
  const user = getCurrentUser();
  const id = getNextId(SHEET_NAMES.AUDIT, 'ID', 'LOG');
  sh.appendRow([id, new Date(), user.Email, action, module, recordId, detail || '']);
}

// ------------------------------------------------------------
// CRUD — SUPPLIER
// ------------------------------------------------------------
function getSuppliers() {
  return sheetToObjects(SHEET_NAMES.SUPPLIER);
}

function addSupplier(nama) {
  requireRole('Supervisor');
  const sh = getSheet(SHEET_NAMES.SUPPLIER);
  const id = getNextId(SHEET_NAMES.SUPPLIER, 'SupplierID', 'SUP');
  sh.appendRow([id, nama, 'Aktif', new Date()]);
  logAudit('CREATE', 'Supplier', id, nama);
  return id;
}

function updateSupplier(supplierId, updates) {
  requireRole('Supervisor');
  updateRowById(SHEET_NAMES.SUPPLIER, 'SupplierID', supplierId, updates);
  logAudit('UPDATE', 'Supplier', supplierId, JSON.stringify(updates));
}

function deleteSupplier(supplierId) {
  requireRole('Admin');
  deleteRowById(SHEET_NAMES.SUPPLIER, 'SupplierID', supplierId);
  logAudit('DELETE', 'Supplier', supplierId, '');
}

// ------------------------------------------------------------
// CRUD — BUYER
// ------------------------------------------------------------
function getBuyers() {
  return sheetToObjects(SHEET_NAMES.BUYER);
}

function addBuyer(nama, kontak) {
  requireRole('Viewer');
  const sh = getSheet(SHEET_NAMES.BUYER);
  const id = getNextId(SHEET_NAMES.BUYER, 'BuyerID', 'BUY');
  sh.appendRow([id, nama, kontak || '', 'Aktif', new Date()]);
  logAudit('CREATE', 'Buyer', id, nama);
  return id;
}

function updateBuyer(buyerId, updates) {
  requireRole('Supervisor');
  updateRowById(SHEET_NAMES.BUYER, 'BuyerID', buyerId, updates);
  logAudit('UPDATE', 'Buyer', buyerId, JSON.stringify(updates));
}

function deleteBuyer(buyerId) {
  requireRole('Admin');
  deleteRowById(SHEET_NAMES.BUYER, 'BuyerID', buyerId);
  logAudit('DELETE', 'Buyer', buyerId, '');
}

// ------------------------------------------------------------
// CRUD — BENCHMARK PRICE
// ------------------------------------------------------------
function getBenchmarkPrices() {
  return sheetToObjects(SHEET_NAMES.BENCHMARK);
}

function addBenchmarkPrice(data) {
  // data = { Tanggal, SpotGold, XAUIDRK, LBMA, Kitco }
  requireRole('Viewer');
  const sh = getSheet(SHEET_NAMES.BENCHMARK);
  const user = getCurrentUser();
  const id = getNextId(SHEET_NAMES.BENCHMARK, 'ID', 'BMK');
  sh.appendRow([
    id, data.Tanggal, data.SpotGold, data.XAUIDRK, data.LBMA, data.Kitco,
    new Date(), user.Email
  ]);
  logAudit('CREATE', 'BenchmarkPrice', id, JSON.stringify(data));
  return id;
}

function updateBenchmarkPrice(id, updates) {
  requireRole('Supervisor');
  updateRowById(SHEET_NAMES.BENCHMARK, 'ID', id, updates);
  logAudit('UPDATE', 'BenchmarkPrice', id, JSON.stringify(updates));
}

function deleteBenchmarkPrice(id) {
  requireRole('Admin');
  deleteRowById(SHEET_NAMES.BENCHMARK, 'ID', id);
  logAudit('DELETE', 'BenchmarkPrice', id, '');
}

// ------------------------------------------------------------
// CRUD — SUPPLIER PRICE (Form 1: Daily Supplier Price)
// ------------------------------------------------------------
function getSupplierPrices() {
  return sheetToObjects(SHEET_NAMES.SUPPLIER_PRICE);
}

/**
 * data = { Tanggal, SupplierID, SupplierNama, Harga1, Harga2,
 *          Harga3, Harga4, Catatan }
 * HargaRataRata dihitung otomatis dari harga yang diisi (>0).
 */
function addSupplierPrice(data) {
  requireRole('Viewer');
  const sh = getSheet(SHEET_NAMES.SUPPLIER_PRICE);
  const user = getCurrentUser();
  const id = getNextId(SHEET_NAMES.SUPPLIER_PRICE, 'ID', 'SPR');

  const harga = [data.Harga1, data.Harga2, data.Harga3, data.Harga4]
    .map(Number)
    .filter(h => h > 0);
  const rata2 = harga.length ? harga.reduce((a, b) => a + b, 0) / harga.length : 0;

  sh.appendRow([
    id, data.Tanggal, data.SupplierID, data.SupplierNama,
    data.Harga1 || '', data.Harga2 || '', data.Harga3 || '', data.Harga4 || '',
    rata2, data.Catatan || '', new Date(), user.Email
  ]);
  logAudit('CREATE', 'SupplierPrice', id, JSON.stringify(data));
  return id;
}

function updateSupplierPrice(id, updates) {
  requireRole('Supervisor');
  // Jika salah satu harga diubah, hitung ulang rata-rata.
  if (updates.Harga1 !== undefined || updates.Harga2 !== undefined ||
      updates.Harga3 !== undefined || updates.Harga4 !== undefined) {
    const current = sheetToObjects(SHEET_NAMES.SUPPLIER_PRICE).find(r => r.ID === id);
    const merged = Object.assign({}, current, updates);
    const harga = [merged.Harga1, merged.Harga2, merged.Harga3, merged.Harga4]
      .map(Number)
      .filter(h => h > 0);
    updates.HargaRataRata = harga.length ? harga.reduce((a, b) => a + b, 0) / harga.length : 0;
  }
  updateRowById(SHEET_NAMES.SUPPLIER_PRICE, 'ID', id, updates);
  logAudit('UPDATE', 'SupplierPrice', id, JSON.stringify(updates));
}

function deleteSupplierPrice(id) {
  requireRole('Admin');
  deleteRowById(SHEET_NAMES.SUPPLIER_PRICE, 'ID', id);
  logAudit('DELETE', 'SupplierPrice', id, '');
}

// ------------------------------------------------------------
// CRUD — OUTSTANDING (Form 2: Piutang Supplier)
// ------------------------------------------------------------
function getOutstandings() {
  return sheetToObjects(SHEET_NAMES.OUTSTANDING);
}

/**
 * data = { Tanggal, SupplierID, SupplierNama, Gramasi,
 *          HargaDasar, Status, Keterangan }
 * Valuasi = Gramasi x HargaDasar, dihitung otomatis.
 */
function addOutstanding(data) {
  requireRole('Viewer');
  const sh = getSheet(SHEET_NAMES.OUTSTANDING);
  const user = getCurrentUser();
  const id = getNextId(SHEET_NAMES.OUTSTANDING, 'ID', 'OUT');
  const valuasi = Number(data.Gramasi || 0) * Number(data.HargaDasar || 0);

  sh.appendRow([
    id, data.Tanggal, data.SupplierID, data.SupplierNama,
    data.Gramasi, data.HargaDasar, valuasi, data.Status || 'Outstanding',
    data.Keterangan || '', new Date(), user.Email
  ]);
  logAudit('CREATE', 'Outstanding', id, JSON.stringify(data));
  return id;
}

function updateOutstanding(id, updates) {
  requireRole('Supervisor');
  if (updates.Gramasi !== undefined || updates.HargaDasar !== undefined) {
    const current = sheetToObjects(SHEET_NAMES.OUTSTANDING).find(r => r.ID === id);
    const merged = Object.assign({}, current, updates);
    updates.Valuasi = Number(merged.Gramasi || 0) * Number(merged.HargaDasar || 0);
  }
  updateRowById(SHEET_NAMES.OUTSTANDING, 'ID', id, updates);
  logAudit('UPDATE', 'Outstanding', id, JSON.stringify(updates));
}

function deleteOutstanding(id) {
  requireRole('Admin');
  deleteRowById(SHEET_NAMES.OUTSTANDING, 'ID', id);
  logAudit('DELETE', 'Outstanding', id, '');
}

// ------------------------------------------------------------
// CRUD — PRICE SETTING (Form 3: Penetapan Harga)
// ------------------------------------------------------------
function getPriceSettings() {
  return sheetToObjects(SHEET_NAMES.PRICE_SETTING);
}

/**
 * data = { Tanggal, HargaSale, HargaBuyback, TanggalBerlaku }
 * HargaMargin = HargaSale - HargaBuyback, dihitung otomatis.
 */
function addPriceSetting(data) {
  requireRole('Supervisor');
  const sh = getSheet(SHEET_NAMES.PRICE_SETTING);
  const user = getCurrentUser();
  const id = getNextId(SHEET_NAMES.PRICE_SETTING, 'ID', 'PRC');
  const margin = Number(data.HargaSale || 0) - Number(data.HargaBuyback || 0);

  sh.appendRow([
    id, data.Tanggal, data.HargaSale, data.HargaBuyback, margin,
    data.TanggalBerlaku, new Date(), user.Email
  ]);
  logAudit('CREATE', 'PriceSetting', id, JSON.stringify(data));
  return id;
}

function updatePriceSetting(id, updates) {
  requireRole('Supervisor');
  if (updates.HargaSale !== undefined || updates.HargaBuyback !== undefined) {
    const current = sheetToObjects(SHEET_NAMES.PRICE_SETTING).find(r => r.ID === id);
    const merged = Object.assign({}, current, updates);
    updates.HargaMargin = Number(merged.HargaSale || 0) - Number(merged.HargaBuyback || 0);
  }
  updateRowById(SHEET_NAMES.PRICE_SETTING, 'ID', id, updates);
  logAudit('UPDATE', 'PriceSetting', id, JSON.stringify(updates));
}

function deletePriceSetting(id) {
  requireRole('Admin');
  deleteRowById(SHEET_NAMES.PRICE_SETTING, 'ID', id);
  logAudit('DELETE', 'PriceSetting', id, '');
}

// ------------------------------------------------------------
// CRUD — INVOICE (Form 4: Input Invoice Penjualan)
// ------------------------------------------------------------
function getInvoices() {
  return sheetToObjects(SHEET_NAMES.INVOICE);
}

/**
 * data = { NomorInvoice, Tanggal, BuyerID, BuyerNama,
 *          Harga, Gram, Status, Catatan }
 * Nominal = Harga x Gram, dihitung otomatis.
 */
function addInvoice(data) {
  requireRole('Viewer');
  const sh = getSheet(SHEET_NAMES.INVOICE);
  const user = getCurrentUser();
  const id = getNextId(SHEET_NAMES.INVOICE, 'ID', 'INV');
  const nominal = Number(data.Harga || 0) * Number(data.Gram || 0);

  sh.appendRow([
    id, data.NomorInvoice, data.Tanggal, data.BuyerID, data.BuyerNama,
    data.Harga, data.Gram, nominal, data.Status || 'Lunas',
    data.Catatan || '', new Date(), user.Email
  ]);
  logAudit('CREATE', 'Invoice', id, JSON.stringify(data));
  return id;
}

function updateInvoice(id, updates) {
  requireRole('Supervisor');
  if (updates.Harga !== undefined || updates.Gram !== undefined) {
    const current = sheetToObjects(SHEET_NAMES.INVOICE).find(r => r.ID === id);
    const merged = Object.assign({}, current, updates);
    updates.Nominal = Number(merged.Harga || 0) * Number(merged.Gram || 0);
  }
  updateRowById(SHEET_NAMES.INVOICE, 'ID', id, updates);
  logAudit('UPDATE', 'Invoice', id, JSON.stringify(updates));
}

function deleteInvoice(id) {
  requireRole('Admin');
  deleteRowById(SHEET_NAMES.INVOICE, 'ID', id);
  logAudit('DELETE', 'Invoice', id, '');
}

// ------------------------------------------------------------
// KALKULASI OTOMATIS (dipakai Dashboard di Fase 5)
// ------------------------------------------------------------
/**
 * Mencari supplier dengan harga rata-rata termurah pada
 * tanggal tertentu (default: hari ini).
 */
function calculateCheapestSupplier(dateStr) {
  const target = dateStr ? formatDateOnly(new Date(dateStr)) : formatDateOnly(new Date());
  const prices = sheetToObjects(SHEET_NAMES.SUPPLIER_PRICE)
    .filter(p => extractDateOnly(p.Tanggal) === target && Number(p.HargaRataRata) > 0);

  if (prices.length === 0) return null;

  prices.sort((a, b) => Number(a.HargaRataRata) - Number(b.HargaRataRata));
  const cheapest = prices[0];

  const benchmark = getTodayBenchmark(target);
  const selisih = benchmark ? Number(cheapest.HargaRataRata) - Number(benchmark.SpotGold) : null;

  return {
    supplierNama: cheapest.SupplierNama,
    harga: Number(cheapest.HargaRataRata),
    selisihDariBenchmark: selisih
  };
}

/**
 * Mengambil data benchmark hari ini (atau tanggal tertentu).
 */
function getTodayBenchmark(dateStr) {
  const target = dateStr ? formatDateOnly(new Date(dateStr)) : formatDateOnly(new Date());
  const list = sheetToObjects(SHEET_NAMES.BENCHMARK)
    .filter(b => extractDateOnly(b.Tanggal) === target);
  return list.length ? list[list.length - 1] : null;
}

/**
 * Rata-rata, tertinggi, terendah margin dari PriceSetting
 * pada tanggal tertentu (default: hari ini).
 */
function calculateMarginStats(dateStr) {
  const target = dateStr ? formatDateOnly(new Date(dateStr)) : formatDateOnly(new Date());
  const list = sheetToObjects(SHEET_NAMES.PRICE_SETTING)
    .filter(p => extractDateOnly(p.Tanggal) === target);

  if (list.length === 0) return { average: 0, highest: 0, lowest: 0 };

  const margins = list.map(p => Number(p.HargaMargin));
  return {
    average: margins.reduce((a, b) => a + b, 0) / margins.length,
    highest: Math.max.apply(null, margins),
    lowest: Math.min.apply(null, margins)
  };
}

/**
 * Total outstanding piutang (gram & nominal) — hanya status "Outstanding".
 */
function calculateTotalOutstanding() {
  const list = sheetToObjects(SHEET_NAMES.OUTSTANDING).filter(o => o.Status === 'Outstanding');
  return {
    totalGram: list.reduce((sum, o) => sum + Number(o.Gramasi || 0), 0),
    totalNominal: list.reduce((sum, o) => sum + Number(o.Valuasi || 0), 0)
  };
}

/**
 * Total penjualan hari ini (nominal, gramasi, jumlah invoice).
 * Sumber: PenjualanTransaksi (data lengkap, termasuk pajak).
 */
function calculateTotalSalesToday(dateStr) {
  const target = dateStr ? formatDateOnly(new Date(dateStr)) : formatDateOnly(new Date());
  const list = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => extractDateOnly(p.Tanggal) === target);

  return {
    totalNominal: list.reduce((sum, p) => sum + Number(p.TotalHarga || 0), 0),
    totalGram: list.reduce((sum, p) => sum + Number(p.Qty || 0), 0),
    jumlahInvoice: list.length
  };
}

function formatDateOnly(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Ambil tanggal (yyyy-MM-dd) dari field Tanggal record TANPA
 * round-trip lewat new Date() — mencegah bug pergeseran 1 hari
 * akibat parsing ulang string ISO sebagai UTC/lokal yang beda
 * dengan Session.getScriptTimeZone(). Aman dipakai untuk semua
 * field Tanggal yang sudah lewat sheetToObjects (Date -> string)
 * maupun yang masih objek Date asli, maupun teks hasil import.
 */
function extractDateOnly(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = String(value || '').trim();
  const isoMatch = str.match(/^(\d{4}-\d{2}-\d{2})/);
  if (isoMatch) return isoMatch[1];

  const dmyMatch = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (dmyMatch) return dmyMatch[1] + '-' + dmyMatch[2].padStart(2, '0') + '-' + dmyMatch[3].padStart(2, '0');

  // Fallback: coba parse Indonesian "21 Juli 2026" atau format lain yang dikenali JS Date
  const bulanID = { januari: 1, februari: 2, maret: 3, april: 4, mei: 5, juni: 6, juli: 7, agustus: 8, september: 9, oktober: 10, november: 11, desember: 12 };
  const idMatch = str.toLowerCase().match(/^(\d{1,2})\s+([a-z]+)\s+(\d{4})/);
  if (idMatch && bulanID[idMatch[2]]) {
    return idMatch[3] + '-' + String(bulanID[idMatch[2]]).padStart(2, '0') + '-' + idMatch[1].padStart(2, '0');
  }

  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return '';
}

/**
 * Parse field Tanggal record jadi timestamp (angka) untuk
 * sorting/perbandingan rentang tanggal (cutoff N hari), tanpa
 * bergantung locale parsing implisit yang rawan salah.
 */
function parseDateValue(value) {
  const iso = extractDateOnly(value);
  return iso ? new Date(iso + 'T00:00:00').getTime() : 0;
}

// ------------------------------------------------------------
// GLOBAL PERIOD FILTER (Dashboard) — helper bersama.
// startDate/endDate berformat 'yyyy-MM-dd', keduanya opsional.
// Kalau tidak dikirim (undefined/null/''), tidak ada filter yang
// diterapkan — supaya fungsi-fungsi di bawah tetap kompatibel
// dengan pemanggil lain (mis. halaman Penjualan) yang belum kirim
// periode. TIDAK mengubah sheet/struktur data apa pun, murni
// filter di memori berdasarkan kolom Tanggal yang sudah ada.
// ------------------------------------------------------------
function isDateInPeriod(value, startDate, endDate) {
  if (!startDate && !endDate) return true;
  const d = extractDateOnly(value);
  if (!d) return false;
  if (startDate && d < extractDateOnly(startDate)) return false;
  if (endDate && d > extractDateOnly(endDate)) return false;
  return true;
}

/**
 * Deret nilai harian (jumlah TotalHarga per hari) dari `startDate`
 * s/d `endDate` (inklusif) — dipakai sparkline KPI Dashboard supaya
 * mengikuti Global Period Filter alih-alih selalu "sejak transaksi
 * pertama".
 */
function buildDailyTotalSeries(list, startDate, endDate) {
  const byDate = {};
  list.forEach(function (p) {
    const d = extractDateOnly(p.Tanggal);
    if (!d) return;
    byDate[d] = (byDate[d] || 0) + Number(p.TotalHarga || 0);
  });
  const result = [];
  const cur = new Date(extractDateOnly(startDate) + 'T00:00:00');
  const end = new Date(extractDateOnly(endDate) + 'T00:00:00');
  while (cur <= end) {
    result.push(byDate[formatDateOnly(cur)] || 0);
    cur.setDate(cur.getDate() + 1);
  }
  return result;
}

// ------------------------------------------------------------
// HELPER "HARI INI" — dipakai untuk recent-list di form (Fase 4)
// ------------------------------------------------------------
function getSupplierPricesToday(dateStr) {
  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  return sheetToObjects(SHEET_NAMES.SUPPLIER_PRICE)
    .filter(p => extractDateOnly(p.Tanggal) === target)
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
}

function getPriceSettingsToday() {
  const target = formatDateOnly(new Date());
  return sheetToObjects(SHEET_NAMES.PRICE_SETTING)
    .filter(p => extractDateOnly(p.Tanggal) === target)
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
}

function getInvoicesToday() {
  const target = formatDateOnly(new Date());
  return sheetToObjects(SHEET_NAMES.INVOICE)
    .filter(inv => extractDateOnly(inv.Tanggal) === target)
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
}

function getOutstandingActive() {
  return sheetToObjects(SHEET_NAMES.OUTSTANDING)
    .filter(o => o.Status === 'Outstanding')
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
}

/**
 * Dipakai form Invoice: cari buyer berdasarkan nama (case-insensitive).
 * Kalau belum ada, otomatis dibuatkan baru. Mengembalikan BuyerID.
 */
function findOrCreateBuyer(nama) {
  requireRole('Viewer');
  const buyers = getBuyers();
  const found = buyers.find(b => String(b.Nama).trim().toLowerCase() === String(nama).trim().toLowerCase());
  if (found) return found.BuyerID;
  return addBuyer(nama.trim(), '');
}

// ------------------------------------------------------------
// FASE 5 — DATA UNTUK CHART & DASHBOARD ANALISIS
// ------------------------------------------------------------

/**
 * Perbandingan harga seluruh supplier hari ini, terurut dari
 * termurah ke termahal. Dipakai untuk horizontal bar chart.
 */
function getSupplierComparisonToday(dateStr) {
  const list = getSupplierPricesToday(dateStr)
    .filter(p => Number(p.HargaRataRata) > 0)
    .sort((a, b) => Number(a.HargaRataRata) - Number(b.HargaRataRata));

  return list.map((p, idx) => ({
    supplierNama: p.SupplierNama,
    harga: Number(p.HargaRataRata),
    isCheapest: idx === 0,
    isMostExpensive: idx === list.length - 1 && list.length > 1
  }));
}

/**
 * Perbandingan benchmark market vs rata-rata harga supplier hari ini.
 * Dipakai untuk bar/line chart perbandingan.
 */
function getBenchmarkComparisonToday(dateStr) {
  const benchmark = getTodayBenchmark(dateStr);
  const suppliers = getSupplierComparisonToday(dateStr);
  const labels = [];
  const values = [];

  if (benchmark) {
    labels.push('Spot Gold', 'LBMA', 'XAUIDRK', 'Kitco');
    values.push(Number(benchmark.SpotGold), Number(benchmark.LBMA), Number(benchmark.XAUIDRK), Number(benchmark.Kitco));
  }
  suppliers.forEach(s => { labels.push(s.supplierNama); values.push(s.harga); });

  return { labels: labels, values: values };
}

/**
 * Outstanding piutang dikelompokkan per supplier (untuk pie chart).
 * Diarahkan ke sumber LIVE (sheet Rekap) — sama seperti Laporan
 * Harian — karena data di sheet Outstanding sudah tidak akurat
 * (hasil import lama). Dipakai Dashboard & AI Insight.
 */
function getOutstandingBySupplier() {
  return getPiutangLive().rows;
}

/**
 * Penjualan (gramasi) hari ini dikelompokkan per buyer (untuk bar chart).
 * Sumber: PenjualanTransaksi.
 */
function getSalesByBuyerToday(dateStr) {
  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  const list = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => extractDateOnly(p.Tanggal) === target);
  const grouped = {};
  list.forEach(p => {
    const key = p.Pembeli;
    grouped[key] = (grouped[key] || 0) + Number(p.Qty || 0);
  });
  return Object.keys(grouped).map(k => ({ buyerNama: k, gram: grouped[k] }));
}

/**
 * Ranking Top Buyer — berdasarkan histori PenjualanTransaksi, opsional
 * difilter oleh Global Period Filter (startDate/endDate 'yyyy-MM-dd').
 */
function getTopBuyers(limit, startDate, endDate) {
  limit = limit || 5;
  const list = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => isDateInPeriod(p.Tanggal, startDate, endDate));
  const grouped = {};
  list.forEach(p => {
    const key = p.Pembeli;
    if (!grouped[key]) grouped[key] = { buyerNama: key, totalGram: 0, totalNominal: 0, jumlahInvoice: 0 };
    grouped[key].totalGram += Number(p.Qty || 0);
    grouped[key].totalNominal += Number(p.TotalHarga || 0);
    grouped[key].jumlahInvoice += 1;
  });
  return Object.values(grouped)
    .sort((a, b) => b.totalNominal - a.totalNominal)
    .slice(0, limit);
}

/**
 * Ranking Top Supplier — berdasarkan histori PembelianTransaksi
 * (data pembelian lengkap, termasuk pajak & diskon), opsional
 * difilter oleh Global Period Filter (startDate/endDate 'yyyy-MM-dd').
 */
function getTopSuppliers(limit, startDate, endDate) {
  limit = limit || 5;
  const list = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL)
    .filter(p => isDateInPeriod(p.Tanggal, startDate, endDate));
  const grouped = {};
  list.forEach(p => {
    const key = p.Seller;
    if (!grouped[key]) grouped[key] = { supplierNama: key, jumlahTransaksi: 0, totalGram: 0, totalHarga: 0 };
    grouped[key].jumlahTransaksi += 1;
    grouped[key].totalGram += Number(p.Qty || 0);
    grouped[key].totalHarga += Number(p.AfterDiskon || 0);
  });
  return Object.values(grouped)
    .map(g => ({
      supplierNama: g.supplierNama,
      jumlahTransaksi: g.jumlahTransaksi,
      totalGram: g.totalGram,
      rataRataHarga: g.jumlahTransaksi ? g.totalHarga / g.jumlahTransaksi : 0
    }))
    .sort((a, b) => b.totalGram - a.totalGram)
    .slice(0, limit);
}

/**
 * DIAGNOSTIK — untuk 30 hari terakhir (skip Sabtu-Minggu, sama
 * seperti getMarginTrendHarian), tampilkan jumlah baris & total
 * nominal Pembelian dan Penjualan per tanggal, DAN hasil kalkulasi
 * marginPercent/marginRp yang sebenarnya dikirim ke chart Dashboard.
 * Dipakai untuk cek kenapa chart Margin Analysis kelihatan cuma
 * keisi sedikit hari.
 */
function diagnosaMarginTrend() {
  const allPenjualan = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const allPembelian = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL);

  Logger.log('Total baris PenjualanTransaksi (semua tanggal): ' + allPenjualan.length);
  Logger.log('Total baris PembelianTransaksi (semua tanggal): ' + allPembelian.length);
  Logger.log('===========================================');

  const trend = getMarginTrendHarian(30);
  trend.forEach(t => {
    Logger.log(t.tanggal + ' | marginPercent=' + (t.average == null ? 'null' : t.average.toFixed(4) + '%') +
      ' | marginRp=' + (t.marginRp == null ? 'null' : 'Rp' + Math.round(t.marginRp).toLocaleString('id-ID')));
  });
}

/**
 * Trend margin (%) per hari — rumus SAMA PERSIS dengan "Margin
 * Rata-rata" di Laporan Harian: (Penjualan - Pembelian) / Pembelian
 * x 100%, dari transaksi PenjualanTransaksi & PembelianTransaksi.
 * Sabtu & Minggu DILEWATI sepenuhnya (tidak dihitung, tidak masuk
 * hasil) karena biasanya libur/tidak ada transaksi — supaya chart
 * "Margin Analysis" di Dashboard tidak penuh titik kosong di
 * akhir pekan. Sheet Penjualan/Pembelian dibaca SEKALI saja lalu
 * dikelompokkan per tanggal di memori — jauh lebih cepat daripada
 * memanggil calculatePenjualanSummaryToday/calculatePembelianSummaryToday
 * berulang kali per hari (yang tiap panggilannya baca ulang sheet
 * dari awal).
 */
function getMarginTrendHarian(days) {
  days = days || 30;
  const allPenjualan = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const allPembelian = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL);

  const pjByDate = {};
  allPenjualan.forEach(p => {
    const d = extractDateOnly(p.Tanggal);
    if (!d) return;
    pjByDate[d] = (pjByDate[d] || 0) + Number(p.TotalHarga || 0);
  });
  const pbByDate = {};
  allPembelian.forEach(p => {
    const d = extractDateOnly(p.Tanggal);
    if (!d) return;
    pbByDate[d] = (pbByDate[d] || 0) + Number(p.TotalHarga || 0);
  });

  const today = new Date();
  const result = [];

  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const dow = d.getDay(); // 0=Minggu, 6=Sabtu
    if (dow === 0 || dow === 6) continue;

    const target = formatDateOnly(d);
    const pjTotal = pjByDate[target] || 0;
    const pbTotal = pbByDate[target] || 0;
    const marginRp = pjTotal - pbTotal;
    const marginPercent = pbTotal ? (marginRp / pbTotal) * 100 : null;
    result.push({ tanggal: target, average: marginPercent, marginRp: (pjTotal || pbTotal) ? marginRp : null });
  }
  return result;
}

/**
 * Trend margin (average/highest/lowest per hari) untuk N hari terakhir.
 */
function getMarginTrend(days) {
  days = days || 7;
  const list = sheetToObjects(SHEET_NAMES.PRICE_SETTING);
  const today = new Date();
  const result = [];

  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const target = formatDateOnly(d);
    const dayList = list.filter(p => extractDateOnly(p.Tanggal) === target);
    const margins = dayList.map(p => Number(p.HargaMargin));
    result.push({
      tanggal: target,
      average: margins.length ? margins.reduce((a, b) => a + b, 0) / margins.length : null
    });
  }
  return result;
}

/**
 * Trend "Barang Tiba" — dibaca LIVE dari sheet F.3.Input_Barang Masuk
 * (spreadsheet Bullion Trading Process), kolom C (Tanggal terima),
 * kolom D (Denominasi), dan kolom E (Qty). Gramasi harian = Denominasi
 * x Qty, dijumlah per hari untuk N hari terakhir (default 30 hari /
 * 1 bulan). Dipakai chart "Barang Tiba" di Dashboard, menggantikan
 * Trend Harga Supplier.
 */
function getBarangTibaTrend(days) {
  days = days || 30;
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = ss.getSheetByName('F.3.Input_Barang Masuk');
  if (!sheet) throw new Error('Sheet "F.3.Input_Barang Masuk" tidak ditemukan.');

  const values = sheet.getDataRange().getValues();
  // Header ada di baris 11 (indeks 10) — data mulai baris 12 (indeks 11).
  // Kolom C (indeks 2) = Tanggal terima, D (indeks 3) = Denominasi,
  // E (indeks 4) = Qty. Gramasi = Denominasi x Qty.
  const rows = values.slice(11);
  const byDate = {};
  rows.forEach(row => {
    const tanggal = row[2];
    const denominasi = Number(row[3]) || 0;
    const qty = Number(row[4]) || 0;
    const gram = denominasi * qty;
    if (!tanggal || !gram) return;
    const d = extractDateOnly(tanggal);
    if (!d) return;
    byDate[d] = (byDate[d] || 0) + gram;
  });

  const today = new Date();
  const result = [];
  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const dow = d.getDay(); // 0=Minggu, 6=Sabtu
    if (dow === 0 || dow === 6) continue;
    const target = formatDateOnly(d);
    result.push({ tanggal: target, total: byDate[target] || null });
  }
  return result;
}

/**
 * Trend harga rata-rata seluruh supplier per hari, untuk N hari terakhir.
 */
function getPriceTrend(days) {
  days = days || 7;
  const list = sheetToObjects(SHEET_NAMES.SUPPLIER_PRICE);
  const today = new Date();
  const result = [];

  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const target = formatDateOnly(d);
    const dayList = list.filter(p => extractDateOnly(p.Tanggal) === target && Number(p.HargaRataRata) > 0);
    const harga = dayList.map(p => Number(p.HargaRataRata));
    result.push({
      tanggal: target,
      average: harga.length ? harga.reduce((a, b) => a + b, 0) / harga.length : null
    });
  }
  return result;
}

/**
 * AI Daily Summary — rule-based (bukan panggil API eksternal),
 * disusun mengikuti pola contoh di spesifikasi awal.
 */
function generateDailySummary(dateStr) {
  const cheapest = calculateCheapestSupplier(dateStr);
  const marginStats = calculateMarginStats(dateStr);
  const sales = calculateTotalSalesToday(dateStr);

  if (!cheapest) {
    return 'Belum ada data harga supplier yang diinput untuk tanggal ini. Input harga di menu Daily Trading untuk melihat ringkasan otomatis.';
  }

  let text = 'Supplier dengan harga paling kompetitif adalah ' + cheapest.supplierNama +
    ' dengan penawaran ' + formatRupiahServer(cheapest.harga) + '/gram';

  if (cheapest.selisihDariBenchmark != null) {
    text += cheapest.selisihDariBenchmark < 0
      ? ', lebih rendah ' + formatRupiahServer(Math.abs(cheapest.selisihDariBenchmark)) + ' dibanding Spot Gold.'
      : ', namun masih ' + formatRupiahServer(cheapest.selisihDariBenchmark) + ' di atas Spot Gold.';
  } else {
    text += '.';
  }

  if (marginStats.average) {
    text += ' Harga tersebut memberikan potensi margin rata-rata ' + formatRupiahServer(marginStats.average) + '/gram';
  }
  if (sales.totalGram) {
    text += ' dengan total penjualan ' + sales.totalGram.toLocaleString('id-ID', { maximumFractionDigits: 2 }) + ' gram pada tanggal ini.';
  } else {
    text += ', dengan belum ada transaksi penjualan tercatat pada tanggal ini.';
  }

  return text;
}

/**
 * Alert otomatis — supplier lebih murah dari benchmark (hijau),
 * harga di atas benchmark (merah), margin negatif (warning).
 */
function getAlerts(dateStr) {
  const alerts = [];
  const cheapest = calculateCheapestSupplier(dateStr);
  const benchmark = getTodayBenchmark(dateStr);
  const marginStats = calculateMarginStats(dateStr);

  if (cheapest && benchmark) {
    if (cheapest.harga < Number(benchmark.SpotGold)) {
      alerts.push({
        type: 'success',
        title: 'Recommended Supplier',
        message: cheapest.supplierNama + ' menawarkan harga di bawah Spot Gold — supplier yang direkomendasikan pada tanggal ini.'
      });
    } else {
      alerts.push({
        type: 'danger',
        title: 'Harga Di Atas Benchmark',
        message: 'Semua supplier pada tanggal ini menawarkan harga di atas atau sama dengan Spot Gold.'
      });
    }
  }

  if (marginStats.lowest < 0) {
    alerts.push({
      type: 'warning',
      title: 'Margin Negatif',
      message: 'Ada penetapan harga dengan margin negatif pada tanggal ini — mohon dicek kembali di menu Daily Trading.'
    });
  }

  return alerts;
}

function formatRupiahServer(num) {
  num = Number(num) || 0;
  return 'Rp' + Math.round(num).toLocaleString('id-ID');
}

// ------------------------------------------------------------
// DASHBOARD — VERSI KESELURUHAN (bukan per tanggal)
// ------------------------------------------------------------

/**
 * Supplier dengan rata-rata harga terendah dari riwayat SupplierPrice,
 * opsional dibatasi Global Period Filter (startDate/endDate 'yyyy-MM-dd').
 */
function getCheapestSupplierOverall(startDate, endDate) {
  const allPrices = sheetToObjects(SHEET_NAMES.SUPPLIER_PRICE)
    .filter(p => Number(p.HargaRataRata) > 0 && isDateInPeriod(p.Tanggal, startDate, endDate));
  const bySupplier = {};
  allPrices.forEach(p => {
    if (!bySupplier[p.SupplierNama]) bySupplier[p.SupplierNama] = [];
    bySupplier[p.SupplierNama].push(Number(p.HargaRataRata));
  });

  let best = null, bestAvg = Infinity;
  Object.keys(bySupplier).forEach(name => {
    const arr = bySupplier[name];
    const avg = arr.reduce((a, b) => a + b, 0) / arr.length;
    if (avg < bestAvg) { bestAvg = avg; best = name; }
  });
  if (!best) return null;

  const latestBenchmark = getLatestBenchmark();
  const selisih = latestBenchmark ? bestAvg - Number(latestBenchmark.SpotGold) : null;
  return { supplierNama: best, harga: bestAvg, selisihDariBenchmark: selisih };
}

/**
 * Entri BenchmarkPrice terbaru (tanpa filter tanggal).
 */
function getLatestBenchmark() {
  const list = sheetToObjects(SHEET_NAMES.BENCHMARK).sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
  return list.length ? list[0] : null;
}

/**
 * Statistik margin dari riwayat PriceSetting, opsional dibatasi
 * Global Period Filter (startDate/endDate 'yyyy-MM-dd').
 */
function calculateMarginStatsOverall(startDate, endDate) {
  const list = sheetToObjects(SHEET_NAMES.PRICE_SETTING)
    .filter(p => isDateInPeriod(p.Tanggal, startDate, endDate));
  if (!list.length) return { average: 0, highest: 0, lowest: 0 };
  const margins = list.map(p => Number(p.HargaMargin));
  return {
    average: margins.reduce((a, b) => a + b, 0) / margins.length,
    highest: Math.max.apply(null, margins),
    lowest: Math.min.apply(null, margins)
  };
}

/**
 * Perbandingan harga supplier — LIVE dari Price Discovery V3 hari ini
 * (fallback ke rata-rata riwayat SupplierPrice lokal kalau live kosong).
 */
function getSupplierComparisonOverall() {
  const benchmarkNames = ['Spot Goldprice', 'XAUIDRK', 'LBMA H-1', 'Kitco H-1'];
  let list = [];

  try {
    const liveData = getSumberDataLive(null);
    list = liveData
      .filter(r => benchmarkNames.indexOf(r.nama) === -1)
      .map(r => {
        const vals = [r.harga1, r.harga2, r.harga3, r.harga4].filter(v => v > 0);
        return vals.length ? { supplierNama: r.nama, harga: vals[vals.length - 1] } : null;
      })
      .filter(r => r);
  } catch (e) {
    list = [];
  }

  if (!list.length) {
    // Fallback: rata-rata seluruh riwayat SupplierPrice lokal.
    const allPrices = sheetToObjects(SHEET_NAMES.SUPPLIER_PRICE).filter(p => Number(p.HargaRataRata) > 0);
    const bySupplier = {};
    allPrices.forEach(p => {
      if (!bySupplier[p.SupplierNama]) bySupplier[p.SupplierNama] = [];
      bySupplier[p.SupplierNama].push(Number(p.HargaRataRata));
    });
    list = Object.keys(bySupplier).map(name => {
      const arr = bySupplier[name];
      return { supplierNama: name, harga: arr.reduce((a, b) => a + b, 0) / arr.length };
    });
  }

  list.sort((a, b) => a.harga - b.harga);
  const marked = list.map((s, idx) => ({
    supplierNama: s.supplierNama,
    harga: s.harga,
    isCheapest: idx === 0,
    isMostExpensive: idx === list.length - 1 && list.length > 1
  }));

  // Dibalik supaya termahal tampil di atas, termurah di bawah pada horizontal bar chart.
  return marked.reverse();
}

/**
 * Perbandingan benchmark terbaru vs rata-rata harga seluruh supplier.
 */
function getBenchmarkComparisonOverall() {
  const benchmark = getLatestBenchmark();
  const suppliers = getSupplierComparisonOverall();
  const labels = [];
  const values = [];

  if (benchmark) {
    labels.push('Spot Gold', 'LBMA', 'XAUIDRK', 'Kitco');
    values.push(Number(benchmark.SpotGold), Number(benchmark.LBMA), Number(benchmark.XAUIDRK), Number(benchmark.Kitco));
  }
  suppliers.forEach(s => { labels.push(s.supplierNama); values.push(s.harga); });

  return { labels: labels, values: values };
}

/**
 * AI Summary versi keseluruhan (dipakai Dashboard).
 */
function generateOverallSummary(startDate, endDate) {
  const cheapest = getCheapestSupplierOverall(startDate, endDate);
  const marginStats = calculateMarginStatsOverall(startDate, endDate);
  const sales = getSalesSummaryAll(startDate, endDate);

  if (!cheapest) {
    return 'Belum ada data harga supplier yang diinput. Input harga di menu Daily Trading untuk melihat ringkasan otomatis.';
  }

  let text = 'Supplier dengan rata-rata harga paling kompetitif secara keseluruhan adalah ' + cheapest.supplierNama +
    ' dengan rata-rata ' + formatRupiahServer(cheapest.harga) + '/gram';

  if (cheapest.selisihDariBenchmark != null) {
    text += cheapest.selisihDariBenchmark < 0
      ? ', lebih rendah ' + formatRupiahServer(Math.abs(cheapest.selisihDariBenchmark)) + ' dibanding Spot Gold terkini.'
      : ', namun masih ' + formatRupiahServer(cheapest.selisihDariBenchmark) + ' di atas Spot Gold terkini.';
  } else {
    text += '.';
  }

  if (marginStats.average) {
    text += ' Margin rata-rata keseluruhan tercatat ' + formatRupiahServer(marginStats.average) + '/gram';
  }
  if (sales.totalGram) {
    text += ' dengan total penjualan keseluruhan ' + sales.totalGram.toLocaleString('id-ID', { maximumFractionDigits: 2 }) + ' gram.';
  } else {
    text += ', dengan belum ada transaksi penjualan tercatat.';
  }

  return text;
}

/**
 * Alert versi keseluruhan (dipakai Dashboard).
 */
function getAlertsOverall(startDate, endDate) {
  const alerts = [];
  const cheapest = getCheapestSupplierOverall(startDate, endDate);
  const benchmark = getLatestBenchmark();
  const marginStats = calculateMarginStatsOverall(startDate, endDate);

  if (cheapest && benchmark) {
    if (cheapest.harga < Number(benchmark.SpotGold)) {
      alerts.push({
        type: 'success',
        title: 'Recommended Supplier',
        message: cheapest.supplierNama + ' memiliki rata-rata harga di bawah Spot Gold terkini — supplier paling direkomendasikan secara keseluruhan.'
      });
    } else {
      alerts.push({
        type: 'danger',
        title: 'Harga Di Atas Benchmark',
        message: 'Rata-rata harga semua supplier berada di atas atau sama dengan Spot Gold terkini.'
      });
    }
  }

  if (marginStats.lowest < 0) {
    alerts.push({
      type: 'warning',
      title: 'Margin Negatif',
      message: 'Ada penetapan harga dengan margin negatif dalam riwayat — mohon dicek kembali di menu Daily Trading.'
    });
  }

  return alerts;
}

// ------------------------------------------------------------
// FASE 6 — AUDIT TRAIL, USER MANAGEMENT, EXPORT
// ------------------------------------------------------------

/**
 * Mengambil entri Audit Trail terbaru (default 50), terbaru dulu.
 */
function getAuditTrail(limit) {
  limit = limit || 50;
  return sheetToObjects(SHEET_NAMES.AUDIT)
    .sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp))
    .slice(0, limit);
}

/**
 * Daftar seluruh user terdaftar (untuk halaman Settings > Role Management).
 */
function getUsers() {
  return sheetToObjects(SHEET_NAMES.USERS);
}

/**
 * Menambah user baru dengan role tertentu. Hanya Admin.
 */
function addUser(email, nama, role) {
  requireRole('Admin');
  const sh = getSheet(SHEET_NAMES.USERS);
  const id = getNextId(SHEET_NAMES.USERS, 'UserID', 'USR');
  sh.appendRow([id, email.trim(), nama.trim(), role, 'Aktif', new Date()]);
  logAudit('CREATE', 'User', id, email + ' (' + role + ')');
  return id;
}

/**
 * Mengubah role seorang user. Hanya Admin.
 */
function updateUserRole(userId, role) {
  requireRole('Admin');
  updateRowById(SHEET_NAMES.USERS, 'UserID', userId, { Role: role });
  logAudit('UPDATE', 'User', userId, 'Role diubah menjadi ' + role);
}

/**
 * Menonaktifkan/mengaktifkan user. Hanya Admin.
 */
function toggleUserStatus(userId, status) {
  requireRole('Admin');
  updateRowById(SHEET_NAMES.USERS, 'UserID', userId, { Status: status });
  logAudit('UPDATE', 'User', userId, 'Status diubah menjadi ' + status);
}

/**
 * URL export database (Google Sheet) ke Excel/PDF — dipakai
 * tombol Export Excel & Export PDF di top header.
 */
function getExportUrls() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  return {
    xlsxUrl: 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx',
    pdfUrl: 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=pdf&size=A4&portrait=false&fitw=true'
  };
}

// ------------------------------------------------------------
// HALAMAN SUPPLIER PRICE — riwayat & chart
// ------------------------------------------------------------
function getSupplierPricesRecent(days) {
  days = days || 30;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  return sheetToObjects(SHEET_NAMES.SUPPLIER_PRICE)
    .filter(p => parseDateValue(p.Tanggal) >= cutoff.getTime())
    .sort((a, b) => parseDateValue(b.Tanggal) - parseDateValue(a.Tanggal));
}

// ------------------------------------------------------------
// HALAMAN BENCHMARK MARKET — trend multi-series
// ------------------------------------------------------------
function getBenchmarkTrend(days) {
  days = days || 7;
  const list = sheetToObjects(SHEET_NAMES.BENCHMARK);
  const today = new Date();
  const result = [];

  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const target = formatDateOnly(d);
    const dayList = list.filter(b => extractDateOnly(b.Tanggal) === target);
    const last = dayList.length ? dayList[dayList.length - 1] : null;
    result.push({
      tanggal: target,
      spotGold: last ? Number(last.SpotGold) : null,
      lbma: last ? Number(last.LBMA) : null,
      xauidrk: last ? Number(last.XAUIDRK) : null,
      kitco: last ? Number(last.Kitco) : null
    });
  }
  return result;
}

// ------------------------------------------------------------
// HALAMAN PENJUALAN — ringkasan & ranking
// ------------------------------------------------------------
/**
 * Ringkasan Penjualan, opsional dibatasi Global Period Filter
 * (startDate/endDate 'yyyy-MM-dd'). Tanpa parameter = seluruh waktu
 * (kompatibel dengan pemanggil lama, mis. halaman Penjualan).
 */
function getSalesSummaryAll(startDate, endDate) {
  const list = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => isDateInPeriod(p.Tanggal, startDate, endDate));
  return {
    totalNominal: list.reduce((sum, p) => sum + Number(p.TotalHarga || 0), 0),
    totalGram: list.reduce((sum, p) => sum + Number(p.Qty || 0), 0),
    jumlahInvoice: list.length
  };
}

function getSalesByBuyerAll(startDate, endDate) {
  const list = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => isDateInPeriod(p.Tanggal, startDate, endDate));
  const grouped = {};
  list.forEach(p => {
    const key = p.Pembeli;
    grouped[key] = (grouped[key] || 0) + Number(p.Qty || 0);
  });
  return Object.keys(grouped)
    .map(k => ({ buyerNama: k, gram: grouped[k] }))
    .sort((a, b) => b.gram - a.gram);
}

/**
 * Total Pembelian (dipakai Dashboard), opsional dibatasi Global
 * Period Filter (startDate/endDate 'yyyy-MM-dd').
 */
function getPembelianSummaryAll(startDate, endDate) {
  const list = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL)
    .filter(p => isDateInPeriod(p.Tanggal, startDate, endDate));
  return {
    totalNominal: list.reduce((sum, p) => sum + Number(p.TotalHarga || 0), 0),
    totalGram: list.reduce((sum, p) => sum + Number(p.Qty || 0), 0),
    jumlahTransaksi: list.length
  };
}

/**
 * Profit Margin = Total Penjualan - Total Pembelian (dipakai kartu
 * "Margin" di Dashboard), opsional dibatasi Global Period Filter.
 */
function calculateProfitMarginOverall(startDate, endDate) {
  const penjualan = getSalesSummaryAll(startDate, endDate);
  const pembelian = getPembelianSummaryAll(startDate, endDate);
  const profitMargin = (penjualan.totalNominal || 0) - (pembelian.totalNominal || 0);
  return {
    profitMargin: profitMargin,
    totalPenjualan: penjualan.totalNominal || 0,
    totalPembelian: pembelian.totalNominal || 0
  };
}

/**
 * Ringkasan kartu "Jumlah Transaksi" Dashboard — gabungan Penjualan +
 * Pembelian pada Global Period Filter yang sedang aktif.
 */
function getTransactionSummary(startDate, endDate) {
  const penjualan = getSalesSummaryAll(startDate, endDate);
  const pembelian = getPembelianSummaryAll(startDate, endDate);
  return {
    totalNominal: (penjualan.totalNominal || 0) + (pembelian.totalNominal || 0),
    jumlahPenjualan: penjualan.jumlahInvoice || 0,
    jumlahPembelian: pembelian.jumlahTransaksi || 0,
    totalTransaksi: (penjualan.jumlahInvoice || 0) + (pembelian.jumlahTransaksi || 0)
  };
}

/**
 * Tanggal transaksi paling awal di antara PenjualanTransaksi &
 * PembelianTransaksi — dipakai sparkline supaya mulai "dari awal".
 */
function getEarliestTransactionDate() {
  const penjualanDates = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL).map(p => parseDateValue(p.Tanggal)).filter(v => v > 0);
  const pembelianDates = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL).map(p => parseDateValue(p.Tanggal)).filter(v => v > 0);
  const all = penjualanDates.concat(pembelianDates);
  if (!all.length) return null;
  return new Date(Math.min.apply(null, all));
}

/**
 * Nilai harian sederhana (angka saja, untuk sparkline mini-chart di
 * kartu KPI Dashboard). Kalau startDate/endDate (Global Period
 * Filter) dikirim, deret mengikuti rentang tersebut; kalau tidak,
 * fallback ke perilaku lama: dari transaksi pertama tercatat s/d
 * hari ini.
 */
function getPenjualanSparklineValues(startDate, endDate) {
  const list = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  let rangeStart, rangeEnd;
  if (startDate && endDate) {
    rangeStart = startDate;
    rangeEnd = endDate;
  } else {
    const earliest = getEarliestTransactionDate();
    if (!earliest) return [0];
    rangeStart = earliest;
    rangeEnd = new Date();
  }
  return buildDailyTotalSeries(list, rangeStart, rangeEnd);
}

function getPembelianSparklineValues(startDate, endDate) {
  const list = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL);
  let rangeStart, rangeEnd;
  if (startDate && endDate) {
    rangeStart = startDate;
    rangeEnd = endDate;
  } else {
    const earliest = getEarliestTransactionDate();
    if (!earliest) return [0];
    rangeStart = earliest;
    rangeEnd = new Date();
  }
  return buildDailyTotalSeries(list, rangeStart, rangeEnd);
}

function getMarginSparklineValues(startDate, endDate) {
  const penjualan = getPenjualanSparklineValues(startDate, endDate);
  const pembelian = getPembelianSparklineValues(startDate, endDate);
  return penjualan.map((v, i) => v - (pembelian[i] || 0));
}

/**
 * Top 5 Merk berdasarkan total nominal Penjualan (field Merk di
 * PenjualanTransaksi — nama brand/produk emas yang terjual),
 * opsional dibatasi Global Period Filter (startDate/endDate 'yyyy-MM-dd').
 */
function getTopMerkPenjualan(limit, startDate, endDate) {
  limit = limit || 5;
  const list = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => p.Merk && isDateInPeriod(p.Tanggal, startDate, endDate));
  const grouped = {};
  let total = 0;
  list.forEach(p => {
    const key = p.Merk;
    grouped[key] = (grouped[key] || 0) + Number(p.TotalHarga || 0);
    total += Number(p.TotalHarga || 0);
  });
  const rows = Object.keys(grouped)
    .map(k => ({ merk: k, nominal: grouped[k] }))
    .sort((a, b) => b.nominal - a.nominal)
    .slice(0, limit);
  rows.forEach(r => { r.percent = total ? (r.nominal / total) * 100 : 0; });
  return { rows: rows, total: total };
}

function getAllInvoices() {
  return sheetToObjects(SHEET_NAMES.INVOICE).sort((a, b) => parseDateValue(b.Tanggal) - parseDateValue(a.Tanggal));
}

// ------------------------------------------------------------
// HALAMAN ANALYTICS — AI Insight (rule-based, 9 poin)
// ------------------------------------------------------------
function getAiInsights() {
  const insights = [];
  const allPrices = sheetToObjects(SHEET_NAMES.SUPPLIER_PRICE).filter(p => Number(p.HargaRataRata) > 0);
  const bySupplier = {};
  allPrices.forEach(p => {
    if (!bySupplier[p.SupplierNama]) bySupplier[p.SupplierNama] = [];
    bySupplier[p.SupplierNama].push(Number(p.HargaRataRata));
  });

  // 1. Supplier Terbaik (rata-rata harga terendah)
  let bestSupplier = null, bestAvg = Infinity;
  Object.keys(bySupplier).forEach(name => {
    const arr = bySupplier[name];
    const avg = arr.reduce((a, b) => a + b, 0) / arr.length;
    if (avg < bestAvg) { bestAvg = avg; bestSupplier = name; }
  });
  insights.push({
    icon: 'workspace_premium', title: 'Supplier Terbaik',
    text: bestSupplier ? bestSupplier + ' memiliki rata-rata harga terendah sepanjang histori (' + formatRupiahServer(bestAvg) + '/gram).' : 'Belum cukup data harga supplier.'
  });

  // 2. Supplier Paling Konsisten (std dev terendah, minimal 2 entri)
  let consistentSupplier = null, lowestStdDev = Infinity;
  Object.keys(bySupplier).forEach(name => {
    const arr = bySupplier[name];
    if (arr.length < 2) return;
    const avg = arr.reduce((a, b) => a + b, 0) / arr.length;
    const variance = arr.reduce((sum, v) => sum + Math.pow(v - avg, 2), 0) / arr.length;
    const stdDev = Math.sqrt(variance);
    if (stdDev < lowestStdDev) { lowestStdDev = stdDev; consistentSupplier = name; }
  });
  insights.push({
    icon: 'verified', title: 'Supplier Paling Konsisten',
    text: consistentSupplier ? consistentSupplier + ' menunjukkan fluktuasi harga paling stabil (deviasi ~' + formatRupiahServer(lowestStdDev) + ').' : 'Butuh minimal 2 entri harga per supplier untuk analisis ini.'
  });

  // 3. Trend Harga (7 hari terakhir vs 7 hari sebelumnya)
  const trend14 = getPriceTrend(14);
  const first7 = trend14.slice(0, 7).map(d => d.average).filter(v => v != null);
  const last7 = trend14.slice(7).map(d => d.average).filter(v => v != null);
  if (first7.length && last7.length) {
    const avgFirst = first7.reduce((a, b) => a + b, 0) / first7.length;
    const avgLast = last7.reduce((a, b) => a + b, 0) / last7.length;
    const pct = avgFirst ? ((avgLast - avgFirst) / avgFirst) * 100 : 0;
    insights.push({
      icon: 'trending_up', title: 'Trend Harga',
      text: 'Harga rata-rata supplier ' + (pct >= 0 ? 'naik' : 'turun') + ' sekitar ' + Math.abs(pct).toFixed(1) + '% dibanding 7 hari sebelumnya.'
    });
  } else {
    insights.push({ icon: 'trending_up', title: 'Trend Harga', text: 'Data belum cukup untuk melihat trend mingguan.' });
  }

  // 4 & 5. Rekomendasi Harga Jual & Potensi Margin
  const cheapest = calculateCheapestSupplier(null);
  const marginStats = calculateMarginStats(null);
  if (cheapest) {
    const suggestedMargin = marginStats.average || 10000;
    insights.push({
      icon: 'sell', title: 'Rekomendasi Harga Jual',
      text: 'Berdasarkan harga termurah hari ini (' + formatRupiahServer(cheapest.harga) + '), harga jual disarankan sekitar ' + formatRupiahServer(cheapest.harga + suggestedMargin) + '/gram untuk menjaga margin rata-rata.'
    });
  } else {
    insights.push({ icon: 'sell', title: 'Rekomendasi Harga Jual', text: 'Input harga supplier hari ini dulu untuk mendapat rekomendasi.' });
  }
  insights.push({
    icon: 'percent', title: 'Potensi Margin',
    text: 'Margin rata-rata hari ini ' + formatRupiahServer(marginStats.average) + '/gram (tertinggi ' + formatRupiahServer(marginStats.highest) + ', terendah ' + formatRupiahServer(marginStats.lowest) + ').'
  });

  // 6. Perubahan Benchmark (hari ini vs kemarin)
  const bTrend = getBenchmarkTrend(2);
  const yesterday = bTrend[0], todayB = bTrend[1];
  if (yesterday && todayB && yesterday.spotGold && todayB.spotGold) {
    const delta = todayB.spotGold - yesterday.spotGold;
    const pct = (delta / yesterday.spotGold) * 100;
    insights.push({
      icon: 'monitoring', title: 'Perubahan Benchmark',
      text: 'Spot Gold ' + (delta >= 0 ? 'naik' : 'turun') + ' ' + formatRupiahServer(Math.abs(delta)) + ' (' + Math.abs(pct).toFixed(2) + '%) dibanding kemarin.'
    });
  } else {
    insights.push({ icon: 'monitoring', title: 'Perubahan Benchmark', text: 'Butuh data benchmark 2 hari berturut-turut untuk membandingkan.' });
  }

  // 7. Risiko Piutang
  const outstandingBySupplier = getOutstandingBySupplier();
  if (outstandingBySupplier.length) {
    const riskiest = outstandingBySupplier.sort((a, b) => b.valuasi - a.valuasi)[0];
    insights.push({
      icon: 'warning', title: 'Risiko Piutang',
      text: 'Konsentrasi piutang terbesar ada pada ' + riskiest.supplierNama + ' senilai ' + formatRupiahServer(riskiest.valuasi) + ' — pertimbangkan percepatan penagihan.'
    });
  } else {
    insights.push({ icon: 'warning', title: 'Risiko Piutang', text: 'Tidak ada outstanding piutang aktif saat ini.' });
  }

  // 8. Buyer Terbesar
  const topBuyers = getTopBuyers(1);
  if (topBuyers.length) {
    insights.push({
      icon: 'emoji_events', title: 'Buyer Terbesar',
      text: topBuyers[0].buyerNama + ' adalah buyer dengan nilai transaksi tertinggi (' + formatRupiahServer(topBuyers[0].totalNominal) + ' dari ' + topBuyers[0].jumlahInvoice + ' invoice).'
    });
  } else {
    insights.push({ icon: 'emoji_events', title: 'Buyer Terbesar', text: 'Belum ada transaksi invoice tercatat.' });
  }

  // 9. Opportunity
  const benchmarkToday = getTodayBenchmark(null);
  if (cheapest && benchmarkToday && cheapest.selisihDariBenchmark != null && cheapest.selisihDariBenchmark < 0) {
    insights.push({
      icon: 'lightbulb', title: 'Opportunity',
      text: 'Harga termurah hari ini ' + formatRupiahServer(Math.abs(cheapest.selisihDariBenchmark)) + ' di bawah Spot Gold — peluang margin ekstra bila volume pembelian ditingkatkan.'
    });
  } else {
    insights.push({ icon: 'lightbulb', title: 'Opportunity', text: 'Belum ada peluang harga signifikan terdeteksi hari ini.' });
  }

  return insights;
}

// ------------------------------------------------------------
// CRUD — PEMBELIAN (transaksi beli emas dari supplier, terpisah dari Piutang)
// ------------------------------------------------------------
function getPembelian() {
  return sheetToObjects(SHEET_NAMES.PEMBELIAN);
}

/**
 * data = { Tanggal, SupplierID, SupplierNama, Gramasi, Harga, Status, Catatan }
 * Nominal = Gramasi x Harga, dihitung otomatis.
 */
function addPembelian(data) {
  requireRole('Viewer');
  const sh = getSheet(SHEET_NAMES.PEMBELIAN);
  const user = getCurrentUser();
  const id = getNextId(SHEET_NAMES.PEMBELIAN, 'ID', 'PMB');
  const nominal = Number(data.Gramasi || 0) * Number(data.Harga || 0);

  sh.appendRow([
    id, data.Tanggal, data.SupplierID, data.SupplierNama,
    data.Gramasi, data.Harga, nominal, data.Status || 'Lunas',
    data.Catatan || '', new Date(), user.Email
  ]);
  logAudit('CREATE', 'Pembelian', id, JSON.stringify(data));
  return id;
}

function updatePembelian(id, updates) {
  requireRole('Supervisor');
  if (updates.Gramasi !== undefined || updates.Harga !== undefined) {
    const current = sheetToObjects(SHEET_NAMES.PEMBELIAN).find(r => r.ID === id);
    const merged = Object.assign({}, current, updates);
    updates.Nominal = Number(merged.Gramasi || 0) * Number(merged.Harga || 0);
  }
  updateRowById(SHEET_NAMES.PEMBELIAN, 'ID', id, updates);
  logAudit('UPDATE', 'Pembelian', id, JSON.stringify(updates));
}

function deletePembelian(id) {
  requireRole('Admin');
  deleteRowById(SHEET_NAMES.PEMBELIAN, 'ID', id);
  logAudit('DELETE', 'Pembelian', id, '');
}

function getPembelianToday() {
  const target = formatDateOnly(new Date());
  return sheetToObjects(SHEET_NAMES.PEMBELIAN)
    .filter(p => extractDateOnly(p.Tanggal) === target)
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
}

function getPembelianRecent(days) {
  days = days || 30;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  return sheetToObjects(SHEET_NAMES.PEMBELIAN)
    .filter(p => parseDateValue(p.Tanggal) >= cutoff.getTime())
    .sort((a, b) => parseDateValue(b.Tanggal) - parseDateValue(a.Tanggal));
}

function calculateTotalPembelianToday(dateStr) {
  const target = dateStr ? formatDateOnly(new Date(dateStr)) : formatDateOnly(new Date());
  const list = sheetToObjects(SHEET_NAMES.PEMBELIAN)
    .filter(p => extractDateOnly(p.Tanggal) === target);
  return {
    totalNominal: list.reduce((sum, p) => sum + Number(p.Nominal || 0), 0),
    totalGram: list.reduce((sum, p) => sum + Number(p.Gramasi || 0), 0),
    jumlahTransaksi: list.length
  };
}

// ------------------------------------------------------------
// CRUD — PENJUALAN TRANSAKSI (struktur lengkap sesuai rekap Excel)
// ------------------------------------------------------------
const PPH_PASAL_22_RATE = 0.0025; // 0.25%

function getPenjualanTransaksi() {
  return sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
}

/**
 * data = { Tanggal, NomorInvoice, HargaPerGram, Qty, Pembeli,
 *          Denominasi, Merk, Status, PaymentTime, Bank,
 *          TanggalPengambilan, EstimasiPenjemputan, Catatan }
 * TotalHarga = HargaPerGram x Qty
 * PphPasal22 = TotalHarga x 0.25%
 * TotalHargaAfterTax = TotalHarga + PphPasal22
 */
function addPenjualanTransaksi(data) {
  requireRole('Viewer');
  const sh = getSheet(SHEET_NAMES.PENJUALAN_FULL);
  const user = getCurrentUser();
  const id = getNextId(SHEET_NAMES.PENJUALAN_FULL, 'ID', 'PJL');

  const totalHarga = Number(data.HargaPerGram || 0) * Number(data.Qty || 0);
  const pph = totalHarga * PPH_PASAL_22_RATE;
  const totalAfterTax = totalHarga + pph;

  sh.appendRow([
    id, data.Tanggal, data.NomorInvoice, data.HargaPerGram, data.Qty,
    totalHarga, pph, totalAfterTax,
    data.Pembeli, data.Denominasi, data.Merk, data.Status || 'V',
    data.PaymentTime || '', data.Bank || '', data.TanggalPengambilan || '', data.EstimasiPenjemputan || '',
    data.Catatan || '', new Date(), user.Email
  ]);
  logAudit('CREATE', 'PenjualanTransaksi', id, JSON.stringify(data));
  return id;
}

function updatePenjualanTransaksi(id, updates) {
  requireRole('Supervisor');
  if (updates.HargaPerGram !== undefined || updates.Qty !== undefined) {
    const current = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL).find(r => r.ID === id);
    const merged = Object.assign({}, current, updates);
    const totalHarga = Number(merged.HargaPerGram || 0) * Number(merged.Qty || 0);
    updates.TotalHarga = totalHarga;
    updates.PphPasal22 = totalHarga * PPH_PASAL_22_RATE;
    updates.TotalHargaAfterTax = totalHarga + updates.PphPasal22;
  }
  updateRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', id, updates);
  logAudit('UPDATE', 'PenjualanTransaksi', id, JSON.stringify(updates));
}

function deletePenjualanTransaksi(id) {
  requireRole('Admin');
  deleteRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', id);
  logAudit('DELETE', 'PenjualanTransaksi', id, '');
}

function getPenjualanTransaksiToday(dateStr) {
  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  return sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => extractDateOnly(p.Tanggal) === target)
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
}

function getPenjualanTransaksiRecent(days) {
  days = days || 30;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  return sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => parseDateValue(p.Tanggal) >= cutoff.getTime())
    .sort((a, b) => parseDateValue(b.Tanggal) - parseDateValue(a.Tanggal));
}

// ------------------------------------------------------------
// CRUD — PEMBELIAN TRANSAKSI (struktur lengkap sesuai rekap Excel)
// ------------------------------------------------------------
const PAJAK_WAPU_RATE = 0.0025; // 0.25%

function getPembelianTransaksi() {
  return sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL);
}

/**
 * data = { Tanggal, NomorPO, Seller, Qty, Harga, AfterDiskon,
 *          InvoiceRef, Catatan }
 * TotalHarga = AfterDiskon x Qty
 * PajakWapu = TotalHarga x 0.25%
 * TotalHargaIncludePajak = TotalHarga - PajakWapu
 */
function addPembelianTransaksi(data) {
  requireRole('Viewer');
  const sh = getSheet(SHEET_NAMES.PEMBELIAN_FULL);
  const user = getCurrentUser();
  const id = getNextId(SHEET_NAMES.PEMBELIAN_FULL, 'ID', 'PBL');

  const afterDiskon = Number(data.AfterDiskon || data.Harga || 0);
  const totalHarga = afterDiskon * Number(data.Qty || 0);
  const pajakWapu = totalHarga * PAJAK_WAPU_RATE;
  const totalInclude = totalHarga - pajakWapu;

  sh.appendRow([
    id, data.Tanggal, data.NomorPO, data.Seller, data.Qty,
    data.Harga, afterDiskon, totalHarga, pajakWapu, totalInclude,
    data.InvoiceRef || '', data.Catatan || '', new Date(), user.Email
  ]);
  logAudit('CREATE', 'PembelianTransaksi', id, JSON.stringify(data));
  return id;
}

function updatePembelianTransaksi(id, updates) {
  requireRole('Supervisor');
  if (updates.Harga !== undefined || updates.AfterDiskon !== undefined || updates.Qty !== undefined) {
    const current = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL).find(r => r.ID === id);
    const merged = Object.assign({}, current, updates);
    const afterDiskon = Number(merged.AfterDiskon || merged.Harga || 0);
    const totalHarga = afterDiskon * Number(merged.Qty || 0);
    updates.TotalHarga = totalHarga;
    updates.PajakWapu = totalHarga * PAJAK_WAPU_RATE;
    updates.TotalHargaIncludePajak = totalHarga - updates.PajakWapu;
  }
  updateRowById(SHEET_NAMES.PEMBELIAN_FULL, 'ID', id, updates);
  logAudit('UPDATE', 'PembelianTransaksi', id, JSON.stringify(updates));
}

function deletePembelianTransaksi(id) {
  requireRole('Admin');
  deleteRowById(SHEET_NAMES.PEMBELIAN_FULL, 'ID', id);
  logAudit('DELETE', 'PembelianTransaksi', id, '');
}

function getPembelianTransaksiToday(dateStr) {
  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  return sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL)
    .filter(p => extractDateOnly(p.Tanggal) === target)
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt));
}

function getPembelianTransaksiRecent(days) {
  days = days || 30;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  return sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL)
    .filter(p => parseDateValue(p.Tanggal) >= cutoff.getTime())
    .sort((a, b) => parseDateValue(b.Tanggal) - parseDateValue(a.Tanggal));
}

// ------------------------------------------------------------
// RINGKASAN — dipakai Dashboard/Analytics (sumber data dialihkan
// ke PenjualanTransaksi & PembelianTransaksi karena lebih lengkap)
// ------------------------------------------------------------
function calculatePenjualanSummaryToday(dateStr) {
  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  const list = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => extractDateOnly(p.Tanggal) === target);
  return {
    totalNominal: list.reduce((sum, p) => sum + Number(p.TotalHarga || 0), 0),
    totalGram: list.reduce((sum, p) => sum + Number(p.Qty || 0), 0),
    jumlahInvoice: list.length
  };
}

/**
 * Mencari baris header sesungguhnya dalam N baris pertama —
 * diperlukan karena beberapa spreadsheet punya kotak ringkasan
 * di atas tabel data (baris 1 bukan header asli). Baris dengan
 * kecocokan kata kunci terbanyak (minimal 2) dianggap header.
 */
function findHeaderRowIndex(values, keywords, maxRowsToScan) {
  maxRowsToScan = Math.min(maxRowsToScan || 20, values.length);
  let bestRow = -1, bestScore = 0;

  for (let r = 0; r < maxRowsToScan; r++) {
    const rowText = values[r].map(c => String(c).trim().toLowerCase()).join(' | ');
    let score = 0;
    keywords.forEach(k => { if (rowText.indexOf(k) !== -1) score++; });
    if (score > bestScore) { bestScore = score; bestRow = r; }
  }
  return bestScore >= 2 ? bestRow : -1;
}

// ------------------------------------------------------------
// IMPORT DARI SPREADSHEET LAMA
// ------------------------------------------------------------
/**
 * Membangun peta {nama_header_ternormalisasi: indeks_kolom} dari
 * baris header, supaya pencarian kolom tidak bergantung posisi
 * tetap (aman meski ada kolom tersembunyi / urutan berbeda).
 */
function buildHeaderMap(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => {
    const norm = String(h).trim().toLowerCase().replace(/\s+/g, ' ');
    if (norm) map[norm] = i;
  });
  return map;
}

/**
 * Mencari indeks kolom berdasarkan kata kunci (semua kata kunci
 * harus muncul di header, dan header TIDAK boleh mengandung kata
 * yang ada di excludeWords). Mengembalikan -1 kalau tidak ketemu.
 */
function findColByKeywords(headerMap, includeWords, excludeWords) {
  excludeWords = excludeWords || [];
  const headers = Object.keys(headerMap);
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    const hasAllInclude = includeWords.every(w => h.indexOf(w) !== -1);
    const hasExclude = excludeWords.some(w => h.indexOf(w) !== -1);
    if (hasAllInclude && !hasExclude) return headerMap[h];
  }
  return -1;
}

function getByCol(row, colIdx) {
  return colIdx === -1 ? '' : row[colIdx];
}

/**
 * Import histori Penjualan dari spreadsheet lama.
 * Jalankan dari editor: importPenjualanLama('ID_SPREADSHEET_LAMA', 'Nama Sheet/Tab').
 * Kolom sumber dicari otomatis berdasarkan nama header (Tanggal,
 * Nomor Invoice, Harga per gram, Qty, Pembeli, Denominasi, Merk,
 * Status, Payment, Bank, Tanggal Pengambilan, Estimasi Penjemputan).
 * Total Harga/PPh/Total After Tax dihitung ulang otomatis dari
 * Harga per Gram x Qty untuk konsistensi dengan data baru.
 */
function importPenjualanLama(sourceSpreadsheetId, sheetName) {
  requireRole('Admin');
  const sourceSs = SpreadsheetApp.openById(sourceSpreadsheetId);
  const sourceSheet = sourceSs.getSheetByName(sheetName);
  if (!sourceSheet) throw new Error('Sheet "' + sheetName + '" tidak ditemukan di spreadsheet sumber.');

  const values = sourceSheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['tanggal', 'invoice', 'pembeli', 'qty', 'harga']);
  if (headerRowIdx === -1) throw new Error('Baris header tidak terdeteksi otomatis di sheet "' + sheetName + '". Cek Logger.log untuk 20 baris pertama, mungkin perlu penyesuaian manual.');
  Logger.log('Baris header terdeteksi di baris ke-' + (headerRowIdx + 1));

  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const rows = values.slice(headerRowIdx + 1);

  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal'], ['pengambilan', 'penjemputan']),
    nomorInvoice: findColByKeywords(headerMap, ['nomor', 'invoice']),
    hargaPerGram: findColByKeywords(headerMap, ['harga'], ['total', 'after', 'simulasi']),
    qty: findColByKeywords(headerMap, ['qty']),
    pembeli: findColByKeywords(headerMap, ['pembeli']),
    denominasi: findColByKeywords(headerMap, ['denominasi']),
    merk: findColByKeywords(headerMap, ['merk']),
    status: findColByKeywords(headerMap, ['status']),
    paymentTime: findColByKeywords(headerMap, ['payment']),
    bank: findColByKeywords(headerMap, ['bank']),
    tglAmbil: findColByKeywords(headerMap, ['pengambilan']),
    estimasi: findColByKeywords(headerMap, ['estimasi'])
  };

  Logger.log('Pemetaan kolom Penjualan: ' + JSON.stringify(col));
  if (col.tanggal === -1 || col.nomorInvoice === -1 || col.hargaPerGram === -1 || col.qty === -1) {
    throw new Error('Kolom wajib (Tanggal/Nomor Invoice/Harga/Qty) tidak ditemukan. Cek Logger.log untuk pemetaan kolom yang terdeteksi, lalu sesuaikan manual jika perlu.');
  }

  const targetSh = getSheet(SHEET_NAMES.PENJUALAN_FULL);
  const user = getCurrentUser();
  let imported = 0, skipped = 0;

  // Hitung ID awal SEKALI saja (bukan per baris) — mencegah timeout.
  const firstId = getNextId(SHEET_NAMES.PENJUALAN_FULL, 'ID', 'PJL');
  let idNum = parseInt(firstId.split('-')[1], 10);
  const outputRows = [];
  const now = new Date();

  rows.forEach(row => {
    const tanggal = getByCol(row, col.tanggal);
    const nomorInvoice = getByCol(row, col.nomorInvoice);
    if (!tanggal || !nomorInvoice) { skipped++; return; }

    const hargaPerGram = Number(getByCol(row, col.hargaPerGram)) || 0;
    const qty = Number(getByCol(row, col.qty)) || 0;
    const id = 'PJL-' + String(idNum).padStart(3, '0');
    idNum++;
    const totalHarga = hargaPerGram * qty;
    const pph = totalHarga * PPH_PASAL_22_RATE;
    const totalAfterTax = totalHarga + pph;

    outputRows.push([
      id, tanggal, nomorInvoice, hargaPerGram, qty,
      totalHarga, pph, totalAfterTax,
      getByCol(row, col.pembeli), getByCol(row, col.denominasi), getByCol(row, col.merk),
      getByCol(row, col.status) || 'V',
      getByCol(row, col.paymentTime), getByCol(row, col.bank),
      getByCol(row, col.tglAmbil), getByCol(row, col.estimasi),
      'Diimpor dari spreadsheet lama', now, user.Email
    ]);
    imported++;
  });

  // Tulis semua baris SEKALIGUS (1 operasi), bukan appendRow per baris.
  if (outputRows.length) {
    const startRow = targetSh.getLastRow() + 1;
    targetSh.getRange(startRow, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }

  logAudit('IMPORT', 'PenjualanTransaksi', '-', imported + ' baris diimpor dari ' + sheetName + ', ' + skipped + ' baris dilewati');
  Logger.log(imported + ' baris berhasil diimpor, ' + skipped + ' baris dilewati (kosong/header/subtotal).');
  return { imported: imported, skipped: skipped, columnMapping: col };
}

/**
 * Import histori Pembelian dari spreadsheet lama.
 * Jalankan dari editor: importPembelianLama('ID_SPREADSHEET_LAMA', 'Nama Sheet/Tab').
 * Kolom sumber dicari otomatis berdasarkan nama header (Tanggal,
 * Nomor Purchase Order, Seller, Qty, Harga, After Diskon, Invoice).
 * Total Harga/Pajak WAPU/Total Include Pajak dihitung ulang otomatis
 * dari After Diskon x Qty untuk konsistensi dengan data baru.
 */
function importPembelianLama(sourceSpreadsheetId, sheetName) {
  requireRole('Admin');
  const sourceSs = SpreadsheetApp.openById(sourceSpreadsheetId);
  const sourceSheet = sourceSs.getSheetByName(sheetName);
  if (!sourceSheet) throw new Error('Sheet "' + sheetName + '" tidak ditemukan di spreadsheet sumber.');

  const values = sourceSheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['tanggal', 'seller', 'qty', 'harga', 'invoice']);
  if (headerRowIdx === -1) throw new Error('Baris header tidak terdeteksi otomatis di sheet "' + sheetName + '". Cek Logger.log untuk 20 baris pertama, mungkin perlu penyesuaian manual.');
  Logger.log('Baris header terdeteksi di baris ke-' + (headerRowIdx + 1));

  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const rows = values.slice(headerRowIdx + 1);

  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    nomorPO: findColByKeywords(headerMap, ['purchase', 'order']),
    seller: findColByKeywords(headerMap, ['seller']),
    qty: findColByKeywords(headerMap, ['qty']),
    harga: findColByKeywords(headerMap, ['harga'], ['diskon', 'total', 'simulasi', 'rilis']),
    afterDiskon: findColByKeywords(headerMap, ['diskon']),
    invoiceRef: findColByKeywords(headerMap, ['invoice'])
  };

  Logger.log('Pemetaan kolom Pembelian: ' + JSON.stringify(col));
  if (col.tanggal === -1 || col.nomorPO === -1 || col.seller === -1 || col.qty === -1) {
    throw new Error('Kolom wajib (Tanggal/Nomor PO/Seller/Qty) tidak ditemukan. Cek Logger.log untuk pemetaan kolom yang terdeteksi, lalu sesuaikan manual jika perlu.');
  }

  const targetSh = getSheet(SHEET_NAMES.PEMBELIAN_FULL);
  const user = getCurrentUser();
  let imported = 0, skipped = 0;

  const firstId = getNextId(SHEET_NAMES.PEMBELIAN_FULL, 'ID', 'PBL');
  let idNum = parseInt(firstId.split('-')[1], 10);
  const outputRows = [];
  const now = new Date();

  rows.forEach(row => {
    const tanggal = getByCol(row, col.tanggal);
    const nomorPO = getByCol(row, col.nomorPO);
    if (!tanggal || !nomorPO) { skipped++; return; }

    const harga = Number(getByCol(row, col.harga)) || 0;
    const afterDiskonRaw = col.afterDiskon !== -1 ? Number(getByCol(row, col.afterDiskon)) : 0;
    const afterDiskon = afterDiskonRaw || harga;
    const qty = Number(getByCol(row, col.qty)) || 0;

    const id = 'PBL-' + String(idNum).padStart(3, '0');
    idNum++;
    const totalHarga = afterDiskon * qty;
    const pajakWapu = totalHarga * PAJAK_WAPU_RATE;
    const totalInclude = totalHarga - pajakWapu;

    outputRows.push([
      id, tanggal, nomorPO, getByCol(row, col.seller), qty,
      harga, afterDiskon, totalHarga, pajakWapu, totalInclude,
      getByCol(row, col.invoiceRef), 'Diimpor dari spreadsheet lama', now, user.Email
    ]);
    imported++;
  });

  if (outputRows.length) {
    const startRow = targetSh.getLastRow() + 1;
    targetSh.getRange(startRow, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }

  logAudit('IMPORT', 'PembelianTransaksi', '-', imported + ' baris diimpor dari ' + sheetName + ', ' + skipped + ' baris dilewati');
  Logger.log(imported + ' baris berhasil diimpor, ' + skipped + ' baris dilewati (kosong/header/subtotal).');
  return { imported: imported, skipped: skipped, columnMapping: col };
}

/**
 * Menghapus HANYA baris yang ditandai "Diimpor dari spreadsheet lama"
 * di PenjualanTransaksi & PembelianTransaksi — dipakai sebelum
 * re-import supaya tidak dobel. Data yang diinput manual lewat form
 * (Catatan-nya bukan teks itu persis) TIDAK akan terhapus.
 * Jalankan ini SEBELUM runImportSekali kalau sebelumnya sempat
 * gagal di tengah jalan (timeout dll).
 */
function clearImportedRowsOnly() {
  [SHEET_NAMES.PENJUALAN_FULL, SHEET_NAMES.PEMBELIAN_FULL].forEach(sheetName => {
    const sh = getSheet(sheetName);
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return;
    const headers = values[0];
    const catatanIdx = headers.indexOf('Catatan');
    if (catatanIdx === -1) return;

    let deleted = 0;
    // Hapus dari bawah ke atas supaya nomor baris tidak berantakan.
    for (let r = values.length - 1; r >= 1; r--) {
      if (values[r][catatanIdx] === 'Diimpor dari spreadsheet lama') {
        sh.deleteRow(r + 1);
        deleted++;
      }
    }
    Logger.log(sheetName + ': ' + deleted + ' baris hasil import lama dihapus.');
  });
}

// ------------------------------------------------------------
// LAPORAN HARIAN — merangkum semua data hari itu jadi satu laporan
// ------------------------------------------------------------

/**
 * Semua entri BenchmarkPrice hari ini, urut kronologis (untuk
 * kolom Harga I/II/III/IV di laporan — tiap submit = 1 kolom).
 */
function getBenchmarkEntriesToday(dateStr) {
  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  return sheetToObjects(SHEET_NAMES.BENCHMARK)
    .filter(b => extractDateOnly(b.Tanggal) === target)
    .sort((a, b) => new Date(a.CreatedAt) - new Date(b.CreatedAt));
}

/**
 * Semua entri PriceSetting hari ini, urut kronologis.
 */
function getPriceSettingEntriesToday(dateStr) {
  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  return sheetToObjects(SHEET_NAMES.PRICE_SETTING)
    .filter(p => extractDateOnly(p.Tanggal) === target)
    .sort((a, b) => new Date(a.CreatedAt) - new Date(b.CreatedAt));
}

/**
 * Ringkasan Piutang per supplier — gramasi & valuasi beserta
 * persentase masing-masing terhadap total (untuk tabel Piutang).
 */
function getPiutangSummary() {
  const list = getOutstandingActive();
  const grouped = {};
  list.forEach(o => {
    const key = o.SupplierNama;
    if (!grouped[key]) grouped[key] = { supplierNama: key, gramasi: 0, valuasi: 0 };
    grouped[key].gramasi += Number(o.Gramasi || 0);
    grouped[key].valuasi += Number(o.Valuasi || 0);
  });

  const rows = Object.values(grouped);
  const totalGramasi = rows.reduce((s, r) => s + r.gramasi, 0);
  const totalValuasi = rows.reduce((s, r) => s + r.valuasi, 0);

  rows.forEach(r => {
    r.gramasiPercent = totalGramasi ? (r.gramasi / totalGramasi) * 100 : 0;
    r.valuasiPercent = totalValuasi ? (r.valuasi / totalValuasi) * 100 : 0;
  });
  rows.sort((a, b) => b.valuasi - a.valuasi);

  return { rows: rows, totalGramasi: totalGramasi, totalValuasi: totalValuasi };
}

/**
 * Total Pembelian hari ini dari PembelianTransaksi (TotalHarga,
 * bukan TotalHargaIncludePajak — konsisten dengan angka "Nominal"
 * di tabel Pembelian pada laporan Excel Anda).
 */
function calculatePembelianSummaryToday(dateStr) {
  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  const list = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL)
    .filter(p => extractDateOnly(p.Tanggal) === target);
  return {
    totalNominal: list.reduce((sum, p) => sum + Number(p.TotalHarga || 0), 0),
    totalGram: list.reduce((sum, p) => sum + Number(p.Qty || 0), 0),
    jumlahTransaksi: list.length
  };
}

// ------------------------------------------------------------
// LIVE READ — Sumber Data & Keputusan Harga dari sheet
// "Price Discovery V3" (spreadsheet kerja Anda yang lain).
// Dibaca langsung tiap kali Laporan Harian dibuka, TIDAK disimpan
// ke database kita — jadi selalu sinkron dengan yang Anda update
// di spreadsheet aslinya.
// ------------------------------------------------------------
const PRICE_DISCOVERY_SOURCE_ID = '13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o';
const PRICE_DISCOVERY_SHEET_NAME = 'Price Discovery V3';

// Sheet "Rekap" — sumber LIVE untuk Piutang (Gold in Transit per supplier).
// Baris & kolom dikonfirmasi manual dari screenshot: baris 8 = nama supplier,
// baris 10 = "Gold in Transit (gr)" (gramasi piutang), baris 13 = "Uang Muka (Rp)"
// (valuasi piutang). Kolom value tiap supplier berselang-seling: R,T,V,X,Z,AB,AD,AF.
const REKAP_SHEET_NAME = 'Rekap';
const REKAP_PIUTANG_ROW = { header: 8, goldInTransit: 10, uangMuka: 13 };
const REKAP_PIUTANG_SUPPLIER_COLS = [
  { labelCol: 18, valueCol: 19 }, // HRTA (R/S)
  { labelCol: 20, valueCol: 21 }, // IGS (T/U)
  { labelCol: 22, valueCol: 23 }, // SJL (V/W)
  { labelCol: 24, valueCol: 25 }, // LOTUS (X/Y)
  { labelCol: 26, valueCol: 27 }, // MKB (Z/AA)
  { labelCol: 28, valueCol: 29 }, // IDN (AB/AC)
  { labelCol: 30, valueCol: 31 }, // Cahaya Baru (AD/AE)
  { labelCol: 32, valueCol: 33 }  // ANTAM (AF/AG)
];

/**
 * Piutang LIVE — dibaca langsung dari sheet "Rekap" (kolom Gold in Transit
 * & Uang Muka per supplier), bukan dari sheet Outstanding database kita.
 * Dipakai khusus di Laporan Harian karena data Outstanding lama sudah
 * tidak akurat (hasil import lama). Struktur objek hasilnya SAMA PERSIS
 * dengan getPiutangSummary() supaya tidak perlu ubah apapun di frontend.
 */
function getPiutangLive() {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = ss.getSheetByName(REKAP_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + REKAP_SHEET_NAME + '" tidak ditemukan.');

  const rows = REKAP_PIUTANG_SUPPLIER_COLS.map(function (colSet) {
    const namaSupplier = sheet.getRange(REKAP_PIUTANG_ROW.header, colSet.labelCol).getValue();
    const gramasi = Number(sheet.getRange(REKAP_PIUTANG_ROW.goldInTransit, colSet.valueCol).getValue()) || 0;
    const valuasi = Number(sheet.getRange(REKAP_PIUTANG_ROW.uangMuka, colSet.valueCol).getValue()) || 0;
    return { supplierNama: String(namaSupplier || '').trim(), gramasi: gramasi, valuasi: valuasi };
  }).filter(function (r) { return r.gramasi > 0 || r.valuasi > 0; });

  const totalGramasi = rows.reduce(function (s, r) { return s + r.gramasi; }, 0);
  const totalValuasi = rows.reduce(function (s, r) { return s + r.valuasi; }, 0);

  rows.forEach(function (r) {
    r.gramasiPercent = totalGramasi ? (r.gramasi / totalGramasi) * 100 : 0;
    r.valuasiPercent = totalValuasi ? (r.valuasi / totalValuasi) * 100 : 0;
  });
  rows.sort(function (a, b) { return b.valuasi - a.valuasi; });

  return { rows: rows, totalGramasi: totalGramasi, totalValuasi: totalValuasi };
}

// Indeks kolom (0-based) di sheet Price Discovery V3 — dikonfirmasi
// manual dari screenshot: B=Tanggal, C=Harga Ke-, E-L=sumber harga,
// M-P=benchmark, Y=Buyback, Z=Keputusan Harga.
const PD_COL = {
  tanggal: 1, hargaKe: 2, pukul: 3,
  hrta: 4, hrtaDisc: 5, igs: 6, sjl: 7, idn: 8, mkb: 9, lotus: 10, apepi: 11,
  spotGoldprice: 12, xauidrk: 13, lbma: 14, kitco: 15,
  kursJual: 16, kursBeli: 17, rpLbma: 18, rpKitco: 19,
  keputusanHarga: 25, buyback: 24
};

/**
 * Mengambil sampai 4 baris (Harga Ke- 1-4) dari Price Discovery V3
 * untuk tanggal tertentu, terurut berdasarkan Harga Ke-.
 */
function getPriceDiscoveryRowsForDate(dateStr) {
  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = ss.getSheetByName(PRICE_DISCOVERY_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + PRICE_DISCOVERY_SHEET_NAME + '" tidak ditemukan.');

  const values = sheet.getDataRange().getValues();
  const matching = values.filter(row => extractDateOnly(row[PD_COL.tanggal]) === target);
  matching.sort((a, b) => (Number(a[PD_COL.hargaKe]) || 0) - (Number(b[PD_COL.hargaKe]) || 0));
  return matching;
}

/**
 * Referensi 8 indikator market (Spot Goldprice, XAUIDRK, LBMA H-1,
 * Kitco H-1, Kurs Jual, Kurs Beli, Rp LBMA, Rp Kitco) — LIVE dari
 * kolom M-T sheet Price Discovery V3, untuk tanggal tertentu
 * (default hari ini). Dipakai sebagai panel referensi di halaman
 * Penetapan Harga (bobot) — TIDAK ikut dihitung dalam rata-rata
 * tertimbang, cuma info pembanding harga pasar saat itu.
 * Reuse getPriceDiscoveryRowsForDate() yang sudah baca sheet
 * sekali & filter per tanggal (efisien, bukan scan seluruh histori).
 */
function getMarketBenchmarkLive(dateStr) {
  const rows = getPriceDiscoveryRowsForDate(dateStr);
  if (!rows.length) return null;

  function lastNonZero(col) {
    for (let i = rows.length - 1; i >= 0; i--) {
      const v = Number(rows[i][col]) || 0;
      if (v) return v;
    }
    return null;
  }

  return {
    spotGold: lastNonZero(PD_COL.spotGoldprice),
    xauIdr: lastNonZero(PD_COL.xauidrk),
    lbma: lastNonZero(PD_COL.lbma),
    kitco: lastNonZero(PD_COL.kitco),
    kursJual: lastNonZero(PD_COL.kursJual),
    kursBeli: lastNonZero(PD_COL.kursBeli),
    rpLbma: lastNonZero(PD_COL.rpLbma),
    rpKitco: lastNonZero(PD_COL.rpKitco)
  };
}

/**
 * Sumber Data (HRTA, HRTA+disc, IGS, SJL, IDN, MKB/STRGLD, LOTUS,
 * APEPI, Spot Goldprice, XAUIDRK, LBMA H-1, Kitco H-1) dengan
 * Harga I-IV, LIVE dari Price Discovery V3.
 */
function getSumberDataLive(dateStr) {
  const rows = getPriceDiscoveryRowsForDate(dateStr);
  if (!rows.length) return [];

  const sources = [
    { nama: 'HRTA', col: PD_COL.hrtaDisc },
    { nama: 'IGS', col: PD_COL.igs },
    { nama: 'SJL', col: PD_COL.sjl },
    { nama: 'IDN', col: PD_COL.idn },
    { nama: 'MKB/STRGLD', col: PD_COL.mkb },
    { nama: 'LOTUS', col: PD_COL.lotus },
    { nama: 'APEPI', col: PD_COL.apepi },
    { nama: 'Spot Goldprice', col: PD_COL.spotGoldprice },
    { nama: 'XAUIDRK', col: PD_COL.xauidrk },
    { nama: 'LBMA H-1', col: PD_COL.lbma },
    { nama: 'Kitco H-1', col: PD_COL.kitco }
  ];

  return sources.map(s => {
    const vals = [0, 1, 2, 3].map(i => (rows[i] ? Number(rows[i][s.col]) || 0 : 0));
    return { nama: s.nama, harga1: vals[0], harga2: vals[1], harga3: vals[2], harga4: vals[3] };
  }).filter(r => r.harga1 || r.harga2 || r.harga3 || r.harga4);
}

/**
 * Keputusan Harga (Sale) & Buyback per Harga Ke-, LIVE dari
 * Price Discovery V3.
 */
function getKeputusanHargaLive(dateStr) {
  const rows = getPriceDiscoveryRowsForDate(dateStr);
  const saleVals = rows.map(r => Number(r[PD_COL.keputusanHarga]) || 0).filter(v => v > 0);
  const buybackVals = rows.map(r => Number(r[PD_COL.buyback]) || 0).filter(v => v > 0);
  return { saleVals: saleVals, buybackVals: buybackVals };
}

/**
 * Data lengkap Laporan Harian — dipanggil frontend untuk
 * merender halaman & teks siap-share.
 */
/**
 * Cari supplier dengan harga terendah dari sumberData live
 * (mengecualikan baris benchmark seperti Spot Goldprice, dst).
 * Dipakai kartu "Supplier Harga Terendah" & teks ringkasan.
 */
function getCheapestFromSumberData(sumberData) {
  const benchmarkNames = ['Spot Goldprice', 'XAUIDRK', 'LBMA H-1', 'Kitco H-1'];
  const supplierOnly = sumberData
    .filter(r => benchmarkNames.indexOf(r.nama) === -1)
    .map(r => ({ nama: r.nama, harga: [r.harga1, r.harga2, r.harga3, r.harga4].filter(v => v > 0) }))
    .filter(r => r.harga.length)
    .map(r => ({ nama: r.nama, harga: r.harga[r.harga.length - 1] })); // harga terakhir hari itu
  return supplierOnly.length
    ? supplierOnly.reduce((min, r) => (r.harga < min.harga ? r : min))
    : null;
}

function getLaporanHarianData(dateStr) {
  const targetDate = dateStr ? new Date(extractDateOnly(dateStr) + 'T00:00:00') : new Date();
  const tanggalLabel = targetDate.getDate() + ' ' + NAMA_BULAN_ID[targetDate.getMonth()] + ' ' + targetDate.getFullYear();

  // --- Sumber Data: LIVE dari Price Discovery V3 ---
  const sumberData = getSumberDataLive(dateStr);
  const cheapestSupplier = getCheapestFromSumberData(sumberData);

  // --- Piutang (LIVE dari sheet Rekap, bukan dari sheet Outstanding) ---
  const piutang = getPiutangLive();

  // --- Penetapan Harga: LIVE dari Price Discovery V3 (Keputusan Harga & Buyback) ---
  const keputusan = getKeputusanHargaLive(dateStr);
  const marginStats = calculateMarginStats(dateStr);

  // --- Penjualan & Pembelian pada tanggal ini ---
  const penjualanRows = getPenjualanTransaksiToday(dateStr);
  const pembelianRows = getPembelianTransaksiToday(dateStr);
  const penjualanSummary = calculatePenjualanSummaryToday(dateStr);
  const pembelianSummary = calculatePembelianSummaryToday(dateStr);

  const profitMargin = (penjualanSummary.totalNominal || 0) - (pembelianSummary.totalNominal || 0);
  const marginPercent = pembelianSummary.totalNominal ? (profitMargin / pembelianSummary.totalNominal) * 100 : 0;

  return {
    tanggalLabel: tanggalLabel,
    sumberData: sumberData,
    cheapestSupplier: cheapestSupplier,
    piutang: piutang,
    penetapanHarga: {
      saleVals: keputusan.saleVals,
      buybackVals: keputusan.buybackVals,
      marginRataRata: marginStats.average,
      marginPercent: marginPercent,
      profitMargin: profitMargin
    },
    penjualan: { rows: penjualanRows, summary: penjualanSummary },
    pembelian: { rows: pembelianRows, summary: pembelianSummary }
  };
}

/**
 * Versi teks siap copy-paste (format seperti update WhatsApp/email
 * harian) — rule-based dari data yang sama dengan getLaporanHarianData().
 */
function generateLaporanHarianText(dateStr) {
  const data = getLaporanHarianData(dateStr);
  const cheapest = data.cheapestSupplier;
  const lines = [];

  lines.push('Update Transaksi Operasional Perdagangan Emas – ' + data.tanggalLabel);
  lines.push('');
  lines.push('Pemasok dan Harga Supplier');
  if (cheapest) {
    lines.push('- ' + cheapest.nama + ' masih menjadi supplier dengan harga paling kompetitif di kisaran ' +
      formatRupiahServer(cheapest.harga) + '/gram, lebih rendah dibandingkan supplier lain maupun benchmark pasar.');
  } else {
    lines.push('- Belum ada data harga supplier pada tanggal ini.');
  }
  lines.push('');

  lines.push('Realisasi Transaksi Harian');
  lines.push('Total transaksi penjualan mencapai ' + (data.penjualan.summary.totalGram / 1000).toLocaleString('id-ID', { maximumFractionDigits: 2 }) +
    ' kg dengan nilai ' + formatRupiahServer(data.penjualan.summary.totalNominal) + '. Rincian penjualan:');
  data.penjualan.rows.forEach(p => {
    lines.push('- ' + p.Pembeli + ': ' + (Number(p.Qty) / 1000).toLocaleString('id-ID', { maximumFractionDigits: 2 }) + ' kg');
  });
  if (data.penjualan.summary.jumlahInvoice) {
    const avgHarga = data.penjualan.summary.totalGram ? data.penjualan.summary.totalNominal / data.penjualan.summary.totalGram : 0;
    lines.push('Harga rata-rata rilis sebesar ' + formatRupiahServer(avgHarga) + '/gram.');
  }
  lines.push('');

  lines.push('Kinerja Profitabilitas');
  lines.push('- Selisih harga beli dan harga jual menghasilkan margin rata-rata sebesar ' + data.penetapanHarga.marginPercent.toFixed(2) + '%.');
  lines.push('- Profit perdagangan tercatat sebesar ' + formatRupiahServer(data.penetapanHarga.profitMargin) + ' dari hasil selisih Pembelian dan Penjualan pada tanggal ini.');
  lines.push('');

  lines.push('Pengelolaan Risiko');
  lines.push('- Total piutang supplier mencapai ' + data.piutang.totalGramasi.toLocaleString('id-ID', { maximumFractionDigits: 0 }) + ' gram senilai ' + formatRupiahServer(data.piutang.totalValuasi) + '.');
  if (data.piutang.rows.length) {
    const top = data.piutang.rows[0];
    lines.push('- Sebesar ' + top.valuasiPercent.toFixed(2) + '% piutang masih terkonsentrasi pada ' + top.supplierNama + '.');
  }

  return lines.join('\n');
}

/**
 * Import histori Piutang dari sheet "Input Barang Masuk".
 * Jalankan dari editor: importPiutangLama('ID_SPREADSHEET_LAMA', 'Nama Sheet/Tab').
 * Kolom dipakai berdasarkan POSISI TETAP (bukan nama header),
 * karena sheet ini punya beberapa kolom dengan nama sama persis
 * ("Harga", "Sisa") di tabel berbeda — sudah dikonfirmasi manual
 * dari screenshot + rumus SUMIF asli: Tanggal=T, Seller=W, Qty=X,
 * Harga=Y, Sisa=AE (indeks 0-based: 19,22,23,24,30).
 */
function importPiutangLama(sourceSpreadsheetId, sheetName) {
  requireRole('Admin');
  const sourceSs = SpreadsheetApp.openById(sourceSpreadsheetId);
  const sourceSheet = sourceSs.getSheetByName(sheetName);
  if (!sourceSheet) throw new Error('Sheet "' + sheetName + '" tidak ditemukan di spreadsheet sumber.');

  const values = sourceSheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['tanggal', 'seller', 'sisa', 'qty', 'harga']);
  if (headerRowIdx === -1) throw new Error('Baris header tidak terdeteksi otomatis di sheet "' + sheetName + '". Cek Logger.log untuk 20 baris pertama.');
  Logger.log('Baris header terdeteksi di baris ke-' + (headerRowIdx + 1));

  const rows = values.slice(headerRowIdx + 1);
  const COL = { tanggal: 19, nomorPO: 20, nomorInvoice: 21, seller: 22, qty: 23, harga: 24, sisa: 30 };

  const targetSh = getSheet(SHEET_NAMES.OUTSTANDING);
  const user = getCurrentUser();
  let imported = 0, skipped = 0;

  const firstId = getNextId(SHEET_NAMES.OUTSTANDING, 'ID', 'OUT');
  let idNum = parseInt(firstId.split('-')[1], 10);
  const outputRows = [];
  const now = new Date();

  rows.forEach(row => {
    const tanggal = row[COL.tanggal];
    const seller = row[COL.seller];
    if (!tanggal || !seller) { skipped++; return; }

    const sisa = Number(row[COL.sisa]) || 0;
    const harga = Number(row[COL.harga]) || 0;
    const qty = Number(row[COL.qty]) || 0;
    const status = sisa > 0 ? 'Outstanding' : 'Lunas';
    const gramasi = sisa > 0 ? sisa : qty;
    const valuasi = gramasi * harga;

    const id = 'OUT-' + String(idNum).padStart(3, '0');
    idNum++;

    outputRows.push([
      id, tanggal, '', seller, gramasi, harga, valuasi, status,
      'Diimpor dari Input Barang Masuk (PO ' + (row[COL.nomorPO] || '-') + ')',
      now, user.Email
    ]);
    imported++;
  });

  if (outputRows.length) {
    const startRow = targetSh.getLastRow() + 1;
    targetSh.getRange(startRow, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }

  logAudit('IMPORT', 'Outstanding', '-', imported + ' baris diimpor dari ' + sheetName + ', ' + skipped + ' baris dilewati');
  Logger.log(imported + ' baris berhasil diimpor, ' + skipped + ' baris dilewati.');
  return { imported: imported, skipped: skipped };
}

/**
 * DIAGNOSTIK — jalankan ini untuk melihat nama PERSIS semua
 * tab/sheet di spreadsheet sumber (kadang beda spasi/karakter
 * dari yang terlihat di tab, misalnya trailing space).
 */
function listSourceSheetNames() {
  const SOURCE_ID = '13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o';
  const ss = SpreadsheetApp.openById(SOURCE_ID);
  const names = ss.getSheets().map(function (sh) { return '"' + sh.getName() + '"'; });
  Logger.log('Total sheet: ' + names.length);
  Logger.log(names.join('\n'));
}

/**
 * RUNNER — jalankan sekali untuk import Piutang dari Input Barang Masuk.
 */
function runImportPiutangSekali() {
  const SOURCE_ID = '13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o';
  const hasil = importPiutangLama(SOURCE_ID, 'F.3.Input_Barang Masuk');
  Logger.log('=== HASIL IMPORT PIUTANG ===');
  Logger.log(JSON.stringify(hasil));
}

/**
 * Menghapus HANYA baris Piutang hasil import (Keterangan diawali
 * "Diimpor dari Input Barang Masuk") — aman dipakai sebelum re-import.
 */
function clearImportedPiutangOnly() {
  const sh = getSheet(SHEET_NAMES.OUTSTANDING);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0];
  const ketIdx = headers.indexOf('Keterangan');
  if (ketIdx === -1) return;

  let deleted = 0;
  for (let r = values.length - 1; r >= 1; r--) {
    if (String(values[r][ketIdx]).indexOf('Diimpor dari Input Barang Masuk') === 0) {
      sh.deleteRow(r + 1);
      deleted++;
    }
  }
  Logger.log(deleted + ' baris piutang hasil import dihapus.');
}

// ------------------------------------------------------------
// LAPORAN BULANAN
// ------------------------------------------------------------
function formatRupiahShortServer(num) {
  num = Number(num) || 0;
  const abs = Math.abs(num);
  if (abs >= 1e9) return (num / 1e9).toLocaleString('id-ID', { maximumFractionDigits: 1 }) + ' M';
  if (abs >= 1e6) return (num / 1e6).toLocaleString('id-ID', { maximumFractionDigits: 1 }) + ' jt';
  return formatRupiahServer(num);
}

const NAMA_BULAN_ID = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

function getMonthKey(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');
}

/**
 * Trend bulanan. "Realisasi" = kumulatif Year-to-Date (dari Januari
 * tahun berjalan s/d bulan tersebut, reset tiap awal tahun) — supaya
 * bisa dibandingkan apple-to-apple dengan Target tahunan. Penjualan/
 * Pembelian/gramasi/Margin tetap dihitung per-bulan (tidak kumulatif).
 */
function getMonthlyTrend(monthsCount) {
  monthsCount = monthsCount || 7;
  const penjualanList = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const pembelianList = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL);
  const today = new Date();
  const result = [];

  for (let i = monthsCount - 1; i >= 0; i--) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const key = getMonthKey(d);
    const pjList = penjualanList.filter(p => extractDateOnly(p.Tanggal).slice(0, 7) === key);
    const pbList = pembelianList.filter(p => extractDateOnly(p.Tanggal).slice(0, 7) === key);
    const pjNominal = pjList.reduce((s, p) => s + Number(p.TotalHarga || 0), 0);
    const pbNominal = pbList.reduce((s, p) => s + Number(p.TotalHarga || 0), 0);
    const pjGram = pjList.reduce((s, p) => s + Number(p.Qty || 0), 0);
    const pbGram = pbList.reduce((s, p) => s + Number(p.Qty || 0), 0);

    // Realisasi kumulatif Year-to-Date: jumlahkan semua transaksi dari
    // 1 Januari tahun yang sama s/d akhir bulan ini.
    const yearStartKey = key.slice(0, 4) + '-01';
    const pjYtd = penjualanList
      .filter(p => { const mk = extractDateOnly(p.Tanggal).slice(0, 7); return mk >= yearStartKey && mk <= key; })
      .reduce((s, p) => s + Number(p.TotalHarga || 0), 0);
    const pbYtd = pembelianList
      .filter(p => { const mk = extractDateOnly(p.Tanggal).slice(0, 7); return mk >= yearStartKey && mk <= key; })
      .reduce((s, p) => s + Number(p.TotalHarga || 0), 0);

    result.push({
      yearMonth: key,
      monthLabel: NAMA_BULAN_ID[d.getMonth()],
      monthLabelFull: NAMA_BULAN_ID[d.getMonth()] + ' ' + d.getFullYear(),
      penjualanNominal: pjNominal,
      pembelianNominal: pbNominal,
      penjualanGram: pjGram,
      pembelianGram: pbGram,
      realisasi: pjYtd + pbYtd,
      margin: pjNominal - pbNominal
    });
  }
  return result;
}

/**
 * Breakdown Pembelian per supplier (Seller) untuk satu bulan
 * tertentu — dipakai chart bar + donut halaman Tren Pembelian.
 */
function getSupplierBreakdownMonth(yearMonth) {
  const target = yearMonth || getMonthKey(new Date());
  const list = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL).filter(p => extractDateOnly(p.Tanggal).slice(0, 7) === target);
  const grouped = {};
  let totalGram = 0;
  list.forEach(p => {
    const key = p.Seller;
    grouped[key] = (grouped[key] || 0) + Number(p.Qty || 0);
    totalGram += Number(p.Qty || 0);
  });
  const rows = Object.keys(grouped)
    .map(k => ({ seller: k, gram: grouped[k] }))
    .sort((a, b) => b.gram - a.gram);
  rows.forEach(r => { r.percent = totalGram ? (r.gram / totalGram) * 100 : 0; });
  return { rows: rows, totalGram: totalGram };
}

/**
 * Ringkasan bulan terpilih + perbandingan MtM terhadap bulan
 * sebelumnya, plus histori 7 bulan untuk chart trend.
 */
function getMonthlySummary(yearMonth) {
  const target = yearMonth || getMonthKey(new Date());
  const trend = getMonthlyTrend(12);
  const idx = trend.findIndex(t => t.yearMonth === target);
  const current = idx !== -1 ? trend[idx] : trend[trend.length - 1];
  const prev = idx > 0 ? trend[idx - 1] : null;
  const mtmPercent = prev && prev.realisasi ? ((current.realisasi - prev.realisasi) / prev.realisasi) * 100 : null;
  const first = trend[Math.max(0, idx - 6)];
  const vsFirstPercent = first && first.realisasi ? ((current.realisasi - first.realisasi) / first.realisasi) * 100 : null;

  return {
    current: current,
    prev: prev,
    first: first,
    mtmPercent: mtmPercent,
    vsFirstPercent: vsFirstPercent,
    trend: trend.slice(Math.max(0, idx - 6), idx + 1)
  };
}

/**
 * Narasi otomatis Laporan Bulanan — rule-based, gaya penulisan
 * mengikuti pola laporan performa (MtM, tren pemulihan, dst).
 */
function generateMonthlyNarrative(yearMonth) {
  const summary = getMonthlySummary(yearMonth);
  const cur = summary.current;
  const prev = summary.prev;
  const first = summary.first;
  const lines = [];

  lines.push('Realisasi per ' + cur.monthLabelFull + ' tercatat sebesar ' + formatRupiahShortServer(cur.realisasi) + '.');

  if (prev && prev.realisasi) {
    const arah = cur.realisasi >= prev.realisasi ? 'meningkat' : 'menurun';
    const mtm = ((cur.realisasi - prev.realisasi) / prev.realisasi) * 100;
    lines.push('Posisi ini ' + arah + ' dibandingkan ' + prev.monthLabelFull + ' sebesar ' + formatRupiahShortServer(prev.realisasi) +
      ', atau tumbuh sekitar ' + mtm.toFixed(2) + '% secara Month-to-Month (MtM).');
  } else {
    lines.push('Belum ada data bulan sebelumnya untuk dibandingkan.');
  }

  if (first && first.yearMonth !== cur.yearMonth && first.realisasi) {
    const arah2 = cur.realisasi >= first.realisasi ? 'lebih tinggi' : 'lebih rendah';
    const pct2 = Math.abs(((cur.realisasi - first.realisasi) / first.realisasi) * 100);
    lines.push('Dibandingkan posisi ' + first.monthLabelFull + ' sebesar ' + formatRupiahShortServer(first.realisasi) +
      ', realisasi ' + cur.monthLabelFull + ' ' + arah2 + ' sekitar ' + pct2.toFixed(2) + '%.');
  }

  const marginArah = cur.margin >= 0 ? 'positif' : 'negatif';
  lines.push('Margin ' + cur.monthLabelFull + ' tercatat ' + marginArah + ' sebesar ' + formatRupiahShortServer(Math.abs(cur.margin)) +
    ' (Penjualan ' + formatRupiahShortServer(cur.penjualanNominal) + ' vs Pembelian ' + formatRupiahShortServer(cur.pembelianNominal) + ').');

  lines.push('');
  lines.push('Insight');

  const trendVals = summary.trend.map(t => t.realisasi);
  const isNaikTerus = trendVals.length > 2 && trendVals.every((v, i) => i === 0 || v >= trendVals[i - 1]);
  lines.push('Tren Pergerakan Realisasi');
  if (isNaikTerus) {
    lines.push('Realisasi menunjukkan tren peningkatan yang konsisten sepanjang periode ' + summary.trend[0].monthLabelFull + ' hingga ' + cur.monthLabelFull + '.');
  } else {
    lines.push('Realisasi berfluktuasi sepanjang periode ' + summary.trend[0].monthLabelFull + ' hingga ' + cur.monthLabelFull +
      ', dengan posisi ' + cur.monthLabelFull + ' sebesar ' + formatRupiahShortServer(cur.realisasi) + '.');
  }

  lines.push('');
  lines.push('Momentum Ke Depan');
  lines.push('Fokus utama ke depan adalah menjaga kesinambungan transaksi, mengoptimalkan volume dari supplier dan buyer eksisting, ' +
    'memperluas potensi mitra baru, serta mempercepat eksekusi transaksi untuk menjaga pertumbuhan yang telah tercapai.');

  return lines.join('\n');
}

/**
 * Narasi halaman 2 — fokus tren Penjualan (gram & nominal).
 */
function generatePenjualanNarrative(yearMonth) {
  const summary = getMonthlySummary(yearMonth);
  const cur = summary.current;
  const prev = summary.prev;
  const lines = [];

  lines.push('Pada ' + cur.monthLabelFull + ', aktivitas penjualan tercatat sebesar ' + cur.penjualanGram.toLocaleString('id-ID', { maximumFractionDigits: 0 }) + ' gram dengan nilai ' + formatRupiahShortServer(cur.penjualanNominal) + '.');

  if (prev && prev.penjualanGram) {
    const arahG = cur.penjualanGram >= prev.penjualanGram ? 'meningkat' : 'menurun';
    const mtmG = ((cur.penjualanGram - prev.penjualanGram) / prev.penjualanGram) * 100;
    lines.push('Secara volume, penjualan ' + arahG + ' dibandingkan ' + prev.monthLabelFull + ' sebesar ' + prev.penjualanGram.toLocaleString('id-ID', { maximumFractionDigits: 0 }) + ' gram, atau ' + Math.abs(mtmG).toFixed(2) + '% secara Month-to-Month (MtM).');
  }
  if (prev && prev.penjualanNominal) {
    const mtmN = ((cur.penjualanNominal - prev.penjualanNominal) / prev.penjualanNominal) * 100;
    lines.push('Secara nominal, penjualan ' + (mtmN >= 0 ? 'tumbuh' : 'turun') + ' sekitar ' + Math.abs(mtmN).toFixed(2) + '% MtM, dari ' + formatRupiahShortServer(prev.penjualanNominal) + ' menjadi ' + formatRupiahShortServer(cur.penjualanNominal) + '.');
  }

  return lines.join('\n');
}

/**
 * Narasi halaman 3 — fokus tren Pembelian & konsentrasi supplier.
 */
function generatePembelianNarrative(yearMonth) {
  const summary = getMonthlySummary(yearMonth);
  const cur = summary.current;
  const prev = summary.prev;
  const breakdown = getSupplierBreakdownMonth(yearMonth || getMonthKey(new Date()));
  const lines = [];

  lines.push('Pada ' + cur.monthLabelFull + ', transaksi pembelian bullion tercatat sebesar ' + cur.pembelianGram.toLocaleString('id-ID', { maximumFractionDigits: 0 }) + ' gram dengan nilai ' + formatRupiahShortServer(cur.pembelianNominal) + '.');

  if (prev && prev.pembelianGram) {
    const arahG = cur.pembelianGram >= prev.pembelianGram ? 'meningkat' : 'menurun';
    const mtmG = ((cur.pembelianGram - prev.pembelianGram) / prev.pembelianGram) * 100;
    lines.push('Volume pembelian ' + arahG + ' dibandingkan ' + prev.monthLabel + ' sebesar ' + Math.abs(mtmG).toFixed(2) + '% secara Month-to-Month (MtM).');
  }

  if (breakdown.rows.length) {
    const top = breakdown.rows[0];
    lines.push('Transaksi pembelian masih didominasi oleh ' + top.seller + ' dengan total pembelian sebesar ' +
      top.gram.toLocaleString('id-ID', { maximumFractionDigits: 0 }) + ' gram atau sekitar ' + top.percent.toFixed(1) + '% dari total pembelian bulan ini.');
    if (breakdown.rows.length > 1) {
      const second = breakdown.rows[1];
      lines.push('Supplier dengan kontribusi terbesar berikutnya adalah ' + second.seller + ' sebesar ' + second.percent.toFixed(1) + '%.');
    }
    lines.push('Konsentrasi pembelian pada satu supplier utama perlu terus dipantau, guna menjaga kesinambungan pasokan dan memitigasi risiko ketergantungan terhadap satu pemasok.');
  } else {
    lines.push('Belum ada data pembelian tercatat pada bulan ini.');
  }

  return lines.join('\n');
}

// ------------------------------------------------------------
// FUNGSI TES CEPAT — jalankan manual dari editor untuk cek
// koneksi & data sudah benar sebelum lanjut ke Fase 3.
// ------------------------------------------------------------
function testConnection() {
  Logger.log('User saat ini: ' + JSON.stringify(getCurrentUser()));
  Logger.log('Daftar Supplier: ' + JSON.stringify(getSuppliers()));
  Logger.log('Total Outstanding: ' + JSON.stringify(calculateTotalOutstanding()));
  Logger.log('Koneksi database berhasil.');
}

/**
 * RUNNER SEKALI-JALAN — jalankan fungsi ini (pilih di dropdown
 * toolbar > Run) untuk mengimpor histori Penjualan & Pembelian
 * dari spreadsheet lama. Aman dijalankan berkali-kali — baris
 * yang sudah masuk tidak akan dobel selama sheet sumbernya sama
 * (tapi tetap cek hasilnya di Logger setelah selesai).
 */
function runImportSekali() {
  const SOURCE_ID = '13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o';

  const hasilPenjualan = importPenjualanLama(SOURCE_ID, 'F.1.Input_Trans_Jual');
  Logger.log('=== HASIL IMPORT PENJUALAN ===');
  Logger.log(JSON.stringify(hasilPenjualan));

  const hasilPembelian = importPembelianLama(SOURCE_ID, 'F.2.Input_Trans_Beli');
  Logger.log('=== HASIL IMPORT PEMBELIAN ===');
  Logger.log(JSON.stringify(hasilPembelian));
}

/**
 * DIAGNOSTIK SLIDES — jalankan ini untuk melihat struktur lengkap
 * (tiap slide + tiap kotak teks/shape) dari file Google Slides
 * template laporan bulanan Anda. Hasilnya dicek di Logger.
 */
function diagnosaGoogleSlides() {
  const SLIDES_ID = '1pFBd6d80kwYiP7QiUtP37mfB1KNBZCWnDCMLpdvyBpo'; // dari link yang Anda kirim
  const presentation = SlidesApp.openById(SLIDES_ID);
  const slides = presentation.getSlides();

  Logger.log('Judul presentasi: ' + presentation.getName());
  Logger.log('Total slide: ' + slides.length);
  Logger.log('===========================================');

  slides.forEach(function (slide, slideIdx) {
    Logger.log('--- SLIDE ' + (slideIdx + 1) + ' (ID: ' + slide.getObjectId() + ') ---');
    const elements = slide.getPageElements();
    elements.forEach(function (el, elIdx) {
      const type = el.getPageElementType();
      let info = '  [' + elIdx + '] Tipe: ' + type;
      try {
        if (type === SlidesApp.PageElementType.SHAPE) {
          const shape = el.asShape();
          const text = shape.getText().asString();
          if (text.trim()) info += ' | Teks: "' + text.replace(/\n/g, ' \\n ') + '"';
        } else if (type === SlidesApp.PageElementType.TABLE) {
          info += ' | (tabel, ' + el.asTable().getNumRows() + 'x' + el.asTable().getNumColumns() + ')';
        } else if (type === SlidesApp.PageElementType.SHEETS_CHART) {
          info += ' | (chart terhubung dari Google Sheets)';
        } else if (type === SlidesApp.PageElementType.IMAGE) {
          info += ' | (gambar)';
        }
      } catch (e) {
        info += ' | (gagal baca: ' + e.message + ')';
      }
      Logger.log(info);
    });
    Logger.log('');
  });
}

/**
 * DIAGNOSTIK SLIDES TAHAP 2 — mencari spreadsheet & chart ID di
 * balik tiap SHEETS_CHART di template Laporan Bulanan.
 */
function diagnosaGoogleSlidesCharts() {
  const SLIDES_ID = '1pFBd6d80kwYiP7QiUtP37mfB1KNBZCWnDCMLpdvyBpo';
  const presentation = SlidesApp.openById(SLIDES_ID);
  const slides = presentation.getSlides();

  slides.forEach(function (slide, slideIdx) {
    const elements = slide.getPageElements();
    elements.forEach(function (el, elIdx) {
      if (el.getPageElementType() === SlidesApp.PageElementType.SHEETS_CHART) {
        const chart = el.asSheetsChart();
        Logger.log('SLIDE ' + (slideIdx + 1) + ' [' + elIdx + '] -> SpreadsheetID: ' + chart.getSpreadsheetId() + ' | ChartID: ' + chart.getChartId());
      }
    });
  });
}

/**
 * DIAGNOSTIK SLIDES TAHAP 3 — mencari range data persis yang
 * dipakai tiap chart di kedua spreadsheet sumber (supaya tahu
 * sel mana yang harus ditulis ulang untuk update chart-nya).
 */
function diagnosaChartRanges() {
  const spreadsheetIds = [
    '13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o',
    '12kEW42AGF7nSuqKjBjFy4CwoaIiD9zN3X4ogfwYgMPk'
  ];
  const targetChartIds = [1189303332, 137024286, 676736328, 1764994655, 1773697314];

  spreadsheetIds.forEach(function (ssId) {
    Logger.log('=========================================');
    Logger.log('SPREADSHEET: ' + ssId);
    const ss = SpreadsheetApp.openById(ssId);
    const sheets = ss.getSheets();

    sheets.forEach(function (sheet) {
      const charts = sheet.getCharts();
      charts.forEach(function (chart) {
        const chartId = chart.getChartId();
        if (targetChartIds.indexOf(chartId) === -1) return; // hanya yang relevan

        Logger.log('--- Chart ID: ' + chartId + ' (di sheet "' + sheet.getName() + '") ---');
        Logger.log('Tipe chart: ' + chart.getOptions().get('chartType'));
        const ranges = chart.getRanges();
        ranges.forEach(function (range, idx) {
          Logger.log('  Range [' + idx + ']: ' + range.getSheet().getName() + '!' + range.getA1Notation());
        });
      });
    });
  });
}

/**
 * DIAGNOSTIK SLIDES TAHAP 4 — dump isi persis tiap range data
 * chart, supaya saya tahu struktur kolomnya sebelum menulis ulang.
 */
function diagnosaChartData() {
  const targets = [
    { ss: '13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o', sheet: 'Rekap', range: 'A80:C86' },
    { ss: '13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o', sheet: 'Lap', range: 'B40:C49' },
    { ss: '13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o', sheet: 'Lap', range: 'B40:F48' },
    { ss: '12kEW42AGF7nSuqKjBjFy4CwoaIiD9zN3X4ogfwYgMPk', sheet: 'Trading ', range: 'A157:E160' }
  ];

  targets.forEach(function (t) {
    Logger.log('=========================================');
    Logger.log(t.ss + ' | ' + t.sheet + '!' + t.range);
    const ss = SpreadsheetApp.openById(t.ss);
    const sheet = ss.getSheetByName(t.sheet);
    const values = sheet.getRange(t.range).getValues();
    values.forEach(function (row, i) {
      Logger.log('  Baris ' + i + ': ' + JSON.stringify(row));
    });
  });
}

// ------------------------------------------------------------
// LAPORAN BULANAN — GENERATE KE GOOGLE SLIDES (aman, tidak
// menyentuh spreadsheet/Slides asli milik user sama sekali —
// data ditulis ke tab BARU khusus 'Lap_Bulanan' yang dibuat user
// sendiri di spreadsheet Bullion Trading Process, bukan ke tab
// Lap/Rekap/Trading yang lama, dan bukan ke database kita)
// ------------------------------------------------------------
const MONTHLY_TEMPLATE_ID = '1pFBd6d80kwYiP7QiUtP37mfB1KNBZCWnDCMLpdvyBpo';
const TARGET_TAHUNAN = 32200000000000; // Rp32,2 Triliun — target realisasi tahunan (manual, isi ulang tiap tahun kalau berubah)
const MONTHLY_DATA_SPREADSHEET_ID = '13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o';
const MONTHLY_DATA_SHEET_NAME = 'Lap_Bulanan';

/**
 * Menyiapkan tab KHUSUS 'Lap_Bulanan' (dibuat user sendiri) di
 * spreadsheet Bullion Trading Process — berisi data + chart untuk
 * laporan bulanan. Dibuat ulang tiap kali generate supaya selalu
 * pakai data terbaru, tidak menumpuk history lama. TIDAK menyentuh
 * tab Lap/Rekap/Trading yang sudah ada.
 */
function prepareMonthlyReportSheet(yearMonth) {
  const db = SpreadsheetApp.openById(MONTHLY_DATA_SPREADSHEET_ID);
  let sheet = db.getSheetByName(MONTHLY_DATA_SHEET_NAME);
  if (sheet) {
    // Bersihkan chart lama supaya tidak menumpuk setiap generate.
    sheet.getCharts().forEach(c => sheet.removeChart(c));
    sheet.clearContents();
  } else {
    sheet = db.insertSheet(MONTHLY_DATA_SHEET_NAME);
  }

  const trend8 = getMonthlyTrend(8);
  const supplierBreakdown = getSupplierBreakdownMonthWithNominal(yearMonth);
  const piutang = getPiutangSummary();

  // Blok A: Realisasi trend (dalam Miliar Rupiah) + baris Target — A1:B10
  sheet.getRange('A1:B1').setValues([['Bulan', 'Realisasi (M)']]);
  const rowsA = trend8.map(t => [t.monthLabel, t.realisasi / 1e9]);
  rowsA.push(['Target', TARGET_TAHUNAN / 1e9]);
  sheet.getRange(2, 1, rowsA.length, 2).setValues(rowsA);

  // Blok B: Gram & nominal trend — D1:H9
  sheet.getRange('D1:H1').setValues([['Bulan', 'Realisasi', 'MtM', 'Penjualan (gr)', 'Pembelian (gr)']]);
  const rowsB = trend8.map((t, i) => {
    const prevT = trend8[i - 1];
    const mtm = i > 0 && prevT.realisasi ? (t.realisasi - prevT.realisasi) / prevT.realisasi : '';
    return [t.monthLabel, t.realisasi, mtm, t.penjualanGram, t.pembelianGram];
  });
  sheet.getRange(2, 4, rowsB.length, 5).setValues(rowsB);

  // Blok C: Supplier breakdown (Pembelian) — J1:L8
  sheet.getRange('J1:L1').setValues([['Supplier', 'Gramasi', 'Nominal']]);
  const topSuppliers = supplierBreakdown.slice(0, 6);
  if (topSuppliers.length) {
    sheet.getRange(2, 10, topSuppliers.length, 3).setValues(topSuppliers.map(s => [s.seller, s.gram, s.nominal]));
  }

  // Blok D: Gold in Transit (Piutang per supplier) — N1:O10
  sheet.getRange('N1:O1').setValues([['Supplier', 'Gramasi']]);
  const piutangRows = (piutang.rows || []).slice(0, 8);
  if (piutangRows.length) {
    sheet.getRange(2, 14, piutangRows.length, 2).setValues(piutangRows.map(r => [r.supplierNama, r.gramasi]));
  }

  SpreadsheetApp.flush();

  const charts = {};

  charts.realisasi = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(1, 1, rowsA.length + 1, 1))
    .addRange(sheet.getRange(1, 2, rowsA.length + 1, 1))
    .setPosition(1, 1, 0, 0)
    .setOption('isStacked', true)
    .setOption('legend', { position: 'none' })
    .setOption('colors', ['#046A38'])
    .setOption('hAxis', { title: 'Bulan', textStyle: { fontSize: 11 } })
    .setOption('vAxis', { title: 'Realisasi (Miliar Rupiah)', textStyle: { fontSize: 10 }, gridlines: { color: '#eee' } })
    .setOption('chartArea', { width: '80%', height: '65%', left: 80 })
    .setOption('bar', { groupWidth: '65%' })
    .build();
  sheet.insertChart(charts.realisasi);

  charts.gram = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(1, 4, rowsB.length + 1, 1))
    .addRange(sheet.getRange(1, 7, rowsB.length + 1, 2))
    .setPosition(1, 5, 0, 0)
    .setOption('isStacked', true)
    .setOption('colors', ['#C9A227', '#046A38'])
    .setOption('legend', { position: 'top', textStyle: { fontSize: 11 } })
    .setOption('hAxis', { title: 'Bulan', textStyle: { fontSize: 11 } })
    .setOption('vAxis', { textStyle: { fontSize: 10 }, gridlines: { color: '#eee' } })
    .setOption('chartArea', { width: '85%', height: '65%', left: 65 })
    .setOption('bar', { groupWidth: '65%' })
    .build();
  sheet.insertChart(charts.gram);

  if (topSuppliers.length) {
    charts.supplierBar = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(1, 10, topSuppliers.length + 1, 1))
      .addRange(sheet.getRange(1, 11, topSuppliers.length + 1, 1))
      .setPosition(1, 10, 0, 0)
      .setOption('isStacked', true)
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#046A38'])
      .setOption('hAxis', { title: 'Supplier', textStyle: { fontSize: 11 } })
      .setOption('vAxis', { textStyle: { fontSize: 10 }, gridlines: { color: '#eee' } })
      .setOption('chartArea', { width: '85%', height: '68%', left: 60 })
      .build();
    sheet.insertChart(charts.supplierBar);

    charts.supplierDonut = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(sheet.getRange(1, 10, topSuppliers.length + 1, 1))
      .addRange(sheet.getRange(1, 11, topSuppliers.length + 1, 1))
      .setPosition(20, 10, 0, 0)
      .setOption('pieHole', 0.5)
      .setOption('colors', ['#046A38', '#1E9E5A', '#6FCB9F', '#C9A227', '#E8C46B', '#94A3B8'])
      .setOption('pieSliceText', 'percentage')
      .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
      .build();
    sheet.insertChart(charts.supplierDonut);
  }

  if (piutangRows.length) {
    charts.transit = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange(1, 14, piutangRows.length + 1, 1))
      .addRange(sheet.getRange(1, 15, piutangRows.length + 1, 1))
      .setPosition(1, 14, 0, 0)
      .setOption('isStacked', true)
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#C9A227'])
      .setOption('hAxis', { textStyle: { fontSize: 10 }, gridlines: { color: '#eee' } })
      .setOption('vAxis', { textStyle: { fontSize: 11 } })
      .setOption('chartArea', { width: '75%', height: '75%', left: 90 })
      .build();
    sheet.insertChart(charts.transit);
  }

  SpreadsheetApp.flush();
  return { sheet: sheet, charts: charts, trend8: trend8 };
}

/**
 * Sama seperti getSupplierBreakdownMonth() tapi ikut menghitung
 * nominal per supplier (dibutuhkan chart Rekap-style).
 */
function getSupplierBreakdownMonthWithNominal(yearMonth) {
  const target = yearMonth || getMonthKey(new Date());
  const list = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL).filter(p => extractDateOnly(p.Tanggal).slice(0, 7) === target);
  const grouped = {};
  list.forEach(p => {
    const key = p.Seller;
    if (!grouped[key]) grouped[key] = { seller: key, gram: 0, nominal: 0 };
    grouped[key].gram += Number(p.Qty || 0);
    grouped[key].nominal += Number(p.TotalHarga || 0);
  });
  return Object.values(grouped).sort((a, b) => b.gram - a.gram);
}

/**
 * Isi teks tiap shape narasi di template, dipetakan sesuai hasil
 * diagnostik struktur slide (index tetap, sudah dikonfirmasi manual).
 * Wording mengikuti pola kalimat template PPTX asli kata-demi-kata,
 * hanya angka/bulan & pilihan kata naik-turun yang menyesuaikan data.
 */
function buildMonthlySlideTexts(yearMonth) {
  const summary = getMonthlySummary(yearMonth);
  const cur = summary.current;
  const prev = summary.prev;
  const first = summary.first;
  const trend = summary.trend; // urut lama -> baru, berakhir di bulan berjalan
  const breakdown = getSupplierBreakdownMonthWithNominal(yearMonth || getMonthKey(new Date()));
  const piutang = getPiutangSummary();

  const R = formatRupiahShortServer;
  const gr = n => Number(n || 0).toLocaleString('id-ID', { maximumFractionDigits: 0 });
  const pctStr = n => Math.abs(n).toFixed(2).replace('.', ',') + '%';

  const naik = !prev || cur.realisasi >= prev.realisasi;
  const mtmPct = prev && prev.realisasi ? ((cur.realisasi - prev.realisasi) / prev.realisasi) * 100 : 0;

  // 3 bulan sebelum bulan berjalan (kalau tersedia), urutan terbaru->lama untuk kalimat pembuka.
  const idxCur = trend.length - 1;
  const back = n => trend[idxCur - n]; // back(1)=bulan lalu, back(2)=2 bulan lalu, dst.
  const b1 = back(1), b2 = back(2), b3 = back(3);

  // Titik terendah/tertinggi dalam window trend (selain bulan berjalan) untuk kalimat "sempat mengalami penurunan/puncak pada ...".
  const others = trend.slice(0, idxCur);
  const extreme = others.length
    ? others.reduce((m, t) => (naik ? (t.realisasi < m.realisasi ? t : m) : (t.realisasi > m.realisasi ? t : m)), others[0])
    : null;

  // Daftar "Bulan sebesar X" dari trend[1] s/d trend[-2] (persis pola "dari Februari ... hingga Mei").
  function historyList(arr) {
    if (arr.length < 2) return '';
    const mids = arr.slice(1, -1).map(t => t.monthLabel + ' sebesar ' + R(t.realisasi));
    const last = arr[arr.length - 1];
    return mids.length
      ? mids.join(', ') + ', hingga ' + last.monthLabel + ' sebesar ' + R(last.realisasi)
      : last.monthLabel + ' sebesar ' + R(last.realisasi);
  }

  const slide1 = {};

  slide1.p8 = 'Realisasi per ' + cur.monthLabelFull + ' menunjukkan tren ' + (naik ? 'perbaikan yang semakin kuat' : 'perlambatan yang perlu diwaspadai') + '. ' +
    'Realisasi tercatat sebesar ' + R(cur.realisasi) + (b1 ? ', ' + (naik ? 'meningkat' : 'menurun') + ' dibandingkan posisi ' + b1.monthLabel + ' sebesar ' + R(b1.realisasi) : '') +
    (b2 ? ', ' + b2.monthLabel + ' sebesar ' + R(b2.realisasi) : '') + (b3 ? ', dan ' + b3.monthLabel + ' sebesar ' + R(b3.realisasi) : '') + '.\n' +
    (extreme
      ? (naik ? 'Peningkatan ini mencerminkan penguatan kinerja secara bertahap setelah sempat mengalami penurunan tajam pada ' : 'Penurunan ini perlu mendapat perhatian, setelah sebelumnya sempat mencapai titik tertinggi pada ') +
        extreme.monthLabel + ' sebesar ' + R(extreme.realisasi) + '.'
      : '');

  slide1.p5 = prev
    ? 'Tren ' + (naik ? 'peningkatan' : 'pergerakan') + ' yang berlangsung sejak ' + trend[0].monthLabelFull + ' hingga ' + cur.monthLabelFull +
      ' menunjukkan bahwa momentum ' + (naik ? 'pemulihan' : 'perubahan') + ' telah terbentuk secara konsisten. Secara Month-to-Month (MtM), realisasi ' + cur.monthLabel + ' ' +
      (naik ? 'tumbuh' : 'turun') + ' sebesar ' + pctStr(mtmPct) + ', yaitu dari ' + R(prev.realisasi) + ' pada ' + prev.monthLabel + ' menjadi ' + R(cur.realisasi) + ' pada ' + cur.monthLabel +
      '. Pertumbuhan tersebut melanjutkan tren ' + (naik ? 'positif' : 'sebelumnya') + ' dari ' + historyList(trend) + '.'
    : 'Belum tersedia data bulan sebelumnya untuk perbandingan MtM.';

  const vsFirstPct = first && first.realisasi && first.yearMonth !== cur.yearMonth ? ((cur.realisasi - first.realisasi) / first.realisasi) * 100 : null;
  slide1.p7 = (first && vsFirstPct !== null)
    ? 'Dibandingkan dengan posisi ' + first.monthLabelFull + ' sebesar ' + R(first.realisasi) + ', realisasi ' + cur.monthLabelFull + ' telah ' +
      (vsFirstPct >= 0 ? 'lebih tinggi' : 'lebih rendah') + ' sebesar ' + R(Math.abs(cur.realisasi - first.realisasi)) + ' atau ' + (vsFirstPct >= 0 ? 'tumbuh' : 'turun') + ' sekitar ' + pctStr(vsFirstPct) +
      '. Kondisi ini menunjukkan bahwa kinerja ' + (vsFirstPct >= 0 ? 'tidak hanya berhasil kembali ke level awal periode, tetapi juga telah melampauinya secara cukup signifikan' : 'masih berada di bawah level awal periode dan perlu mendapat perhatian') + '.\n\n' +
      (function () {
        const targetPct = TARGET_TAHUNAN ? (cur.realisasi / TARGET_TAHUNAN) * 100 : 0;
        const gap = TARGET_TAHUNAN - cur.realisasi;
        const gapPct = TARGET_TAHUNAN ? (gap / TARGET_TAHUNAN) * 100 : 0;
        let s = 'Dari sisi pencapaian target, realisasi ' + cur.monthLabel + ' telah mencapai sekitar ' + targetPct.toFixed(2).replace('.', ',') + '% dari target sebesar ' + R(TARGET_TAHUNAN) + '.';
        if (gap > 0) {
          s += ' Dengan demikian, masih terdapat gap sebesar ' + R(gap) + ' atau sekitar ' + gapPct.toFixed(2).replace('.', ',') + '% yang perlu dipenuhi.';
          if (prev) {
            const prevGapPct = TARGET_TAHUNAN ? ((TARGET_TAHUNAN - prev.realisasi) / TARGET_TAHUNAN) * 100 : 0;
            s += ' Dibandingkan posisi ' + prev.monthLabel + ', gap terhadap target telah ' + (gapPct < prevGapPct ? 'menurun' : 'melebar') + ' dari ' + prevGapPct.toFixed(2).replace('.', ',') + '% menjadi ' + gapPct.toFixed(2).replace('.', ',') + '%.';
          }
          s += ' Percepatan realisasi dan kesinambungan transaksi tetap perlu dijaga agar sisa target dapat dicapai sesuai rencana.';
        } else {
          s += ' Target tahunan telah tercapai, melampaui sebesar ' + R(Math.abs(gap)) + '.';
        }
        return s;
      })()
    : '';

  const levelBlock = first
    ? (vsFirstPct >= 0 ? 'Kinerja Semakin Melampaui Level ' + first.monthLabel : 'Kinerja Masih di Bawah Level ' + first.monthLabel) + '\n' +
      'Realisasi ' + cur.monthLabel + ' sebesar ' + R(cur.realisasi) + ' ' + (vsFirstPct >= 0 ? 'telah melampaui' : 'masih berada di bawah') + ' realisasi ' + first.monthLabel + ' sebesar ' + R(first.realisasi) +
      '. Terdapat ' + (vsFirstPct >= 0 ? 'peningkatan' : 'selisih') + ' sebesar ' + R(Math.abs(cur.realisasi - first.realisasi)) + ' atau ' + pctStr(vsFirstPct) + ' dibandingkan posisi awal periode.\n\n'
    : '';

  const targetPctForInsight = TARGET_TAHUNAN ? (cur.realisasi / TARGET_TAHUNAN) * 100 : 0;
  const gapForInsight = TARGET_TAHUNAN - cur.realisasi;
  const gapPctForInsight = TARGET_TAHUNAN ? (gapForInsight / TARGET_TAHUNAN) * 100 : 0;
  const prevGapPctForInsight = prev && TARGET_TAHUNAN ? ((TARGET_TAHUNAN - prev.realisasi) / TARGET_TAHUNAN) * 100 : null;
  const gapBlock = gapForInsight > 0
    ? 'Gap Target Semakin Menyempit\n' +
      'Realisasi ' + cur.monthLabel + ' telah mencapai sekitar ' + targetPctForInsight.toFixed(2).replace('.', ',') + '% dari target sebesar ' + R(TARGET_TAHUNAN) + '. Masih terdapat gap sebesar ' + R(gapForInsight) + ' atau sekitar ' + gapPctForInsight.toFixed(2).replace('.', ',') + '% yang perlu dipenuhi.' +
      (prevGapPctForInsight !== null ? ' Dibandingkan posisi ' + prev.monthLabel + ', gap terhadap target telah ' + (gapPctForInsight < prevGapPctForInsight ? 'menurun' : 'melebar') + ' dari ' + prevGapPctForInsight.toFixed(2).replace('.', ',') + '% menjadi ' + gapPctForInsight.toFixed(2).replace('.', ',') + '%.' : '') +
      ' Akselerasi tetap diperlukan agar target dapat tercapai sesuai rencana.\n\n'
    : 'Target Tahunan Tercapai\n' +
      'Realisasi ' + cur.monthLabel + ' sebesar ' + R(cur.realisasi) + ' telah melampaui target tahunan sebesar ' + R(TARGET_TAHUNAN) + '.\n\n';

  slide1.p6 = (naik ? 'Tren Pemulihan Semakin Kuat' : 'Tren Perlambatan Perlu Diwaspadai') + '\n' +
    'Realisasi menunjukkan tren ' + (naik ? 'pemulihan yang semakin kuat' : 'perlambatan yang cukup nyata') + '. Sejak ' + trend[0].monthLabelFull + ' hingga ' + cur.monthLabelFull + ', realisasi ' +
    (naik ? 'terus meningkat secara konsisten' : 'bergerak fluktuatif') + ', dari ' + R(trend[0].realisasi) + ' menjadi ' + R(cur.realisasi) + '. Pada ' + cur.monthLabel + ', realisasi ' + (naik ? 'tumbuh' : 'turun') +
    ' sebesar ' + (prev ? pctStr(mtmPct) : '-') + ' secara Month-to-Month (MtM)' + (prev ? ' dibandingkan ' + prev.monthLabel + ' sebesar ' + R(prev.realisasi) : '') + '.\n\n' +
    levelBlock +
    gapBlock +
    'Momentum Pertumbuhan Perlu Dipertahankan\n' +
    'Peningkatan realisasi perlu terus dijaga agar tidak terjadi perlambatan pada periode berikutnya. Fokus utama ke depan adalah menjaga kesinambungan transaksi, mengoptimalkan volume dari supplier dan buyer eksisting, memperluas potensi mitra baru, serta mempercepat eksekusi transaksi.';

  const slide2 = {};
  const totalGramCur = cur.penjualanGram + cur.pembelianGram;
  const totalGramPrev = prev ? prev.penjualanGram + prev.pembelianGram : 0;
  const naikVol = !prev || totalGramCur >= totalGramPrev;
  const mtmVolPct = prev && totalGramPrev ? ((totalGramCur - totalGramPrev) / totalGramPrev) * 100 : 0;
  const pjPct = totalGramCur ? (cur.penjualanGram / totalGramCur) * 100 : 0;
  const pbPct = totalGramCur ? (cur.pembelianGram / totalGramCur) * 100 : 0;

  slide2.p8 = 'Pada ' + cur.monthLabelFull + ', aktivitas transaksi menunjukkan ' + (naikVol ? 'perbaikan' : 'perlambatan') + ' dibandingkan bulan sebelumnya. Secara volume, total transaksi tercatat sebesar ' +
    gr(totalGramCur) + ' gram' + (prev ? ', ' + (naikVol ? 'meningkat' : 'menurun') + ' dibandingkan ' + prev.monthLabel + ' sebesar ' + gr(totalGramPrev) + ' gram atau ' + (naikVol ? 'tumbuh' : 'turun') + ' sebesar ' + pctStr(mtmVolPct) + ' secara Month-to-Month (MtM)' : '') + '.\n' +
    'Dari sisi komposisi, volume penjualan tercatat sebesar ' + gr(cur.penjualanGram) + ' gram' +
    (prev ? ', ' + (cur.penjualanGram >= prev.penjualanGram ? 'meningkat' : 'menurun') + ' sebesar ' + gr(Math.abs(cur.penjualanGram - prev.penjualanGram)) + ' gram dari posisi ' + prev.monthLabel + ' sebesar ' + gr(prev.penjualanGram) + ' gram' : '') + '. ' +
    'Volume pembelian tercatat sebesar ' + gr(cur.pembelianGram) + ' gram. Dengan demikian, penjualan berkontribusi sebesar ' + pjPct.toFixed(0) + '% dan pembelian ' + pbPct.toFixed(0) + '% terhadap total volume transaksi ' + cur.monthLabel + '.';

  slide2.p9 = 'Pada ' + cur.monthLabelFull + ', aktivitas transaksi secara nominal menunjukkan ' + (naik ? 'peningkatan' : 'penurunan') + ' dibandingkan bulan sebelumnya. Penjualan tercatat sebesar ' + formatRupiahServer(cur.penjualanNominal) +
    (prev && prev.penjualanNominal ? ', ' + (cur.penjualanNominal >= prev.penjualanNominal ? 'tumbuh' : 'turun') + ' sekitar ' + pctStr(((cur.penjualanNominal - prev.penjualanNominal) / prev.penjualanNominal) * 100) + ' secara Month-to-Month (MtM) dari ' + prev.monthLabel + ' sebesar ' + formatRupiahServer(prev.penjualanNominal) : '') + '. ' +
    'Sementara itu, pembelian tercatat sebesar ' + formatRupiahServer(cur.pembelianNominal) +
    (prev && prev.pembelianNominal ? ', ' + (cur.pembelianNominal >= prev.pembelianNominal ? 'meningkat' : 'menurun') + ' sebesar ' + pctStr(((cur.pembelianNominal - prev.pembelianNominal) / prev.pembelianNominal) * 100) + ' MtM dari ' + formatRupiahServer(prev.pembelianNominal) : '') + '.\n' +
    'Sejalan dengan hal tersebut, margin ' + cur.monthLabel + ' tercatat ' + (cur.margin >= 0 ? 'positif' : 'negatif') + ' sebesar ' + R(Math.abs(cur.margin)) +
    (prev ? ' dibandingkan margin ' + prev.monthLabel + ' sebesar ' + R(Math.abs(prev.margin)) : '') + '. Total nilai transaksi pembelian dan penjualan mencapai ' + formatRupiahServer(cur.realisasi) + '.';

  slide2.p5 = extreme
    ? 'Peningkatan pada ' + cur.monthLabel + (naikVol ? ' menunjukkan adanya pemulihan aktivitas perdagangan' : ' masih menunjukkan pelemahan aktivitas perdagangan') + '. ' +
      'Komposisi pembelian dan penjualan yang tetap seimbang mencerminkan bahwa perputaran transaksi dan pengelolaan persediaan masih terjaga. Ke depan, momentum perlu dipertahankan melalui penguatan pasokan supplier, optimalisasi permintaan buyer, dan percepatan eksekusi transaksi.'
    : 'Komposisi pembelian dan penjualan yang tetap seimbang mencerminkan bahwa perputaran transaksi dan pengelolaan persediaan masih terjaga.';

  const marginRatioCur = cur.penjualanNominal ? (cur.margin / cur.penjualanNominal) * 100 : 0;
  const marginRatioPrev = prev && prev.penjualanNominal ? (prev.margin / prev.penjualanNominal) * 100 : null;
  slide2.p6 = 'Meskipun nominal transaksi mengalami ' + (naik ? 'peningkatan' : 'perubahan') + ', rasio margin terhadap penjualan tercatat sekitar ' + marginRatioCur.toFixed(2).replace('.', ',') + '% pada ' + cur.monthLabel +
    (marginRatioPrev !== null ? ' (' + marginRatioPrev.toFixed(2).replace('.', ',') + '% pada ' + prev.monthLabel + ')' : '') +
    '. Secara keseluruhan, kinerja ' + cur.monthLabel + ' mencerminkan ' + (naik ? 'pemulihan aktivitas transaksi yang positif' : 'perlambatan aktivitas transaksi') +
    ', namun optimalisasi harga beli, harga jual, dan spread tetap perlu dilakukan agar pertumbuhan volume dan nominal transaksi dapat memberikan peningkatan margin yang lebih maksimal.';

  const slide3 = {};
  if (breakdown.length) {
    const totalGramSupplier = breakdown.reduce((s, r) => s + r.gram, 0);
    const withPct = breakdown.map(r => ({ ...r, pct: totalGramSupplier ? (r.gram / totalGramSupplier) * 100 : 0 }));
    const top = withPct[0];
    slide3.p6 = 'Berdasarkan grafik pembelian terhadap supplier pada bulan ' + cur.monthLabelFull + ', transaksi pembelian bullion masih didominasi oleh ' + top.seller + ' dengan total pembelian sebesar ' +
      gr(top.gram) + ' gram atau sekitar ' + top.pct.toFixed(1) + '% dari total pembelian. Dominasi ini menunjukkan bahwa ' + top.seller + ' masih menjadi supplier utama dalam pemenuhan kebutuhan pasokan bullion.\n\n' +
      (withPct.length > 1 ? 'Supplier dengan kontribusi terbesar berikutnya adalah ' +
        withPct.slice(1, 3).map(r => r.seller + ' sebesar ' + gr(r.gram) + ' gram atau sekitar ' + r.pct.toFixed(1) + '%').join(', diikuti oleh ') +
        '. Kedua supplier tersebut menjadi kontributor penting setelah ' + top.seller + ', meskipun selisih volume pembeliannya masih cukup jauh dibandingkan supplier utama.\n\n' : '') +
      (withPct.length > 3 ? 'Sementara itu, kontribusi dari supplier lainnya masih relatif kecil, yaitu ' +
        withPct.slice(3).map(r => r.seller + ' sebesar ' + r.pct.toFixed(1) + '%').join(', ') + ' dari total pembelian.\n\n' : '') +
      'Secara keseluruhan, struktur pembelian bullion pada ' + cur.monthLabelFull + ' masih terkonsentrasi pada ' + top.seller +
      '. Oleh karena itu, diperlukan pemantauan terhadap konsentrasi pembelian supplier guna menjaga kesinambungan pasokan, mendorong diversifikasi sumber pembelian, serta memitigasi risiko ketergantungan terhadap satu supplier utama.';
  } else {
    slide3.p6 = 'Belum ada data pembelian tercatat pada bulan ini.';
  }

  const piutangRows = (piutang.rows || []).filter(r => r.gramasi > 0).sort((a, b) => b.gramasi - a.gramasi);
  const zeroRows = (piutang.rows || []).filter(r => !r.gramasi);
  if (piutangRows.length) {
    const topP = piutangRows[0];
    const totalHutang = (piutang.rows || []).reduce((s, r) => s + r.gramasi, 0);
    slide3.p8 = 'Berdasarkan data hutang supplier per akhir ' + cur.monthLabelFull + ', hutang terbesar tercatat kepada ' + topP.supplierNama + ' sebesar ' + gr(topP.gramasi) +
      ' gram, sehingga menjadi komponen hutang supplier paling dominan.' +
      (piutangRows.length > 1 ? ' Selanjutnya, hutang kepada ' + piutangRows.slice(1).map(r => r.supplierNama + ' sebesar ' + gr(r.gramasi) + ' gram').join(', ') + '.' : '') +
      ' Total Hutang Supplier pada bulan ' + cur.monthLabel + ' sebesar ' + gr(totalHutang) + ' gram.\n' +
      (zeroRows.length ? 'Sementara itu, ' + zeroRows.map(r => r.supplierNama).join(', ') + ' tidak memiliki outstanding hutang atau tercatat 0 gram.\n' : '') +
      'Kondisi ini menunjukkan bahwa kewajiban pemenuhan hutang emas masih terkonsentrasi pada ' + topP.supplierNama + ', sehingga perlu dilakukan monitoring penyelesaian hutang secara berkala untuk menjaga kelancaran operasional dan hubungan kerja sama dengan supplier.';
  } else {
    slide3.p8 = 'Tidak terdapat outstanding hutang supplier per akhir ' + cur.monthLabelFull + '.';
  }

  return { slide1: slide1, slide2: slide2, slide3: slide3 };
}

/**
 * Set teks shape ke-N di slide tertentu (index sesuai urutan
 * getPageElements(), sudah dikonfirmasi lewat diagnostik manual).
 */
function setShapeTextByIndex(slide, index, text) {
  if (!text) return;
  const elements = slide.getPageElements();
  if (index >= elements.length) return;
  const el = elements[index];
  if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
  el.asShape().getText().setText(text);
}

/**
 * Ganti chart lama (SHEETS_CHART) di posisi tertentu dengan chart
 * baru dari sheet kita sendiri, di posisi & ukuran yang sama persis.
 */
function replaceChartAtIndex(slide, index, newEmbeddedChart) {
  const elements = slide.getPageElements();
  if (index >= elements.length) return;
  const el = elements[index];
  if (el.getPageElementType() !== SlidesApp.PageElementType.SHEETS_CHART) return;
  const left = el.getLeft(), top = el.getTop(), width = el.getWidth(), height = el.getHeight();
  el.remove();
  slide.insertSheetsChart(newEmbeddedChart, left, top, width, height);
}

/**
 * FUNGSI UTAMA — dipanggil dari tombol Download di halaman
 * Laporan Bulanan. Menghasilkan PDF base64 untuk didownload client.
 */
function generateMonthlyReportSlides(yearMonth) {
  requireRole('Viewer');
  const target = yearMonth || getMonthKey(new Date());

  const prepared = prepareMonthlyReportSheet(target);
  const texts = buildMonthlySlideTexts(target);

  const templateFile = DriveApp.getFileById(MONTHLY_TEMPLATE_ID);
  const copyFile = templateFile.makeCopy('Laporan Bulanan Perdagangan Emas - ' + target);
  const presentation = SlidesApp.openById(copyFile.getId());
  const slides = presentation.getSlides();

  // Slide 1 — Kinerja Perdagangan Emas
  setShapeTextByIndex(slides[0], 5, texts.slide1.p5);
  setShapeTextByIndex(slides[0], 6, texts.slide1.p6);
  setShapeTextByIndex(slides[0], 7, texts.slide1.p7);
  setShapeTextByIndex(slides[0], 8, texts.slide1.p8);
  replaceChartAtIndex(slides[0], 9, prepared.charts.realisasi);

  // Slide 2 — Tren Transaksi Penjualan
  setShapeTextByIndex(slides[1], 5, texts.slide2.p5);
  setShapeTextByIndex(slides[1], 6, texts.slide2.p6);
  setShapeTextByIndex(slides[1], 8, texts.slide2.p8);
  setShapeTextByIndex(slides[1], 9, texts.slide2.p9);
  replaceChartAtIndex(slides[1], 10, prepared.charts.gram);
  replaceChartAtIndex(slides[1], 11, prepared.charts.gram);

  // Slide 3 — Tren Transaksi Pembelian
  setShapeTextByIndex(slides[2], 6, texts.slide3.p6);
  setShapeTextByIndex(slides[2], 8, texts.slide3.p8);
  if (prepared.charts.supplierBar) replaceChartAtIndex(slides[2], 7, prepared.charts.supplierBar);
  if (prepared.charts.supplierDonut) replaceChartAtIndex(slides[2], 10, prepared.charts.supplierDonut);
  if (prepared.charts.transit) replaceChartAtIndex(slides[2], 11, prepared.charts.transit);

  presentation.saveAndClose();
  // Beri jeda supaya chart yang baru disisipkan sempat selesai di-render
  // di sisi server Google sebelum di-export (kalau tidak, chart bisa
  // tampil kosong di hasil export).
  Utilities.sleep(5000);

  // Rapikan file duplikat ke folder khusus (supaya gampang ditemukan & tidak numpuk di root Drive).
  const folderName = 'Laporan Bulanan Perdagangan Emas - Auto Generated';
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  folder.addFile(copyFile);
  DriveApp.getRootFolder().removeFile(copyFile);

  // Export ke PPTX (bisa dibuka di PowerPoint maupun di-upload ulang ke Slides).
  const exportUrl = 'https://docs.google.com/presentation/d/' + copyFile.getId() + '/export/pptx';
  const exportResponse = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  const pptxBlob = exportResponse.getBlob();
  const base64 = Utilities.base64Encode(pptxBlob.getBytes());

  const monthLabelFull = prepared.trend8[prepared.trend8.length - 1].monthLabelFull;
  return {
    base64: base64,
    fileName: 'Laporan Bulanan Perdagangan Emas – ' + monthLabelFull + '.pptx',
    slidesUrl: 'https://docs.google.com/presentation/d/' + copyFile.getId() + '/edit'
  };
}

// ------------------------------------------------------------
// SUPPLIER PRICE TREND ANALYSIS
// ------------------------------------------------------------
const SUPPLIER_PRICE_SOURCES = [
  { key: 'HRTA', col: PD_COL.hrtaDisc },
  { key: 'IGS', col: PD_COL.igs },
  { key: 'SJL', col: PD_COL.sjl },
  { key: 'IDN', col: PD_COL.idn },
  { key: 'MKB/STRGLD', col: PD_COL.mkb },
  { key: 'LOTUS', col: PD_COL.lotus },
  { key: 'APEPI', col: PD_COL.apepi },
  { key: 'Penetapan Harga', col: PD_COL.keputusanHarga }
];

/**
 * Baca SEKALI seluruh Price Discovery V3 (bukan per-tanggal) lalu
 * susun jadi time-series per supplier untuk rentang tanggal
 * tertentu — jauh lebih cepat daripada memanggil getSumberDataLive
 * berkali-kali dalam loop.
 */
function getSupplierPriceTimeSeriesRaw(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = ss.getSheetByName(PRICE_DISCOVERY_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + PRICE_DISCOVERY_SHEET_NAME + '" tidak ditemukan.');
  const values = sheet.getDataRange().getValues();

  const start = extractDateOnly(startDateStr);
  const end = extractDateOnly(endDateStr);

  const byDate = {};
  values.forEach(row => {
    const d = extractDateOnly(row[PD_COL.tanggal]);
    if (!d || d < start || d > end) return;
    const hargaKe = Number(row[PD_COL.hargaKe]) || 0;
    if (!byDate[d]) byDate[d] = {};
    SUPPLIER_PRICE_SOURCES.forEach(s => {
      const v = Number(row[s.col]) || 0;
      if (v > 0) {
        if (!byDate[d][s.key] || hargaKe >= byDate[d][s.key].hargaKe) {
          byDate[d][s.key] = { hargaKe: hargaKe, value: v };
        }
      }
    });
  });

  const dates = Object.keys(byDate).sort();
  const series = {};
  SUPPLIER_PRICE_SOURCES.forEach(s => {
    series[s.key] = dates.map(d => (byDate[d][s.key] ? byDate[d][s.key].value : null));
  });

  return { dates: dates, series: series, supplierList: SUPPLIER_PRICE_SOURCES.map(s => s.key) };
}

/**
 * Bungkus time-series harian jadi Mingguan/Bulanan (ambil nilai
 * terakhir yang tersedia dalam tiap bucket).
 */
function getSupplierPriceTimeSeries(startDateStr, endDateStr, frequency) {
  const raw = getSupplierPriceTimeSeriesRaw(startDateStr, endDateStr);
  if (!frequency || frequency === 'daily' || !raw.dates.length) return raw;

  function bucketKey(dateStr) {
    if (frequency === 'monthly') return dateStr.slice(0, 7);
    const d = new Date(dateStr + 'T00:00:00');
    const day = (d.getDay() + 6) % 7; // 0 = Senin
    d.setDate(d.getDate() - day);
    return formatDateOnly(d);
  }

  const buckets = [];
  const bucketToIdx = {};
  raw.dates.forEach((d, i) => {
    const k = bucketKey(d);
    if (bucketToIdx[k] === undefined) buckets.push(k);
    bucketToIdx[k] = i;
  });

  const series = {};
  raw.supplierList.forEach(key => { series[key] = buckets.map(k => raw.series[key][bucketToIdx[k]]); });

  return { dates: buckets, series: series, supplierList: raw.supplierList };
}

/**
 * KPI ringkas + insight otomatis untuk periode yang dipilih.
 */
function getSupplierPriceKpiAndInsights(startDateStr, endDateStr) {
  const data = getSupplierPriceTimeSeriesRaw(startDateStr, endDateStr);
  const n = data.dates.length;
  if (!n) return null;

  const lastIdx = n - 1;
  const lastPrices = {};
  data.supplierList.forEach(k => { lastPrices[k] = data.series[k][lastIdx]; });
  const validLast = data.supplierList.filter(k => lastPrices[k] != null);

  let cheapestToday = null, mostExpensiveToday = null;
  validLast.forEach(k => {
    if (!cheapestToday || lastPrices[k] < lastPrices[cheapestToday]) cheapestToday = k;
    if (!mostExpensiveToday || lastPrices[k] > lastPrices[mostExpensiveToday]) mostExpensiveToday = k;
  });
  const selisihTerbesar = (cheapestToday && mostExpensiveToday) ? lastPrices[mostExpensiveToday] - lastPrices[cheapestToday] : 0;

  let sumAll = 0, countAll = 0;
  data.supplierList.forEach(k => data.series[k].forEach(v => { if (v != null) { sumAll += v; countAll++; } }));
  const avgAll = countAll ? sumAll / countAll : 0;

  let spreadSum = 0, spreadCount = 0;
  for (let i = 0; i < n; i++) {
    const dayVals = data.supplierList.map(k => data.series[k][i]).filter(v => v != null);
    if (dayVals.length >= 2) { spreadSum += Math.max.apply(null, dayVals) - Math.min.apply(null, dayVals); spreadCount++; }
  }
  const avgSpread = spreadCount ? spreadSum / spreadCount : 0;

  const insights = data.supplierList.map(key => {
    const vals = data.series[key].filter(v => v != null);
    if (!vals.length) return null;
    const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
    const variance = vals.reduce((a, b) => a + Math.pow(b - avg, 2), 0) / vals.length;
    const stddev = Math.sqrt(variance);
    let cheapestDays = 0;
    for (let i = 0; i < n; i++) {
      const dayVals = data.supplierList.map(k => data.series[k][i]).filter(v => v != null);
      if (!dayVals.length) continue;
      const min = Math.min.apply(null, dayVals);
      if (data.series[key][i] === min) cheapestDays++;
    }
    const firstVal = vals[0], lastVal = vals[vals.length - 1];
    const trendPct = firstVal ? ((lastVal - firstVal) / firstVal) * 100 : 0;
    return { key: key, avg: avg, stddev: stddev, cheapestDays: cheapestDays, trendPct: trendPct };
  }).filter(x => x);

  if (!insights.length) return null;

  const termurahPeriode = insights.reduce((m, r) => (r.avg < m.avg ? r : m), insights[0]);
  const palingStabil = insights.reduce((m, r) => (r.stddev < m.stddev ? r : m), insights[0]);
  const palingFluktuatif = insights.reduce((m, r) => (r.stddev > m.stddev ? r : m), insights[0]);
  const palingSeringTermurah = insights.reduce((m, r) => (r.cheapestDays > m.cheapestDays ? r : m), insights[0]);
  const trenNaik = insights.filter(r => r.trendPct > 1).sort((a, b) => b.trendPct - a.trendPct);
  const trenTurun = insights.filter(r => r.trendPct < -1).sort((a, b) => a.trendPct - b.trendPct);

  return {
    jumlahSupplier: data.supplierList.length,
    cheapestToday: cheapestToday, cheapestTodayPrice: cheapestToday ? lastPrices[cheapestToday] : 0,
    mostExpensiveToday: mostExpensiveToday, mostExpensiveTodayPrice: mostExpensiveToday ? lastPrices[mostExpensiveToday] : 0,
    selisihTerbesar: selisihTerbesar,
    avgAll: avgAll,
    avgSpread: avgSpread,
    termurahPeriode: termurahPeriode,
    palingStabil: palingStabil,
    palingFluktuatif: palingFluktuatif,
    palingSeringTermurah: palingSeringTermurah,
    trenNaik: trenNaik.slice(0, 3),
    trenTurun: trenTurun.slice(0, 3)
  };
}

/**
 * DIAGNOSTIK — cek struktur sheet Master_Customer.
 * Cuma membaca, tidak mengubah apa pun.
 */
function diagnosaMasterCustomer() {
  const ss = SpreadsheetApp.openById('13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o');
  const sheet = ss.getSheetByName('Master_Customer');

  if (!sheet) {
    Logger.log('Sheet "Master_Customer" TIDAK ditemukan di spreadsheet itu juga.');
    Logger.log('Nama tab-nya mungkin sedikit beda (spasi/kapital). Menjalankan listSourceSheetNames() bisa bantu cek nama persis semua tab di spreadsheet itu.');
    return;
  }

  Logger.log('Total baris: ' + sheet.getLastRow() + ' | Total kolom: ' + sheet.getLastColumn());
  Logger.log('--- Scan 15 baris pertama (cari baris header yang sebenarnya) ---');
  const scanRange = sheet.getRange(1, 1, Math.min(15, sheet.getLastRow()), sheet.getLastColumn());
  const scanValues = scanRange.getValues();
  scanValues.forEach((row, i) => {
    const nonEmpty = row.filter(c => c !== '' && c !== null).length;
    Logger.log('Baris ' + (i + 1) + ' (' + nonEmpty + ' sel terisi): ' + JSON.stringify(row));
  });
}

// ------------------------------------------------------------
// PROFILE MITRA — Fase 1: Sheet MITRA_PROFILE, Daftar Mitra, Overview
// Semua fungsi ini HANYA MEMBACA sheet existing (Master_Customer,
// PenjualanTransaksi, PembelianTransaksi, Outstanding) — tidak ada
// satupun yang menulis/mengubah sheet-sheet tersebut.
// ------------------------------------------------------------
const MASTER_CUSTOMER_SPREADSHEET_ID = '13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o';
const MASTER_CUSTOMER_SHEET_NAME = 'Master_Customer';
const MASTER_CUSTOMER_HEADER_ROW = 11; // header sebenarnya ada di baris 11, bukan baris 1

/**
 * Baca Master_Customer (READ-ONLY, tidak pernah menulis ke sini).
 * Header ada di baris 11 (bukan baris 1), jadi dibaca manual.
 */
function getMasterCustomerData() {
  const ss = SpreadsheetApp.openById(MASTER_CUSTOMER_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MASTER_CUSTOMER_SHEET_NAME);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= MASTER_CUSTOMER_HEADER_ROW) return [];

  const headers = sheet.getRange(MASTER_CUSTOMER_HEADER_ROW, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const dataRange = sheet.getRange(MASTER_CUSTOMER_HEADER_ROW + 1, 1, lastRow - MASTER_CUSTOMER_HEADER_ROW, lastCol).getValues();

  return dataRange
    .filter(row => row[1]) // kolom B = Nama Customer, skip baris kosong
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
      return obj;
    });
}

/**
 * Deteksi Jenis Mitra otomatis dari histori transaksi (READ-ONLY).
 * Tidak menulis apa pun ke PenjualanTransaksi/PembelianTransaksi.
 */
function detectJenisMitra(namaPerusahaan) {
  if (!namaPerusahaan) return 'Belum ada transaksi';
  const nama = String(namaPerusahaan).trim().toLowerCase();

  const adaSebagaiSupplier = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL)
    .some(p => String(p.Seller || '').trim().toLowerCase() === nama);
  const adaSebagaiBuyer = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .some(p => String(p.Pembeli || '').trim().toLowerCase() === nama);

  if (adaSebagaiSupplier && adaSebagaiBuyer) return 'Supplier & Buyer';
  if (adaSebagaiSupplier) return 'Supplier';
  if (adaSebagaiBuyer) return 'Buyer';
  return 'Belum ada transaksi';
}

/**
 * Ringkasan transaksi (total nominal, gram, jumlah transaksi,
 * tanggal terakhir) untuk satu mitra — dari sheet transaksi existing.
 */
function getMitraTransaksiSummary(namaPerusahaan) {
  const nama = String(namaPerusahaan || '').trim().toLowerCase();

  const pembelianRows = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL)
    .filter(p => String(p.Seller || '').trim().toLowerCase() === nama);
  const penjualanRows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => String(p.Pembeli || '').trim().toLowerCase() === nama);

  const pembelian = {
    totalNominal: pembelianRows.reduce((s, p) => s + Number(p.TotalHarga || 0), 0),
    totalGram: pembelianRows.reduce((s, p) => s + Number(p.Qty || 0), 0),
    jumlahTransaksi: pembelianRows.length,
    terakhir: pembelianRows.length ? pembelianRows.map(p => extractDateOnly(p.Tanggal)).sort().pop() : null,
    rows: pembelianRows.sort((a, b) => extractDateOnly(b.Tanggal).localeCompare(extractDateOnly(a.Tanggal)))
  };
  const penjualan = {
    totalNominal: penjualanRows.reduce((s, p) => s + Number(p.TotalHarga || 0), 0),
    totalGram: penjualanRows.reduce((s, p) => s + Number(p.Qty || 0), 0),
    jumlahTransaksi: penjualanRows.length,
    terakhir: penjualanRows.length ? penjualanRows.map(p => extractDateOnly(p.Tanggal)).sort().pop() : null,
    rows: penjualanRows.sort((a, b) => extractDateOnly(b.Tanggal).localeCompare(extractDateOnly(a.Tanggal)))
  };

  const piutangRows = sheetToObjects(SHEET_NAMES.OUTSTANDING)
    .filter(o => String(o.SupplierNama || '').trim().toLowerCase() === nama);
  const piutang = {
    totalGramasi: piutangRows.reduce((s, o) => s + Number(o.Gramasi || 0), 0),
    totalValuasi: piutangRows.reduce((s, o) => s + Number(o.Valuasi || 0), 0),
    rows: piutangRows
  };

  return { pembelian: pembelian, penjualan: penjualan, piutang: piutang };
}

/**
 * Daftar seluruh mitra untuk halaman Daftar Mitra — gabungan
 * Master_Customer (baca saja) + MITRA_PROFILE (baca saja) +
 * ringkasan transaksi otomatis.
 */
function getMitraList() {
  const masterList = getMasterCustomerData();
  const profileList = sheetToObjects(SHEET_NAMES.MITRA_PROFILE);
  const profileByName = {};
  profileList.forEach(p => { profileByName[String(p.NamaPerusahaan || '').trim().toLowerCase()] = p; });

  // Baca SEKALI seluruh transaksi, lalu agregasi per nama (bukan re-scan per mitra — penting untuk performa dengan ratusan/ribuan mitra).
  const pembelianAgg = {};
  sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL).forEach(p => {
    const key = String(p.Seller || '').trim().toLowerCase();
    if (!key) return;
    if (!pembelianAgg[key]) pembelianAgg[key] = 0;
    pembelianAgg[key] += Number(p.TotalHarga || 0);
  });
  const penjualanAgg = {};
  sheetToObjects(SHEET_NAMES.PENJUALAN_FULL).forEach(p => {
    const key = String(p.Pembeli || '').trim().toLowerCase();
    if (!key) return;
    if (!penjualanAgg[key]) penjualanAgg[key] = 0;
    penjualanAgg[key] += Number(p.TotalHarga || 0);
  });

  // Baca SEKALI seluruh MITRA_PKS, ambil PKS dengan TanggalBerakhir paling jauh per Mitra_ID.
  const pksByMitraId = {};
  sheetToObjects(SHEET_NAMES.MITRA_PKS).forEach(r => {
    if (!r.Mitra_ID) return;
    const akhir = r.TanggalBerakhir ? extractDateOnly(r.TanggalBerakhir) : '';
    if (!pksByMitraId[r.Mitra_ID] || akhir > pksByMitraId[r.Mitra_ID].akhir) {
      pksByMitraId[r.Mitra_ID] = { akhir: akhir };
    }
  });
  function pksInfoFor(mitraId) {
    if (!mitraId || !pksByMitraId[mitraId] || !pksByMitraId[mitraId].akhir) return { statusPKS: '-', sisaMasaPKS: '-' };
    const st = computePKSStatus(pksByMitraId[mitraId].akhir);
    return { statusPKS: st.label, sisaMasaPKS: st.sisaHari != null ? st.sisaHari + ' hari' : '-' };
  }

  const daftar = masterList.map(m => {
    const nama = String(m['Nama Customer'] || '').trim();
    const key = nama.toLowerCase();
    const profile = profileByName[key] || {};
    const adaSupplier = pembelianAgg.hasOwnProperty(key);
    const adaBuyer = penjualanAgg.hasOwnProperty(key);
    const jenisOtomatis = adaSupplier && adaBuyer ? 'Supplier & Buyer' : adaSupplier ? 'Supplier' : adaBuyer ? 'Buyer' : 'Belum ada transaksi';
    const jenisMitra = (profile.JenisMitra_Override && String(profile.JenisMitra_Override).trim())
      ? profile.JenisMitra_Override
      : jenisOtomatis;
    const pksInfo = pksInfoFor(profile.Mitra_ID);

    return {
      mitraId: profile.Mitra_ID || null,
      namaPerusahaan: nama,
      jenisMitra: jenisMitra,
      kategori: profile.Kategori || '-',
      statusAktif: profile.StatusAktif || 'Aktif',
      kota: m['Kota'] || '-',
      picUtama: m['PIC'] || '-',
      totalPembelian: adaSupplier ? pembelianAgg[key] : null,
      totalPenjualan: adaBuyer ? penjualanAgg[key] : null,
      statusPKS: pksInfo.statusPKS,
      sisaMasaPKS: pksInfo.sisaMasaPKS,
      companyVisitTerakhir: '-'
    };
  });

  // Mitra yang ditambah manual (Tambah Mitra) tapi belum ada di Master_Customer.
  const masterNames = masterList.map(m => String(m['Nama Customer'] || '').trim().toLowerCase());
  profileList.forEach(profile => {
    const nama = String(profile.NamaPerusahaan || '').trim();
    const key = nama.toLowerCase();
    if (!nama || masterNames.indexOf(key) > -1) return; // sudah ada di Master_Customer, skip
    const adaSupplier = pembelianAgg.hasOwnProperty(key);
    const adaBuyer = penjualanAgg.hasOwnProperty(key);
    const jenisOtomatis = adaSupplier && adaBuyer ? 'Supplier & Buyer' : adaSupplier ? 'Supplier' : adaBuyer ? 'Buyer' : 'Belum ada transaksi';
    const jenisMitra = (profile.JenisMitra_Override && String(profile.JenisMitra_Override).trim()) ? profile.JenisMitra_Override : jenisOtomatis;
    const pksInfo = pksInfoFor(profile.Mitra_ID);
    daftar.push({
      mitraId: profile.Mitra_ID || null,
      namaPerusahaan: nama,
      jenisMitra: jenisMitra,
      kategori: profile.Kategori || '-',
      statusAktif: profile.StatusAktif || 'Aktif',
      kota: '-',
      picUtama: '-',
      totalPembelian: adaSupplier ? pembelianAgg[key] : null,
      totalPenjualan: adaBuyer ? penjualanAgg[key] : null,
      statusPKS: pksInfo.statusPKS,
      sisaMasaPKS: pksInfo.sisaMasaPKS,
      companyVisitTerakhir: '-'
    });
  });

  return daftar;
}

/**
 * Detail satu mitra (tab Overview + ringkasan Transaksi) —
 * dicari berdasarkan nama perusahaan (dipakai sebagai identifier
 * di URL/state halaman, karena tidak semua mitra tentu sudah
 * punya Mitra_ID di MITRA_PROFILE).
 */
function getMitraDetail(namaPerusahaan) {
  const masterList = getMasterCustomerData();
  const master = masterList.find(m => String(m['Nama Customer'] || '').trim().toLowerCase() === String(namaPerusahaan || '').trim().toLowerCase());
  if (!master) return null;

  const nama = String(master['Nama Customer'] || '').trim();
  const profileList = sheetToObjects(SHEET_NAMES.MITRA_PROFILE);
  const profile = profileList.find(p => String(p.NamaPerusahaan || '').trim().toLowerCase() === nama.toLowerCase()) || {};

  const jenisMitra = (profile.JenisMitra_Override && String(profile.JenisMitra_Override).trim())
    ? profile.JenisMitra_Override
    : detectJenisMitra(nama);

  const summary = getMitraTransaksiSummary(nama);

  return {
    mitraId: profile.Mitra_ID || null,
    namaPerusahaan: nama,
    jenisMitra: jenisMitra,
    kategori: profile.Kategori || '-',
    statusAktif: profile.StatusAktif || 'Aktif',
    // Dari Master_Customer (read-only):
    alamat: master['Alamat'] || '-',
    kota: master['Kota'] || '-',
    npwp: master['NPWP'] || '-',
    picUtama: master['PIC'] || '-',
    hpPIC: master['HP PIC'] || '-',
    emailPIC: master['Email PIC'] || '-',
    phone: master['Phone'] || '-',
    namaBank: master['Nama Bank'] || '-',
    noRekening: master['No Rekening'] || '-',
    // Dari MITRA_PROFILE (read-only, field tambahan):
    website: profile.Website || '-',
    jenisProduk: profile.JenisProduk || '-',
    merekEmas: profile.MerekEmas || '-',
    lokasiPabrik: profile.LokasiPabrik || '-',
    lokasiGudang: profile.LokasiGudang || '-',
    catatan: profile.Catatan || '-',
    // Ringkasan transaksi:
    summary: summary
  };
}

/**
 * Tambah/update baris di MITRA_PROFILE (satu-satunya sheet yang
 * ditulis di fase ini — Master_Customer & sheet transaksi TIDAK
 * PERNAH ditulis oleh fungsi manapun di modul Profile Mitra ini).
 */
function saveMitraProfile(data) {
  requireRole('Supervisor');
  const sheet = getSheet(SHEET_NAMES.MITRA_PROFILE);
  const rows = sheetToObjects(SHEET_NAMES.MITRA_PROFILE);
  const nama = String(data.namaPerusahaan || '').trim();
  if (!nama) throw new Error('Nama Perusahaan wajib diisi.');

  const existingIdx = rows.findIndex(r => String(r.NamaPerusahaan || '').trim().toLowerCase() === nama.toLowerCase());
  const now = new Date();

  if (existingIdx > -1) {
    const rowNum = existingIdx + 2; // +1 header, +1 karena 0-based
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = headers.map(h => {
      if (h === 'UpdatedAt') return now;
      if (h === 'CreatedAt') return rows[existingIdx].CreatedAt;
      if (h === 'NamaPerusahaan') return nama;
      const fieldMap = { JenisMitra_Override: 'jenisMitraOverride', Kategori: 'kategori', StatusAktif: 'statusAktif', Website: 'website', JenisProduk: 'jenisProduk', MerekEmas: 'merekEmas', LokasiPabrik: 'lokasiPabrik', LokasiGudang: 'lokasiGudang', Catatan: 'catatan', Mitra_ID: 'mitraId' };
      const key = fieldMap[h];
      return key && data[key] !== undefined ? data[key] : rows[existingIdx][h];
    });
    sheet.getRange(rowNum, 1, 1, headers.length).setValues([rowData]);
    logAudit('UPDATE', 'MitraProfile', nama, 'Profil mitra diperbarui');
    return { updated: true, mitraId: rowData[headers.indexOf('Mitra_ID')] };
  } else {
    const newId = 'MTR-' + String(rows.length + 1).padStart(3, '0');
    sheet.appendRow([
      newId, nama, data.jenisMitraOverride || '', data.kategori || '', data.statusAktif || 'Aktif',
      data.website || '', data.jenisProduk || '', data.merekEmas || '', data.lokasiPabrik || '',
      data.lokasiGudang || '', data.catatan || '', now, now
    ]);
    logAudit('CREATE', 'MitraProfile', nama, 'Profil mitra baru dibuat (' + newId + ')');
    return { created: true, mitraId: newId };
  }
}

// ------------------------------------------------------------
// PROFILE MITRA — Fase 2: PKS & Legal
// Sheet MITRA_PKS (baru, sudah dibuat via setupMitraSheets()).
// Sheet existing (Master_Customer, PenjualanTransaksi, dst)
// TIDAK disentuh oleh modul ini.
// ------------------------------------------------------------

/**
 * Pastikan mitra ini punya baris di MITRA_PROFILE (bikin otomatis
 * kalau belum ada, supaya punya Mitra_ID yang stabil untuk relasi
 * ke MITRA_PKS/MITRA_PKS_DOCUMENTS/MITRA_COMPANY_VISIT). Ini SATU-
 * SATUNYA sheet yang ditulis di sini — bukan Master_Customer.
 */
function getOrCreateMitraId(namaPerusahaan) {
  const nama = String(namaPerusahaan || '').trim();
  if (!nama) throw new Error('Nama perusahaan kosong.');
  const sheet = getSheet(SHEET_NAMES.MITRA_PROFILE);
  const rows = sheetToObjects(SHEET_NAMES.MITRA_PROFILE);
  const existing = rows.find(r => String(r.NamaPerusahaan || '').trim().toLowerCase() === nama.toLowerCase());
  if (existing && existing.Mitra_ID) return existing.Mitra_ID;

  const now = new Date();
  const newId = 'MTR-' + String(rows.length + 1).padStart(3, '0');
  sheet.appendRow([newId, nama, '', '', 'Aktif', '', '', '', '', '', '', now, now]);
  return newId;
}

/**
 * Hitung status & sisa masa berlaku PKS berdasarkan tanggal hari ini.
 * Hijau (Aktif) >60 hari, Kuning (Aktif) 31-60 hari,
 * Oranye (Akan Berakhir) 1-30 hari, Merah (Berakhir) <=0 hari.
 */
function computePKSStatus(tanggalBerakhir) {
  if (!tanggalBerakhir) return { label: '-', colorClass: 'none', sisaHari: null };
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const akhir = new Date(extractDateOnly(tanggalBerakhir) + 'T00:00:00');
  const sisaHari = Math.round((akhir - today) / 86400000);

  let label, colorClass;
  if (sisaHari <= 0) { label = 'Berakhir'; colorClass = 'nonaktif'; }
  else if (sisaHari <= 30) { label = 'Akan Berakhir'; colorClass = 'oranye'; }
  else if (sisaHari <= 60) { label = 'Aktif'; colorClass = 'kuning'; }
  else { label = 'Aktif'; colorClass = 'aktif'; }

  return { label: label, colorClass: colorClass, sisaHari: sisaHari };
}

/**
 * Daftar PKS milik satu mitra (terurut, PKS paling baru duluan),
 * masing-masing sudah dilengkapi status & sisa masa berlaku
 * otomatis serta reminder mana yang sudah "kena" (H-90/60/30/14/7).
 */
function getMitraPKSList(namaPerusahaan) {
  const mitraId = getOrCreateMitraIdReadOnly(namaPerusahaan);
  if (!mitraId) return [];

  const rows = sheetToObjects(SHEET_NAMES.MITRA_PKS).filter(r => r.Mitra_ID === mitraId);
  const REMINDER_DAYS = [90, 60, 30, 14, 7];

  return rows.map(r => {
    const status = computePKSStatus(r.TanggalBerakhir);
    const mulai = r.TanggalMulai ? extractDateOnly(r.TanggalMulai) : null;
    const akhir = r.TanggalBerakhir ? extractDateOnly(r.TanggalBerakhir) : null;
    let jangkaWaktuHari = null;
    if (mulai && akhir) {
      jangkaWaktuHari = Math.round((new Date(akhir + 'T00:00:00') - new Date(mulai + 'T00:00:00')) / 86400000);
    }
    const reminderKena = status.sisaHari != null ? REMINDER_DAYS.filter(d => status.sisaHari <= d && status.sisaHari > 0) : [];

    return {
      pksId: r.PKS_ID,
      nomorPKS: r.NomorPKS || '-',
      namaPKS: r.NamaPKS || '-',
      tanggalMulai: mulai,
      tanggalBerakhir: akhir,
      jangkaWaktuHari: jangkaWaktuHari,
      catatanLegal: r.CatatanLegal || '-',
      status: status.label,
      colorClass: status.colorClass,
      sisaHari: status.sisaHari,
      reminderKena: reminderKena
    };
  }).sort((a, b) => (b.tanggalBerakhir || '').localeCompare(a.tanggalBerakhir || ''));
}

/**
 * Versi read-only dari getOrCreateMitraId — TIDAK membuat baris
 * baru, cuma cari (dipakai saat membaca/menampilkan, bukan saat
 * menyimpan, supaya tidak menumpuk baris kosong tanpa sengaja).
 */
function getOrCreateMitraIdReadOnly(namaPerusahaan) {
  const nama = String(namaPerusahaan || '').trim().toLowerCase();
  const rows = sheetToObjects(SHEET_NAMES.MITRA_PROFILE);
  const existing = rows.find(r => String(r.NamaPerusahaan || '').trim().toLowerCase() === nama);
  return existing ? existing.Mitra_ID : null;
}

/**
 * Ringkasan PKS aktif (yang paling relevan) untuk satu mitra —
 * dipakai kartu ringkas di Daftar Mitra & Overview.
 */
function getMitraPKSSummary(namaPerusahaan) {
  const list = getMitraPKSList(namaPerusahaan);
  if (!list.length) return { status: '-', sisaHari: null, colorClass: 'none' };
  // Ambil yang tanggal berakhirnya paling jauh (PKS paling relevan/terbaru).
  const aktifTerbaru = list[0];
  return { status: aktifTerbaru.status, sisaHari: aktifTerbaru.sisaHari, colorClass: aktifTerbaru.colorClass };
}

/**
 * Tambah/update satu PKS. Mitra_ID diambil/dibuat otomatis dari
 * nama perusahaan (satu-satunya penulisan ke MITRA_PROFILE terjadi
 * di sini kalau memang belum ada baris untuk mitra ini).
 */
function saveMitraPKS(data) {
  requireRole('Supervisor');
  const nama = String(data.namaPerusahaan || '').trim();
  if (!nama) throw new Error('Nama perusahaan wajib diisi.');
  const mitraId = getOrCreateMitraId(nama);

  const sheet = getSheet(SHEET_NAMES.MITRA_PKS);
  const now = new Date();

  if (data.pksId) {
    const rows = sheetToObjects(SHEET_NAMES.MITRA_PKS);
    const idx = rows.findIndex(r => r.PKS_ID === data.pksId);
    if (idx === -1) throw new Error('PKS tidak ditemukan.');
    const rowNum = idx + 2;
    sheet.getRange(rowNum, 1, 1, 8).setValues([[
      data.pksId, mitraId, data.nomorPKS || '', data.namaPKS || '',
      data.tanggalMulai || '', data.tanggalBerakhir || '', data.catatanLegal || '', rows[idx].CreatedAt
    ]]);
    logAudit('UPDATE', 'MitraPKS', nama, 'PKS diperbarui (' + data.pksId + ')');
    return { updated: true, pksId: data.pksId };
  } else {
    const existingRows = sheetToObjects(SHEET_NAMES.MITRA_PKS);
    const newId = 'PKS-' + String(existingRows.length + 1).padStart(4, '0');
    sheet.appendRow([newId, mitraId, data.nomorPKS || '', data.namaPKS || '', data.tanggalMulai || '', data.tanggalBerakhir || '', data.catatanLegal || '', now]);
    logAudit('CREATE', 'MitraPKS', nama, 'PKS baru dibuat (' + newId + ')');
    return { created: true, pksId: newId };
  }
}

// ------------------------------------------------------------
// PROFILE MITRA — Fase 3: Dokumen PKS (upload ke Google Drive)
// Sheet MITRA_PKS_DOCUMENTS (baru, sudah dibuat via setupMitraSheets()).
// File disimpan di Drive; metadata di sheet baru ini saja.
// Sheet existing tidak disentuh.
// ------------------------------------------------------------
const MITRA_DOCS_FOLDER_NAME = 'Profile Mitra - Dokumen PKS';

function getMitraDocsRootFolder() {
  const folders = DriveApp.getFoldersByName(MITRA_DOCS_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(MITRA_DOCS_FOLDER_NAME);
}

function getMitraDocsFolderFor(namaPerusahaan) {
  const root = getMitraDocsRootFolder();
  const safeName = String(namaPerusahaan || 'Tanpa Nama').replace(/[\\/:*?"<>|]/g, '-');
  const subFolders = root.getFoldersByName(safeName);
  return subFolders.hasNext() ? subFolders.next() : root.createFolder(safeName);
}

/**
 * Upload dokumen PKS baru. base64Data adalah isi file dalam base64
 * (dikirim dari client via FileReader). Metadata disimpan ke
 * MITRA_PKS_DOCUMENTS (sheet baru) — tidak menyentuh sheet lain.
 */
function uploadMitraDocument(payload) {
  requireRole('Supervisor');
  const nama = String(payload.namaPerusahaan || '').trim();
  if (!nama) throw new Error('Nama perusahaan wajib diisi.');
  if (!payload.base64Data || !payload.fileName) throw new Error('File tidak valid.');

  const mitraId = getOrCreateMitraId(nama);
  const folder = getMitraDocsFolderFor(nama);

  const bytes = Utilities.base64Decode(payload.base64Data);
  const blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream', payload.fileName);
  const file = folder.createFile(blob);

  const sheet = getSheet(SHEET_NAMES.MITRA_PKS_DOCUMENTS);
  const existingRows = sheetToObjects(SHEET_NAMES.MITRA_PKS_DOCUMENTS);
  const newId = 'DOC-' + String(existingRows.length + 1).padStart(4, '0');
  const now = new Date();
  const user = getCurrentUser();

  sheet.appendRow([
    newId, mitraId, payload.pksId || '', payload.namaDokumen || payload.fileName, payload.jenisDokumen || 'Dokumen Pendukung',
    payload.versi || '1.0', payload.tanggalDokumen || '', now, (user && user.Nama) || (user && user.Email) || 'Admin',
    file.getId(), 'Aktif', payload.catatan || ''
  ]);

  logAudit('CREATE', 'MitraDocument', nama, 'Dokumen "' + (payload.namaDokumen || payload.fileName) + '" diupload (' + newId + ')');
  return { created: true, docId: newId, driveFileId: file.getId(), viewUrl: file.getUrl() };
}

/**
 * Daftar dokumen milik satu mitra, lengkap dengan link Drive.
 */
function getMitraDocuments(namaPerusahaan) {
  const mitraId = getOrCreateMitraIdReadOnly(namaPerusahaan);
  if (!mitraId) return [];

  const rows = sheetToObjects(SHEET_NAMES.MITRA_PKS_DOCUMENTS).filter(r => r.Mitra_ID === mitraId);
  return rows.map(r => {
    let viewUrl = '', downloadUrl = '', fileExists = true;
    try {
      const file = DriveApp.getFileById(r.DriveFileID);
      viewUrl = file.getUrl();
      downloadUrl = 'https://drive.google.com/uc?export=download&id=' + r.DriveFileID;
    } catch (e) {
      fileExists = false;
    }
    return {
      docId: r.Doc_ID,
      pksId: r.PKS_ID || '',
      namaDokumen: r.NamaDokumen || '-',
      jenisDokumen: r.JenisDokumen || '-',
      versi: r.Versi || '-',
      tanggalDokumen: r.TanggalDokumen ? extractDateOnly(r.TanggalDokumen) : '-',
      tanggalUpload: r.TanggalUpload ? extractDateOnly(r.TanggalUpload) : '-',
      pengunggah: r.Pengunggah || '-',
      statusDokumen: r.StatusDokumen || 'Aktif',
      catatan: r.Catatan || '-',
      viewUrl: viewUrl,
      downloadUrl: downloadUrl,
      fileExists: fileExists
    };
  }).sort((a, b) => (b.tanggalUpload || '').localeCompare(a.tanggalUpload || ''));
}

/**
 * Edit metadata dokumen (TIDAK menyentuh file-nya di Drive).
 */
function updateMitraDocumentInfo(data) {
  requireRole('Supervisor');
  const sheet = getSheet(SHEET_NAMES.MITRA_PKS_DOCUMENTS);
  const rows = sheetToObjects(SHEET_NAMES.MITRA_PKS_DOCUMENTS);
  const idx = rows.findIndex(r => r.Doc_ID === data.docId);
  if (idx === -1) throw new Error('Dokumen tidak ditemukan.');
  const rowNum = idx + 2;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updated = Object.assign({}, rows[idx], {
    NamaDokumen: data.namaDokumen || rows[idx].NamaDokumen,
    JenisDokumen: data.jenisDokumen || rows[idx].JenisDokumen,
    Versi: data.versi || rows[idx].Versi,
    TanggalDokumen: data.tanggalDokumen || rows[idx].TanggalDokumen,
    Catatan: data.catatan !== undefined ? data.catatan : rows[idx].Catatan
  });
  const rowValues = headers.map(h => updated[h]);
  sheet.getRange(rowNum, 1, 1, headers.length).setValues([rowValues]);
  logAudit('UPDATE', 'MitraDocument', data.docId, 'Info dokumen diperbarui');
  return { updated: true };
}

/**
 * Arsipkan dokumen (ubah status saja, file di Drive TIDAK dihapus).
 */
function archiveMitraDocument(docId) {
  requireRole('Supervisor');
  const sheet = getSheet(SHEET_NAMES.MITRA_PKS_DOCUMENTS);
  const rows = sheetToObjects(SHEET_NAMES.MITRA_PKS_DOCUMENTS);
  const idx = rows.findIndex(r => r.Doc_ID === docId);
  if (idx === -1) throw new Error('Dokumen tidak ditemukan.');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('StatusDokumen') + 1;
  sheet.getRange(idx + 2, statusCol).setValue('Diarsipkan');
  logAudit('ARCHIVE', 'MitraDocument', docId, 'Dokumen diarsipkan');
  return { archived: true };
}

/**
 * DIAGNOSTIK — cek 5 baris terakhir PenjualanTransaksi &
 * PembelianTransaksi, bandingkan format Tanggal-nya dengan
 * tanggal hari ini yang dipakai sistem. Cuma membaca.
 */
function diagnosaTransaksiHariIni() {
  const today = formatDateOnly(new Date());
  Logger.log('Tanggal hari ini (yang dicari sistem): ' + today);
  Logger.log('===========================================');

  const pj = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  Logger.log('Total baris PenjualanTransaksi: ' + pj.length);
  Logger.log('--- 5 baris TERAKHIR (raw) ---');
  pj.slice(-5).forEach((row, i) => {
    Logger.log('Baris ' + i + ': Tanggal="' + row.Tanggal + '" (tipe: ' + typeof row.Tanggal + ') | extractDateOnly hasil: "' + extractDateOnly(row.Tanggal) + '" | Pembeli=' + row.Pembeli + ' | TotalHarga=' + row.TotalHarga);
  });

  Logger.log('===========================================');
  const pb = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL);
  Logger.log('Total baris PembelianTransaksi: ' + pb.length);
  Logger.log('--- 5 baris TERAKHIR (raw) ---');
  pb.slice(-5).forEach((row, i) => {
    Logger.log('Baris ' + i + ': Tanggal="' + row.Tanggal + '" (tipe: ' + typeof row.Tanggal + ') | extractDateOnly hasil: "' + extractDateOnly(row.Tanggal) + '" | Seller=' + row.Seller + ' | TotalHarga=' + row.TotalHarga);
  });
}

/**
 * DIAGNOSTIK KHUSUS — dump SEMUA baris PenjualanTransaksi untuk satu
 * tanggal tertentu (ID, NomorInvoice, Qty, Harga, Pembeli), supaya
 * bisa dibandingkan baris-per-baris dengan sheet sumber (F.1.Input_Trans_Jual)
 * untuk mencari duplikat Nomor Invoice yang ke-sync dua kali dengan
 * harga/nominal berbeda (biasanya karena harga di source diedit
 * setelah sync pertama jalan). GANTI tanggal di bawah kalau perlu.
 */
function diagnosaPenjualanTanggalTertentu() {
  const TANGGAL_CEK = '2026-07-29'; // ganti tanggal di sini kalau mau cek tanggal lain
  const rows = getPenjualanTransaksiToday(TANGGAL_CEK);
  Logger.log('Tanggal dicek: ' + TANGGAL_CEK);
  Logger.log('Total baris ditemukan: ' + rows.length);
  Logger.log('===========================================');

  const byInvoice = {};
  rows.forEach(r => {
    const key = String(r.NomorInvoice || '').trim();
    if (!byInvoice[key]) byInvoice[key] = [];
    byInvoice[key].push(r);
  });

  rows.forEach(r => {
    Logger.log('ID=' + r.ID + ' | NomorInvoice="' + r.NomorInvoice + '" | Qty=' + r.Qty + ' | HargaPerGram=' + r.HargaPerGram + ' | TotalHarga=' + r.TotalHarga + ' | Pembeli=' + r.Pembeli + ' | Catatan=' + r.Catatan);
  });

  Logger.log('===========================================');
  Logger.log('--- CEK DUPLIKAT NOMOR INVOICE ---');
  let adaDuplikat = false;
  Object.keys(byInvoice).forEach(key => {
    if (byInvoice[key].length > 1) {
      adaDuplikat = true;
      Logger.log('DOBEL! Nomor Invoice "' + key + '" muncul ' + byInvoice[key].length + ' kali:');
      byInvoice[key].forEach(r => Logger.log('   -> ID=' + r.ID + ' | Qty=' + r.Qty + ' | Harga=' + r.HargaPerGram + ' | Catatan=' + r.Catatan));
    }
  });
  if (!adaDuplikat) Logger.log('Tidak ditemukan Nomor Invoice yang dobel persis. Kemungkinan selisih dari sumber lain (misal baris ter-skip atau format Nomor Invoice sedikit beda).');

  const totalGram = rows.reduce((s, r) => s + Number(r.Qty || 0), 0);
  Logger.log('===========================================');
  Logger.log('Total Gramasi (dari database aplikasi): ' + totalGram + ' gram');
}

// ------------------------------------------------------------
// SYNC TRANSAKSI BARU — menarik baris baru dari sheet sumber
// (F.1.Input_Trans_Jual / F.2.Input_Trans_Beli di spreadsheet
// Bullion Trading Process) ke PenjualanTransaksi/PembelianTransaksi
// TANPA dobel data (dicek dari Nomor Invoice / Nomor PO yang
// sudah ada). Sheet sumber tetap read-only, tidak pernah ditulis.
// ------------------------------------------------------------

/**
 * Sync Penjualan: cari baris baru di F.1.Input_Trans_Jual yang
 * Nomor Invoice-nya BELUM ada di PenjualanTransaksi, lalu import.
 * SEKALIGUS mengecek baris yang SUDAH ada — kalau datanya di sheet
 * sumber berubah (harga, qty, pembeli, dll, misalnya karena diedit
 * setelah sync pertama jalan), baris di database ikut diperbarui. Update
 * otomatis ini HANYA berlaku untuk baris hasil sync/import sebelumnya
 * (Catatan diawali "Sync otomatis" atau "Diimpor dari spreadsheet
 * lama") — baris yang diinput manual lewat form TIDAK akan tertimpa.
 */
function syncTransaksiBaruPenjualan() {
  requireRole('Supervisor');
  const sourceSs = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sourceSheet = sourceSs.getSheetByName('F.1.Input_Trans_Jual');
  if (!sourceSheet) throw new Error('Sheet "F.1.Input_Trans_Jual" tidak ditemukan.');

  const values = sourceSheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['tanggal', 'invoice', 'pembeli', 'qty', 'harga']);
  if (headerRowIdx === -1) throw new Error('Baris header tidak terdeteksi di F.1.Input_Trans_Jual.');

  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const rows = values.slice(headerRowIdx + 1);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal'], ['pengambilan', 'penjemputan']),
    nomorInvoice: findColByKeywords(headerMap, ['nomor', 'invoice']),
    hargaPerGram: findColByKeywords(headerMap, ['harga'], ['total', 'after', 'simulasi']),
    qty: findColByKeywords(headerMap, ['qty']),
    pembeli: findColByKeywords(headerMap, ['pembeli']),
    denominasi: findColByKeywords(headerMap, ['denominasi']),
    merk: findColByKeywords(headerMap, ['merk']),
    status: findColByKeywords(headerMap, ['status']),
    paymentTime: findColByKeywords(headerMap, ['payment']),
    bank: findColByKeywords(headerMap, ['bank']),
    tglAmbil: findColByKeywords(headerMap, ['pengambilan']),
    estimasi: findColByKeywords(headerMap, ['estimasi'])
  };
  if (col.tanggal === -1 || col.nomorInvoice === -1 || col.hargaPerGram === -1 || col.qty === -1) {
    throw new Error('Kolom wajib tidak ditemukan di F.1.Input_Trans_Jual.');
  }

  const existingRows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const existingByInvoice = {};
  existingRows.forEach(r => { existingByInvoice[String(r.NomorInvoice || '').trim()] = r; });

  const targetSh = getSheet(SHEET_NAMES.PENJUALAN_FULL);
  const user = getCurrentUser();
  const firstId = getNextId(SHEET_NAMES.PENJUALAN_FULL, 'ID', 'PJL');
  let idNum = parseInt(firstId.split('-')[1], 10);
  const outputRows = [];
  const now = new Date();
  let skipped = 0, updated = 0;
  const autoSyncCatatan = ['Sync otomatis dari sheet sumber', 'Diimpor dari spreadsheet lama'];

  rows.forEach(row => {
    const tanggal = getByCol(row, col.tanggal);
    const nomorInvoice = String(getByCol(row, col.nomorInvoice) || '').trim();
    const hargaPerGram = Number(getByCol(row, col.hargaPerGram)) || 0;
    const qty = Number(getByCol(row, col.qty)) || 0;

    // Validasi baris nyata (bukan baris kosong/marker seperti "end", "-", dst).
    if (!tanggal || !nomorInvoice || hargaPerGram <= 0 || qty <= 0) { skipped++; return; }

    const existing = existingByInvoice[nomorInvoice];
    const newPembeli = getByCol(row, col.pembeli);
    const newDenominasi = getByCol(row, col.denominasi);
    const newMerk = getByCol(row, col.merk);
    const newStatus = getByCol(row, col.status) || 'V';
    const newPaymentTime = getByCol(row, col.paymentTime);
    const newBank = getByCol(row, col.bank);
    const newTglAmbil = getByCol(row, col.tglAmbil);
    const newEstimasi = getByCol(row, col.estimasi);

    if (existing) {
      // Baris sudah ada — cek apakah datanya berubah di sumber, dan
      // hanya update kalau baris ini memang hasil sync/import (bukan
      // input manual lewat form) supaya tidak menimpa data manual.
      const bolehAutoUpdate = autoSyncCatatan.indexOf(String(existing.Catatan || '')) !== -1;
      if (!bolehAutoUpdate) return;

      const berubah =
        Number(existing.HargaPerGram) !== hargaPerGram ||
        Number(existing.Qty) !== qty ||
        String(existing.Pembeli || '') !== String(newPembeli || '') ||
        String(existing.Denominasi || '') !== String(newDenominasi || '') ||
        String(existing.Merk || '') !== String(newMerk || '') ||
        String(existing.Status || '') !== String(newStatus || '') ||
        String(existing.Bank || '') !== String(newBank || '');

      if (berubah) {
        const totalHarga = hargaPerGram * qty;
        const pph = totalHarga * PPH_PASAL_22_RATE;
        const totalAfterTax = totalHarga + pph;
        updateRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', existing.ID, {
          HargaPerGram: hargaPerGram, Qty: qty, TotalHarga: totalHarga, PphPasal22: pph, TotalHargaAfterTax: totalAfterTax,
          Pembeli: newPembeli, Denominasi: newDenominasi, Merk: newMerk, Status: newStatus,
          PaymentTime: newPaymentTime, Bank: newBank, TanggalPengambilan: newTglAmbil, EstimasiPenjemputan: newEstimasi
        });
        updated++;
      }
      return;
    }

    const id = 'PJL-' + String(idNum).padStart(3, '0');
    idNum++;
    const totalHarga = hargaPerGram * qty;
    const pph = totalHarga * PPH_PASAL_22_RATE;
    const totalAfterTax = totalHarga + pph;

    outputRows.push([
      id, tanggal, nomorInvoice, hargaPerGram, qty,
      totalHarga, pph, totalAfterTax,
      newPembeli, newDenominasi, newMerk, newStatus,
      newPaymentTime, newBank, newTglAmbil, newEstimasi,
      'Sync otomatis dari sheet sumber', now, user.Email
    ]);
    existingByInvoice[nomorInvoice] = { NomorInvoice: nomorInvoice }; // cegah baris sama dobel di run yang sama
  });

  if (outputRows.length) {
    const startRow = targetSh.getLastRow() + 1;
    targetSh.getRange(startRow, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }

  logAudit('SYNC', 'PenjualanTransaksi', '-', outputRows.length + ' baris baru, ' + updated + ' baris diperbarui dari F.1.Input_Trans_Jual');
  return { imported: outputRows.length, updated: updated };
}

/**
 * Sync Pembelian: cari baris baru di F.2.Input_Trans_Beli yang
 * Nomor PO-nya BELUM ada di PembelianTransaksi, lalu import.
 * SEKALIGUS mengecek baris yang SUDAH ada — kalau datanya di sheet
 * sumber berubah (harga, qty, seller, dll), baris di database ikut
 * diperbarui. Update otomatis ini HANYA berlaku untuk baris hasil
 * sync/import sebelumnya (Catatan diawali "Sync otomatis" atau
 * "Diimpor dari spreadsheet lama") — baris input manual tidak tertimpa.
 */
function syncTransaksiBaruPembelian() {
  requireRole('Supervisor');
  const sourceSs = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sourceSheet = sourceSs.getSheetByName('F.2.Input_Trans_Beli');
  if (!sourceSheet) throw new Error('Sheet "F.2.Input_Trans_Beli" tidak ditemukan.');

  const values = sourceSheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['tanggal', 'purchase', 'seller', 'qty', 'harga']);
  if (headerRowIdx === -1) throw new Error('Baris header tidak terdeteksi di F.2.Input_Trans_Beli.');

  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const rows = values.slice(headerRowIdx + 1);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    nomorPO: findColByKeywords(headerMap, ['nomor', 'purchase']),
    seller: findColByKeywords(headerMap, ['seller']),
    qty: findColByKeywords(headerMap, ['qty']),
    afterDiskon: findColByKeywords(headerMap, ['after', 'diskon']),
    harga: findColByKeywords(headerMap, ['harga'], ['after', 'total', 'simulasi']),
    invoice: findColByKeywords(headerMap, ['invoice']),
    catatan: findColByKeywords(headerMap, ['catatan'])
  };
  if (col.tanggal === -1 || col.nomorPO === -1 || col.qty === -1) {
    throw new Error('Kolom wajib tidak ditemukan di F.2.Input_Trans_Beli.');
  }

  const existingRows = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL);
  const existingByPO = {};
  existingRows.forEach(r => { existingByPO[String(r.NomorPO || '').trim()] = r; });

  const targetSh = getSheet(SHEET_NAMES.PEMBELIAN_FULL);
  const user = getCurrentUser();
  const firstId = getNextId(SHEET_NAMES.PEMBELIAN_FULL, 'ID', 'PBL');
  let idNum = parseInt(firstId.split('-')[1], 10);
  const outputRows = [];
  const now = new Date();
  let skipped = 0, updated = 0;
  const autoSyncCatatan = ['Sync otomatis dari sheet sumber', 'Diimpor dari spreadsheet lama'];

  rows.forEach(row => {
    const tanggal = getByCol(row, col.tanggal);
    const nomorPO = String(getByCol(row, col.nomorPO) || '').trim();
    const qty = Number(getByCol(row, col.qty)) || 0;
    const harga = Number(getByCol(row, col.harga)) || 0;
    const afterDiskonRaw = col.afterDiskon !== -1 ? Number(getByCol(row, col.afterDiskon)) : 0;
    const afterDiskon = afterDiskonRaw || harga;
    const newSeller = getByCol(row, col.seller);
    const newInvoiceRef = getByCol(row, col.invoice);
    const newCatatanSumber = getByCol(row, col.catatan) || 'Sync otomatis dari sheet sumber';

    if (!tanggal || !nomorPO || qty <= 0 || afterDiskon <= 0) { skipped++; return; }

    const existing = existingByPO[nomorPO];
    if (existing) {
      const bolehAutoUpdate = autoSyncCatatan.indexOf(String(existing.Catatan || '')) !== -1;
      if (!bolehAutoUpdate) return;

      const berubah =
        Number(existing.Harga) !== harga ||
        Number(existing.AfterDiskon) !== afterDiskon ||
        Number(existing.Qty) !== qty ||
        String(existing.Seller || '') !== String(newSeller || '') ||
        String(existing.InvoiceRef || '') !== String(newInvoiceRef || '');

      if (berubah) {
        const totalHarga = afterDiskon * qty;
        const pajakWapu = totalHarga * PAJAK_WAPU_RATE;
        const totalInclude = totalHarga - pajakWapu;
        updateRowById(SHEET_NAMES.PEMBELIAN_FULL, 'ID', existing.ID, {
          Harga: harga, AfterDiskon: afterDiskon, Qty: qty,
          TotalHarga: totalHarga, PajakWapu: pajakWapu, TotalHargaIncludePajak: totalInclude,
          Seller: newSeller, InvoiceRef: newInvoiceRef
        });
        updated++;
      }
      return;
    }

    const id = 'PBL-' + String(idNum).padStart(3, '0');
    idNum++;
    const totalHarga = afterDiskon * qty;
    const pajakWapu = totalHarga * PAJAK_WAPU_RATE;
    const totalInclude = totalHarga - pajakWapu;

    outputRows.push([
      id, tanggal, nomorPO, newSeller, qty,
      harga, afterDiskon, totalHarga, pajakWapu, totalInclude,
      newInvoiceRef, newCatatanSumber,
      now, user.Email
    ]);
    existingByPO[nomorPO] = { NomorPO: nomorPO }; // cegah baris sama dobel di run yang sama
  });

  if (outputRows.length) {
    const startRow = targetSh.getLastRow() + 1;
    targetSh.getRange(startRow, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }

  logAudit('SYNC', 'PembelianTransaksi', '-', outputRows.length + ' baris baru, ' + updated + ' baris diperbarui dari F.2.Input_Trans_Beli');
  return { imported: outputRows.length, updated: updated };
}

/**
 * Dipanggil dari tombol "Sync Transaksi" di halaman Settings —
 * jalankan sync Penjualan & Pembelian sekaligus.
 */
function syncTransaksiBaru() {
  const pj = syncTransaksiBaruPenjualan();
  const pb = syncTransaksiBaruPembelian();
  return { penjualan: pj.imported, penjualanUpdate: pj.updated, pembelian: pb.imported, pembelianUpdate: pb.updated };
}

/**
 * JALANKAN INI SEKALI dari editor (pilih function ini di dropdown,
 * lalu Run) untuk memasang trigger otomatis — sync akan jalan
 * sendiri setiap 30 menit tanpa perlu klik tombol lagi.
 * Aman dijalankan berkali-kali (trigger lama dihapus dulu supaya
 * tidak dobel-dobel jalan).
 */
function installSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncTransaksiBaru') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncTransaksiBaru')
    .timeBased()
    .everyMinutes(30)
    .create();
  Logger.log('Trigger otomatis terpasang — syncTransaksiBaru() akan jalan sendiri tiap 30 menit.');
}

/**
 * Jalankan ini kalau mau MATIKAN sync otomatis (kembali ke manual
 * lewat tombol saja).
 */
function uninstallSyncTrigger() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncTransaksiBaru') { ScriptApp.deleteTrigger(t); removed++; }
  });
  Logger.log(removed + ' trigger sync otomatis dihapus.');
}

// ------------------------------------------------------------
// PENETAPAN HARGA (BOBOT SUPPLIER) — mengikuti prinsip Gold
// Studio: tiap supplier punya Bobot (%) yang totalnya harus 100%
// di antara supplier yang aktif. Harga per supplier dikalikan
// bobotnya masing-masing untuk mendapat System Reference Price
// (rata-rata tertimbang), lalu Harga Jual = Ref x (1 + Margin%).
// TIDAK ADA approval — begitu disubmit, langsung tersimpan
// sebagai harga resmi (beda dengan Gold Studio yang pakai alur
// submit -> approve).
// ------------------------------------------------------------

/**
 * Menyamakan baris BobotSupplier dengan daftar Supplier terkini —
 * supplier baru yang belum punya baris bobot otomatis ditambahkan
 * (bobot awal 0, status Aktif). Tidak menghapus baris bobot milik
 * supplier yang sudah tidak ada (biar histori tidak hilang).
 */
// Indikator market (Rupiah) yang boleh ikut dihitung sebagai komponen
// bobot di Penetapan Harga, sama seperti supplier biasa — harganya
// otomatis terisi live (bukan input manual). LBMA H-1 & Kitco H-1
// SENGAJA TIDAK dimasukkan ke sini karena dalam USD ($) — kalau
// dicampur langsung ke rata-rata tertimbang Rupiah/gram hasilnya
// salah secara matematis. Versi Rupiah-nya (Rp LBMA/Rp Kitco) yang
// dipakai sebagai gantinya.
const MARKET_INDICATOR_DEFS = [
  { id: 'MKT-SPOTGOLD', nama: 'Spot Goldprice', key: 'spotGold' },
  { id: 'MKT-XAUIDRK', nama: 'XAUIDRK', key: 'xauIdr' },
  { id: 'MKT-KURSJUAL', nama: 'Kurs Jual', key: 'kursJual' },
  { id: 'MKT-KURSBELI', nama: 'Kurs Beli', key: 'kursBeli' },
  { id: 'MKT-RPLBMA', nama: 'Rp LBMA', key: 'rpLbma' },
  { id: 'MKT-RPKITCO', nama: 'Rp Kitco', key: 'rpKitco' }
];

function ensureBobotRows_() {
  const suppliers = getSuppliers();
  const bobotRows = sheetToObjects(SHEET_NAMES.BOBOT);
  const existingIds = new Set(bobotRows.map(b => String(b.SupplierID)));
  const sh = getSheet(SHEET_NAMES.BOBOT);

  const toAppend = [];
  suppliers.forEach(s => {
    if (!existingIds.has(String(s.SupplierID))) {
      toAppend.push([s.SupplierID, s.Nama, 0, 'true']);
    }
  });
  MARKET_INDICATOR_DEFS.forEach(m => {
    if (!existingIds.has(m.id)) {
      toAppend.push([m.id, m.nama, 0, 'false']);
    }
  });
  if (toAppend.length) {
    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, toAppend.length, 4).setValues(toAppend);
  }
}

/**
 * Daftar bobot supplier (otomatis sinkron dengan daftar Supplier
 * terkini). Dipakai halaman Penetapan Harga untuk konfigurasi
 * bobot maupun untuk input harga harian.
 */
function getBobotSupplierList() {
  ensureBobotRows_();
  const marketIds = new Set(MARKET_INDICATOR_DEFS.map(m => m.id));
  const marketKeyById = {};
  MARKET_INDICATOR_DEFS.forEach(m => { marketKeyById[m.id] = m.key; });

  return sheetToObjects(SHEET_NAMES.BOBOT).map(b => ({
    supplierId: b.SupplierID,
    nama: b.Nama,
    bobot: Number(b.Bobot) || 0,
    aktif: String(b.Aktif) === 'true',
    isMarket: marketIds.has(String(b.SupplierID)),
    marketKey: marketKeyById[String(b.SupplierID)] || null
  }));
}

/**
 * Menyimpan konfigurasi bobot. data = [{ supplierId, bobot, aktif }].
 * Total bobot dari yang aktif WAJIB = 100 (toleransi 0.01) —
 * kalau tidak, ditolak dan tidak ada yang tersimpan.
 */
function saveBobotSupplierList(data) {
  requireRole('Supervisor');
  const aktifList = data.filter(b => b.aktif);
  const total = aktifList.reduce((s, b) => s + (Number(b.bobot) || 0), 0);
  if (Math.abs(total - 100) > 0.01) {
    throw new Error('Total bobot supplier aktif harus 100%. Saat ini: ' + total.toFixed(2) + '%');
  }

  const sh = getSheet(SHEET_NAMES.BOBOT);
  const rows = sheetToObjects(SHEET_NAMES.BOBOT);
  data.forEach(b => {
    const idx = rows.findIndex(r => String(r.SupplierID) === String(b.supplierId));
    if (idx === -1) return;
    const rowNum = idx + 2;
    sh.getRange(rowNum, 3).setValue(Number(b.bobot) || 0);
    sh.getRange(rowNum, 4).setValue(b.aktif ? 'true' : 'false');
  });

  logAudit('UPDATE', 'BobotSupplier', '-', 'Total bobot diset ulang: ' + total.toFixed(2) + '%');
  return { ok: true };
}

/**
 * Menghitung System Reference Price (rata-rata tertimbang) dan
 * Harga Jual dari daftar { supplierId, harga, diskon, bobot, aktif }
 * TANPA menyimpan apapun — dipakai untuk preview real-time di
 * frontend sebelum disubmit. Harga Base = Harga - Diskon, dan
 * Harga Base inilah yang dipakai sebagai basis rata-rata tertimbang
 * (bukan harga mentahnya), supaya selisih diskon (kalau ada, misal
 * HRTA) ikut diperhitungkan sebelum dirata-ratakan.
 */
function previewPenetapanHargaBobot(hargaSupplier, margin) {
  const withBase = hargaSupplier.map(h => {
    const hargaBase = Number(h.harga || 0) - Number(h.diskon || 0);
    return Object.assign({}, h, { hargaBase: hargaBase });
  });
  const aktif = withBase.filter(h => h.aktif && Number(h.hargaBase) > 0);
  const totalBobot = aktif.reduce((s, h) => s + (Number(h.bobot) || 0), 0);
  if (!totalBobot) return { sysRefPrice: 0, hargaJual: 0, totalBobot: 0 };

  let weighted = 0;
  aktif.forEach(h => { weighted += Number(h.hargaBase) * (Number(h.bobot) / totalBobot); });
  const ref = Math.round(weighted);
  const m = Number(margin) || 0;
  const hargaJual = Math.round(ref * (1 + m / 100));
  return { sysRefPrice: ref, hargaJual: hargaJual, totalBobot: totalBobot };
}

/**
 * Submit Penetapan Harga — TANPA approval, langsung tersimpan
 * sebagai harga resmi begitu disubmit. data = {
 *   hargaSupplier: [{ supplierId, nama, harga, bobot, aktif }],
 *   margin: number
 * }
 * Validasi: total bobot dari supplier yang diisi harganya (aktif
 * & harga > 0) harus = 100% (toleransi 0.01) sebelum bisa disimpan.
 */
function submitPenetapanHargaBobot(data) {
  requireRole('Supervisor');
  const items = (data.hargaSupplier || []).map(h => {
    const hargaBase = Number(h.harga || 0) - Number(h.diskon || 0);
    return Object.assign({}, h, { hargaBase: hargaBase });
  });
  const aktif = items.filter(h => h.aktif && Number(h.hargaBase) > 0);
  if (!aktif.length) throw new Error('Masukkan minimal satu harga supplier.');

  const totalBobot = aktif.reduce((s, h) => s + (Number(h.bobot) || 0), 0);
  if (Math.abs(totalBobot - 100) > 0.1) {
    throw new Error('Total bobot supplier yang diisi harganya harus 100%. Saat ini: ' + totalBobot.toFixed(2) + '%');
  }

  const margin = Number(data.margin) || 0;
  const preview = previewPenetapanHargaBobot(data.hargaSupplier, margin);
  const ref = preview.sysRefPrice;
  const hargaJual = preview.hargaJual;

  const sh = getSheet(SHEET_NAMES.PENETAPAN_BOBOT);
  const user = getCurrentUser();
  const id = getNextId(SHEET_NAMES.PENETAPAN_BOBOT, 'ID', 'PHB');
  const now = new Date();

  sh.appendRow([
    id, formatDateOnly(now), now, user.Email,
    JSON.stringify(items), ref, margin, hargaJual, now
  ]);

  logAudit('CREATE', 'PenetapanHargaBobot', id, 'Sys Ref ' + ref + ', Margin ' + margin + '%, Harga Jual ' + hargaJual);
  return { ok: true, id: id, sysRefPrice: ref, margin: margin, hargaJual: hargaJual };
}

/**
 * Riwayat Penetapan Harga (terbaru dulu), dengan detail per
 * supplier di-parse dari JSON. Dipakai tabel riwayat di halaman
 * Penetapan Harga.
 */
function getPenetapanHargaBobotRecent(limit) {
  limit = limit || 30;
  const rows = sheetToObjects(SHEET_NAMES.PENETAPAN_BOBOT)
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt))
    .slice(0, limit);

  return rows.map(r => {
    let detail = [];
    try { detail = JSON.parse(r.HargaSupplierJSON || '[]'); } catch (e) {}
    return {
      id: r.ID,
      tanggal: r.Tanggal,
      waktu: r.Waktu,
      inputBy: r.InputBy,
      sysRefPrice: Number(r.SysRefPrice) || 0,
      margin: Number(r.Margin) || 0,
      hargaJual: Number(r.HargaJual) || 0,
      detail: detail
    };
  });
}

/**
 * Penetapan Harga (bobot) terbaru — dipakai kartu ringkas di
 * Dashboard/halaman lain kalau diperlukan.
 */
function getPenetapanHargaBobotLatest() {
  const list = getPenetapanHargaBobotRecent(1);
  return list.length ? list[0] : null;
}

// ============================================================
// LOGO PEGADAIAN — disimpan sebagai base64 supaya bisa disisipkan
// langsung ke email (inline image via cid), tanpa perlu hosting
// file terpisah. Dipakai di header form Confirmation Deal Email.
// ============================================================
const CONFIRMATION_DEAL_LOGO_BASE64 = '/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/2wBDAQMDAwQDBAgEBAgQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBD/wAARCABdAL8DASIAAhEBAxEB/8QAHQAAAgIDAQEBAAAAAAAAAAAAAAgHCQQFBgEDAv/EAEQQAAEDAwIEAwYEAwQHCQAAAAECAwQFBhEABwgSITEJE0EUIjJRYXEVQoGRI2KhUnOCkhYXJTNysfEYOENjZKKytMH/xAAcAQABBAMBAAAAAAAAAAAAAAAAAQIDBgQFBwj/xAAyEQABAwMDAgMGBQUAAAAAAAABAAIDBAUREiExQVEGE2EHFCIycYEVIzNCoVJikbHB/9oADAMBAAIRAxEAPwC09SwgZV0Hz1Hl/bnRqC4umUdCZU5OAsk+4zn5/M/TXRX5cQtu2pNQQf45AbYH86ug/bqf00uDji33FPOrK1rJUok5JJ7/ANdcu8eeLpbK0UdJ+o4bn+kJwC2FVuOu1l1TtRqkh4KyC3zcqMfIJGB/TWvbWtpQW24ptYOQtBIIPz6YI/TX50a4HPVz1MnmyOJd3KCcLoIu7lUsaGupViqB+nRhlxEpRUQnsAlXfOewOdcVcvErX7+aUbQfVSKaTyYQB7SvH9s/l+ycffUJ75XY9VbhFtR3CIVLI50g/G+QCSf+EEAfUHXF2pWnaLWW1En2d8+U8M9ACeh/Q9ddqsLLm2y/mTHLhkDsO2eVyK9ePHC5fh8W0ecE9cqXpc2ZPd8+dMfku4xzvOFasfLJ9NZtEui47bfTIoValwlJOeVtw8iu3dJ909vlrWaxKtUUUqnSJ7ic+SjIHzV2SP1OtRE6Xzg6M4dnlZr6v3cGfV9SpYqHGHUrat9+nVCiM1KvFA9ncbX5bZz+Z5Ppj+Xv/L30vV4b27mX4+p24LtnoaJymLDdMdhA+QSgjP8AiKj9dcgn2ur1RHM4FSZrqUFS3EoHMo4GVKwEjJAyTgDUjXDtZa23Mamu7hV6oSpVYipnRItDZQtsslRGTIcISTkdQlPqCD1xrp9I6rkgDZHZx1VCr71efEGtzH6YWeuBj1PVRozJkMPJkMvusupPMHW1qStJ+YIIOdSft5xJbobfSGmo9efrVLSAhUCorK8p9eRz40H9SPodaFFQ2elIUw9ad1RD+SQ1WGXl/QltbSU9foo41nyNpmqpZ8zcCyq+3No0J1TDzNRQIkptwJ5ykEktu9P7JBPYJ7ZlijlYdUblrLcyvpHmagly4bnGf5B5T47P7sWzu7b4q9EdS1KjhKJsJ1Q82M4fQ/MHHQ9j++pIaOUZ6fpqrjZ/cuobXX7T7qjuq9mDiWKg0D0ejKOF5HqUj3k/UffVoFMkR5cFqXEdS6y+kOIWk5CgRkEfcY1uqOo89u/IXb/B/iE36kLpf1GbHsfUDplZWjRo1mK3o0aNGhCNGjRoQjRo0aEI0aNGhCive5xwUqmt9QlclRPyOEnUQanfd2hOVK1jLYSVOQHA+QP7J6K/YHP6agjr6j/przf7SoJIrsXvGzgMHuno16kZUAR66816O+ueM2cCo5N2kBKFckhyTclSkvqK3Fz3+ZR9cLIzrWKOFKI9DnXXbpUBy375qDHlcjEpz2tgj4SlfU4+ysj9Nc3TYL1TqLMBlJKnnAnI9B+Y/oNejaCeOS3xyN+XSvJ1ypJW3Z8JHxaj/tTHDdW9CjvLVzKcaQpX3Kc51z24S1ot8AKwFyEJI+YwT/z10qUJbSEIAASAkD5Adtai7qcqp0CSy22VuNgOoAHUlJyR+2qnSyNFY154yum18L3W10QHxaVxFi2q1ddwxYc55KKc2hyXUXep8qG0nndV278oIHf3iOmddRPuCbulQ7gjT8idS311uktAZ8uIEpbfYSAeyWksrA+TSu+cjqOH/bmhbk27ctDRecag1qousRghTYccXCT/ABFhCSpOQtwJCsZ6NYPxHUnUnhHiWLWoN1Sd1WUs090Ov+bBShDjXULbUpThASpJUk5+eupwRSOiGjgrQW2y1jqFhiALH51/EOuwGP7eUosZp9+Q1GiAvOvLDTKEjKlrV0AA+eSNSt+JQI9ysbRyp7aaA3CNvyXgrLaai4sOKlf4ZQQkKP5E9CB119NraDYzm9M+nJuqHDp8BU1dDnOqAZcfBxHI5iArl5uYZPXkGpRTwNMSEJkK3RJWv31KNPyVE+p/ifr/ANdEEEmCYsc91HaLHUsiLqfDiXYduOAf+pWZdLm0ybKp1RYLUiK65FfaPdK0kpI/cHVnexEx6pbO2fMkPqedco0VS1q7qPlp66RPe604g3SgW7aleauKr1WNEjznIuMLn58rJAJwVAJWoZ6En551YRYNDZte0KVbkZCUtUyK1ESE9j5aQnI/bWVb4yx71cfAFvloa6pb+wYHcZ5x9gt/o0aNbVdXRo0aNCEaNGjQhGjRo0IRo0aNCF8pLKZDK2VoCkrBBCuxB7jUC35t3OtqU5MgMrepqzzhSepYyfhV9PkR+up/1+HGm3QQ4nmBGCD21WfEPhmm8QweXPs4cHslGyU/uAQe+j5Y657Y0wlZ2ts6rLU8aeYryjkrjL5Ovzx2/prVo2WtNKwpT1QWAclJeGFD5HCe2uQS+y24xyaY3gjulylj3E28/wBP6ciNDQfxWNlUJaUk8xPdtQHUpPz9DqNqNt3V7GmPM3TT1xKv8Km1/wDho/lI6K+pHTGrDqPZ1u0BOKTTW2FHopz4lq+6jk6/Fw2Pa91whAuCkMTWk5KfMQOZH/Crun9Mavtn8HVlBbzSyTZJ6Y2Hp3VTuXhOjrqsVwGJB16JDPv9dGDjOOmNNnL4YNtZDxdZcq8ZJ7NtSgUj/Mkn+us2icOm2dHdD66fIqK0nKfbX+dI/wAIASf1B1C3whWa93ABQiwVB+FxGEhdxbXXOmLKvC3qBPk0dglT7zDRIjH1xjqpPX0Bx69NcMuTLlNgOS33m8dOZxShj7atwZo9PjxkQ48dLTDaeRDbYCUpT8gAMAaim+eFfaC9pjtQft5VMmPHmXJprpjqUo9yUj3CfqU51cIbRJDA2MPyRz6qs3j2fzSZkoJcZ5aePt2VcPKT0KNZLEqc6URGJEtwunlS004o86j+UJ9T9Bp5m+BfaYuEOVW58Y6ATWgf38vUm2HsBtXtyQ9bdsMol4x7XIUX5A+y15Kf8OBpzLZLnd2Fpbd7ObmZPzpA0ehKhLhZ4cpdrTBuPfMHyKkpJFMgrHvReYdXXPk4R0A/KM56no10cYb+5J14mM2n4c/bX0AwMDW3hiELQ0LrtptMNopxTw8Dr1J6kr3Ro0alWzRo0aNCEaNGjQhGjRo0IRo0aNCEmZ8UrYYLwm170UCcA+xxxn93tZdO8Trh6nSeWdT7upqDgB56nNuIyfo24T/TSG8INiWrudxC2tY970pNRotRE32mMVqQF8kR1xPVJBGFJSeh9NPbvVwD8PDe2dyVi07Zk2/VqVS5M6JLYnOrSHGm1OBK0LUUqSeXB9RnuNZLmMY7SVWKGtrq6EzNxgJndvd0bC3WoSLm2+uWJWqc4ooL0dXVtYAJQtJwpCgFD3VAHqNdRzDVUXhnXrWqTvrULejvq/Dq5QpD0pjqUF1gpU25j5jmWnP8/wBTpwd8OPHbDZGtuWjKgTriuJhvnkQqZyBuKSMpS86sgIUQR7oClAEEgZ1E+Ih2AtlTXON9OJpjhM4VAa9BzpP9o/Eh2v3Gr8K1rmt6oWhPqDwYjSJb6H4KnFdEIU8nlKCo9AVICckddNLU7ki0SmyqvVpcaJDiNrdefeWEIbQkEqUpROAABk6Y5rmncLLp6yGpaXxuyB/C3WRoz9DpIrt8UzbKlVv8PtCwK3cNPbd5HaiqQ3DbUnp7zSFBSljqfj5O2tlE8UHZ2VUYFMYsm8HHJrjLRX7NHCWluEDlVl7J5SoZIyPlpTE/sohc6QnSH7pzNGuF3o3et7Y7bqp7l3REnSqdS1MpcZhISt5xTrqWkhIUpI+JY7kaUib4rlmoqaWYG0Vdegk4Lz1SYaeA/uwFJ/8AeNAY53CfUV8FOcSOwnwByde6hfh14pduOJCLNcs16TFn0xKVTaXObDcllKiQlfQlC0HHxJUQCcHHTPZ7p7rWptBa828L0qrFPpkFKeZxzKlOLVnlaQkdVLURgJH/ACBwhBGxUzamMx+aHbLrpjnkxlvflbSVq6egBzpbNoeOna3ejcaJtha1FuSPUpwkKaemR2kMkMoUtWSlwqGUpOOmosl+Kjt3MqcimJ2wuRFIdbU2mcZDHn5IwCY+cBOT38zP00tPh+lK+La2HUqylbFUWk49DEdI/pjUzYtiXLSVF2Dp42U5yHHBVwevAoH11CXEPxT2nw3U+j1C76DWamiuOPMxk0xltXKpsAq51OLSkdxj179Oml5pviu2U7UlNVLaCvMQMnD8eosPOkehLZCEj/P+p1GI3OGQtlNcIIH+W52CnxU6lCec5I+g0ru8/HztJs9ue9ttV6VcFQkQCymqyoTbfkwitIVykLWFLISpKjyjscDJ6ambaPd6y977RavOxaqJlLdWWVgoKHo7yQOZp1B+FYCknHqFAgkHSRcYh4RadxFyUblWZuDIuFTEF+oGhyYzcCeFpAR5ocUHObkCUqKeXIA9eukaN8FJV1BbCJIXD79VYjTqpCq0SPOgulxiUyiQ0vlICm1AFJ6/MEay9R/fe4lubPbYVDcOpQpX4JQae2+Y8VsKdDWEhCEpJAz1CepAHz6aUyo+LBZSZraKHtFX5cQ/G7JnsR3B9m0hwH/MNDWOfwE+WuhpwPNOMp89eEgdzqC+Hri7214h5E2k26qXS63Bb892k1JtKH1M5A81spUpLiQSAcHIynIGdZnEbxNWnw20ui1W8KJV6m3XJLsWOimIaKkKbQFEr8xaRjBGOukLXA4wn+9xGLzg74e6mjI17pM694n+zNNoMKpUa27hqlTmtKcVTPKbYMRQUQEPvFRQCQAr3OfAUM41u9lPEH243jrqLTm0eRaNZfaW5HTUn0OxX+QFSkpeTykKCQThSUg4OCT00/ynYyQoW3Skc/QHhVsbF7rvbI7oUTc5iioqy6MJAENbxaDnmsONfEAcY8zP6aYDdvxHNwt0bRqViW5ZUC3BXmPYn5LEtyVJLS8hxtpPKkBSweXPU4zgA9dRvwO0SiXRxP2fRbho8OpQJAnl2NMYQ8yvlhPKTlCwUnBAIyO41bvQdr9ubad9rt2xLfpb5A/iw6Ywyvoc/ElIOp5iGv3CrNmpKmenLWSaWk8dUi3A7sVdOz1u3hxC33QnoEyNbsr8FpklJbfLCW/OW6tPdHP5aEpSRnAJI+HKqbFXPtQ9u+1e/EKqZUaOfaanJbTGVI9tnrUFIDqU9VI5itR9DgA9Mg3LX5bSq/ZFxW/Dd5H6xS5cBta+oSt1laAenU9VDVOfDKNp6FvFDo3EHb0aRbzqX6XLRPC0NQJoICHXeUggJWgoJOAnnyegOkY7XkqW4UvujoIWEAevH3WbxWXRsJet7QLl2DpwpUJ+CW6nERTVwmkyUKHI6hvASkqScHkwAUZ7nTRcU279wv8AArt45IlPIqO4ESmx57qDkuMoY811JOfzlCM985VnvqQby218PCxqCK9ctD27jRnWvNaDVRW+4+jp/u223CtzuPhB1xviA2dSXeF6wqzYMBqNbNu1CIuIxHQrkagyI6kMnqSUpHMgdfVQGl1AkbKQ08lNHM/UCSNw1c94fXDbYF/bfVbcq/rUplfemz3aZCROR5iI7LQSF4QegWpSj73cAJxg50tXELYluba8U9Ysm0KcIFHplYp3s0YOKWloONx3VAFRJxzLVgegOB0A03Phx73WFRNqqzt1cVyUqjVKj1GRVEpny0M+0RHEpUXElZAIQUqCsE46HsdKdxJ3lbl/cWVdu20aqxUqROrVOTHlsElt7ymo7Syg4GU87agD2IGR0OlGoOOQsWSOBtJE+MjORlWqb+TdpGtrqp/rvXCFo4aXObl83K6pLgW2hIT7yllaRypT1P2zpF7g4j+Ah+lSrUhcN1QXT1ocYRMYpsRh5KcEJcQ4Xg6D2PUgjv1103iqVOpBO3FGDjqaa85UpTiOX3FSEBhKCT8wha8fcnW94T6LwrU/hkp96X1AsN2qBuWuuTK2iO7IaeS6sJQUucyk+55fKhHQ8wIyTpgGBlbCoqHVM5gj0gNHJHKXTw8K3MpPFNRafTXiI9Yp1RhvBfdbKWFPIBx0CuZlB6fXHfUleKZetQkbg2nYCJPLTqdSlVZ9tJ+OS64psFQ9eVtv3T/5itQ9wR1mjUbiut2sVCpRINPaFWX58hxLTTaDDeCclRAT6ADUm+J/bslG7Vs3eI/+z61b/s6JAPuqeYeWVJz2yEONn656Zxp5GXha2Nz2W1wz+7B+iYjYXg/2aPDrRVXht7S6rWK7R2qpUJ8pBMkOutlwJbcB5mggKCQEEA8uT3Ok88P9Qb4trXabQcNx6qkDOegiO6ebYbiV2qqPDRRK/X77otNcotGapdXaky0Nux5TTXllBbJ5jzEAowCVBQx3wEa8P9aTxcWutoEocj1RaeYYODEdxoaThwKyZ2wtlpjHgfRP1xW3Bw1W/b1IqPERSolYTFfdfotMU2p5997lCVlDaSARjGVK90ZHX00kW92/nCFuNt9UbfsnYSZblwhpP4TU2YMOMGHgoE+aWXSVIKQR+YnPp31jeI1NmPcTsyFUpUgxabRoDUNPNzeW2pBUsoB9SpSj9Tqft+aTwsWdwmzmbOp9htVGoUJlNDltNxnalKfV5auZK8F1ThBJUon3euSANNaNIBTqmR9VJM0YaG9+StF4UFfdU9uLbCistAQaggFR5UqPmtqwO2ThOT68o+WoX8Q3rxazx/6Ck/8Aw1IfhZVuk0i9r4jT6jGjP1CDBYitvOpSp9fmOHlQD1UcZOB1xqO/EM/72U/zPiNPpJGP7saUbPOVBKdVsjx0Ksz3If28b2hqKd1H4TVqGlpRVTMJDSmSkDl6dSSSAAnKicAAnGkVqHEj4f8ASoUq2Kdw2vT6dzKbMtFJitqU2c+82tx0PD1xkpPTXd+JzVawxs3YVKhLdTTJ1UC5vKDylbUbLKVH5ZUsgHuRn8udaXgat3hkOxsy7tw4dmSa9GnSvxd6vFhao7CSPKTyv5CEFvr7owcnOT001owNSzqmV0tT7q3Aw3kpd+EGrRaZxf2euzFS2aTUKtMiRUTAPP8AYXGXeVDgSSObl5QcEjIOCdM74rQ5bN28KgM/i03OP7hGlo2Vr1pM8btFuKiSYUS2jeE56E6nljxWoig/5ZSDhKEcpTgdMZHz0yPiny2KhYG20+HKYksP1SW408yoKQ4ksIIUkgkEEdcg6e75wsCnI/DphnOCfosXgY4VNm9ydoW9zb7tFqv1CpTpkUomvKLLDbLpSny208oycElRyfQYGk64h7Ro+32/F72fbMcxKbSqu61CaSo/wWiAoIBJJwAvHU9hqybw2yr/ALK9IzjH4rU//sK0gfFHTHqzxa7g0yMMuv1txKQAT8MdKvT6JOkjPxkKO508bLdE+NuCcb9U72xPh4w9kN1KNufH3Qk1dyjpkBENdMQylwusLayVBZIwHCeg9NOCyhSG8KGDk56519NGsZzi45KuVNSxUjdEIwF8Zba3WS2jOScHr6f/AL9tKPxA+HjaO8Ffk3ta1xrtWvzVlyaBHD8SW4e7imwUlCyB1Uk+9nJGckt9o0NcW8IqKWKqbplGVXbQPCclCeP9Kt4GFQD/ALxFMpJbeX16gKccUlPTPXB+2nBvK1turW2Wl2ZuDIBsql0RMCe9UFc/LCbQEcyygZ5gADzJGQeoHQalHUb787ZN7ybdVvbR64JdFbrMVIXMjIC1ICHErKSgkBSVcuFDI6E/fTtRed1jNoYaKN3kN5H+Ui6fDu27vGea7tjxHUmbbDp89AdZbmOtNZ7B1DqUnHbmUB1HXGlpuOxrXpPEWxt3t5cL1dpEe4oFKi1FXIoyXedpDqx5fu8vnF3GOnKB1PU60e7FgP7N3i/abNeXUS4wh1yQ2yYwWFDqkoC1Z746k6djgU4T7VQ5S99K/WnKpNiEu0yneyhpiI+Ukh1R51FxSR8PwgHBxkDWTktGSqtC1ldOIoWaTnlNhv7w72nxC2WbOuwOxSw97TAqEcgvQ38EcyQeikkHCknGR8iAQpNH8J+UmtJcuHeJp2loUfdhUnklFORgc61lIJGRnlOM9tWIaNYokcFbJrbTVB1SNyVX3dPhUsz6+6uz90BSrffUlXsk6D7XKY6dQl4KSF4PUcw++e+mh3S4bLR3h2vg7aXmlxwUuOwiDUmeVEiK+035YeRkEdQPeQfdI6fXUx6NBkcTnKI7bTRBwa35uVXhTvCgkt14SKtvC0/SUuZLcekeXKW3n4edThQlRGeuCB/TUnbHeH1E2R3cp26lP3Ik1MU1MpDVPegIbyl5pbY5nAo9UhecgdcacHRpTK4qOO00kRBazcJbuJrgxtjiOciV92tPW/dFPZMZmosth1DzOSoNPN5HMEqUSFAhQyfn0hGwfClo9Pqpm7kblqqkRJ6RKRB9kU50/O6pSlAZ9E/uNP8A6NIJHAYT5bZSzSea9u6Q/bbwzqtt9unRr7a3hQ7At+pt1CIy3S+WS4lCshtay5yjmGUqIHY9Brtt+OARjfXdF3c6buU/RnXY8SOYTVPS+kBhOAecrBOft003WjR5jkgtdKG6A3bOfuo33e2QtXejbh3ba72XFxFhpTMllSUvxn2xht5BIICh16YIIUoHodJq14Ts1dbzM3ljmkJX7vJRcS+T7+ZyBX2GNWJaNAkcEs9tpqhwdI3JCr+vvwqoVQqjTm3+5rlNpZZbQ9GqkT2p3zAkBa0rQpA94gq5SkAFWB06CY93uCtjdnaXb3a16/HqYiwIzUZqY3BS4ZYRHQzlSOYBGeQKwCe+NM9o0vmuKRtrpWBzWt2dyoo4eNjRsBthD22jV9dbRDlyZPta2AwV+c4VkcgJxjOM566hm5+AGJcm+dU3pf3Ad/2rOemuUtVOQpKedgtcvOV9cZ5skemm90aaHkElSvoYHsbG5uw4X//Z';

// ============================================================
// PENJUALAN — CONFIRMATION DEAL EMAIL (manual, tombol per transaksi)
// CARA PASANG: buka Code.gs, scroll ke PALING BAWAH file, paste
// seluruh isi file ini di baris paling akhir. Tidak mengubah
// fungsi/sheet yang sudah ada.
// ============================================================

/**
 * Cari Email PIC buyer dari sheet Master_Customer (READ-ONLY),
 * dicocokkan dengan nama Pembeli di transaksi Penjualan.
 */
function getBuyerEmailByName_(namaBuyer) {
  const nama = String(namaBuyer || '').trim().toLowerCase();
  if (!nama) return '';
  const master = getMasterCustomerData();
  const found = master.find(m => String(m['Nama Customer'] || '').trim().toLowerCase() === nama);
  return found ? String(found['Email PIC'] || '').trim() : '';
}

const CONFIRMATION_DEAL_TEMPLATE =
  'Yth. Bapak/Ibu {{NAMA_BUYER}},\n\n' +
  'Terima kasih atas kepercayaan yang telah diberikan kepada PT Pegadaian.\n\n' +
  'Dengan ini kami mengonfirmasi bahwa transaksi pembelian emas telah disepakati (deal) dengan rincian sebagai berikut:\n\n' +
  'Detail Transaksi\n' +
  '- Nomor Invoice : {{NO_INVOICE}}\n' +
  '- Tanggal Deal : {{TANGGAL_DEAL}}\n' +
  '- Buyer : {{NAMA_BUYER}}\n' +
  '- Produk (Merk) : {{MERK}}\n' +
  '- Berat (Gramasi) : {{GRAMASI}} gram\n' +
  '- Harga Deal : {{HARGA_DEAL}}\n' +
  '- Nilai Transaksi : {{NILAI_TRANSAKSI}}\n' +
  '- Lokasi Pengambilan : {{LOKASI_PENGAMBILAN}}\n\n' +
  'Segala perubahan harga setelah dokumen ini diterbitkan hanya dapat dilakukan berdasarkan persetujuan tertulis antara PT Pegadaian dan Buyer.\n\n' +
  'Dokumen ini menjadi dasar penerbitan invoice dan pelaksanaan proses transaksi sesuai ketentuan yang berlaku.\n\n' +
  'Apabila terdapat pertanyaan atau memerlukan informasi lebih lanjut, silakan menghubungi PIC PT Pegadaian.\n\n' +
  'Hormat kami,\n' +
  'PT Pegadaian - Divisi Bisnis Bulion\n\n' +
  '---\n' +
  'Email ini dibuat secara otomatis oleh sistem. Mohon tidak membalas email ini.';

// Detail perusahaan yang tampil di header/footer form — sesuaikan di
// sini kalau ada perubahan alamat, kontak, atau lokasi pengambilan.
const CONFIRMATION_DEAL_COMPANY = {
  namaResmi: 'PT PEGADAIAN',
  divisi: 'Divisi Bisnis Bulion',
  alamatGedung: 'The Gade Tower',
  alamatJalan: 'Jl. Kramat Raya No.162, Jakarta Pusat',
  email: 'bisnis.bulion@pegadaian.co.id',
  website: 'www.pegadaian.co.id',
  lokasiPengambilanDefault: 'The Gade Tower'
};

/**
 * Isi template TEKS POLOS dengan data satu baris PenjualanTransaksi —
 * dipakai sebagai plainBody (fallback untuk klien email yang tidak
 * bisa render HTML). Nilai Transaksi dipakai TotalHargaAfterTax
 * (nominal final setelah PPh Pasal 22) karena itu angka yang jadi
 * dasar pembayaran buyer.
 */
function buildConfirmationDealBody_(row) {
  const tanggalLabel = formatTanggalDealLabel_(row.Tanggal);

  return CONFIRMATION_DEAL_TEMPLATE
    .replace(/{{NAMA_BUYER}}/g, row.Pembeli || '-')
    .replace(/{{NO_INVOICE}}/g, row.NomorInvoice || '-')
    .replace(/{{TANGGAL_DEAL}}/g, tanggalLabel)
    .replace(/{{MERK}}/g, row.Merk || '-')
    .replace(/{{GRAMASI}}/g, Number(row.Qty || 0).toLocaleString('id-ID', { maximumFractionDigits: 2 }))
    .replace(/{{HARGA_DEAL}}/g, formatRupiahServer(row.HargaPerGram))
    .replace(/{{NILAI_TRANSAKSI}}/g, formatRupiahServer(row.TotalHargaAfterTax))
    .replace(/{{LOKASI_PENGAMBILAN}}/g, CONFIRMATION_DEAL_COMPANY.lokasiPengambilanDefault);
}

function formatTanggalDealLabel_(tanggalValue) {
  const iso = extractDateOnly(tanggalValue);
  if (!iso) return String(tanggalValue || '-');
  const d = new Date(iso + 'T00:00:00');
  return d.getDate() + ' ' + NAMA_BULAN_ID[d.getMonth()] + ' ' + d.getFullYear();
}

/**
 * Satu baris tabel "Detail Transaksi" pada form HTML — label di
 * kiri (abu-abu, bold), value di kanan, dipisah garis tipis.
 */
function confirmationDealHtmlRow_(label, value) {
  return '<tr>' +
    '<td style="padding:9px 12px;border-bottom:1px solid #E7EAEE;font-size:13.5px;font-weight:700;color:#4B5563;width:38%;background:#F8F9FA;">' + label + '</td>' +
    '<td style="padding:9px 12px;border-bottom:1px solid #E7EAEE;font-size:13.5px;color:#1B2430;">' + value + '</td>' +
    '</tr>';
}

/**
 * Membangun form HTML "PRICE DEAL CONFIRMATION FORM" sesuai desain
 * resmi — header perusahaan, tabel detail transaksi, teks konfirmasi,
 * catatan, dan footer copyright. Semua data transaksi diambil dari
 * baris PenjualanTransaksi yang sama dengan versi teks polos, supaya
 * kedua versi selalu konsisten.
 */
function buildConfirmationDealHtml_(row) {
  const c = CONFIRMATION_DEAL_COMPANY;
  const tanggalLabel = formatTanggalDealLabel_(row.Tanggal);
  const gramasiLabel = Number(row.Qty || 0).toLocaleString('id-ID', { maximumFractionDigits: 2 }) + ' gram';
  const tahun = new Date().getFullYear();

  const rows =
    confirmationDealHtmlRow_('Nomor Invoice', row.NomorInvoice || '-') +
    confirmationDealHtmlRow_('Tanggal Deal', tanggalLabel) +
    confirmationDealHtmlRow_('Buyer', row.Pembeli || '-') +
    confirmationDealHtmlRow_('Produk (Merk)', row.Merk || '-') +
    confirmationDealHtmlRow_('Berat (Gramasi)', gramasiLabel) +
    confirmationDealHtmlRow_('Harga Deal', formatRupiahServer(row.HargaPerGram)) +
    confirmationDealHtmlRow_('Nilai Transaksi', '<b>' + formatRupiahServer(row.TotalHargaAfterTax) + '</b>') +
    confirmationDealHtmlRow_('Lokasi Pengambilan', c.lokasiPengambilanDefault);

  return '' +
    '<div style="max-width:650px;margin:0 auto;font-family:Arial,Helvetica,sans-serif;color:#1B2430;">' +

      // Header: identitas kiri, alamat/kontak kanan
      '<table width="100%" style="border-collapse:collapse;">' +
        '<tr>' +
          '<td style="vertical-align:top;">' +
            '<img src="cid:logo_pegadaian" alt="Pegadaian" style="height:46px;display:block;margin-bottom:4px;" />' +
            '<div style="font-size:11px;color:#046A38;font-style:italic;">Mengatasi Masalah Tanpa Masalah</div>' +
          '</td>' +
          '<td style="vertical-align:top;text-align:right;font-size:12px;line-height:1.7;color:#333;">' +
            '<div style="font-size:15px;font-weight:800;color:#1B2430;">' + c.namaResmi + '</div>' +
            '<div style="font-weight:700;color:#046A38;">' + c.divisi + '</div>' +
            '<div>' + c.alamatGedung + '</div>' +
            '<div>' + c.alamatJalan + '</div>' +
            '<div>Email: ' + c.email + '</div>' +
            '<div>Website: ' + c.website + '</div>' +
          '</td>' +
        '</tr>' +
      '</table>' +

      '<div style="border-top:3px solid #046A38;margin:14px 0 22px;"></div>' +

      // Judul form
      '<div style="text-align:center;margin-bottom:18px;">' +
        '<div style="font-size:23px;font-weight:800;letter-spacing:0.3px;color:#1B2430;">PRICE DEAL CONFIRMATION FORM</div>' +
        '<div style="font-size:13.5px;color:#6B7684;margin-top:4px;">Formulir Konfirmasi Kesepakatan Harga Jual Beli Emas</div>' +
      '</div>' +

      '<p style="font-size:13.5px;line-height:1.6;color:#333;">Dokumen ini merupakan konfirmasi resmi atas kesepakatan harga transaksi jual beli emas antara PT Pegadaian dan Buyer. Seluruh informasi di bawah ini telah disepakati oleh kedua belah pihak dan menjadi dasar pelaksanaan transaksi.</p>' +

      '<div style="font-size:14.5px;font-weight:800;color:#046A38;margin:22px 0 8px;">DETAIL TRANSAKSI</div>' +
      '<table width="100%" style="border-collapse:collapse;border:1px solid #E7EAEE;">' + rows + '</table>' +

      '<div style="font-size:14.5px;font-weight:800;color:#046A38;margin:22px 0 8px;">KONFIRMASI</div>' +
      '<p style="font-size:13.5px;line-height:1.6;color:#333;">Dengan ini kami menyatakan bahwa rincian harga transaksi sebagaimana tercantum pada dokumen ini telah diperiksa dan disepakati oleh kedua belah pihak.</p>' +
      '<p style="font-size:13.5px;line-height:1.6;color:#333;">Segala perubahan harga setelah dokumen ini diterbitkan hanya dapat dilakukan berdasarkan persetujuan tertulis antara PT Pegadaian dan Buyer.</p>' +
      '<p style="font-size:13.5px;line-height:1.6;color:#333;">Dokumen ini menjadi dasar penerbitan invoice dan pelaksanaan proses transaksi sesuai ketentuan yang berlaku.</p>' +

      '<div style="font-size:13px;font-weight:700;margin-top:14px;color:#1B2430;">Catatan:</div>' +
      '<ul style="font-size:12.5px;line-height:1.6;color:#333;margin:6px 0 0;padding-left:18px;">' +
        '<li>Harga berlaku sesuai waktu deal.</li>' +
        '<li>Transaksi akan diproses setelah seluruh persyaratan dipenuhi.</li>' +
        '<li>Dokumen ini dibuat secara otomatis oleh Gold Trading Management System - PT Pegadaian Divisi Bisnis Bulion.</li>' +
      '</ul>' +

      '<div style="border-top:2px solid #046A38;margin-top:22px;padding-top:10px;text-align:center;font-size:11px;color:#6B7684;">' +
        'Copyright &copy; ' + tahun + ' PT Pegadaian - Divisi Bisnis Bulion. All Rights Reserved.' +
      '</div>' +
    '</div>';
}

/**
 * Dipicu MANUAL dari tombol "Kirim Confirmation Deal" / ikon mail
 * di baris transaksi Penjualan (baik di "Penjualan Hari Ini" atau
 * di tabel "Riwayat Penjualan"). Email buyer dicari dari
 * Master_Customer (kolom "Email PIC"), dicocokkan dari field
 * Pembeli di transaksi — KECUALI jika overrideEmail diisi (dipakai
 * untuk testing), maka email itu yang dipakai dan lookup ke
 * Master_Customer dilewati sepenuhnya. Kalau tidak ada override
 * DAN tidak ketemu juga di Master_Customer, dilempar error yang
 * jelas supaya user tahu harus lengkapi dulu atau isi email manual.
 * Email dikirim sebagai HTML (form "Price Deal Confirmation") dengan
 * versi teks polos sebagai fallback (plainBody) untuk klien email
 * yang tidak mendukung HTML.
 * BELUM ada lampiran invoice PDF — menyusul saat generatornya
 * sudah dibuat terpisah.
 */
// Alamat & nama pengirim resmi untuk Confirmation Deal Email. Alamat
// ini WAJIB sudah terdaftar & terverifikasi sebagai "Send mail as"
// di akun Gmail yang menjalankan script ini (Settings > Accounts and
// Import > Send mail as) — kalau belum, GmailApp akan menolak/mengabaikan
// opsi "from" ini dan tetap kirim dari email akun asli.
const CONFIRMATION_DEAL_SENDER_EMAIL = 'bisnis.bulion@pegadaian.co.id';
const CONFIRMATION_DEAL_SENDER_NAME = 'PT Pegadaian - Divisi Bisnis Bulion';

/**
 * Data awal untuk modal "Kirim Confirmation Deal" di frontend —
 * mengembalikan nomor invoice, nama buyer, dan email default (hasil
 * lookup Master_Customer) supaya field "To" bisa terisi otomatis
 * tapi tetap bisa diedit/ditambah oleh user sebelum kirim.
 */
function getConfirmationDealPrefill(transaksiId) {
  requireRole('Viewer');
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const row = rows.find(r => r.ID === transaksiId);
  if (!row) throw new Error('Transaksi Penjualan dengan ID "' + transaksiId + '" tidak ditemukan.');
  return {
    nomorInvoice: row.NomorInvoice || transaksiId,
    pembeli: row.Pembeli || '-',
    emailDefault: getBuyerEmailByName_(row.Pembeli) || ''
  };
}

/**
 * Dipicu dari modal "Kirim Confirmation Deal". toOverride & ccOverride
 * boleh berisi lebih dari satu alamat email dipisah koma (mis.
 * "a@x.com, b@x.com") — GmailApp.sendEmail menerima format ini
 * langsung tanpa perlu diparse manual. Kalau toOverride kosong,
 * email diambil otomatis dari Master_Customer (kolom Email PIC)
 * berdasarkan nama Pembeli di transaksi.
 */
function sendConfirmationDealEmail(transaksiId, toOverride, ccOverride) {
  requireRole('Viewer');
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const row = rows.find(r => r.ID === transaksiId);
  if (!row) throw new Error('Transaksi Penjualan dengan ID "' + transaksiId + '" tidak ditemukan.');

  const manualTo = String(toOverride || '').trim();
  let to = manualTo;
  if (!to) {
    to = getBuyerEmailByName_(row.Pembeli);
  }
  if (!to) {
    throw new Error('Email PIC untuk buyer "' + row.Pembeli + '" tidak ditemukan di Master_Customer. Lengkapi dulu kolom Email PIC di Master_Customer, atau isi email tujuan manual di form kirim.');
  }
  const cc = String(ccOverride || '').trim();

  const subject = 'Confirmation Deal — Invoice ' + (row.NomorInvoice || transaksiId);
  const plainBody = buildConfirmationDealBody_(row);
  const htmlBody = buildConfirmationDealHtml_(row);
  const logoBlob = Utilities.newBlob(Utilities.base64Decode(CONFIRMATION_DEAL_LOGO_BASE64), 'image/jpeg', 'logo_pegadaian.jpg');

  const sendOptions = {
    htmlBody: htmlBody,
    inlineImages: { logo_pegadaian: logoBlob },
    from: CONFIRMATION_DEAL_SENDER_EMAIL,
    name: CONFIRMATION_DEAL_SENDER_NAME
  };
  if (cc) sendOptions.cc = cc;

  GmailApp.sendEmail(to, subject, plainBody, sendOptions);
  logAudit('SEND_EMAIL', 'PenjualanTransaksi', transaksiId, 'Confirmation Deal dikirim ke ' + to + (cc ? ' (CC: ' + cc + ')' : '') + ' dari ' + CONFIRMATION_DEAL_SENDER_EMAIL);
  return { sent: true, to: to, cc: cc };
}

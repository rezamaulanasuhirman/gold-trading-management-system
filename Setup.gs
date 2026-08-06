/**
 * ============================================================
 * DAILY GOLD TRADING MANAGEMENT SYSTEM
 * Divisi Bullion PT Pegadaian
 * Setup.gs — Script Inisialisasi Database
 * ============================================================
 * CARA PAKAI:
 * 1. Buka https://script.new (atau Extensions > Apps Script
 *    dari Google Sheet kosong).
 * 2. Hapus isi default, paste seluruh isi file ini.
 * 3. Simpan project dengan nama "Gold Trading Management System".
 * 4. Pilih function "setupDatabase" di dropdown toolbar, lalu
 *    klik Run (ikon play). Izinkan permission saat diminta.
 * 5. Cek Log (Ctrl+Enter / View > Logs) untuk melihat
 *    Spreadsheet ID yang baru dibuat — simpan ID ini, akan
 *    dipakai di file Code.gs pada fase berikutnya.
 * ============================================================
 */

/**
 * Fungsi utama — jalankan ini SEKALI saja.
 * Membuat spreadsheet baru + seluruh sheet + header + data awal.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.create('DB - Gold Trading Management System');
  const ssId = ss.getId();

  Logger.log('============================================');
  Logger.log('Spreadsheet berhasil dibuat!');
  Logger.log('Spreadsheet ID: ' + ssId);
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('============================================');
  Logger.log('SIMPAN Spreadsheet ID di atas — akan dipakai di Code.gs');

  createSupplierSheet(ss);
  createBuyerSheet(ss);
  createBenchmarkPriceSheet(ss);
  createSupplierPriceSheet(ss);
  createOutstandingSheet(ss);
  createPriceSettingSheet(ss);
  createInvoiceSheet(ss);
  createUsersSheet(ss);
  createAuditTrailSheet(ss);
  createSettingsSheet(ss);

  // Hapus sheet default "Sheet1" yang otomatis dibuat Google
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  // Simpan Spreadsheet ID ke Script Properties supaya Code.gs
  // bisa langsung menemukannya otomatis tanpa hardcode.
  PropertiesService.getScriptProperties().setProperty('DB_SPREADSHEET_ID', ssId);

  Logger.log('Semua sheet + data awal selesai dibuat.');
  Logger.log('Spreadsheet ID juga sudah disimpan otomatis di Script Properties (key: DB_SPREADSHEET_ID).');
}

// ------------------------------------------------------------
// 1. SUPPLIER
// ------------------------------------------------------------
function createSupplierSheet(ss) {
  const sh = ss.insertSheet('Supplier');
  const headers = ['SupplierID', 'Nama', 'Status', 'CreatedAt'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);

  const suppliers = ['Emasku', 'IGS', 'Marva', 'IDN Bullion', 'StarGold', 'Lotus', 'APEPI'];
  const now = new Date();
  const rows = suppliers.map((name, i) => ['SUP-' + String(i + 1).padStart(3, '0'), name, 'Aktif', now]);
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.autoResizeColumns(1, headers.length);
}

// ------------------------------------------------------------
// 2. BUYER
// ------------------------------------------------------------
function createBuyerSheet(ss) {
  const sh = ss.insertSheet('Buyer');
  const headers = ['BuyerID', 'Nama', 'Kontak', 'Status', 'CreatedAt'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);
  sh.autoResizeColumns(1, headers.length);
  // Data buyer diisi manual lewat form nanti — tidak perlu seed data.
}

// ------------------------------------------------------------
// 3. BENCHMARK PRICE (Spot Gold, XAUIDRK, LBMA, Kitco harian)
// ------------------------------------------------------------
function createBenchmarkPriceSheet(ss) {
  const sh = ss.insertSheet('BenchmarkPrice');
  const headers = ['ID', 'Tanggal', 'SpotGold', 'XAUIDRK', 'LBMA', 'Kitco', 'CreatedAt', 'CreatedBy'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);
  sh.autoResizeColumns(1, headers.length);
}

// ------------------------------------------------------------
// 4. SUPPLIER PRICE (Daily Supplier Price — Form 1)
// ------------------------------------------------------------
function createSupplierPriceSheet(ss) {
  const sh = ss.insertSheet('SupplierPrice');
  const headers = [
    'ID', 'Tanggal', 'SupplierID', 'SupplierNama',
    'Harga1', 'Harga2', 'Harga3', 'Harga4', 'HargaRataRata',
    'Catatan', 'CreatedAt', 'CreatedBy'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);
  sh.autoResizeColumns(1, headers.length);
}

// ------------------------------------------------------------
// 5. OUTSTANDING (Piutang Supplier — Form 2)
// ------------------------------------------------------------
function createOutstandingSheet(ss) {
  const sh = ss.insertSheet('Outstanding');
  const headers = [
    'ID', 'Tanggal', 'SupplierID', 'SupplierNama',
    'Gramasi', 'HargaDasar', 'Valuasi', 'Status',
    'Keterangan', 'CreatedAt', 'CreatedBy'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);
  sh.autoResizeColumns(1, headers.length);
}

// ------------------------------------------------------------
// 6. PRICE SETTING (Penetapan Harga — Form 3)
// ------------------------------------------------------------
function createPriceSettingSheet(ss) {
  const sh = ss.insertSheet('PriceSetting');
  const headers = [
    'ID', 'Tanggal', 'HargaSale', 'HargaBuyback', 'HargaMargin',
    'TanggalBerlaku', 'CreatedAt', 'CreatedBy'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);
  sh.autoResizeColumns(1, headers.length);
}

// ------------------------------------------------------------
// 7. INVOICE (Input Invoice Penjualan — Form 4)
// ------------------------------------------------------------
function createInvoiceSheet(ss) {
  const sh = ss.insertSheet('Invoice');
  const headers = [
    'ID', 'NomorInvoice', 'Tanggal', 'BuyerID', 'BuyerNama',
    'Harga', 'Gram', 'Nominal', 'Status', 'Catatan',
    'CreatedAt', 'CreatedBy'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);
  sh.autoResizeColumns(1, headers.length);
}

// ------------------------------------------------------------
// 8. USERS (Role Management)
// ------------------------------------------------------------
function createUsersSheet(ss) {
  const sh = ss.insertSheet('Users');
  const headers = ['UserID', 'Email', 'Nama', 'Role', 'Status', 'CreatedAt'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);

  // User pertama = admin, otomatis pakai email akun yang menjalankan script ini.
  const myEmail = Session.getActiveUser().getEmail() || 'admin@pegadaian.co.id';
  sh.getRange(2, 1, 1, headers.length).setValues([[
    'USR-001', myEmail, 'Administrator', 'Admin', 'Aktif', new Date()
  ]]);
  sh.autoResizeColumns(1, headers.length);
}

// ------------------------------------------------------------
// 9. AUDIT TRAIL
// ------------------------------------------------------------
function createAuditTrailSheet(ss) {
  const sh = ss.insertSheet('AuditTrail');
  const headers = ['ID', 'Timestamp', 'User', 'Action', 'Module', 'RecordID', 'Detail'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);
  sh.autoResizeColumns(1, headers.length);
}

// ------------------------------------------------------------
// 10. SETTINGS
// ------------------------------------------------------------
function createSettingsSheet(ss) {
  const sh = ss.insertSheet('Settings');
  const headers = ['Key', 'Value', 'Description'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);

  const defaults = [
    ['APP_NAME', 'Daily Gold Trading Management System', 'Nama aplikasi'],
    ['DIVISI', 'Bullion PT Pegadaian', 'Divisi pemilik aplikasi'],
    ['DEFAULT_MARGIN_WARNING', '0', 'Ambang batas margin dianggap negatif (Rp/gram)'],
    ['THEME', 'light', 'Tema default: light / dark']
  ];
  sh.getRange(2, 1, defaults.length, 3).setValues(defaults);
  sh.autoResizeColumns(1, headers.length);
}

// ------------------------------------------------------------
// MODUL TAMBAHAN — PEMBELIAN (dijalankan sekali saja, terpisah
// dari setupDatabase() karena database sudah pernah dibuat)
// ------------------------------------------------------------
/**
 * Jalankan fungsi ini SEKALI dari editor Apps Script (pilih
 * "addPembelianModule" di dropdown toolbar, klik Run) untuk
 * menambahkan sheet "Pembelian" ke database yang sudah ada.
 * Tidak akan menghapus atau mengubah sheet lain yang sudah ada.
 */
function addPembelianModule() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (!id) throw new Error('DB_SPREADSHEET_ID belum diset — jalankan setupDatabase() dulu.');
  const ss = SpreadsheetApp.openById(id);

  if (ss.getSheetByName('Pembelian')) {
    Logger.log('Sheet "Pembelian" sudah ada, tidak dibuat ulang.');
    return;
  }

  const sh = ss.insertSheet('Pembelian');
  const headers = [
    'ID', 'Tanggal', 'SupplierID', 'SupplierNama',
    'Gramasi', 'Harga', 'Nominal', 'Status',
    'Catatan', 'CreatedAt', 'CreatedBy'
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatHeader(sh, headers.length);
  sh.autoResizeColumns(1, headers.length);

  Logger.log('Sheet "Pembelian" berhasil dibuat.');
}

// ------------------------------------------------------------
// MODUL TAMBAHAN — PENJUALAN & PEMBELIAN LENGKAP
// (struktur sesuai rekap Excel — jalankan sekali saja)
// ------------------------------------------------------------
/**
 * Jalankan fungsi ini SEKALI dari editor (pilih
 * "addPenjualanPembelianFullModule" di dropdown toolbar, klik Run)
 * untuk menambahkan sheet "PenjualanTransaksi" dan
 * "PembelianTransaksi" ke database yang sudah ada.
 */
function addPenjualanPembelianFullModule() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (!id) throw new Error('DB_SPREADSHEET_ID belum diset — jalankan setupDatabase() dulu.');
  const ss = SpreadsheetApp.openById(id);

  if (!ss.getSheetByName('PenjualanTransaksi')) {
    const sh1 = ss.insertSheet('PenjualanTransaksi');
    const headers1 = [
      'ID', 'Tanggal', 'NomorInvoice', 'HargaPerGram', 'Qty',
      'TotalHarga', 'PphPasal22', 'TotalHargaAfterTax',
      'Pembeli', 'Denominasi', 'Merk', 'Status',
      'PaymentTime', 'Bank', 'TanggalPengambilan', 'EstimasiPenjemputan',
      'Catatan', 'CreatedAt', 'CreatedBy'
    ];
    sh1.getRange(1, 1, 1, headers1.length).setValues([headers1]);
    formatHeader(sh1, headers1.length);
    sh1.autoResizeColumns(1, headers1.length);
    Logger.log('Sheet "PenjualanTransaksi" berhasil dibuat.');
  } else {
    Logger.log('Sheet "PenjualanTransaksi" sudah ada, dilewati.');
  }

  if (!ss.getSheetByName('PembelianTransaksi')) {
    const sh2 = ss.insertSheet('PembelianTransaksi');
    const headers2 = [
      'ID', 'Tanggal', 'NomorPO', 'Seller', 'Qty',
      'Harga', 'AfterDiskon', 'TotalHarga', 'PajakWapu',
      'TotalHargaIncludePajak', 'InvoiceRef', 'Catatan',
      'CreatedAt', 'CreatedBy'
    ];
    sh2.getRange(1, 1, 1, headers2.length).setValues([headers2]);
    formatHeader(sh2, headers2.length);
    sh2.autoResizeColumns(1, headers2.length);
    Logger.log('Sheet "PembelianTransaksi" berhasil dibuat.');
  } else {
    Logger.log('Sheet "PembelianTransaksi" sudah ada, dilewati.');
  }
}

// ------------------------------------------------------------
// HELPER — format header (bold, background hijau emerald, font putih)
// ------------------------------------------------------------
function formatHeader(sheet, numCols) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setFontWeight('bold');
  range.setBackground('#046A38'); // emerald green Pegadaian
  range.setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
}

// ------------------------------------------------------------
// MODUL PROFILE MITRA — jalankan sekali. Cek dulu tiap sheet,
// TIDAK PERNAH menimpa/mengubah sheet yang sudah ada.
// ------------------------------------------------------------
function setupMitraSheets() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (!id) throw new Error('DB_SPREADSHEET_ID belum diset — jalankan setupDatabase() dulu.');
  const ss = SpreadsheetApp.openById(id);

  if (!ss.getSheetByName('MITRA_PROFILE')) {
    const sh = ss.insertSheet('MITRA_PROFILE');
    const headers = [
      'Mitra_ID', 'NamaPerusahaan', 'JenisMitra_Override', 'Kategori', 'StatusAktif',
      'Website', 'JenisProduk', 'MerekEmas', 'LokasiPabrik', 'LokasiGudang', 'Catatan',
      'CreatedAt', 'UpdatedAt'
    ];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader(sh, headers.length);
    sh.autoResizeColumns(1, headers.length);
    Logger.log('Sheet "MITRA_PROFILE" berhasil dibuat.');
  } else {
    Logger.log('Sheet "MITRA_PROFILE" sudah ada, dilewati (tidak diubah).');
  }

  if (!ss.getSheetByName('MITRA_PKS')) {
    const sh = ss.insertSheet('MITRA_PKS');
    const headers = [
      'PKS_ID', 'Mitra_ID', 'NomorPKS', 'NamaPKS', 'TanggalMulai', 'TanggalBerakhir',
      'CatatanLegal', 'CreatedAt'
    ];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader(sh, headers.length);
    sh.autoResizeColumns(1, headers.length);
    Logger.log('Sheet "MITRA_PKS" berhasil dibuat.');
  } else {
    Logger.log('Sheet "MITRA_PKS" sudah ada, dilewati (tidak diubah).');
  }

  if (!ss.getSheetByName('MITRA_PKS_DOCUMENTS')) {
    const sh = ss.insertSheet('MITRA_PKS_DOCUMENTS');
    const headers = [
      'Doc_ID', 'Mitra_ID', 'PKS_ID', 'NamaDokumen', 'JenisDokumen', 'Versi',
      'TanggalDokumen', 'TanggalUpload', 'Pengunggah', 'DriveFileID', 'StatusDokumen', 'Catatan'
    ];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader(sh, headers.length);
    sh.autoResizeColumns(1, headers.length);
    Logger.log('Sheet "MITRA_PKS_DOCUMENTS" berhasil dibuat.');
  } else {
    Logger.log('Sheet "MITRA_PKS_DOCUMENTS" sudah ada, dilewati (tidak diubah).');
  }

  if (!ss.getSheetByName('MITRA_COMPANY_VISIT')) {
    const sh = ss.insertSheet('MITRA_COMPANY_VISIT');
    const headers = [
      'Visit_ID', 'Mitra_ID', 'TanggalKunjungan', 'Waktu', 'Lokasi', 'TujuanKunjungan',
      'PesertaInternal', 'PesertaMitra', 'PICMitra', 'Agenda', 'RingkasanHasil', 'TemuanUtama',
      'KebutuhanMitra', 'PotensiKerjaSama', 'KendalaIsu', 'Risiko', 'Kesepakatan', 'NextAction',
      'PICInternal', 'TargetPenyelesaian', 'Prioritas', 'StatusTindakLanjut', 'TanggalFollowUp',
      'LampiranDriveID', 'CreatedAt'
    ];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader(sh, headers.length);
    sh.autoResizeColumns(1, headers.length);
    Logger.log('Sheet "MITRA_COMPANY_VISIT" berhasil dibuat.');
  } else {
    Logger.log('Sheet "MITRA_COMPANY_VISIT" sudah ada, dilewati (tidak diubah).');
  }

  Logger.log('=== Setup Profile Mitra selesai. Sheet existing (Supplier, Buyer, PenjualanTransaksi, PembelianTransaksi, Outstanding, dst) tidak disentuh sama sekali. ===');
}

// ------------------------------------------------------------
// MODUL PENETAPAN HARGA (BOBOT SUPPLIER) — mengikuti prinsip
// Gold Studio: bobot per supplier (total 100%) -> System
// Reference Price (rata-rata tertimbang) -> Harga Jual =
// Ref x (1 + Margin%). TANPA approval, langsung tersimpan.
// Jalankan SEKALI dari editor untuk menambah 2 sheet baru ke
// database yang sudah ada. Tidak menyentuh sheet lain.
// ------------------------------------------------------------
function addPenetapanHargaBobotModule() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (!id) throw new Error('DB_SPREADSHEET_ID belum diset — jalankan setupDatabase() dulu.');
  const ss = SpreadsheetApp.openById(id);

  if (!ss.getSheetByName('BobotSupplier')) {
    const sh = ss.insertSheet('BobotSupplier');
    const headers = ['SupplierID', 'Nama', 'Bobot', 'Aktif'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader(sh, headers.length);
    sh.autoResizeColumns(1, headers.length);
    Logger.log('Sheet "BobotSupplier" berhasil dibuat.');
  } else {
    Logger.log('Sheet "BobotSupplier" sudah ada, dilewati (tidak diubah).');
  }

  if (!ss.getSheetByName('PenetapanHargaBobot')) {
    const sh2 = ss.insertSheet('PenetapanHargaBobot');
    const headers2 = [
      'ID', 'Tanggal', 'Waktu', 'InputBy', 'HargaSupplierJSON',
      'SysRefPrice', 'Margin', 'HargaJual', 'CreatedAt'
    ];
    sh2.getRange(1, 1, 1, headers2.length).setValues([headers2]);
    formatHeader(sh2, headers2.length);
    sh2.autoResizeColumns(1, headers2.length);
    Logger.log('Sheet "PenetapanHargaBobot" berhasil dibuat.');
  } else {
    Logger.log('Sheet "PenetapanHargaBobot" sudah ada, dilewati (tidak diubah).');
  }

  Logger.log('=== Setup Penetapan Harga (Bobot Supplier) selesai. Sheet existing lain tidak disentuh sama sekali. ===');
}

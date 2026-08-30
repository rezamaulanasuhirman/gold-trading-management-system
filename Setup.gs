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

// ------------------------------------------------------------
// MODUL KELENGKAPAN DOKUMEN SUPPLIER TRADING BULION
// Enhancement pada Monitoring Supplier PKS (halaman Profile Mitra,
// tab baru "Kelengkapan Dokumen"). Master checklist dokumen (~20
// item, jarang berubah) SUDAH DIPINDAH jadi konstanta backend
// (CHECKLIST_MASTER_DOKUMEN di Code.gs) — lebih maintainable
// daripada sheet terpisah untuk data referensi sekecil ini, dan
// mengurangi satu sheet + satu read per request. Jadi function
// setup ini SEKARANG HANYA membuat SATU sheet baru: idempotent,
// cek getSheetByName dulu, TIDAK PERNAH menimpa/menghapus data.
// Jalankan SEKALI dari editor Apps Script.
// ------------------------------------------------------------
function addSupplierDocumentChecklistModule() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (!id) throw new Error('DB_SPREADSHEET_ID belum diset — jalankan setupDatabase() dulu.');
  const ss = SpreadsheetApp.openById(id);

  if (!ss.getSheetByName('SupplierDocumentChecklist')) {
    const sh = ss.insertSheet('SupplierDocumentChecklist');
    const headers = [
      'ID', 'SupplierNama', 'Kode', 'Status', 'TanggalDiterima',
      'TanggalBerlaku', 'Catatan', 'LinkDokumen', 'UpdatedAt', 'UpdatedBy'
    ];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader(sh, headers.length);
    sh.autoResizeColumns(1, headers.length);
    Logger.log('Sheet "SupplierDocumentChecklist" berhasil dibuat (KOSONG — isi lewat importInitialSupplierDocumentData() sekali, lalu diedit langsung dari UI Detail Supplier).');
  } else {
    Logger.log('Sheet "SupplierDocumentChecklist" sudah ada, dilewati (tidak diubah, tidak ditimpa).');
  }

  Logger.log('=== Setup Kelengkapan Dokumen Supplier selesai. Sheet existing lain (Supplier, MITRA_*, Master_Customer, dst) tidak disentuh sama sekali. ===');
}

// ------------------------------------------------------------
// MIGRASI AWAL (SEKALI JALAN) — bukan bagian dari flow operasional
// aplikasi. Fungsi ini HANYA untuk memindahkan data checklist dari
// file referensi "PEMENUHAN DOKUMEN SUPPLIER TRADING BULION" ke
// sheet SupplierDocumentChecklist SATU KALI di awal. TIDAK ada
// tombol/menu di UI yang memanggil fungsi ini — jalankan manual
// dari dropdown fungsi di editor Apps Script (pilih
// "importInitialSupplierDocumentData", klik Run) kalau diperlukan
// migrasi ulang/debugging. Setelah migrasi, seluruh operasional
// (baca & update status dokumen) memakai database Gold Trading
// (SupplierDocumentChecklist) langsung — TIDAK bergantung ke file
// Excel/spreadsheet referensi lagi.
//
// Tidak bergantung pada nama tab tertentu (ANTAM/AMMAN/dst) — sheet
// APAPUN di spreadsheet sumber yang punya baris header mengandung
// kata "Status" akan diproses, dan nama tab dipakai sebagai nama
// supplier. Baris yang tidak match dengan checklist master dilaporkan
// di Logger untuk diperiksa manual, TIDAK ditebak/dikarang.
// ------------------------------------------------------------
function importInitialSupplierDocumentData() {
  requireRole('Admin');
  const SOURCE_SPREADSHEET_ID = '1goQHUeWB-qTrN5hZ6uQE2w_YseDglAfGhBR5cwp78Qo'; // file referensi PEMENUHAN DOKUMEN SUPPLIER TRADING BULION

  const sourceSs = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const sheets = sourceSs.getSheets();
  const user = getCurrentUser();
  const now = new Date();

  let itemsImported = 0;
  const sheetsProcessed = [];
  const sheetsSkipped = [];
  const unmatchedRows = [];

  sheets.forEach(sh => {
    const supplierName = sh.getName().trim();
    const values = sh.getDataRange().getValues();

    let headerRowIdx = -1, colStatus = -1, colLabel = 0, colTerima = -1, colBerlaku = -1, colCatatan = -1, colLink = -1;
    for (let r = 0; r < Math.min(15, values.length); r++) {
      const rowText = values[r].map(c => normalizeChecklistText(c));
      const idx = rowText.findIndex(c => c.indexOf('status') > -1);
      if (idx > -1) {
        headerRowIdx = r;
        colStatus = idx;
        rowText.forEach((c, ci) => {
          if (c.indexOf('nama dokumen') > -1 || c === 'dokumen' || c.indexOf('checklist') > -1 || c.indexOf('item') > -1) colLabel = ci;
          if (c.indexOf('terima') > -1) colTerima = ci;
          if (c.indexOf('berlaku') > -1 || c.indexOf('expired') > -1 || c.indexOf('expiry') > -1) colBerlaku = ci;
          if (c.indexOf('catatan') > -1 || c.indexOf('keterangan') > -1) colCatatan = ci;
          if (c.indexOf('link') > -1 || c.indexOf('url') > -1) colLink = ci;
        });
        break;
      }
    }
    if (headerRowIdx === -1) { sheetsSkipped.push(supplierName + ' (tidak ada header "Status" — dilewati)'); return; }

    let importedFromThisSheet = 0;
    for (let r = headerRowIdx + 1; r < values.length; r++) {
      const row = values[r];
      const label = row[colLabel];
      if (!label || !String(label).trim()) continue;
      const kode = matchChecklistCode(label);
      if (!kode) { unmatchedRows.push({ supplier: supplierName, baris: r + 1, teks: String(label) }); continue; }

      const rawStatus = row[colStatus];
      const status = normalizeChecklistStatus(rawStatus);
      const tanggalDiterima = colTerima > -1 ? row[colTerima] : '';
      const tanggalBerlaku = colBerlaku > -1 ? row[colBerlaku] : '';
      const catatan = colCatatan > -1 ? row[colCatatan] : '';
      const linkDokumen = colLink > -1 ? row[colLink] : '';

      saveSupplierDocumentChecklistItemInternal(supplierName, kode, status, tanggalDiterima, tanggalBerlaku, catatan, linkDokumen, now, user.Email);
      itemsImported++;
      importedFromThisSheet++;
    }
    if (importedFromThisSheet > 0) sheetsProcessed.push(supplierName + ' (' + importedFromThisSheet + ' item)');
    else sheetsSkipped.push(supplierName + ' (0 item cocok)');
  });

  logAudit('IMPORT', 'SupplierDocumentChecklist', SOURCE_SPREADSHEET_ID, itemsImported + ' item diimpor dari ' + sheetsProcessed.length + ' sheet (migrasi awal).');

  Logger.log('=== Migrasi Initial Supplier Document Data selesai ===');
  Logger.log('Total item diimpor: ' + itemsImported);
  Logger.log('Sheet diproses: ' + (sheetsProcessed.join(' | ') || '-'));
  Logger.log('Sheet dilewati: ' + (sheetsSkipped.join(' | ') || '-'));
  if (unmatchedRows.length) {
    Logger.log('Baris tidak dikenali (isi manual lewat UI Detail Supplier):');
    unmatchedRows.forEach(u => Logger.log('  ' + u.supplier + ' baris ' + u.baris + ': "' + u.teks + '"'));
  }
  return { itemsImported: itemsImported, sheetsProcessed: sheetsProcessed, sheetsSkipped: sheetsSkipped, unmatchedRows: unmatchedRows };
}

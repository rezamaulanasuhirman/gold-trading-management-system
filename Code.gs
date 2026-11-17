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
  PENETAPAN_BOBOT: 'PenetapanHargaBobot',
  MONITORING_PKS: 'MonitoringPKSSupplier',
  MASTER_HUT: 'Master_HUT',
  HUT_EMAIL_LOG: 'HUT_EmailLog',
  EGR_SESI_SNAPSHOT: 'EGR_Sesi_Snapshot',
  MONITORING_PO: 'MonitoringDeliveryPO',
  KINERJA_BULANAN_MANUAL: 'KinerjaBulananManual',
  KINERJA_BULANAN_TEMPLATE: 'Kinerja Bulanan'
};

// 8 dokumen kelengkapan PKS yang dipantau per supplier. Key di sini = nama
// kolom di sheet MonitoringPKSSupplier, values di setiap kolom hanya boleh
// salah satu dari MONITORING_PKS_STATES.
const MONITORING_PKS_DOCS = [
  { key: 'EDD', label: 'EDD (DIV Kepatuhan)' },
  { key: 'LDD', label: 'LDD (Div Legal)' },
  { key: 'RencanaBisnis', label: 'Rencana Bisnis (Dept Ops Trading)' },
  { key: 'StudiKelayakan', label: 'Studi Kelayakan Pendayagunaan Aset Tetap' },
  { key: 'KajianRisiko', label: 'Kajian Risiko' },
  { key: 'KajianLegal', label: 'Kajian Legal' },
  { key: 'KajianKepatuhan', label: 'Kajian Kepatuhan' },
  { key: 'PaktaIntegritas', label: 'Pakta Integritas' }
];
const MONITORING_PKS_STATES = ['Belum', 'Proses', 'Selesai', 'N/A'];

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
  if (e && e.parameter && e.parameter.debugPriceV3Header === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugPriceV3Header_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugMpksDump === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugMpksDump_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugDedupeMpks === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugDedupeMpks_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugDbUrl === '1') {
    const dbId = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
    return ContentService.createTextOutput(JSON.stringify({ dbSpreadsheetId: dbId, url: dbId ? 'https://docs.google.com/spreadsheets/d/' + dbId + '/edit' : null })).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugBulkAddSupervisors === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugBulkAddSupervisors_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugLaporanHarian === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugLaporanHarian_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugListSheetsWithPayment === '1') {
    const ss = getDb();
    const sheets = ss.getSheets();
    const result = sheets.map(function (sh) {
      const lastCol = sh.getLastColumn();
      const lastRow = sh.getLastRow();
      let headers = [];
      let hasPaymentHeader = false;
      if (lastCol > 0 && lastRow > 0) {
        headers = sh.getRange(1, 1, 1, Math.min(lastCol, 40)).getValues()[0];
        hasPaymentHeader = headers.some(function (h) { return String(h || '').trim() === 'Payment'; });
      }
      return { name: sh.getName(), lastRow: lastRow, lastCol: lastCol, headers: headers, hasPaymentHeader: hasPaymentHeader };
    });
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugRkAddAliasRerun === '1') {
    const buyerNama = e.parameter.buyer || '';
    const alias = e.parameter.alias || '';
    const sh = rkAliasSheet_();
    const id = 'ALIAS-' + new Date().getTime();
    sh.appendRow([id, buyerNama, alias, new Date(), 'debug-endpoint']);
    logAudit('CREATE', 'Rekonsiliasi.Alias', id, buyerNama + ' -> ' + alias);
    const recon = rkRunReconciliationEngine_();
    return ContentService.createTextOutput(JSON.stringify({ aliasId: id, buyerNama: buyerNama, alias: alias, recon: recon })).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugRkInvoiceRaw === '1') {
    const nomorInvoice = e.parameter.invoice || '';
    return ContentService.createTextOutput(JSON.stringify(_debugRkInvoiceRaw_(nomorInvoice))).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugRkInvoiceCheck === '1') {
    const buyerSubstr = e.parameter.buyer || '';
    return ContentService.createTextOutput(JSON.stringify(_debugRkInvoiceCheck_(buyerSubstr))).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugRkResetAll === '1') {
    return ContentService.createTextOutput(JSON.stringify(rkResetAllData_('debug-endpoint'))).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugRkResetTanggal === '1') {
    // URL param kadang kebawa karakter liar (trailing backslash dsb dari
    // browser/copy-paste) — ambil 10 karakter pertama saja (format YYYY-MM-DD).
    const tanggalRaw = String(e.parameter.tanggal || '2026-09-09').trim();
    const tanggal = tanggalRaw.slice(0, 10);
    return ContentService.createTextOutput(JSON.stringify(rkResetHasilUntukTanggal_(tanggal))).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugRkPaymentColumn === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugRkPaymentColumn_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugPrintV3Audit === '1') {
    return ContentService.createTextOutput(JSON.stringify(auditPrintV3Diagnostic())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugPphTest === '1') {
    const tanggal = e.parameter.tanggal || null;
    const tahap = e.parameter.tahap || 1;
    let out;
    try {
      out = {
        trading: getPenetapanHargaTradingReguler(tanggal, tahap),
        antam: getPenetapanHargaAntam(tanggal),
        tahapOptions: getPenetapanHargaTahapOptions(tanggal)
      };
    } catch (err) {
      out = { error: err.message, stack: err.stack };
    }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugPphFolder === '1') {
    let out;
    try {
      const folder = DriveApp.getFolderById(e.parameter.folderId);
      const files = folder.getFiles();
      const list = [];
      while (files.hasNext()) {
        const f = files.next();
        list.push({ id: f.getId(), name: f.getName(), mimeType: f.getMimeType(), size: f.getSize() });
      }
      out = { ok: true, files: list };
    } catch (err) {
      out = { ok: false, error: err.message };
    }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugPphAsset === '1') {
    let out;
    try {
      const file = DriveApp.getFileById(e.parameter.fileId);
      const blob = file.getBlob();
      const bytes = blob.getBytes();
      out = { ok: true, mimeType: blob.getContentType(), byteLength: bytes.length, base64Preview: Utilities.base64Encode(bytes).slice(0, 200) };
    } catch (err) {
      out = { ok: false, error: err.message };
    }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugSyncCheck === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugSyncCheck_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugFixPembelianTanggal === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugFixPembelianTanggal_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugLaporanRevisiInvestigasi === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugLaporanRevisiInvestigasi_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugFixPembelianPOFull === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugFixPembelianPOFull_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugInstallHutTrigger === '1') {
    installHutNotificationTrigger_();
    return ContentService.createTextOutput(JSON.stringify({ installed: true, status: isHutTriggerInstalled_() })).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugHutTestEmail === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugHutTestEmail_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugEgrSaldo === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugEgrSaldo_(e.parameter.tanggal))).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugPjlPickupCheck === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugPjlPickupCheck_())).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.debugF1SourceCheck === '1') {
    return ContentService.createTextOutput(JSON.stringify(_debugF1SourceCheck_())).setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gold Trading Management System')
    .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/favicon3.ico')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// TEMP DIAGNOSTIC — investigasi untuk revisi Laporan Harian: header
// Input_Jual_EGR/Input_Beli_EGR (untuk detail baris ETF) + baris 8
// sheet Rekap kolom R-AG (untuk cek kenapa ANTAM tidak muncul di
// Piutang). Read-only. (dihapus setelah selesai)
function _debugLaporanRevisiInvestigasi_() {
  const result = {};

  [EGR_SHEET.JUAL, EGR_SHEET.BELI].forEach(name => {
    try {
      const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
      const sheet = getEgrSheet_(ss, name);
      if (!sheet) { result[name] = { error: 'Sheet tidak ditemukan' }; return; }
      const values = sheet.getDataRange().getValues();
      const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
      if (headerRowIdx === -1) { result[name] = { error: 'Header tidak terdeteksi' }; return; }
      result[name] = {
        headerRowIdx: headerRowIdx,
        headerRow: values[headerRowIdx],
        baris3Contoh: values.slice(headerRowIdx + 1, headerRowIdx + 4)
      };
    } catch (err) {
      result[name] = { error: err.message };
    }
  });

  try {
    const ss2 = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
    const rekap = ss2.getSheetByName(REKAP_SHEET_NAME);
    if (!rekap) {
      result.rekapPiutang = { error: 'Sheet Rekap tidak ditemukan' };
    } else {
      result.rekapPiutang = REKAP_PIUTANG_SUPPLIER_COLS.map(colSet => ({
        labelCol: colSet.labelCol,
        label: rekap.getRange(REKAP_PIUTANG_ROW.header, colSet.labelCol).getValue(),
        goldInTransit: rekap.getRange(REKAP_PIUTANG_ROW.goldInTransit, colSet.valueCol).getValue(),
        uangMuka: rekap.getRange(REKAP_PIUTANG_ROW.uangMuka, colSet.valueCol).getValue()
      }));
      // Scan lebih lebar (kolom 15-40) di baris header, siapa tahu ANTAM
      // sebenarnya di kolom lain / posisinya bergeser.
      result.rekapHeaderRowScan = [];
      for (let c = 15; c <= 40; c++) {
        const v = rekap.getRange(REKAP_PIUTANG_ROW.header, c).getValue();
        if (v) result.rekapHeaderRowScan.push({ col: c, value: v });
      }
    }
  } catch (err) {
    result.rekapPiutang = { error: err.message };
  }

  return result;
}

// TEMP FIX — koreksi PENUH (Tanggal, Seller, Qty, Harga, AfterDiskon,
// TotalHarga, PajakWapu, TotalHargaIncludePajak, InvoiceRef) untuk baris
// PembelianTransaksi dengan Nomor PO 00620/00621. Datanya dibaca LANGSUNG
// dari sheet sumber F.2.Input_Trans_Beli (dicocokkan by Nomor PO persis,
// pakai header asli, bukan hardcode nama supplier) — supaya baris yang
// PO-nya dipakai ulang di sheet sumber untuk transaksi berbeda kembali
// merefleksikan transaksi yang BENAR-BENAR ada di sumber saat ini, bukan
// data lama dari transaksi sebelumnya yang kebetulan pakai PO sama.
// (dihapus setelah selesai)
function _debugFixPembelianPOFull_() {
  const targetPOs = ['00620/PO-00017/2026', '00621/PO-00017/2026'];

  const sourceSs = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sourceSheet = sourceSs.getSheetByName('F.2.Input_Trans_Beli');
  if (!sourceSheet) throw new Error('Sheet "F.2.Input_Trans_Beli" tidak ditemukan.');

  const values = sourceSheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['tanggal', 'purchase', 'seller', 'qty', 'harga']);
  if (headerRowIdx === -1) throw new Error('Baris header tidak terdeteksi di F.2.Input_Trans_Beli.');
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    nomorPO: findColByKeywords(headerMap, ['nomor', 'purchase']),
    seller: findColByKeywords(headerMap, ['seller']),
    qty: findColByKeywords(headerMap, ['qty']),
    afterDiskon: findColByKeywords(headerMap, ['after', 'diskon']),
    harga: findColByKeywords(headerMap, ['harga'], ['after', 'total', 'simulasi']),
    invoice: findColByKeywords(headerMap, ['invoice'])
  };

  const sourceRows = values.slice(headerRowIdx + 1);
  const existingRows = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL);
  const results = [];

  targetPOs.forEach(po => {
    const srcRow = sourceRows.find(row => String(getByCol(row, col.nomorPO) || '').trim() === po);
    if (!srcRow) { results.push({ nomorPO: po, error: 'Tidak ditemukan di sheet sumber F.2.Input_Trans_Beli' }); return; }

    const existing = existingRows.find(r => String(r.NomorPO || '').trim() === po);
    if (!existing) { results.push({ nomorPO: po, error: 'Tidak ditemukan di database PembelianTransaksi' }); return; }

    const tanggal = getByCol(srcRow, col.tanggal);
    const seller = getByCol(srcRow, col.seller);
    const qty = Number(getByCol(srcRow, col.qty)) || 0;
    const harga = Number(getByCol(srcRow, col.harga)) || 0;
    const afterDiskonRaw = col.afterDiskon !== -1 ? Number(getByCol(srcRow, col.afterDiskon)) : 0;
    const afterDiskon = afterDiskonRaw || harga;
    const invoiceRef = col.invoice !== -1 ? getByCol(srcRow, col.invoice) : existing.InvoiceRef;
    const totalHarga = afterDiskon * qty;
    const pajakWapu = totalHarga * PAJAK_WAPU_RATE;
    const totalInclude = totalHarga - pajakWapu;

    const sebelum = { tanggal: extractDateOnly(existing.Tanggal), seller: existing.Seller, qty: existing.Qty, harga: existing.Harga, totalHarga: existing.TotalHarga };

    updateRowById(SHEET_NAMES.PEMBELIAN_FULL, 'ID', existing.ID, {
      Tanggal: tanggal, Seller: seller, Qty: qty, Harga: harga, AfterDiskon: afterDiskon,
      TotalHarga: totalHarga, PajakWapu: pajakWapu, TotalHargaIncludePajak: totalInclude,
      InvoiceRef: invoiceRef
    });

    results.push({
      nomorPO: po, id: existing.ID,
      sebelum: sebelum,
      sesudah: { tanggal: extractDateOnly(tanggal), seller: seller, qty: qty, harga: harga, afterDiskon: afterDiskon, totalHarga: totalHarga }
    });
  });

  logAudit('UPDATE', 'PembelianTransaksi', targetPOs.join(','), 'Koreksi penuh data (root cause: PO dipakai ulang di sumber)');
  return results;
}

// TEMP FIX — koreksi manual Tanggal untuk PBL-620/PBL-621 (PO 00620/00621)
// jadi 10 Agustus 2026, sesuai konfirmasi user dari screenshot sheet
// sumber. Baris ini dilindungi dari auto-sync karena Catatan-nya sudah
// diubah manual jadi "PEMENUHAN ETF", jadi dikoreksi langsung di sini
// satu kali saja — bukan lewat sync biasa. (dihapus setelah selesai)
function _debugFixPembelianTanggal_() {
  const targetDate = new Date('2026-08-10T00:00:00+07:00');
  const results = [];
  ['PBL-620', 'PBL-621'].forEach(id => {
    try {
      updateRowById(SHEET_NAMES.PEMBELIAN_FULL, 'ID', id, { Tanggal: targetDate });
      results.push({ id: id, aksi: 'Tanggal diupdate ke 2026-08-10' });
    } catch (err) {
      results.push({ id: id, error: err.message });
    }
  });
  logAudit('UPDATE', 'PembelianTransaksi', 'PBL-620,PBL-621', 'Koreksi manual Tanggal ke 2026-08-10');
  return results;
}

// TEMP DIAGNOSTIC — replikasi logic syncTransaksiBaruPenjualan/Pembelian
// TAPI read-only (tidak menulis apa pun), untuk lihat 15 baris TERAKHIR
// dari sheet sumber F.1/F.2: nilai mentah, hasil deteksi kolom, apakah
// lolos validasi, dan apakah sudah ada di database. (dihapus setelah selesai)
function _debugSyncCheck_() {
  const result = {};

  // --- F.1.Input_Trans_Jual (Penjualan) ---
  try {
    const sourceSs = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
    const sourceSheet = sourceSs.getSheetByName('F.1.Input_Trans_Jual');
    if (!sourceSheet) throw new Error('Sheet "F.1.Input_Trans_Jual" tidak ditemukan.');

    const values = sourceSheet.getDataRange().getValues();
    const headerRowIdx = findHeaderRowIndex(values, ['tanggal', 'invoice', 'pembeli', 'qty', 'harga']);
    if (headerRowIdx === -1) throw new Error('Header tidak terdeteksi.');
    const headerMap = buildHeaderMap(values[headerRowIdx]);
    const col = {
      tanggal: findColByKeywords(headerMap, ['tanggal'], ['pengambilan', 'penjemputan']),
      nomorInvoice: findColByKeywords(headerMap, ['nomor', 'invoice']),
      hargaPerGram: findColByKeywords(headerMap, ['harga'], ['total', 'after', 'simulasi']),
      qty: findColByKeywords(headerMap, ['qty']),
      pembeli: findColByKeywords(headerMap, ['pembeli'])
    };
    const existingByInvoice = {};
    sheetToObjects(SHEET_NAMES.PENJUALAN_FULL).forEach(r => { existingByInvoice[String(r.NomorInvoice || '').trim()] = true; });

    const rows = values.slice(headerRowIdx + 1);
    const rowsBerisi = rows.filter(row => {
      const t = getByCol(row, col.tanggal);
      return t && String(t).trim() !== '' && String(t).trim().toLowerCase() !== 'end';
    });
    result.penjualan = {
      totalBarisSumber: rows.length,
      totalBarisBerisiData: rowsBerisi.length,
      kolomTerdeteksi: col,
      baris15Terakhir: rowsBerisi.slice(-15).map(row => {
        const tanggal = getByCol(row, col.tanggal);
        const nomorInvoice = String(getByCol(row, col.nomorInvoice) || '').trim();
        const hargaPerGram = Number(getByCol(row, col.hargaPerGram)) || 0;
        const qty = Number(getByCol(row, col.qty)) || 0;
        return {
          tanggal: tanggal, nomorInvoice: nomorInvoice, hargaPerGram: hargaPerGram, qty: qty,
          pembeli: getByCol(row, col.pembeli),
          lolosValidasi: !!(tanggal && nomorInvoice && hargaPerGram > 0 && qty > 0),
          sudahAdaDiDatabase: !!existingByInvoice[nomorInvoice]
        };
      })
    };
  } catch (err) {
    result.penjualan = { error: err.message };
  }

  // --- F.2.Input_Trans_Beli (Pembelian) ---
  try {
    const sourceSs2 = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
    const sourceSheet2 = sourceSs2.getSheetByName('F.2.Input_Trans_Beli');
    if (!sourceSheet2) throw new Error('Sheet "F.2.Input_Trans_Beli" tidak ditemukan.');

    const values2 = sourceSheet2.getDataRange().getValues();
    const headerRowIdx2 = findHeaderRowIndex(values2, ['tanggal', 'purchase', 'seller', 'qty', 'harga']);
    if (headerRowIdx2 === -1) throw new Error('Header tidak terdeteksi.');
    const headerMap2 = buildHeaderMap(values2[headerRowIdx2]);
    const col2 = {
      tanggal: findColByKeywords(headerMap2, ['tanggal']),
      nomorPO: findColByKeywords(headerMap2, ['nomor', 'purchase']),
      seller: findColByKeywords(headerMap2, ['seller']),
      qty: findColByKeywords(headerMap2, ['qty']),
      harga: findColByKeywords(headerMap2, ['harga'], ['after', 'total', 'simulasi']),
      afterDiskon: findColByKeywords(headerMap2, ['after', 'diskon'])
    };
    const existingByPO = {};
    sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL).forEach(r => { existingByPO[String(r.NomorPO || '').trim()] = true; });

    const rows2 = values2.slice(headerRowIdx2 + 1);
    const rows2Berisi = rows2.filter(row => {
      const t = getByCol(row, col2.tanggal);
      return t && String(t).trim() !== '' && String(t).trim().toLowerCase() !== 'end';
    });
    result.pembelian = {
      totalBarisSumber: rows2.length,
      totalBarisBerisiData: rows2Berisi.length,
      kolomTerdeteksi: col2,
      baris15Terakhir: rows2Berisi.slice(-15).map(row => {
        const tanggal = getByCol(row, col2.tanggal);
        const nomorPO = String(getByCol(row, col2.nomorPO) || '').trim();
        const qty = Number(getByCol(row, col2.qty)) || 0;
        const harga = Number(getByCol(row, col2.harga)) || 0;
        const afterDiskonRaw = col2.afterDiskon !== -1 ? Number(getByCol(row, col2.afterDiskon)) : 0;
        const afterDiskon = afterDiskonRaw || harga;
        return {
          tanggal: tanggal, nomorPO: nomorPO, harga: harga, afterDiskon: afterDiskon, qty: qty,
          seller: getByCol(row, col2.seller),
          lolosValidasi: !!(tanggal && nomorPO && qty > 0 && afterDiskon > 0),
          sudahAdaDiDatabase: !!existingByPO[nomorPO]
        };
      })
    };
  } catch (err) {
    result.pembelian = { error: err.message };
  }

  return result;
}

// TEMP DIAGNOSTIC — versi JSON dari diagnosaTransaksiHariIni(), untuk cek
// kenapa transaksi yang baru diinput tidak muncul di Laporan Harian.
// Read-only. (dihapus setelah selesai)
function _debugLaporanHarian_() {
  const today = formatDateOnly(new Date());
  const pj = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const pb = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL);

  const targetInvoices = ['00620/INV-00017/2026', '00621/INV-00017/2026'];
  const targetPOs = ['00620/PO-00017/2026', '00621/PO-00017/2026'];

  function mapPj(row) {
    return {
      ID: row.ID, NomorInvoice: row.NomorInvoice,
      Tanggal: row.Tanggal, tipeTanggal: typeof row.Tanggal,
      extractDateOnlyHasil: extractDateOnly(row.Tanggal),
      cocokHariIni: extractDateOnly(row.Tanggal) === today,
      Pembeli: row.Pembeli, TotalHarga: row.TotalHarga, CreatedAt: row.CreatedAt, Catatan: row.Catatan
    };
  }
  function mapPb(row) {
    return {
      ID: row.ID, NomorPO: row.NomorPO,
      Tanggal: row.Tanggal, tipeTanggal: typeof row.Tanggal,
      extractDateOnlyHasil: extractDateOnly(row.Tanggal),
      cocokHariIni: extractDateOnly(row.Tanggal) === today,
      Seller: row.Seller, TotalHarga: row.TotalHarga, CreatedAt: row.CreatedAt, Catatan: row.Catatan
    };
  }

  const pjByCreatedAtDesc = pj.slice().sort((a, b) => new Date(b.CreatedAt || 0) - new Date(a.CreatedAt || 0));
  const pbByCreatedAtDesc = pb.slice().sort((a, b) => new Date(a.CreatedAt || 0) - new Date(b.CreatedAt || 0)).reverse();

  return {
    tanggalHariIniMenurutSistem: today,
    totalBarisPenjualanTransaksi: pj.length,
    baris5TerakhirByPosisiFisik: pj.slice(-5).map(mapPj),
    baris5TerbaruByCreatedAt: pjByCreatedAtDesc.slice(0, 5).map(mapPj),
    cariInvoiceTarget: targetInvoices.map(inv => {
      const found = pj.find(r => String(r.NomorInvoice || '').trim() === inv);
      return found ? mapPj(found) : { NomorInvoice: inv, ditemukan: false };
    }),
    totalBarisPembelianTransaksi: pb.length,
    baris5TerakhirPembelianByPosisiFisik: pb.slice(-5).map(mapPb),
    baris5TerbaruPembelianByCreatedAt: pbByCreatedAtDesc.slice(0, 5).map(mapPb),
    cariPOTarget: targetPOs.map(po => {
      const found = pb.find(r => String(r.NomorPO || '').trim() === po);
      return found ? mapPb(found) : { NomorPO: po, ditemukan: false };
    })
  };
}

// TEMP — daftarkan 5 orang yang punya akses Editor di spreadsheet DB
// sebagai role Supervisor di sheet Users aplikasi (idempotent: kalau
// sudah terdaftar dan rolenya Supervisor/Admin, dilewati; kalau sudah
// terdaftar tapi rolenya lebih rendah, di-upgrade ke Supervisor; kalau
// belum terdaftar, ditambahkan baru). (dihapus setelah selesai)
function _debugBulkAddSupervisors_() {
  const targets = [
    { email: 'anggun.aprilia@pegadaian.co.id', nama: 'Anggun Aprilia Purnamasari' },
    { email: 'eva.anggytanie@pegadaian.co.id', nama: 'Eva Anggytanie' },
    { email: 'frida.adrias@pegadaian.co.id', nama: 'Frida Sakinah Adrias' },
    { email: 'muhammad.ansar2@pegadaian.co.id', nama: 'Muhammad Arya Ansar' },
    { email: 'muhammad.febrian@pegadaian.co.id', nama: 'Muhammad Reza Febrian' }
  ];

  const existing = sheetToObjects(SHEET_NAMES.USERS);
  const results = [];

  targets.forEach(t => {
    const found = existing.find(u => String(u.Email || '').toLowerCase().trim() === t.email.toLowerCase());
    if (!found) {
      const id = addUser(t.email, t.nama, 'Supervisor');
      results.push({ email: t.email, aksi: 'Ditambahkan baru sebagai Supervisor (' + id + ')' });
    } else if (ROLE_LEVEL[found.Role] >= ROLE_LEVEL['Supervisor']) {
      results.push({ email: t.email, aksi: 'Sudah role "' + found.Role + '", dilewati (tidak diturunkan)' });
    } else {
      updateUserRole(found.UserID, 'Supervisor');
      results.push({ email: t.email, aksi: 'Role di-upgrade dari "' + found.Role + '" ke Supervisor' });
    }
  });

  return results;
}

// TEMP CLEANUP — hapus baris duplikat di MonitoringPKSSupplier (SupplierID
// sama muncul >1x), sisakan baris dengan UpdatedAt paling baru. Hanya
// menyentuh sheet MonitoringPKSSupplier (dibuat sendiri oleh fitur ini,
// bukan sheet sumber eksternal). (dihapus setelah selesai)
function _debugDedupeMpks_() {
  const sh = getSheet(SHEET_NAMES.MONITORING_PKS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idColIdx = headers.indexOf('SupplierID');
  const updatedAtColIdx = headers.indexOf('UpdatedAt');

  const bestRowBySupplierId = {};
  for (let r = 1; r < values.length; r++) {
    const sid = String(values[r][idColIdx]);
    const current = bestRowBySupplierId[sid];
    if (current === undefined || new Date(values[r][updatedAtColIdx] || 0) >= new Date(values[current][updatedAtColIdx] || 0)) {
      bestRowBySupplierId[sid] = r;
    }
  }

  const rowsToDelete = [];
  for (let r = 1; r < values.length; r++) {
    const sid = String(values[r][idColIdx]);
    if (r !== bestRowBySupplierId[sid]) rowsToDelete.push(r + 1);
  }

  rowsToDelete.sort((a, b) => b - a).forEach(rowNum => sh.deleteRow(rowNum));

  return { totalRowsBefore: values.length - 1, rowsDeleted: rowsToDelete, totalRowsAfter: values.length - 1 - rowsToDelete.length };
}

// TEMP DIAGNOSTIC — dump seluruh isi sheet MonitoringPKSSupplier + jumlah
// baris Supplier aktif dengan nama mengandung 'amman', untuk cek kenapa
// checklist PT AMMAN tidak tersimpan lengkap. Read-only.
// (dihapus setelah selesai)
function _debugMpksDump_() {
  const sh = getSheet(SHEET_NAMES.MONITORING_PKS);
  const values = sh.getDataRange().getValues();
  const suppliers = sheetToObjects(SHEET_NAMES.SUPPLIER).filter(s => /amman/i.test(s.Nama || ''));
  return {
    headers: values[0],
    totalRows: values.length - 1,
    allRows: values.slice(1),
    supplierAmmanEntries: suppliers
  };
}

// TEMP DIAGNOSTIC — cek header kolom sheet Price V3 untuk cari kolom XAU
// (dihapus setelah selesai, read-only tidak mengubah apa pun)
function _debugPriceV3Header_() {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = ss.getSheetByName(PRICE_DISCOVERY_SHEET_NAME);
  if (!sheet) return { error: 'Sheet tidak ditemukan' };
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const result = headerRow.map(function (h, i) {
    return { kolom: String.fromCharCode(65 + i), indexNolBased: i, header: h };
  });
  return result;
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

/**
 * Set RecordID yang pernah tercatat SEND_EMAIL di AuditTrail untuk
 * module tertentu — dipakai untuk status "Sudah dikirim"/"Belum dikirim"
 * pada tabel Riwayat Penjualan/Pembelian, TANPA perlu kolom baru di
 * sheet transaksi (status diturunkan dari log yang sudah ada).
 */
function getDealEmailSentMap_(moduleName) {
  const rows = sheetToObjects(SHEET_NAMES.AUDIT);
  const map = {};
  rows.forEach(r => {
    if (r.Action === 'SEND_EMAIL' && r.Module === moduleName && r.RecordID) map[r.RecordID] = true;
  });
  return map;
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

// ------------------------------------------------------------
// MONITORING KELENGKAPAN DOKUMEN PKS SUPPLIER
// Satu baris per supplier di sheet MonitoringPKSSupplier. Baris
// dibuat otomatis (lazy) saat pertama kali ada field yang diisi —
// tidak menulis apa pun saat hanya dibaca (GET).
// ------------------------------------------------------------

/**
 * Hitung % Progres & Status dari 8 kolom checklist. Kolom bernilai
 * 'N/A' dikeluarkan dari penyebut (tidak dihitung sebagai belum).
 */
// Belum = 0 poin, Proses = 50 poin (dihitung sebagai progres berjalan,
// bukan 0), Selesai = 100 poin. N/A dikeluarkan dari penyebut.
const MONITORING_PKS_POIN = { 'Belum': 0, 'Proses': 50, 'Selesai': 100 };

function computeMonitoringPksProgres_(row) {
  const applicable = MONITORING_PKS_DOCS.filter(d => row[d.key] !== 'N/A');
  const totalPoin = applicable.reduce((sum, d) => sum + (MONITORING_PKS_POIN[row[d.key]] || 0), 0);
  const progres = applicable.length ? Math.round(totalPoin / applicable.length) : 0;
  let status;
  if (applicable.length === 0) status = 'Belum Mulai';
  else if (progres === 100) status = 'Lengkap';
  else if (progres === 0) status = 'Belum Mulai';
  else status = 'Dalam Proses';
  return { progres: progres, status: status };
}

/**
 * Daftar seluruh supplier (dari sheet Supplier, tidak diubah) di-join
 * dengan data monitoring PKS-nya. Supplier yang belum punya baris di
 * MonitoringPKSSupplier ditampilkan dengan nilai default (semua
 * 'Belum'), TANPA menulis apa pun ke sheet (murni baca).
 */
function getMonitoringPksList() {
  const suppliers = sheetToObjects(SHEET_NAMES.SUPPLIER).filter(s => s.Status === 'Aktif');
  const monitoringRows = sheetToObjects(SHEET_NAMES.MONITORING_PKS);
  const monitoringBySupplierId = {};
  // Kalau ada baris duplikat untuk SupplierID yang sama (seharusnya tidak
  // terjadi, tapi jaga-jaga), pakai yang UpdatedAt-nya paling baru.
  monitoringRows.forEach(r => {
    const current = monitoringBySupplierId[r.SupplierID];
    if (!current || new Date(r.UpdatedAt || 0) >= new Date(current.UpdatedAt || 0)) {
      monitoringBySupplierId[r.SupplierID] = r;
    }
  });

  return suppliers.map(s => {
    const existing = monitoringBySupplierId[s.SupplierID];
    const row = { SupplierID: s.SupplierID, SupplierNama: s.Nama };
    MONITORING_PKS_DOCS.forEach(d => { row[d.key] = (existing && existing[d.key]) || 'Belum'; });
    const computed = computeMonitoringPksProgres_(row);
    return {
      supplierId: s.SupplierID,
      supplierNama: s.Nama,
      docs: MONITORING_PKS_DOCS.map(d => ({ key: d.key, label: d.label, value: row[d.key] })),
      progres: existing ? Number(existing.Progres) || computed.progres : computed.progres,
      status: existing ? (existing.Status || computed.status) : computed.status,
      linkDokumen: (existing && existing.LinkDokumen) || '',
      pic: (existing && existing.PIC) || '',
      singkatan: (existing && existing.Singkatan) || '',
      jenisKerjaSama: (existing && existing.JenisKerjaSama) || '',
      batasWaktu: (existing && existing.BatasWaktu) ? extractDateOnly(existing.BatasWaktu) : '',
      teleponPic: (existing && existing.TeleponPIC) || ''
    };
  });
}

/**
 * Cari baris supplier di MonitoringPKSSupplier; buat baru (lazy) kalau
 * belum ada. Mengembalikan { sheet, headers, rowNum, rowValuesObj }.
 */
function getOrCreateMonitoringPksRow_(supplierId, supplierNama) {
  const sh = getSheet(SHEET_NAMES.MONITORING_PKS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idColIdx = headers.indexOf('SupplierID');
  const updatedAtColIdx = headers.indexOf('UpdatedAt');

  // Kalau (seharusnya tidak terjadi) ada beberapa baris untuk SupplierID
  // yang sama, pakai baris dengan UpdatedAt paling baru — jangan asal
  // ambil match pertama, supaya update tidak "nyasar" ke baris basi.
  let best = null;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idColIdx]) === String(supplierId)) {
      if (!best || new Date(values[r][updatedAtColIdx] || 0) >= new Date(values[best][updatedAtColIdx] || 0)) {
        best = r;
      }
    }
  }
  if (best !== null) {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = values[best][i]; });
    return { sh: sh, headers: headers, rowNum: best + 1, obj: obj };
  }

  const newRow = headers.map(h => {
    if (h === 'SupplierID') return supplierId;
    if (h === 'SupplierNama') return supplierNama || '';
    if (MONITORING_PKS_DOCS.some(d => d.key === h)) return 'Belum';
    if (h === 'Progres') return 0;
    if (h === 'Status') return 'Belum Mulai';
    return '';
  });
  sh.appendRow(newRow);
  const obj = {};
  headers.forEach((h, i) => { obj[h] = newRow[i]; });
  return { sh: sh, headers: headers, rowNum: sh.getLastRow(), obj: obj };
}

/**
 * Simpan SATU baris monitoring PKS supplier sekaligus (8 checklist +
 * Link Dokumen + PIC), dipanggil dari tombol "Simpan" per baris di
 * halaman. Menghitung ulang % Progres & Status otomatis.
 */
function saveMonitoringPksRow(supplierId, supplierNama, payload) {
  requireRole('Supervisor');
  payload = payload || {};

  MONITORING_PKS_DOCS.forEach(d => {
    const value = payload[d.key];
    if (MONITORING_PKS_STATES.indexOf(value) === -1) throw new Error('Nilai status tidak valid untuk ' + d.label + ': ' + value);
  });

  const found = getOrCreateMonitoringPksRow_(supplierId, supplierNama);
  MONITORING_PKS_DOCS.forEach(d => { found.obj[d.key] = payload[d.key]; });
  found.obj.LinkDokumen = payload.linkDokumen || '';
  found.obj.PIC = payload.pic || '';

  const computed = computeMonitoringPksProgres_(found.obj);
  found.obj.Progres = computed.progres;
  found.obj.Status = computed.status;
  found.obj.UpdatedAt = new Date();

  const rowValues = found.headers.map(h => found.obj[h]);
  found.sh.getRange(found.rowNum, 1, 1, found.headers.length).setValues([rowValues]);
  logAudit('UPDATE', 'MonitoringPKSSupplier', supplierId, 'Checklist disimpan (' + computed.progres + '%)');

  return { progres: computed.progres, status: computed.status };
}

/**
 * Simpan data Narahubung (PIC) + Target Penyelesaian Kajian utk satu
 * supplier — dipakai modal "Edit Data Supplier" di Monitoring PKS
 * Supplier. Field ini yang otomatis dipakai Draft Nota Dinas (Preview &
 * Download Word, lihat validateNodinDataLengkap_). TIDAK menyentuh
 * checklist/Progres/Status.
 */
function saveMonitoringPksMeta(supplierId, supplierNama, payload) {
  requireRole('Supervisor');
  payload = payload || {};

  const found = getOrCreateMonitoringPksRow_(supplierId, supplierNama);
  found.obj.PIC = String(payload.namaPic || '').trim();
  found.obj.TeleponPIC = String(payload.teleponPic || '').trim();
  found.obj.BatasWaktu = payload.batasWaktu || '';
  found.obj.UpdatedAt = new Date();

  const rowValues = found.headers.map(h => found.obj[h] !== undefined ? found.obj[h] : '');
  found.sh.getRange(found.rowNum, 1, 1, found.headers.length).setValues([rowValues]);
  logAudit('UPDATE', 'MonitoringPKSSupplier', supplierId, 'Data narahubung/target penyelesaian disimpan');

  return {
    pic: found.obj.PIC,
    teleponPic: found.obj.TeleponPIC,
    batasWaktu: found.obj.BatasWaktu ? extractDateOnly(found.obj.BatasWaktu) : ''
  };
}

// 5 kajian yang bisa dimohonkan lewat modal "Tambah Supplier". Kunci di
// sini HARUS salah satu key di MONITORING_PKS_DOCS.
const MONITORING_PKS_INTAKE_KAJIAN = ['EDD', 'LDD', 'KajianRisiko', 'KajianLegal', 'KajianKepatuhan'];

/**
 * Tambah supplier baru sekaligus intake data PKS-nya (dipanggil dari
 * modal "Tambah Supplier" di halaman Monitoring PKS Supplier).
 * Kajian yang dicentang -> 'Proses', yang tidak dicentang -> 'N/A'.
 * RencanaBisnis/StudiKelayakan/PaktaIntegritas TIDAK disentuh di sini,
 * tetap default 'Belum' (mengikuti logika existing checklist manual).
 */
function addSupplierWithPksIntake(payload) {
  requireRole('Supervisor');
  payload = payload || {};
  const nama = String(payload.nama || '').trim();
  if (!nama) throw new Error('Nama perusahaan wajib diisi.');

  const supplierId = addSupplier(nama);
  const found = getOrCreateMonitoringPksRow_(supplierId, nama);

  const kajianDicentang = payload.kajianDicentang || {};
  MONITORING_PKS_INTAKE_KAJIAN.forEach(key => { found.obj[key] = kajianDicentang[key] ? 'Proses' : 'N/A'; });

  found.obj.LinkDokumen = payload.linkDokumen || '';
  found.obj.PIC = payload.namaPic || '';
  found.obj.Singkatan = payload.singkatan || '';
  found.obj.JenisKerjaSama = payload.jenisKerjaSama || '';
  found.obj.BatasWaktu = payload.batasWaktu || '';
  found.obj.TeleponPIC = payload.teleponPic || '';

  const computed = computeMonitoringPksProgres_(found.obj);
  found.obj.Progres = computed.progres;
  found.obj.Status = computed.status;
  found.obj.UpdatedAt = new Date();

  const rowValues = found.headers.map(h => found.obj[h] !== undefined ? found.obj[h] : '');
  found.sh.getRange(found.rowNum, 1, 1, found.headers.length).setValues([rowValues]);
  logAudit('CREATE', 'MonitoringPKSSupplier', supplierId, 'Supplier baru + intake PKS: ' + nama);

  return { supplierId: supplierId };
}

function joinDaftarIndonesia_(items) {
  items = (items || []).filter(Boolean);
  if (items.length === 0) return '';
  if (items.length === 1) return items[0];
  if (items.length === 2) return items[0] + ' dan ' + items[1];
  return items.slice(0, -1).join(', ') + ', dan ' + items[items.length - 1];
}

function formatTanggalPanjangIndo_(dateStr) {
  if (!dateStr) return '-';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return String(dateStr);
  return d.getDate() + ' ' + NAMA_BULAN_ID[d.getMonth()] + ' ' + d.getFullYear();
}

// Nodin dipecah jadi 2 dokumen independen (tidak lagi digabung jadi 1
// nota dinas): NODIN DUE DILIGENCE (EDD/LDD) & NODIN KAJIAN DRAFT PKS
// (Kajian Risiko/Legal/Kepatuhan). Masing-masing punya checklist,
// preview, & tombol download sendiri-sendiri — TIDAK PERNAH digabung
// jadi satu file, dan pilihan checklist di modal Nodin ini HANYA
// memengaruhi isi dokumen, TIDAK PERNAH mengubah checklist/Progres yang
// tersimpan di tabel Monitoring PKS.
const NODIN_DD_KEYS = ['EDD', 'LDD'];
const NODIN_DD_LABELS = { EDD: 'Enhanced Due Diligence (EDD)', LDD: 'Legal Due Diligence (LDD)' };

const NODIN_PKS_KEYS = ['KajianRisiko', 'KajianLegal', 'KajianKepatuhan'];
const NODIN_PKS_LABELS = { KajianRisiko: 'Kajian Risiko', KajianLegal: 'Kajian Legal', KajianKepatuhan: 'Kajian Kepatuhan' };
const NODIN_PKS_LABELS_SINGKAT = { KajianRisiko: 'Risiko', KajianLegal: 'Legal', KajianKepatuhan: 'Kepatuhan' };

/**
 * Variant formatTanggalPanjangIndo_ KHUSUS Nodin — TIDAK PERNAH
 * mengembalikan "-" utk tanggal kosong (beda dgn formatTanggalPanjangIndo_
 * yg dipakai laporan EGR & fungsi lain, sengaja TIDAK diubah supaya tidak
 * memengaruhi fitur lain). Placeholder ini hanya utk PREVIEW — download
 * sungguhan diblokir lebih dulu oleh validateNodinDataLengkap_.
 */
function formatTanggalPanjangIndoAtauKosong_(dateStr) {
  if (!dateStr) return '(target penyelesaian belum diisi)';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return String(dateStr);
  return d.getDate() + ' ' + NAMA_BULAN_ID[d.getMonth()] + ' ' + d.getFullYear();
}

/**
 * Data umum yang dipakai KEDUA jenis Nodin: nama perusahaan (SELALU
 * mengikuti baris Supplier yang dipilih, tidak pernah statis), link
 * Google Drive, target penyelesaian (mentah, blm diformat), narahubung.
 * Halaman ini khusus Monitoring PKS SUPPLIER — kedudukan perusahaan
 * selalu "pemasok (supplier)".
 */
function buildNodinBaseData_(supplierId) {
  const supplier = sheetToObjects(SHEET_NAMES.SUPPLIER).find(s => s.SupplierID === supplierId);
  if (!supplier) throw new Error('Supplier tidak ditemukan.');

  const monitoringRows = sheetToObjects(SHEET_NAMES.MONITORING_PKS).filter(r => r.SupplierID === supplierId);
  monitoringRows.sort((a, b) => new Date(b.UpdatedAt || 0) - new Date(a.UpdatedAt || 0));
  const row = monitoringRows[0] || {};

  return {
    namaPerusahaan: supplier.Nama,
    kedudukanLabel: 'pemasok (supplier)',
    row: row,
    linkDokumen: row.LinkDokumen || '',
    batasWaktuRaw: row.BatasWaktu || '',
    namaPic: row.PIC || '',
    teleponPic: row.TeleponPIC || ''
  };
}

/**
 * Field yang WAJIB terisi sebelum Nodin boleh di-download (target
 * penyelesaian & narahubung) — kalau kosong, TIDAK ditampilkan sbg "-"
 * di dokumen, tapi diblokir dgn validasi (lihat G. LOGIKA DOKUMEN).
 */
function validateNodinDataLengkap_(base) {
  const missing = [];
  if (!base.batasWaktuRaw) missing.push('Target Penyelesaian Kajian');
  if (!base.namaPic) missing.push('Nama Narahubung');
  if (!base.teleponPic) missing.push('Nomor Kontak Narahubung');
  return missing;
}

function buildNodinDDData_(supplierId, overrides) {
  const base = buildNodinBaseData_(supplierId);
  overrides = overrides || {};
  // Pilihan checklist di modal Nodin (overrides.kajianKeys) HANYA
  // memengaruhi isi dokumen — kalau tidak dikirim, fallback ke kajian yg
  // sudah berstatus Proses/Selesai di checklist Monitoring PKS tersimpan.
  const kajianKeys = Array.isArray(overrides.kajianKeys)
    ? NODIN_DD_KEYS.filter(k => overrides.kajianKeys.indexOf(k) !== -1)
    : NODIN_DD_KEYS.filter(k => base.row[k] && base.row[k] !== 'N/A');
  base.kajianKeys = kajianKeys;
  base.daftarKajianText = joinDaftarIndonesia_(kajianKeys.map(k => NODIN_DD_LABELS[k]));
  return base;
}

function buildNodinPksData_(supplierId, overrides) {
  const base = buildNodinBaseData_(supplierId);
  overrides = overrides || {};
  const kajianKeys = Array.isArray(overrides.kajianKeys)
    ? NODIN_PKS_KEYS.filter(k => overrides.kajianKeys.indexOf(k) !== -1)
    : NODIN_PKS_KEYS.filter(k => base.row[k] && base.row[k] !== 'N/A');
  base.kajianKeys = kajianKeys;
  base.daftarKajianText = joinDaftarIndonesia_(kajianKeys.map(k => NODIN_PKS_LABELS[k]));
  base.daftarKajianSingkatText = joinDaftarIndonesia_(kajianKeys.map(k => NODIN_PKS_LABELS_SINGKAT[k]));
  return base;
}

function buildNodinDDJudul_(data) {
  if (!data.kajianKeys.length) return 'Permohonan Pelaksanaan Due Diligence terhadap ' + data.namaPerusahaan;
  return 'Permohonan Pelaksanaan ' + data.daftarKajianText + ' terhadap ' + data.namaPerusahaan;
}

function buildNodinPksJudul_(data) {
  if (!data.kajianKeys.length) return 'Permohonan Kajian atas Draft PKS dengan ' + data.namaPerusahaan;
  return 'Permohonan Kajian ' + data.daftarKajianSingkatText + ' atas Draft PKS dengan ' + data.namaPerusahaan;
}

/**
 * NODIN DUE DILIGENCE — hanya EDD/LDD. Struktur: preamble (tidak
 * bernomor), points (1 topik = 1 poin bernomor urut, tanpa nomor yang
 * hilang/meloncat krn urutan array otomatis mengikuti isi), closing
 * (tidak bernomor). Kalau tidak ada kajian dicentang, hanya
 * preamble+closing yang dikembalikan (dokumen sungguhan tidak pernah
 * dibuat utk kondisi ini — tombol download diblokir di frontend/backend).
 */
function buildNodinDDPointGroups_(data) {
  const preamble = [
    { type: 'text', text: 'Dalam rangka menunjang operasional Perdagangan Emas Bulion serta pelaksanaan kerja sama antara PT Pegadaian dengan ' + data.namaPerusahaan + ' selaku ' + data.kedudukanLabel + ', bersama ini kami sampaikan hal-hal sebagai berikut:' }
  ];

  const points = [];
  if (data.kajianKeys.length) {
    points.push([
      { type: 'text', text: 'Sehubungan dengan rencana pelaksanaan kerja sama antara PT Pegadaian dengan ' + data.namaPerusahaan + ', kami bermaksud mengajukan permohonan ' + data.daftarKajianText + '. Pelaksanaan kajian tersebut diperlukan untuk memastikan bahwa aspek calon mitra telah sesuai dengan ketentuan internal PT Pegadaian, peraturan perundang-undangan yang berlaku, prinsip tata kelola perusahaan yang baik, serta aspek kepatuhan dan manajemen risiko.' }
    ]);
  }

  if (data.kajianKeys.indexOf('LDD') !== -1) {
    points.push([
      { type: 'text', text: 'Pelaksanaan Legal Due Diligence (LDD) dimaksudkan untuk memperoleh penilaian secara komprehensif terhadap aspek legal ' + data.namaPerusahaan + ', antara lain terkait legalitas perusahaan, kewenangan bertindak, struktur perizinan, kepatuhan terhadap peraturan perundang-undangan, potensi sengketa atau permasalahan hukum, serta aspek hukum lainnya yang dapat memengaruhi pelaksanaan kerja sama maupun menimbulkan risiko hukum bagi PT Pegadaian.' }
    ]);
  }
  if (data.kajianKeys.indexOf('EDD') !== -1) {
    points.push([
      { type: 'text', text: 'Kami mengajukan permohonan pelaksanaan Enhanced Due Diligence (EDD) dalam kerangka Program Kepatuhan Anti Pencucian Uang dan Pencegahan Pendanaan Terorisme (APU-PPT). Kajian tersebut diperlukan sebagai penerapan prinsip kehati-hatian, mengingat perdagangan emas merupakan sektor berisiko tinggi terhadap potensi tindak pidana pencucian uang. Melalui pelaksanaan EDD, diharapkan dapat dilakukan identifikasi dan verifikasi secara mendalam terhadap profil, legalitas, kepatuhan, integritas, beneficial ownership, dan profil risiko ' + data.namaPerusahaan + '. Pelaksanaan EDD juga diperlukan untuk mendukung penerapan prinsip Mengenali Pengguna Jasa serta pemeriksaan terhadap daftar terorisme, pendanaan proliferasi senjata pemusnah massal, dan indikasi aktivitas keuangan mencurigakan yang dapat menimbulkan risiko reputasi maupun kepatuhan bagi PT Pegadaian.' }
    ]);
  }

  if (data.kajianKeys.length) {
    points.push([
      { type: 'text', text: 'Sebagai bahan penelaahan, kami telah melampirkan dokumen legalitas perusahaan serta dokumen pendukung lainnya yang dapat diakses melalui tautan berikut:' },
      { type: 'link', text: data.linkDokumen || '(link belum diisi)', url: data.linkDokumen }
    ]);
    points.push([
      { type: 'text', text: 'Kami mengharapkan hasil ' + data.daftarKajianText + ' dapat kami terima paling lambat pada ' + formatTanggalPanjangIndoAtauKosong_(data.batasWaktuRaw) + '.' }
    ]);
    points.push([
      { type: 'text', text: 'Adapun untuk koordinasi teknis lebih lanjut, Bapak/Ibu dapat menghubungi narahubung kami, Sdr. ' + (data.namaPic || '(nama narahubung belum diisi)') + ' (' + (data.teleponPic || 'nomor kontak belum diisi') + ').' }
    ]);
  }

  const closing = [
    { type: 'text', text: 'Demikian disampaikan dengan hormat untuk diketahui dan ditindaklanjuti sebagaimana mestinya. Atas perhatian dan tindak lanjut Bapak/Ibu, kami ucapkan terima kasih.' }
  ];

  return { preamble: preamble, points: points, closing: closing };
}

/**
 * NODIN KAJIAN DRAFT PKS — hanya Kajian Risiko/Legal/Kepatuhan. Struktur
 * sama dengan buildNodinDDPointGroups_ (lihat komentarnya).
 */
function buildNodinPksPointGroups_(data) {
  const preamble = [
    { type: 'text', text: 'Dalam rangka menunjang operasional Perdagangan Emas Bulion serta rencana pelaksanaan kerja sama antara PT Pegadaian dengan ' + data.namaPerusahaan + ' selaku ' + data.kedudukanLabel + ', bersama ini kami sampaikan hal-hal sebagai berikut:' }
  ];

  const points = [];
  if (data.kajianKeys.length) {
    points.push([
      { type: 'text', text: 'Sehubungan dengan rencana Penandatanganan Perjanjian Kerja Sama antara PT Pegadaian dengan ' + data.namaPerusahaan + ', kami bermaksud mengajukan permohonan kajian dan review terhadap draft Perjanjian Kerja Sama dimaksud.' }
    ]);
  }

  if (data.kajianKeys.indexOf('KajianRisiko') !== -1) {
    points.push([
      { type: 'text', text: 'Kami mengajukan permohonan kajian dan review aspek risiko terhadap draft Perjanjian Kerja Sama dengan ' + data.namaPerusahaan + '. Kajian tersebut diperlukan untuk mengidentifikasi, mengukur, dan memitigasi potensi risiko yang dapat timbul dalam pelaksanaan kerja sama, termasuk risiko operasional, hukum, kepatuhan, reputasi, dan risiko lain yang relevan.' }
    ]);
  }
  if (data.kajianKeys.indexOf('KajianLegal') !== -1) {
    points.push([
      { type: 'text', text: 'Kami mengajukan permohonan kajian dan review aspek legal terhadap draft Perjanjian Kerja Sama dengan ' + data.namaPerusahaan + '. Kajian tersebut diperlukan untuk memastikan bahwa ketentuan dalam draft PKS telah memberikan perlindungan hukum yang memadai bagi PT Pegadaian serta sesuai dengan ketentuan internal dan peraturan perundang-undangan yang berlaku.' }
    ]);
  }
  if (data.kajianKeys.indexOf('KajianKepatuhan') !== -1) {
    points.push([
      { type: 'text', text: 'Kami mengajukan permohonan kajian dan review aspek kepatuhan terhadap draft Perjanjian Kerja Sama dengan ' + data.namaPerusahaan + '. Kajian tersebut diperlukan untuk memastikan bahwa rencana kerja sama dan substansi perjanjian telah sesuai dengan ketentuan internal PT Pegadaian, prinsip tata kelola perusahaan yang baik, serta ketentuan regulator yang berlaku.' }
    ]);
  }

  if (data.kajianKeys.length) {
    points.push([
      { type: 'text', text: 'Sebagai bahan penelaahan, draft Perjanjian Kerja Sama serta dokumen pendukung lainnya dapat diakses melalui tautan berikut:' },
      { type: 'link', text: data.linkDokumen || '(link belum diisi)', url: data.linkDokumen }
    ]);
    points.push([
      { type: 'text', text: 'Kami mengharapkan hasil ' + data.daftarKajianText + ' dapat kami terima paling lambat pada ' + formatTanggalPanjangIndoAtauKosong_(data.batasWaktuRaw) + ' sebagai bahan penyelesaian draft Perjanjian Kerja Sama.' }
    ]);
    points.push([
      { type: 'text', text: 'Adapun untuk koordinasi teknis lebih lanjut, Bapak/Ibu dapat menghubungi narahubung kami, Sdr. ' + (data.namaPic || '(nama narahubung belum diisi)') + ' (' + (data.teleponPic || 'nomor kontak belum diisi') + ').' }
    ]);
  }

  const closing = [
    { type: 'text', text: 'Demikian disampaikan dengan hormat untuk diketahui dan ditindaklanjuti sebagaimana mestinya. Atas perhatian dan tindak lanjut Bapak/Ibu, kami ucapkan terima kasih.' }
  ];

  return { preamble: preamble, points: points, closing: closing };
}

// Susun preamble/points/closing (dgn nomor urut otomatis) jadi larik
// baris teks polos — dipakai KEDUA jenis Nodin utk preview.
function renderNodinPreviewLines_(groups) {
  const lines = [];
  groups.preamble.forEach(p => lines.push(p.text));
  groups.points.forEach((point, idx) => {
    const num = (idx + 1) + '. ';
    point.forEach((p, pi) => lines.push((pi === 0 ? num : '') + p.text));
  });
  groups.closing.forEach(p => lines.push(p.text));
  return lines;
}

/**
 * Tulis judul (Perihal, tebal) + preamble/points/closing ke body Google
 * Doc — dipakai KEDUA jenis Nodin. Format (font/ukuran/spasi/rata kiri
 * kanan) SAMA PERSIS dengan template Nodin existing, tidak diubah.
 */
function writeNodinDocxBody_(doc, judul, groups) {
  const body = doc.getBody();

  function styleParagraph(p) {
    p.setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
      [DocumentApp.Attribute.FONT_SIZE]: 22,
      [DocumentApp.Attribute.LINE_SPACING]: 1.15
    });
    p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    p.setSpacingAfter(10);
  }

  let isFirstParagraph = true;
  function addParagraph(prefix, item, bold) {
    const paragraph = isFirstParagraph ? body.getParagraphs()[0] : body.appendParagraph('');
    isFirstParagraph = false;
    if (prefix) paragraph.appendText(prefix);
    if (item.type === 'link') {
      const text = paragraph.appendText(item.text);
      if (item.url) text.setLinkUrl(item.url);
    } else {
      paragraph.appendText(item.text);
    }
    styleParagraph(paragraph);
    if (bold) paragraph.editAsText().setBold(true);
  }

  addParagraph('Perihal: ', { type: 'text', text: judul }, true);
  groups.preamble.forEach(p => addParagraph('', p));
  groups.points.forEach((point, idx) => {
    const num = (idx + 1) + '. ';
    point.forEach((p, pi) => addParagraph(pi === 0 ? num : '', p));
  });
  groups.closing.forEach(p => addParagraph('', p));
}

// Export Google Doc sementara -> base64 docx, lalu langsung hapus (trash)
// doc-nya — tidak meninggalkan file permanen. Dipakai KEDUA jenis Nodin.
function exportGoogleDocAsDocxBase64_(doc) {
  doc.saveAndClose();
  const docId = doc.getId();

  // DriveApp.getAs(MimeType.MICROSOFT_WORD) TIDAK didukung Google untuk
  // konversi Google Doc -> docx ("Converting ... is not supported").
  // Cara yang benar: panggil endpoint export Drive API v3 langsung
  // pakai token OAuth milik eksekusi ini.
  const exportUrl = 'https://www.googleapis.com/drive/v3/files/' + docId +
    '/export?mimeType=' + encodeURIComponent('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    DriveApp.getFileById(docId).setTrashed(true);
    throw new Error('Gagal export ke docx (HTTP ' + response.getResponseCode() + '): ' + response.getContentText());
  }
  const base64 = Utilities.base64Encode(response.getContent());
  DriveApp.getFileById(docId).setTrashed(true);
  return base64;
}

// Nama file aman dari karakter yang tidak valid utk filename Windows/macOS.
function nodinSafeFileNamePart_(text) {
  return String(text || '').replace(/[\\/:*?"<>|]+/g, ' ').trim();
}

/** Preview teks Nodin Due Diligence (tanpa membuat file apa pun). */
function getNodinDueDiligencePreview(supplierId, overrides) {
  requireRole('Viewer');
  const data = buildNodinDDData_(supplierId, overrides);
  const groups = buildNodinDDPointGroups_(data);
  const judul = buildNodinDDJudul_(data);
  return {
    namaPerusahaan: data.namaPerusahaan,
    judul: judul,
    kajianKeys: data.kajianKeys,
    paragraphs: renderNodinPreviewLines_(groups),
    missingFields: validateNodinDataLengkap_(data)
  };
}

/** Preview teks Nodin Kajian Draft PKS (tanpa membuat file apa pun). */
function getNodinKajianPksPreview(supplierId, overrides) {
  requireRole('Viewer');
  const data = buildNodinPksData_(supplierId, overrides);
  const groups = buildNodinPksPointGroups_(data);
  const judul = buildNodinPksJudul_(data);
  return {
    namaPerusahaan: data.namaPerusahaan,
    judul: judul,
    kajianKeys: data.kajianKeys,
    paragraphs: renderNodinPreviewLines_(groups),
    missingFields: validateNodinDataLengkap_(data)
  };
}

/** Buat file .docx Nodin Due Diligence (EDD/LDD). */
function generateNodinDueDiligenceDocx(supplierId, overrides) {
  requireRole('Supervisor');
  const data = buildNodinDDData_(supplierId, overrides);
  if (!data.kajianKeys.length) throw new Error('Pilih minimal satu jenis kajian (EDD/LDD).');
  const missing = validateNodinDataLengkap_(data);
  if (missing.length) throw new Error('Lengkapi data berikut sebelum download: ' + missing.join(', ') + '.');

  const groups = buildNodinDDPointGroups_(data);
  const judul = buildNodinDDJudul_(data);
  const doc = DocumentApp.create('TEMP_Nodin_DD_' + supplierId + '_' + Date.now());
  writeNodinDocxBody_(doc, judul, groups);
  const base64 = exportGoogleDocAsDocxBase64_(doc);

  logAudit('CREATE', 'MonitoringPKSSupplier', supplierId, 'Download draft Nodin Due Diligence: ' + data.namaPerusahaan + ' (' + data.kajianKeys.join(',') + ')');

  return {
    base64: base64,
    filename: 'Nodin-Due-Diligence-' + nodinSafeFileNamePart_(data.namaPerusahaan) + '.docx'
  };
}

/** Buat file .docx Nodin Kajian Draft PKS (Risiko/Legal/Kepatuhan). */
function generateNodinKajianPksDocx(supplierId, overrides) {
  requireRole('Supervisor');
  const data = buildNodinPksData_(supplierId, overrides);
  if (!data.kajianKeys.length) throw new Error('Pilih minimal satu jenis kajian (Risiko/Legal/Kepatuhan).');
  const missing = validateNodinDataLengkap_(data);
  if (missing.length) throw new Error('Lengkapi data berikut sebelum download: ' + missing.join(', ') + '.');

  const groups = buildNodinPksPointGroups_(data);
  const judul = buildNodinPksJudul_(data);
  const doc = DocumentApp.create('TEMP_Nodin_PKS_' + supplierId + '_' + Date.now());
  writeNodinDocxBody_(doc, judul, groups);
  const base64 = exportGoogleDocAsDocxBase64_(doc);

  logAudit('CREATE', 'MonitoringPKSSupplier', supplierId, 'Download draft Nodin Kajian Draft PKS: ' + data.namaPerusahaan + ' (' + data.kajianKeys.join(',') + ')');

  return {
    base64: base64,
    filename: 'Nodin-Kajian-Draft-PKS-' + nodinSafeFileNamePart_(data.namaPerusahaan) + '.docx'
  };
}

/**
 * Cek apakah nama supplier tertentu punya riwayat transaksi di sheet
 * Input_Beli_EGR (spreadsheet EGR eksternal, PRICE_DISCOVERY_SOURCE_ID —
 * bukan sheet Pembelian aplikasi ini). Dicocokkan lewat kolom Seller
 * (teks bebas, sama persis dgn cara Pembelian mengisi datalist Seller
 * dari daftar nama supplier) — bukan lewat SupplierID, karena sheet
 * eksternal ini memang tidak punya kolom SupplierID sama sekali.
 */
function egrBeliHasRiwayatUntukSeller_(namaSupplier) {
  const namaKey = String(namaSupplier || '').trim().toLowerCase();
  if (!namaKey) return false;

  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.BELI);
  if (!sheet) return false;

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return false;
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const sellerCol = findColByKeywords(headerMap, ['seller']);
  if (sellerCol === -1) return false;

  return values.slice(headerRowIdx + 1).some(function (row) {
    const seller = String(getByCol(row, sellerCol) || '').trim().toLowerCase();
    return seller === namaKey;
  });
}

/**
 * Hapus supplier BERDASARKAN SupplierID (bukan nomor baris/nama) —
 * deleteRowById mencari baris yg cocok persis lewat ID lalu menghapus
 * SATU baris itu saja, tidak menggeser/mengubah data supplier lain.
 * Riwayat transaksi di Input_Beli_EGR (sheet eksternal, dikaitkan lewat
 * nama bukan SupplierID) TIDAK PERNAH disentuh/dihapus oleh fungsi ini —
 * dicek dulu SEBELUM baris dihapus supaya frontend bisa menampilkan
 * peringatan tambahan kalau riwayatnya memang ada.
 */
function deleteSupplier(supplierId) {
  requireRole('Admin');
  const supplier = sheetToObjects(SHEET_NAMES.SUPPLIER).find(s => s.SupplierID === supplierId);
  if (!supplier) throw new Error('Supplier tidak ditemukan.');
  const namaSupplier = supplier.Nama;

  const adaRiwayatEgr = egrBeliHasRiwayatUntukSeller_(namaSupplier);

  deleteRowById(SHEET_NAMES.SUPPLIER, 'SupplierID', supplierId);
  logAudit('DELETE', 'Supplier', supplierId, namaSupplier);

  return { namaSupplier: namaSupplier, adaRiwayatEgr: adaRiwayatEgr };
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
  // Date invalid (mis. Invalid Date) bikin Utilities.formatDate() THROW —
  // jangan sampai satu nilai tanggal yang rusak menjatuhkan seluruh
  // fungsi pemanggil.
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
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
    // Cell Date yang isinya corrupt/invalid (mis. Invalid Date) bikin
    // Utilities.formatDate() THROW — jangan sampai satu baris rusak
    // menjatuhkan seluruh fungsi yang memanggil extractDateOnly.
    if (isNaN(value.getTime())) return '';
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
// dengan pemanggil lain (mis. halaman Penjualan, Laporan Harian)
// yang belum kirim periode. TIDAK mengubah sheet/struktur data apa
// pun, murni filter di memori berdasarkan kolom Tanggal yang sudah ada.
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

// ------------------------------------------------------------
// GOLD IN TRANSIT — MONITORING PEMENUHAN PO (halaman Piutang Supplier)
// Sumber data: sheet EKSTERNAL "F.3.Input_Barang Masuk" (spreadsheet
// Bullion Trading Process, PRICE_DISCOVERY_SOURCE_ID) — TIDAK ada
// sheet/database baru yang ditulis, murni baca. Sheet sumber berisi DUA
// tabel independen berdampingan pada baris yang sama (bukan 1 baris = 1
// relasi): blok KIRI = daftar penerimaan/BAST (Tanggal terima, Gramasi,
// Nomor BAST IN, No Invoice, Ref PO), blok KANAN = daftar master PO
// (Tanggal, Nomor Purchase Order, Nomor Invoice, Seller, Qty (Gram)).
// Dihubungkan lewat Ref PO (kiri) === Nomor Purchase Order (kanan).
// Posisi kolom dicari via NAMA HEADER (bukan index hardcode) supaya
// tahan kalau sheet sumber sedikit berubah. TIDAK PERNAH memakai kolom
// "(1) Sisa"/"Delivery"/"Aging Order" di sheet (snapshot HARI INI) utk
// hitung posisi historis — Gold in Transit SELALU dihitung ulang dari
// baris mentah berdasarkan tanggal posisi yang dipilih user (tervalidasi
// cocok persis dgn angka riil posisi 30 Juni 2026: Total PO 10.112.500 gr,
// Diterima 9.682.500 gr, Gold in Transit 430.000 gr, 73 PO Outstanding,
// 3 Datang Sebagian).
// ------------------------------------------------------------
const F3_SHEET_NAME = 'F.3.Input_Barang Masuk';
const F3_MONTHS_EN = { january: 1, february: 2, march: 3, april: 4, may: 5, june: 6, july: 7, august: 8, september: 9, october: 10, november: 11, december: 12 };

function f3NormHeader_(s) {
  return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
}
function f3FindColExact_(headerRow, target) {
  const t = f3NormHeader_(target);
  for (let i = 0; i < headerRow.length; i++) { if (f3NormHeader_(headerRow[i]) === t) return i; }
  return -1;
}
function f3FindColContains_(headerRow, target) {
  const t = f3NormHeader_(target);
  for (let i = 0; i < headerRow.length; i++) { if (f3NormHeader_(headerRow[i]).indexOf(t) !== -1) return i; }
  return -1;
}
// Sama seperti f3FindColExact_ tapi mulai cari dari kolom ke-`fromIdx` —
// perlu krn header "Harga" muncul 3x di sheet (blok kiri/penerimaan &
// blok kanan/PO masing-masing punya kolom "Harga" sendiri, nama sama
// persis), jadi pencarian harus dibatasi ke blok kanan (>= kolom Tanggal PO).
function f3FindColExactFrom_(headerRow, target, fromIdx) {
  const t = f3NormHeader_(target);
  for (let i = fromIdx; i < headerRow.length; i++) { if (f3NormHeader_(headerRow[i]) === t) return i; }
  return -1;
}
// "18 March 2025" -> "2025-03-18". Baris kosong/placeholder sheet
// (mis. "0 January 1900") sengaja menghasilkan '' (invalid, diabaikan).
// Sel di Google Sheets bisa berupa objek Date SUNGGUHAN (bukan cuma
// teks tampilan) — getDataRange().getValues() akan mengembalikan Date,
// BUKAN string "18 March 2025", utk kolom yg diformat sbg tanggal.
function f3ParseDate_(value) {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(value || '').trim();
  const m = s.match(/^(\d{1,2})\s+([a-zA-Z]+)\s+(\d{4})/);
  if (!m) return '';
  const d = parseInt(m[1], 10), mo = F3_MONTHS_EN[m[2].toLowerCase()], y = parseInt(m[3], 10);
  if (!mo || !d || !y) return '';
  return y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
}
// Normalisasi Nomor PO utk pencocokan (Ref PO harus === Nomor PO):
// hapus spasi awal/akhir, samakan kapital. Angka/garis miring/tanda
// hubung/tahun TIDAK diubah.
function f3NormPO_(str) {
  return String(str || '').trim().toUpperCase();
}

/**
 * Baca & parse sheet F.3.Input_Barang Masuk jadi { pos, delivs }.
 * pos = daftar PO master (blok kanan), delivs = daftar penerimaan/BAST
 * (blok kiri). Baris kosong/placeholder (Seller="0", Nomor PO kosong,
 * BAST kosong, dsb — banyak ditemukan di bagian bawah sheet) disingkirkan.
 */
function readF3BarangMasuk_() {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = ss.getSheetByName(F3_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + F3_SHEET_NAME + '" tidak ditemukan di spreadsheet sumber.');

  const values = sheet.getDataRange().getValues();
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(values.length, 20); r++) {
    const row = values[r];
    if (f3FindColExact_(row, 'tanggal terima') !== -1 && f3FindColContains_(row, 'purchase order') !== -1) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx === -1) throw new Error('Baris header tidak ditemukan di sheet "' + F3_SHEET_NAME + '".');
  const headerRow = values[headerRowIdx];

  const col = {
    tanggalTerima: f3FindColExact_(headerRow, 'tanggal terima'),
    gramasiTerima: f3FindColExact_(headerRow, 'gramasi'),
    nomorBAST: f3FindColExact_(headerRow, 'nomor bast in'),
    noInvoiceDelivery: f3FindColExact_(headerRow, 'no invoice'),
    refPO: f3FindColExact_(headerRow, 'ref po'),
    tanggalPO: f3FindColExact_(headerRow, 'tanggal'),
    nomorPO: f3FindColContains_(headerRow, 'purchase order'),
    nomorInvoicePO: f3FindColExact_(headerRow, 'nomor invoice'),
    supplier: f3FindColExact_(headerRow, 'seller'),
    qtyPO: f3FindColExact_(headerRow, 'qty (gram)')
  };
  const missing = Object.keys(col).filter(function (k) { return col[k] === -1; });
  if (missing.length) throw new Error('Kolom berikut tidak ditemukan di sheet "' + F3_SHEET_NAME + '": ' + missing.join(', '));
  // "Harga" muncul 3x di sheet (nama sama persis di blok kiri & kanan) —
  // ambil KHUSUS yg ada di blok kanan/PO (>= kolom Tanggal PO), dipakai
  // utk nilai Rupiah Gold in Transit. Opsional: kalau tak ketemu, nilai
  // Rupiah cuma jadi 0 (tidak menggagalkan seluruh fitur gramasi).
  const colHargaPO = f3FindColExactFrom_(headerRow, 'harga', col.tanggalPO);

  const pos = [];
  const delivs = [];
  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];

    const seller = String(getByCol(row, col.supplier) || '').trim();
    const nomorPO = String(getByCol(row, col.nomorPO) || '').trim();
    if (seller && seller !== '0' && nomorPO) {
      const tanggalPO = f3ParseDate_(getByCol(row, col.tanggalPO));
      const qty = parseAngkaSel_(getByCol(row, col.qtyPO)) || 0;
      if (tanggalPO && qty > 0) {
        pos.push({
          nomorPO: nomorPO, nomorPOKey: f3NormPO_(nomorPO), tanggalPO: tanggalPO,
          supplier: seller, qty: qty,
          harga: colHargaPO !== -1 ? (parseAngkaSel_(getByCol(row, colHargaPO)) || 0) : 0,
          nomorInvoice: String(getByCol(row, col.nomorInvoicePO) || '').trim()
        });
      }
    }

    const nomorBAST = String(getByCol(row, col.nomorBAST) || '').trim();
    if (nomorBAST) {
      const refPO = String(getByCol(row, col.refPO) || '').trim();
      const tanggalTerima = f3ParseDate_(getByCol(row, col.tanggalTerima));
      const gramasi = parseAngkaSel_(getByCol(row, col.gramasiTerima)) || 0;
      if (refPO && tanggalTerima && gramasi > 0) {
        delivs.push({
          refPOKey: f3NormPO_(refPO), tanggalTerima: tanggalTerima, gramasi: gramasi,
          nomorBAST: nomorBAST, nomorInvoice: String(getByCol(row, col.noInvoiceDelivery) || '').trim()
        });
      } else if (!refPO) {
        // Penerimaan tanpa Referensi PO — tandai "Belum Terpetakan",
        // TIDAK mengurangi Gold in Transit PO manapun (poin 8 spec).
        delivs.push({ refPOKey: '', tanggalTerima: tanggalTerima, gramasi: gramasi, nomorBAST: nomorBAST, nomorInvoice: '', belumTerpetakan: true });
      }
    }
  }

  return { pos: pos, delivs: delivs };
}

function goldInTransitTotals_(posisi, pos, delivs) {
  const totalPO = pos.reduce(function (s, p) { return s + (p.tanggalPO <= posisi ? p.qty : 0); }, 0);
  const totalDiterima = delivs.reduce(function (s, d) { return s + (!d.belumTerpetakan && d.tanggalTerima <= posisi ? d.gramasi : 0); }, 0);
  return { totalPO: totalPO, totalDiterima: totalDiterima, goldInTransit: Math.max(totalPO - totalDiterima, 0) };
}

/** Rincian per-PO pada tanggal posisi tertentu — dasar utk Ringkasan & Detail. */
function goldInTransitPerPO_(posisi, pos, delivs) {
  const delivByPO = {};
  delivs.forEach(function (d) {
    if (d.belumTerpetakan || d.tanggalTerima > posisi) return;
    if (!delivByPO[d.refPOKey]) delivByPO[d.refPOKey] = [];
    delivByPO[d.refPOKey].push(d);
  });

  const result = [];
  pos.forEach(function (p) {
    if (p.tanggalPO > posisi) return;
    const deliveries = (delivByPO[p.nomorPOKey] || []).slice().sort(function (a, b) { return a.tanggalTerima < b.tanggalTerima ? -1 : (a.tanggalTerima > b.tanggalTerima ? 1 : 0); });
    let diterima = 0;
    const historiPenerimaan = deliveries.map(function (d) {
      diterima += d.gramasi;
      return { tanggalTerima: d.tanggalTerima, nomorBAST: d.nomorBAST, nomorInvoice: d.nomorInvoice, gramasi: d.gramasi, kumulatif: diterima, sisaSetelah: Math.max(p.qty - diterima, 0) };
    });
    const goldInTransit = Math.max(p.qty - diterima, 0);
    let status = 'Belum Datang';
    if (diterima > 0 && diterima < p.qty) status = 'Sebagian';
    else if (diterima >= p.qty && p.qty > 0) status = 'Selesai';
    const aging = Math.round((new Date(posisi + 'T00:00:00').getTime() - new Date(p.tanggalPO + 'T00:00:00').getTime()) / 86400000);
    const tanggalTerakhirTerima = deliveries.length ? deliveries[deliveries.length - 1].tanggalTerima : '';

    result.push({
      nomorPO: p.nomorPO, tanggalPO: p.tanggalPO, supplier: p.supplier,
      nomorInvoice: p.nomorInvoice || (deliveries[0] ? deliveries[0].nomorInvoice : ''),
      qtyPO: p.qty, diterima: diterima, goldInTransit: goldInTransit,
      harga: p.harga || 0,
      nilaiGoldInTransit: goldInTransit * (p.harga || 0),
      persenPemenuhan: p.qty > 0 ? Math.min(100, (diterima / p.qty) * 100) : 0,
      tanggalTerakhirTerima: tanggalTerakhirTerima, aging: Math.max(aging, 0), status: status,
      historiPenerimaan: historiPenerimaan
    });
  });
  return result;
}

function goldInTransitLastDataDate_(pos, delivs) {
  let last = '';
  pos.forEach(function (p) { if (p.tanggalPO > last) last = p.tanggalPO; });
  delivs.forEach(function (d) { if (d.tanggalTerima && d.tanggalTerima > last) last = d.tanggalTerima; });
  return last || formatDateOnly(new Date());
}

function goldInTransitTrend_(pos, delivs, lastDataDate) {
  const allDates = pos.map(function (p) { return p.tanggalPO; }).concat(delivs.map(function (d) { return d.tanggalTerima; })).filter(Boolean);
  if (!allDates.length) return [];
  const minDate = allDates.reduce(function (a, b) { return a < b ? a : b; });
  const minMonthKey = minDate.slice(0, 7);
  const maxMonthKey = lastDataDate.slice(0, 7);

  const months = [];
  let y = parseInt(minMonthKey.slice(0, 4), 10), m = parseInt(minMonthKey.slice(5, 7), 10);
  const yEnd = parseInt(maxMonthKey.slice(0, 4), 10), mEnd = parseInt(maxMonthKey.slice(5, 7), 10);
  while (y < yEnd || (y === yEnd && m <= mEnd)) {
    months.push(y + '-' + String(m).padStart(2, '0'));
    m++; if (m > 12) { m = 1; y++; }
  }

  return months.map(function (mk) {
    const isLastMonth = mk === maxMonthKey;
    const yy = parseInt(mk.slice(0, 4), 10), mm = parseInt(mk.slice(5, 7), 10);
    const akhirBulan = formatDateOnly(new Date(yy, mm, 0));
    const posisiBulan = isLastMonth ? lastDataDate : akhirBulan;
    const totals = goldInTransitTotals_(posisiBulan, pos, delivs);
    return { bulan: mk, posisi: posisiBulan, isSementara: isLastMonth && posisiBulan !== akhirBulan, goldInTransit: totals.goldInTransit };
  });
}

/**
 * Endpoint utama halaman "Monitoring Pemenuhan PO — Gold in Transit"
 * (tab Ringkasan Gold in Transit & Detail Outstanding PO di halaman
 * Piutang Supplier). tanggalPosisi kosong -> pakai tanggal data
 * terakhir yang tersedia di sheet. Mengembalikan SELURUH data (semua
 * PO, termasuk histori penerimaan per PO) — filter Supplier/Status/
 * Aging/Nomor PO/Tahun diterapkan di CLIENT (dataset ~1300-an PO, aman
 * di-cache di browser spt pola Monitoring Delivery PO).
 */
function getGoldInTransitData(tanggalPosisi) {
  requireRole('Viewer');
  const data = readF3BarangMasuk_();
  const lastDataDate = goldInTransitLastDataDate_(data.pos, data.delivs);
  const posisi = tanggalPosisi ? extractDateOnly(tanggalPosisi) : lastDataDate;

  const perPO = goldInTransitPerPO_(posisi, data.pos, data.delivs);
  const totals = goldInTransitTotals_(posisi, data.pos, data.delivs);
  const poOutstanding = perPO.filter(function (r) { return r.goldInTransit > 0; });
  const poSebagian = perPO.filter(function (r) { return r.status === 'Sebagian'; });

  const supplierMap = {};
  poOutstanding.forEach(function (r) {
    if (!supplierMap[r.supplier]) supplierMap[r.supplier] = { supplier: r.supplier, goldInTransit: 0, poCount: 0 };
    supplierMap[r.supplier].goldInTransit += r.goldInTransit;
    supplierMap[r.supplier].poCount++;
  });
  const supplierBreakdown = Object.keys(supplierMap).map(function (k) {
    const s = supplierMap[k];
    s.persen = totals.goldInTransit > 0 ? (s.goldInTransit / totals.goldInTransit) * 100 : 0;
    return s;
  }).sort(function (a, b) { return b.goldInTransit - a.goldInTransit; });

  const supplierList = Array.from(new Set(data.pos.map(function (p) { return p.supplier; }).filter(Boolean))).sort();
  const tahunList = Array.from(new Set(data.pos.map(function (p) { return p.tanggalPO.slice(0, 4); }).filter(Boolean))).sort();

  return {
    posisi: posisi,
    lastDataDate: lastDataDate,
    supplierList: supplierList,
    tahunList: tahunList,
    scorecard: {
      goldInTransit: totals.goldInTransit,
      rupiahGoldInTransit: poOutstanding.reduce(function (s, r) { return s + r.nilaiGoldInTransit; }, 0),
      totalPO: perPO.length,
      totalOrder: totals.totalPO,
      sudahDiterima: totals.totalDiterima,
      poOutstanding: poOutstanding.length,
      poSebagian: poSebagian.length
    },
    trend: goldInTransitTrend_(data.pos, data.delivs, lastDataDate),
    supplierBreakdown: supplierBreakdown,
    detail: perPO
  };
}

// ------------------------------------------------------------
// PIUTANG SUPPLIER EGR — klon "Piutang Supplier Trading" di atas, SUMBER
// DATA BEDA SAJA (sheet "Input_Barang EGR", spreadsheet EGR yang sama —
// PRICE_DISCOVERY_SOURCE_ID), logic perhitungan 100% SAMA (reuse
// goldInTransitTotals_/goldInTransitPerPO_/goldInTransitLastDataDate_/
// goldInTransitTrend_ apa adanya — TIDAK ada logic baru yang ditulis di
// sini, murni fungsi baca+parse sheet yang berbeda). Read-only, tidak
// pernah menulis ke Input_Barang EGR. Data Trading (F.3.Input_Barang
// Masuk) TIDAK PERNAH dicampur ke sini.
//
// Struktur sheet sumber (dikonfirmasi dari screenshot user) SAMA seperti
// F.3.Input_Barang Masuk — dua blok berdampingan: blok KIRI = penerimaan
// (Tanggal terima, Gramasi, Nomor BAST IN, No Invoice, Referensi
// Transaksi), blok KANAN = master PO/transaksi (Tanggal, Nomor Purchase
// Order, Nomor Invoice, Seller, Qty (Gram), Harga). Dihubungkan lewat
// Referensi Transaksi (kiri) === Nomor Purchase Order (kanan). SATU-
// SATUNYA beda nama header dari versi Trading: kolom penghubung di blok
// kiri bernama "Referensi Transaksi", BUKAN "Ref PO".
// ------------------------------------------------------------
const EGR_BARANG_SHEET_NAME = 'Input_Barang EGR';

/**
 * Baca & parse sheet "Input_Barang EGR" jadi { pos, delivs } — bentuk
 * data IDENTIK dengan readF3BarangMasuk_() supaya bisa dipakai langsung
 * oleh goldInTransitTotals_/goldInTransitPerPO_/goldInTransitTrend_ yang
 * sama persis. Kalau header wajib tidak ditemukan, lempar error yang
 * jelas menyebutkan nama kolom yang hilang — TIDAK membuat nilai asumsi.
 */
function readEgrBarangMasuk_() {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = ss.getSheetByName(EGR_BARANG_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + EGR_BARANG_SHEET_NAME + '" tidak ditemukan di spreadsheet sumber.');

  const values = sheet.getDataRange().getValues();
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(values.length, 20); r++) {
    const row = values[r];
    if (f3FindColExact_(row, 'tanggal terima') !== -1 && f3FindColContains_(row, 'purchase order') !== -1) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx === -1) throw new Error('Baris header tidak ditemukan di sheet "' + EGR_BARANG_SHEET_NAME + '".');
  const headerRow = values[headerRowIdx];

  const col = {
    tanggalTerima: f3FindColExact_(headerRow, 'tanggal terima'),
    gramasiTerima: f3FindColExact_(headerRow, 'gramasi'),
    nomorBAST: f3FindColExact_(headerRow, 'nomor bast in'),
    noInvoiceDelivery: f3FindColExact_(headerRow, 'no invoice'),
    refPO: f3FindColExact_(headerRow, 'referensi transaksi'),
    tanggalPO: f3FindColExact_(headerRow, 'tanggal'),
    nomorPO: f3FindColContains_(headerRow, 'purchase order'),
    nomorInvoicePO: f3FindColExact_(headerRow, 'nomor invoice'),
    supplier: f3FindColExact_(headerRow, 'seller'),
    qtyPO: f3FindColExact_(headerRow, 'qty (gram)')
  };
  const missing = Object.keys(col).filter(function (k) { return col[k] === -1; });
  if (missing.length) throw new Error('Kolom berikut tidak ditemukan di sheet "' + EGR_BARANG_SHEET_NAME + '": ' + missing.join(', '));
  const colHargaPO = f3FindColExactFrom_(headerRow, 'harga', col.tanggalPO);

  const pos = [];
  const delivs = [];
  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];

    const seller = String(getByCol(row, col.supplier) || '').trim();
    const nomorPO = String(getByCol(row, col.nomorPO) || '').trim();
    if (seller && seller !== '0' && nomorPO) {
      const tanggalPO = f3ParseDate_(getByCol(row, col.tanggalPO));
      const qty = parseAngkaSel_(getByCol(row, col.qtyPO)) || 0;
      if (tanggalPO && qty > 0) {
        pos.push({
          nomorPO: nomorPO, nomorPOKey: f3NormPO_(nomorPO), tanggalPO: tanggalPO,
          supplier: seller, qty: qty,
          harga: colHargaPO !== -1 ? (parseAngkaSel_(getByCol(row, colHargaPO)) || 0) : 0,
          nomorInvoice: String(getByCol(row, col.nomorInvoicePO) || '').trim()
        });
      }
    }

    const nomorBAST = String(getByCol(row, col.nomorBAST) || '').trim();
    if (nomorBAST) {
      const refPO = String(getByCol(row, col.refPO) || '').trim();
      const tanggalTerima = f3ParseDate_(getByCol(row, col.tanggalTerima));
      const gramasi = parseAngkaSel_(getByCol(row, col.gramasiTerima)) || 0;
      if (refPO && tanggalTerima && gramasi > 0) {
        delivs.push({
          refPOKey: f3NormPO_(refPO), tanggalTerima: tanggalTerima, gramasi: gramasi,
          nomorBAST: nomorBAST, nomorInvoice: String(getByCol(row, col.noInvoiceDelivery) || '').trim()
        });
      } else if (!refPO) {
        delivs.push({ refPOKey: '', tanggalTerima: tanggalTerima, gramasi: gramasi, nomorBAST: nomorBAST, nomorInvoice: '', belumTerpetakan: true });
      }
    }
  }

  return { pos: pos, delivs: delivs };
}

/**
 * Endpoint "Piutang Supplier EGR" — klon getGoldInTransitData() di atas,
 * HANYA beda sumber baca (readEgrBarangMasuk_ vs readF3BarangMasuk_).
 * Seluruh perhitungan (totals/per-PO/trend/supplier breakdown) memanggil
 * fungsi SHARED yang sama persis dengan versi Trading — bukan duplikat
 * logic, jadi dijamin identik.
 */
function getGoldInTransitDataEgr(tanggalPosisi) {
  requireRole('Viewer');
  const data = readEgrBarangMasuk_();
  const lastDataDate = goldInTransitLastDataDate_(data.pos, data.delivs);
  const posisi = tanggalPosisi ? extractDateOnly(tanggalPosisi) : lastDataDate;

  const perPO = goldInTransitPerPO_(posisi, data.pos, data.delivs);
  const totals = goldInTransitTotals_(posisi, data.pos, data.delivs);
  const poOutstanding = perPO.filter(function (r) { return r.goldInTransit > 0; });
  const poSebagian = perPO.filter(function (r) { return r.status === 'Sebagian'; });

  const supplierMap = {};
  poOutstanding.forEach(function (r) {
    if (!supplierMap[r.supplier]) supplierMap[r.supplier] = { supplier: r.supplier, goldInTransit: 0, poCount: 0 };
    supplierMap[r.supplier].goldInTransit += r.goldInTransit;
    supplierMap[r.supplier].poCount++;
  });
  const supplierBreakdown = Object.keys(supplierMap).map(function (k) {
    const s = supplierMap[k];
    s.persen = totals.goldInTransit > 0 ? (s.goldInTransit / totals.goldInTransit) * 100 : 0;
    return s;
  }).sort(function (a, b) { return b.goldInTransit - a.goldInTransit; });

  const supplierList = Array.from(new Set(data.pos.map(function (p) { return p.supplier; }).filter(Boolean))).sort();
  const tahunList = Array.from(new Set(data.pos.map(function (p) { return p.tanggalPO.slice(0, 4); }).filter(Boolean))).sort();

  return {
    posisi: posisi,
    lastDataDate: lastDataDate,
    supplierList: supplierList,
    tahunList: tahunList,
    scorecard: {
      goldInTransit: totals.goldInTransit,
      rupiahGoldInTransit: poOutstanding.reduce(function (s, r) { return s + r.nilaiGoldInTransit; }, 0),
      totalPO: perPO.length,
      totalOrder: totals.totalPO,
      sudahDiterima: totals.totalDiterima,
      poOutstanding: poOutstanding.length,
      poSebagian: poSebagian.length
    },
    trend: goldInTransitTrend_(data.pos, data.delivs, lastDataDate),
    supplierBreakdown: supplierBreakdown,
    detail: perPO
  };
}

/**
 * Card "Piutang per Supplier" Dashboard — dipakai bareng Global Period
 * Filter. Bentuk hasilnya (rows[].supplierNama/gramasi/valuasi/
 * valuasiPercent, totalGramasi, totalValuasi) DIBUAT SAMA seperti
 * getPiutangLive() (sumber Rekap) yang digantikannya, supaya frontend
 * (renderPiutangSupplierChart) tidak perlu diubah bentuknya — hanya
 * SUMBER datanya diganti ke Gold in Transit (posisi per tanggal,
 * reuse readF3BarangMasuk_/goldInTransitPerPO_ — SAMA PERSIS dengan
 * menu "Piutang Supplier" / Monitoring Pemenuhan PO), bukan cell
 * snapshot statis di sheet Rekap — supaya card ini beneran bisa
 * dihitung "per tanggal" (posisi akhir periode filter), bukan cuma
 * kondisi hari ini.
 * tanggalPosisi kosong/null -> posisi = data terakhir yang tersedia
 * (persis seperti getGoldInTransitData tanpa parameter).
 */
function getPiutangSupplierDashboard(tanggalPosisi) {
  const data = readF3BarangMasuk_();
  const lastDataDate = goldInTransitLastDataDate_(data.pos, data.delivs);
  const posisi = tanggalPosisi ? extractDateOnly(tanggalPosisi) : lastDataDate;

  const perPO = goldInTransitPerPO_(posisi, data.pos, data.delivs);
  const outstanding = perPO.filter(function (r) { return r.goldInTransit > 0; });

  const bySupplier = {};
  outstanding.forEach(function (r) {
    if (!bySupplier[r.supplier]) bySupplier[r.supplier] = { supplierNama: r.supplier, gramasi: 0, valuasi: 0 };
    bySupplier[r.supplier].gramasi += r.goldInTransit;
    bySupplier[r.supplier].valuasi += (r.nilaiGoldInTransit || 0);
  });

  const rows = Object.values(bySupplier);
  const totalGramasi = rows.reduce(function (s, r) { return s + r.gramasi; }, 0);
  const totalValuasi = rows.reduce(function (s, r) { return s + r.valuasi; }, 0);
  rows.forEach(function (r) {
    r.gramasiPercent = totalGramasi ? (r.gramasi / totalGramasi) * 100 : 0;
    r.valuasiPercent = totalValuasi ? (r.valuasi / totalValuasi) * 100 : 0;
  });
  rows.sort(function (a, b) { return b.valuasi - a.valuasi; });

  return { rows: rows, totalGramasi: totalGramasi, totalValuasi: totalValuasi, posisi: posisi, lastDataDate: lastDataDate };
}

/**
 * Card "Piutang per Supplier" Dashboard — tab EGR. Klon
 * getPiutangSupplierDashboard() di atas, HANYA beda sumber baca
 * (readEgrBarangMasuk_ vs readF3BarangMasuk_) — logic agregasi per
 * supplier 100% sama. TIDAK PERNAH melempar error ke client (dibungkus
 * try/catch) — kalau sheet "Input_Barang EGR" gagal dibaca/kosong,
 * kembalikan rows kosong + empty:true supaya widget menampilkan pesan
 * "Belum terdapat data..." alih-alih membuat seluruh card Dashboard error.
 */
function getPiutangSupplierDashboardEgr(tanggalPosisi) {
  try {
    const data = readEgrBarangMasuk_();
    const lastDataDate = goldInTransitLastDataDate_(data.pos, data.delivs);
    const posisi = tanggalPosisi ? extractDateOnly(tanggalPosisi) : lastDataDate;

    const perPO = goldInTransitPerPO_(posisi, data.pos, data.delivs);
    const outstanding = perPO.filter(function (r) { return r.goldInTransit > 0; });

    const bySupplier = {};
    outstanding.forEach(function (r) {
      if (!bySupplier[r.supplier]) bySupplier[r.supplier] = { supplierNama: r.supplier, gramasi: 0, valuasi: 0 };
      bySupplier[r.supplier].gramasi += r.goldInTransit;
      bySupplier[r.supplier].valuasi += (r.nilaiGoldInTransit || 0);
    });

    const rows = Object.values(bySupplier);
    const totalGramasi = rows.reduce(function (s, r) { return s + r.gramasi; }, 0);
    const totalValuasi = rows.reduce(function (s, r) { return s + r.valuasi; }, 0);
    rows.forEach(function (r) {
      r.gramasiPercent = totalGramasi ? (r.gramasi / totalGramasi) * 100 : 0;
      r.valuasiPercent = totalValuasi ? (r.valuasi / totalValuasi) * 100 : 0;
    });
    rows.sort(function (a, b) { return b.valuasi - a.valuasi; });

    return { rows: rows, totalGramasi: totalGramasi, totalValuasi: totalValuasi, posisi: posisi, lastDataDate: lastDataDate, empty: rows.length === 0 };
  } catch (e) {
    return { rows: [], totalGramasi: 0, totalValuasi: 0, posisi: tanggalPosisi ? extractDateOnly(tanggalPosisi) : '', lastDataDate: '', empty: true, error: e.message || String(e) };
  }
}

/**
 * Card "Piutang per Supplier" Dashboard — tab Gabungan (Trading + EGR).
 * Menjumlahkan hasil agregasi per-supplier dari KEDUA sumber (posisi yang
 * SAMA dikirim ke keduanya — readF3BarangMasuk_ untuk Trading,
 * readEgrBarangMasuk_ untuk EGR — masing-masing lewat goldInTransitPerPO_
 * yang sama persis, TIDAK ada logic hitung baru). Supplier yang sama di
 * kedua sumber digabung berdasarkan NAMA SUPPLIER yang sama persis
 * (case/spasi-insensitive, BUKAN kemiripan/fuzzy — tidak ada tabel alias
 * supplier terpisah di aplikasi ini, hanya Master_Alias_Buyer utk BUYER).
 * Kalau EGR gagal dibaca/kosong, Gabungan TETAP menampilkan angka Trading
 * apa adanya (nilai EGR jadi 0) — tidak pernah membuat seluruh card error.
 */
function getPiutangSupplierDashboardCombined(tanggalPosisi) {
  const tradingData = readF3BarangMasuk_();
  const tradingLast = goldInTransitLastDataDate_(tradingData.pos, tradingData.delivs);

  let egrData = { pos: [], delivs: [] };
  let egrLast = '';
  let egrEmpty = true;
  let egrError = null;
  try {
    egrData = readEgrBarangMasuk_();
    egrLast = goldInTransitLastDataDate_(egrData.pos, egrData.delivs);
  } catch (e) {
    egrError = e.message || String(e);
  }

  const lastDataDate = egrLast > tradingLast ? egrLast : tradingLast;
  const posisi = tanggalPosisi ? extractDateOnly(tanggalPosisi) : (lastDataDate || formatDateOnly(new Date()));

  function bySupplierFrom_(pos, delivs) {
    const perPO = goldInTransitPerPO_(posisi, pos, delivs);
    const outstanding = perPO.filter(function (r) { return r.goldInTransit > 0; });
    const map = {};
    outstanding.forEach(function (r) {
      const key = String(r.supplier || '').trim().toLowerCase().replace(/\s+/g, ' ');
      if (!key) return;
      if (!map[key]) map[key] = { key: key, supplierNama: r.supplier, gramasi: 0, valuasi: 0 };
      map[key].gramasi += r.goldInTransit;
      map[key].valuasi += (r.nilaiGoldInTransit || 0);
    });
    return map;
  }

  const tradingMap = bySupplierFrom_(tradingData.pos, tradingData.delivs);
  const egrMap = bySupplierFrom_(egrData.pos, egrData.delivs);
  egrEmpty = Object.keys(egrMap).length === 0;

  const tradingTotalGramasi = Object.values(tradingMap).reduce(function (s, r) { return s + r.gramasi; }, 0);
  const tradingTotalValuasi = Object.values(tradingMap).reduce(function (s, r) { return s + r.valuasi; }, 0);
  const egrTotalGramasi = Object.values(egrMap).reduce(function (s, r) { return s + r.gramasi; }, 0);
  const egrTotalValuasi = Object.values(egrMap).reduce(function (s, r) { return s + r.valuasi; }, 0);

  const combinedMap = {};
  Object.keys(tradingMap).forEach(function (k) {
    combinedMap[k] = {
      key: k, supplierNama: tradingMap[k].supplierNama,
      gramasi: tradingMap[k].gramasi, valuasi: tradingMap[k].valuasi,
      tradingGramasi: tradingMap[k].gramasi, tradingValuasi: tradingMap[k].valuasi,
      egrGramasi: 0, egrValuasi: 0
    };
  });
  Object.keys(egrMap).forEach(function (k) {
    if (!combinedMap[k]) {
      combinedMap[k] = {
        key: k, supplierNama: egrMap[k].supplierNama, gramasi: 0, valuasi: 0,
        tradingGramasi: 0, tradingValuasi: 0, egrGramasi: 0, egrValuasi: 0
      };
    }
    combinedMap[k].gramasi += egrMap[k].gramasi;
    combinedMap[k].valuasi += egrMap[k].valuasi;
    combinedMap[k].egrGramasi = egrMap[k].gramasi;
    combinedMap[k].egrValuasi = egrMap[k].valuasi;
  });

  const rows = Object.values(combinedMap);
  const totalGramasi = tradingTotalGramasi + egrTotalGramasi;
  const totalValuasi = tradingTotalValuasi + egrTotalValuasi;
  rows.forEach(function (r) {
    r.gramasiPercent = totalGramasi ? (r.gramasi / totalGramasi) * 100 : 0;
    r.valuasiPercent = totalValuasi ? (r.valuasi / totalValuasi) * 100 : 0;
  });
  rows.sort(function (a, b) { return b.valuasi - a.valuasi; });

  return {
    rows: rows, totalGramasi: totalGramasi, totalValuasi: totalValuasi,
    tradingGramasi: tradingTotalGramasi, tradingValuasi: tradingTotalValuasi,
    egrGramasi: egrTotalGramasi, egrValuasi: egrTotalValuasi,
    egrEmpty: egrEmpty, egrError: egrError,
    posisi: posisi, lastDataDate: lastDataDate
  };
}

// ============================================================
// PRINT PENETAPAN HARGA — read-only, dibaca LANGSUNG dari sheet
// "Print_ V3" & sumber aslinya di spreadsheet Price Discovery V3
// (PRICE_DISCOVERY_SOURCE_ID). TIDAK menulis apa pun ke spreadsheet
// sumber, TIDAK ada formula/angka baru yang di-hardcode selain
// yang sudah diverifikasi manual terhadap formula sheet aslinya.
//
// HASIL AUDIT (mapping formula, per tanggal audit) — WAJIB dibaca
// ulang & diverifikasi ke sheet sumber kalau strukturnya berubah:
//
// Trading Reguler — kunci record = Tanggal + Tahap ("Harga Ke-"):
//   Sheet "Price V3" kolom: B=Tanggal, C=Harga Ke-(Tahap 1-5),
//   D=Pukul, E=Emasku(header asli "HRTA"), I=IDN Bullion,
//   J=StarGold(header "MKB/STRGLD"), K=Lotus Archi(header "LOTUS"),
//   M=APEPI, N=Spot Goldprice, O=XAUIDRK,
//   T=LBMA H-1 (kolom "Rp LBMA" -- SUDAH dalam Rupiah, BUKAN kolom
//     P yang harga USD mentah),
//   U=Kitco H-1 (kolom "Rp Kitco" -- juga Rupiah, BUKAN kolom Q),
//   Z=Harga Penetapan (header asli "Keputusan Harga" -- nilai
//     manual/keputusan, BUKAN average/min/max dari 9 sumber di atas),
//   AA=Harga TE (header asli "Buyback"),
//   AB=Gramasi Deposito (header asli "Deposito").
//   IGS & Marva justru dari sheet TERPISAH "New Price Discovery"
//   (kunci sama: kolom B=Tanggal, C=Tahap), kolom G=IGS, H=Marva
//   (header asli kolom H = "SJL" -- dikonfirmasi user: SJL memang
//   nama legal/emiten untuk merek dagang "Marva", bukan kesalahan
//   data).
//   Maksimal Disc = Harga Penetapan x (1 - persen), PERSIS formula
//   sumber: Manop 0,1% / Kadep 0,2% / Kadiv 0,3%.
//
// ANTAM — kunci record = Tanggal + Gramasi (denominasi), TIDAK ADA
//   dimensi Tahap:
//   Sheet "Price ANTAM" kolom: B=Tanggal, C=Pukul, D=Gramasi
//   (PENTING: header row sheet ini SALAH LABEL -- tertulis
//   C="Gramasi", D="Pukul", tapi data & formula sumber MEMBUKTIKAN
//   sebaliknya: C isinya jam ("8.30"), D isinya angka gramasi
//   1/2/3/5/10/25/50/100/250/500/1000. Kolom yang dipakai di sini
//   mengikuti FORMULA & DATA ASLI, bukan teks header sheet-nya),
//   E=Harga acuan ("ANTAM"), G=Harga Jual (header asli "Keputusan
//   Harga"). Kolom F ("Buyback") formulanya rusak (#REF!) di
//   sumbernya -- tidak dipakai di sini.
//   Pukul ANTAM di Print_ V3 di-HARDCODE "8.30" langsung di sheet
//   (bukan formula) -- direplikasi sebagai teks statis.
//   Deposito Emas untuk tab ANTAM memakai sel Tahap Trading Reguler
//   (Print_ V3 blok ANTAM me-reuse cell yang sama persis dengan
//   blok Trading) -- karena tidak ada selector Tahap independen di
//   tab ANTAM, di sini dipakai Tahap 1 tetap sebagai default.
//
// Piutang Supplier — SENGAJA TIDAK memakai formula SUMIF kolom
//   Sisa/Total F.3.Input_Barang Masuk seperti di Print_ V3 (yang
//   all-time, tidak terikat tanggal). Sesuai keputusan eksplisit
//   user, di menu ini dipakai getPiutangSupplierDashboard()
//   (logic Gold in Transit, sama seperti card Dashboard) supaya
//   bisa dilihat posisi per Tanggal Penetapan yang dipilih --
//   konsekuensinya angka Piutang di sini TIDAK akan identik dengan
//   cetakan Print_ V3 untuk field ini secara spesifik (field lain
//   semuanya identik).
// ============================================================

const PH_TAHAP_LABELS = ['Penetapan Harga', 'Penetapan Harga II', 'Penetapan Harga III', 'Penetapan Harga IV', 'Penetapan Harga V'];
const PH_ANTAM_DENOMINASI = [1, 2, 3, 5, 10, 25, 50, 100, 250, 500, 1000];

// Indeks kolom 0-based (hasil getValues()) -- lihat komentar audit di atas.
const PH_PV3_COL = {
  tanggal: 1, tahap: 2, pukul: 3, emasku: 4, idn: 8, stargold: 9, lotus: 10,
  apepi: 12, spot: 13, xauidrk: 14, lbma: 19, kitco: 20,
  hargaPenetapan: 25, hargaTE: 26, deposito: 27
};
const PH_NPD_COL = { tanggal: 1, tahap: 2, igs: 6, marva: 7 };
const PH_ANTAM_COL = { tanggal: 1, pukul: 2, gramasi: 3, acuan: 4, keputusan: 6 };

function phReadFullSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" tidak ditemukan di spreadsheet sumber Penetapan Harga.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  return sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
}

function phNum_(row, col) {
  if (!row) return 0;
  const v = Number(row[col]);
  return isNaN(v) ? 0 : v;
}

/**
 * Tahap Penetapan Harga (1-5) yang benar-benar ADA datanya untuk
 * satu tanggal tertentu di sheet "Price V3" (Trading Reguler) --
 * dipakai supaya dropdown Tahap di menu tidak menampilkan opsi
 * yang datanya kosong.
 */
function getPenetapanHargaTahapOptions(tanggalStr) {
  const target = extractDateOnly(tanggalStr || formatDateOnly(new Date()));
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const rows = phReadFullSheet_(ss, 'Price V3');

  const tahapSet = {};
  rows.forEach(function (r) {
    if (extractDateOnly(r[PH_PV3_COL.tanggal]) !== target) return;
    const t = Number(r[PH_PV3_COL.tahap]);
    if (t >= 1 && t <= 5) tahapSet[t] = true;
  });

  return Object.keys(tahapSet)
    .map(Number)
    .sort(function (a, b) { return a - b; })
    .map(function (t) { return { tahap: t, label: PH_TAHAP_LABELS[t - 1] }; });
}

// Asset visual "Download Laporan" — file asli yang di-upload user ke
// Google Drive, dipakai APA ADANYA (di-crop via CSS di sisi client,
// TIDAK diproses/diedit di server) untuk header & footer PDF Penetapan
// Harga. Dibaca lewat DriveApp (server-side, akun yang men-deploy web
// app ini) supaya tidak perlu hosting/URL publik.
//   - PDFPH_FOOTER_FILE_ID: foto wave hijau-teal-gold utk pita footer
//     dekoratif (revisi: gambar gold bar di header DIHAPUS total,
//     sisa kosong/white space di kanan atas sesuai instruksi user).
//   - PDFPH_LOGO_FILE_ID: logo resmi Pegadaian (lockup 3 lingkaran + wordmark).
const PDFPH_FOOTER_FILE_ID = '1-JfzFR5X1yGEdogeUJCHOTWgLvsEoeVT';
const PDFPH_LOGO_FILE_ID = '1afCndeegBJmehOcS7JT8038dkr-o-qeH';

function phFileToDataUri_(fileId) {
  const blob = DriveApp.getFileById(fileId).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

/**
 * Data URI (base64) kedua asset visual PDF Penetapan Harga di atas,
 * dalam SATU panggilan. Dipanggil LAZY oleh client hanya saat tombol
 * "Download Laporan" diklik (bukan saat halaman dibuka), lalu
 * di-cache di browser.
 */
function getPdfphAssets() {
  return { footer: phFileToDataUri_(PDFPH_FOOTER_FILE_ID), logo: phFileToDataUri_(PDFPH_LOGO_FILE_ID) };
}

/**
 * Data lengkap "Print Penetapan Harga" — Trading Reguler, untuk
 * satu Tanggal Penetapan + Tahap Penetapan tertentu.
 */
function getPenetapanHargaTradingReguler(tanggalStr, tahap) {
  const target = extractDateOnly(tanggalStr || formatDateOnly(new Date()));
  const t = Number(tahap) || 1;
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);

  const pv3Rows = phReadFullSheet_(ss, 'Price V3');
  const pv3Match = pv3Rows.find(function (r) {
    return extractDateOnly(r[PH_PV3_COL.tanggal]) === target && Number(r[PH_PV3_COL.tahap]) === t;
  });

  const npdRows = phReadFullSheet_(ss, 'New Price Discovery');
  const npdMatch = npdRows.find(function (r) {
    return extractDateOnly(r[PH_NPD_COL.tanggal]) === target && Number(r[PH_NPD_COL.tahap]) === t;
  });

  const referensi = [
    { sumber: 'Emasku', jenis: 'Fisik', harga: phNum_(pv3Match, PH_PV3_COL.emasku), satuan: 'Rupiah' },
    { sumber: 'IGS', jenis: 'Fisik', harga: phNum_(npdMatch, PH_NPD_COL.igs), satuan: 'Rupiah' },
    { sumber: 'Marva', jenis: 'Fisik', harga: phNum_(npdMatch, PH_NPD_COL.marva), satuan: 'Rupiah' },
    { sumber: 'IDN Bullion', jenis: 'Fisik', harga: phNum_(pv3Match, PH_PV3_COL.idn), satuan: 'Rupiah' },
    { sumber: 'StarGold', jenis: 'Fisik', harga: phNum_(pv3Match, PH_PV3_COL.stargold), satuan: 'Rupiah' },
    { sumber: 'Lotus Archi', jenis: 'Fisik', harga: phNum_(pv3Match, PH_PV3_COL.lotus), satuan: 'Rupiah' },
    { sumber: 'APEPI', jenis: 'Fisik', harga: phNum_(pv3Match, PH_PV3_COL.apepi), satuan: 'Rupiah' },
    { sumber: 'Spot Goldprice', jenis: 'Fisik', harga: phNum_(pv3Match, PH_PV3_COL.spot), satuan: 'Rupiah' },
    { sumber: 'XAU/IDR', jenis: 'Derivative', harga: phNum_(pv3Match, PH_PV3_COL.xauidrk), satuan: 'Rupiah' },
    { sumber: 'LBMA H-1', jenis: 'Fisik', harga: phNum_(pv3Match, PH_PV3_COL.lbma), satuan: 'Rupiah' },
    { sumber: 'Kitco H-1', jenis: 'Fisik', harga: phNum_(pv3Match, PH_PV3_COL.kitco), satuan: 'Rupiah' }
  ];

  const hargaPenetapan = phNum_(pv3Match, PH_PV3_COL.hargaPenetapan);
  const gramasiDeposito = phNum_(pv3Match, PH_PV3_COL.deposito);
  const hargaTE = phNum_(pv3Match, PH_PV3_COL.hargaTE);

  return {
    tanggal: target,
    tahap: t,
    tahapLabel: PH_TAHAP_LABELS[t - 1] || ('Tahap ' + t),
    pukul: pv3Match ? pv3Match[PH_PV3_COL.pukul] : null,
    hasData: !!pv3Match,
    referensi: referensi,
    hargaJual: {
      hargaPenetapan: hargaPenetapan,
      maksDiscManop: hargaPenetapan * (1 - 0.001),
      maksDiscKadep: hargaPenetapan * (1 - 0.002),
      maksDiscKadiv: hargaPenetapan * (1 - 0.003)
    },
    deposito: { gramasi: gramasiDeposito, hargaTE: hargaTE, valuasi: gramasiDeposito * hargaTE },
    piutang: getPiutangSupplierDashboard(target)
  };
}

/**
 * Data lengkap "Print Penetapan Harga" — ANTAM, untuk satu Tanggal
 * Penetapan (tanpa dimensi Tahap), per 11 denominasi.
 */
function getPenetapanHargaAntam(tanggalStr) {
  const target = extractDateOnly(tanggalStr || formatDateOnly(new Date()));
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);

  const antamRows = phReadFullSheet_(ss, 'Price ANTAM');
  const byGramasi = {};
  antamRows.forEach(function (r) {
    if (extractDateOnly(r[PH_ANTAM_COL.tanggal]) !== target) return;
    const g = Number(r[PH_ANTAM_COL.gramasi]);
    if (g > 0) byGramasi[g] = r;
  });

  const denominasi = PH_ANTAM_DENOMINASI.map(function (g) {
    const r = byGramasi[g];
    return { gramasi: g, acuan: phNum_(r, PH_ANTAM_COL.acuan), hargaJual: phNum_(r, PH_ANTAM_COL.keputusan) };
  });
  const hasData = denominasi.some(function (d) { return d.acuan || d.hargaJual; });

  // Deposito ikut Tahap 1 (lihat catatan audit di atas -- blok ANTAM di
  // Print_ V3 me-reuse cell Tahap Trading Reguler, tidak independen).
  const pv3Rows = phReadFullSheet_(ss, 'Price V3');
  const pv3Match = pv3Rows.find(function (r) {
    return extractDateOnly(r[PH_PV3_COL.tanggal]) === target && Number(r[PH_PV3_COL.tahap]) === 1;
  });
  const gramasiDeposito = phNum_(pv3Match, PH_PV3_COL.deposito);
  const hargaTE = phNum_(pv3Match, PH_PV3_COL.hargaTE);

  return {
    tanggal: target,
    pukul: '08.30',
    hasData: hasData,
    denominasi: denominasi,
    deposito: { gramasi: gramasiDeposito, hargaTE: hargaTE, valuasi: gramasiDeposito * hargaTE },
    piutang: getPiutangSupplierDashboard(target)
  };
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
 * Ranking Top Buyer — berdasarkan seluruh histori PenjualanTransaksi.
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
 * (data pembelian lengkap, termasuk pajak & diskon).
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
 * Supplier dengan rata-rata harga terendah dari SELURUH riwayat
 * SupplierPrice (bukan cuma satu tanggal).
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
 * Statistik margin dari SELURUH riwayat PriceSetting.
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
// HALAMAN PENJUALAN — ringkasan & ranking seluruh waktu
// ------------------------------------------------------------
function getSalesSummaryAll(startDate, endDate) {
  const list = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => isDateInPeriod(p.Tanggal, startDate, endDate));
  return {
    totalNominal: list.reduce((sum, p) => sum + Number(p.TotalHarga || 0), 0),
    totalGram: list.reduce((sum, p) => sum + Number(p.Qty || 0), 0),
    jumlahInvoice: list.length
  };
}

/**
 * Menjumlah kolom "Total Harga" (dicari by-keyword, aman terhadap variasi
 * teks header) dari sheet ETF di spreadsheet sumber (PRICE_DISCOVERY_SOURCE_ID)
 * — read-only, sheet sumber TIDAK PERNAH ditulis. Baris tanpa nilai (0/kosong)
 * dilewati (dianggap baris kosong di akhir tabel, bukan transaksi nyata).
 * dateStr opsional (yyyy-MM-dd) — kalau diisi, hanya baris dengan kolom
 * Tanggal yang cocok yang dihitung (dipakai Laporan Harian); kalau kosong,
 * seluruh baris dihitung (dipakai Dashboard, seperti semula).
 */
function readEgrSheetTotal_(sheetName, dateStr) {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" tidak ditemukan.');

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) throw new Error('Baris header tidak terdeteksi di ' + sheetName + '.');

  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const totalCol = findColByKeywords(headerMap, ['total', 'harga']);
  if (totalCol === -1) throw new Error('Kolom Total Harga tidak ditemukan di ' + sheetName + '.');
  // Kolom gramasi bersifat opsional — kalau tidak ada, gram tetap 0 (tidak error).
  const gramCol = findColByKeywords(headerMap, ['qty']) !== -1
    ? findColByKeywords(headerMap, ['qty'])
    : findColByKeywords(headerMap, ['gram']);
  // Kolom tanggal juga opsional — kalau diminta filter tanggal tapi kolomnya
  // tidak ketemu, filter dilewati (aman) daripada melempar error.
  const tanggalCol = dateStr ? findColByKeywords(headerMap, ['tanggal']) : -1;
  const targetDate = dateStr ? extractDateOnly(dateStr) : null;

  let nominal = 0, count = 0, gram = 0;
  values.slice(headerRowIdx + 1).forEach(row => {
    if (targetDate && tanggalCol !== -1 && extractDateOnly(getByCol(row, tanggalCol)) !== targetDate) return;
    const val = Number(getByCol(row, totalCol)) || 0;
    if (val > 0) {
      nominal += val;
      count++;
      if (gramCol !== -1) gram += Number(getByCol(row, gramCol)) || 0;
    }
  });
  return { nominal: nominal, count: count, gram: gram };
}

/**
 * Sama seperti readEgrSheetTotal_ tapi menerima RENTANG tanggal
 * (startDate/endDate 'yyyy-MM-dd', dua-duanya opsional) alih-alih
 * satu tanggal — dipakai Global Period Filter Dashboard. Dibuat
 * sebagai fungsi terpisah (bukan mengubah readEgrSheetTotal_) supaya
 * pemanggil lama (Laporan Harian, single-date) sama sekali tidak
 * ikut berubah.
 */
function readEgrSheetTotalRange_(sheetName, startDate, endDate) {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" tidak ditemukan.');

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) throw new Error('Baris header tidak terdeteksi di ' + sheetName + '.');

  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const totalCol = findColByKeywords(headerMap, ['total', 'harga']);
  if (totalCol === -1) throw new Error('Kolom Total Harga tidak ditemukan di ' + sheetName + '.');
  const gramCol = findColByKeywords(headerMap, ['qty']) !== -1
    ? findColByKeywords(headerMap, ['qty'])
    : findColByKeywords(headerMap, ['gram']);
  const tanggalCol = (startDate || endDate) ? findColByKeywords(headerMap, ['tanggal']) : -1;

  let nominal = 0, count = 0, gram = 0;
  values.slice(headerRowIdx + 1).forEach(row => {
    if (tanggalCol !== -1 && !isDateInPeriod(getByCol(row, tanggalCol), startDate, endDate)) return;
    const val = Number(getByCol(row, totalCol)) || 0;
    if (val > 0) {
      nominal += val;
      count++;
      if (gramCol !== -1) gram += Number(getByCol(row, gramCol)) || 0;
    }
  });
  return { nominal: nominal, count: count, gram: gram };
}

/**
 * Total Redemption EGR (Input_Redem_EGR) pada rentang tanggal — versi
 * AMAN dari readEgrSheetTotalRange_: kalau sheet/header/kolom redemption
 * tidak ada, mengembalikan 0 (tidak melempar error) supaya Dashboard,
 * Kinerja Bulanan, dan pemakai getDashboardKategoriBreakdown lain tidak
 * ikut gagal hanya karena sumber ketiga ini bermasalah.
 * Kriteria baris & kolom SAMA dengan readEgrSheetTotalRange_ (Total Harga
 * > 0, gram dari kolom Qty).
 */
function readEgrRedemTotalRangeSafe_(startDate, endDate) {
  const zero = { nominal: 0, count: 0, gram: 0, tersedia: false };
  try {
    const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
    if (!getEgrSheet_(ss, EGR_SHEET.REDEM)) return zero;
    const res = readEgrSheetTotalRange_(EGR_SHEET.REDEM, startDate, endDate);
    res.tersedia = true;
    return res;
  } catch (err) {
    Logger.log('readEgrRedemTotalRangeSafe_ gagal: ' + err.message);
    return zero;
  }
}

/**
 * Baris transaksi Redemption EGR (Input_Redem_EGR) pada satu tanggal.
 * Kriteria baris SAMA dengan perhitungan inventory kumulatif
 * (getEgrInventoryKumulatifSebelum_): Total Harga > 0 — supaya
 * Inventory "After" hari ini = "Before" hari berikutnya.
 * Nominal = Total Harga (bukan after tax), sama seperti Restock.
 */
function readEgrRedemRowsForDate_(dateStr) {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.REDEM);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return [];
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    nomor: findColByKeywordsMulti_(headerMap, [['nomor', 'tc'], ['nomor']]),
    hargaPerGram: findColByKeywords(headerMap, ['harga'], ['total', 'after', 'buyback']),
    qty: findColByKeywords(headerMap, ['qty']),
    totalHarga: findColByKeywords(headerMap, ['total', 'harga']),
    totalAfterTax: findColByKeywords(headerMap, ['after', 'tax']),
    pembeli: findColByKeywords(headerMap, ['pembeli'])
  };
  if (col.tanggal === -1 || col.totalHarga === -1 || col.qty === -1) return [];

  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  const rows = [];
  values.slice(headerRowIdx + 1).forEach(row => {
    const tanggal = getByCol(row, col.tanggal);
    if (!tanggal || extractDateOnly(tanggal) !== target) return;
    const totalHarga = Number(getByCol(row, col.totalHarga)) || 0;
    if (totalHarga <= 0) return;
    rows.push({
      NomorTC: col.nomor !== -1 ? (getByCol(row, col.nomor) || '-') : '-',
      HargaPerGram: col.hargaPerGram !== -1 ? (Number(getByCol(row, col.hargaPerGram)) || 0) : 0,
      Qty: Number(getByCol(row, col.qty)) || 0,
      TotalHarga: totalHarga,
      TotalAfterTax: col.totalAfterTax !== -1 ? (Number(getByCol(row, col.totalAfterTax)) || 0) : 0,
      Pembeli: col.pembeli !== -1 ? (getByCol(row, col.pembeli) || '-') : '-'
    });
  });
  return rows;
}

/**
 * Baris detail (bukan cuma total) transaksi ETF Penjualan pada satu
 * tanggal, dari sheet Input_Jual_EGR (read-only). Bentuk objeknya
 * dibuat SAMA seperti baris PenjualanTransaksi (NomorInvoice,
 * HargaPerGram, Qty, TotalHarga, Pembeli) supaya bisa langsung
 * digabung ke tabel Detail Penjualan di Laporan Harian, ditambah
 * Kategori: 'ETF'.
 */
function readEgrJualRowsForDate_(dateStr) {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return [];
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    // Header di Input_Subs_EGR = "Nomor TC" (nomor transaksi EGR);
    // "Nomor Invoice" tetap dicoba duluan utk kompatibilitas sheet lama.
    invoice: findColByKeywordsMulti_(headerMap, [['nomor', 'invoice'], ['nomor', 'tc']]),
    hargaPerGram: findColByKeywords(headerMap, ['harga'], ['total', 'after']),
    qty: findColByKeywords(headerMap, ['qty']),
    totalHarga: findColByKeywords(headerMap, ['total', 'harga']),
    pembeli: findColByKeywords(headerMap, ['pembeli'])
  };
  if (col.tanggal === -1 || col.totalHarga === -1) return [];

  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  const rows = [];
  values.slice(headerRowIdx + 1).forEach(row => {
    const tanggal = getByCol(row, col.tanggal);
    if (!tanggal || extractDateOnly(tanggal) !== target) return;
    const totalHarga = Number(getByCol(row, col.totalHarga)) || 0;
    if (totalHarga <= 0) return;
    rows.push({
      NomorInvoice: col.invoice !== -1 ? getByCol(row, col.invoice) : '-',
      HargaPerGram: col.hargaPerGram !== -1 ? (Number(getByCol(row, col.hargaPerGram)) || 0) : 0,
      Qty: col.qty !== -1 ? (Number(getByCol(row, col.qty)) || 0) : 0,
      TotalHarga: totalHarga,
      Pembeli: col.pembeli !== -1 ? (getByCol(row, col.pembeli) || 'EGR') : 'EGR',
      Kategori: 'ETF'
    });
  });
  return rows;
}

/**
 * Sama seperti readEgrJualRowsForDate_ tapi menerima RENTANG tanggal
 * (dipakai Kinerja & Profitabilitas Bulanan utk hitung "Jumlah buyer
 * aktif" gabungan Trading+EGR — getDashboardKategoriBreakdown tidak
 * mengekspos daftar Pembeli, cuma total). Fungsi terpisah, TIDAK
 * mengubah readEgrJualRowsForDate_ — pemanggil lama (Laporan Harian)
 * sama sekali tidak ikut berubah. Pola sama seperti readEgrSheetTotalRange_.
 */
function readEgrJualRowsForRange_(startDate, endDate) {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return [];
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    totalHarga: findColByKeywords(headerMap, ['total', 'harga']),
    pembeli: findColByKeywords(headerMap, ['pembeli'])
  };
  if (col.tanggal === -1 || col.totalHarga === -1) return [];

  const rows = [];
  values.slice(headerRowIdx + 1).forEach(row => {
    const tanggal = getByCol(row, col.tanggal);
    if (!tanggal || !isDateInPeriod(tanggal, startDate, endDate)) return;
    const totalHarga = Number(getByCol(row, col.totalHarga)) || 0;
    if (totalHarga <= 0) return;
    rows.push({ Pembeli: col.pembeli !== -1 ? (getByCol(row, col.pembeli) || 'EGR') : 'EGR' });
  });
  return rows;
}

// ------------------------------------------------------------
// SPK PENGIRIMAN — sumber "Biaya distribusi ditanggung Perusahaan"
// di Kinerja & Profitabilitas Bulanan. Sheet "SPK" di spreadsheet
// eksternal PRICE_DISCOVERY_SOURCE_ID (SAMA spreadsheet dengan Rekap/
// EGR/Price Discovery V3) — dikonfirmasi user, data tabelnya mulai
// dari kolom L (kolom sebelum L diabaikan sepenuhnya, kemungkinan
// blok lain di sheet yang sama — pola sama seperti sheet "Rekap" yang
// juga punya beberapa blok berdampingan). Kolom dicari berdasar NAMA
// HEADER ("Tanggal Kirim", "Estimasi Biaya"), BUKAN posisi kolom
// tetap — pola sama persis dengan reader sheet eksternal lain di file
// ini (readF3BarangMasuk_, readEgrBarangMasuk_, dst), supaya tahan
// kalau urutan kolom di sheet sumber berubah.
// ------------------------------------------------------------
const SPK_SHEET_NAME = 'SPK';
const SPK_START_COL_ = 11; // 0-based index -> kolom L (A=0)

function readSpkPengirimanRows_() {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = ss.getSheetByName(SPK_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SPK_SHEET_NAME + '" tidak ditemukan di spreadsheet sumber.');

  const allValues = sheet.getDataRange().getValues();
  const values = allValues.map(function (row) { return row.slice(SPK_START_COL_); }); // mulai kolom L
  // Sheet "SPK" punya blok form/kalkulator kosong di ~20 baris pertama
  // (dikonfirmasi via debug: header "Pengiriman | Tanggal Kirim | ...
  // | Estimasi Biaya | ..." baru muncul di baris ke-21) — default
  // findHeaderRowIndex cuma scan 20 baris, jadi di sini eksplisit
  // dilebarkan ke 40 baris supaya header sungguhan ikut ke-scan.
  const headerRowIdx = findHeaderRowIndex(values, ['tanggal kirim', 'estimasi biaya'], 40);
  if (headerRowIdx === -1) throw new Error('Baris header ("Tanggal Kirim"/"Estimasi Biaya") tidak terdeteksi di sheet "' + SPK_SHEET_NAME + '" (mulai kolom L).');

  const headerRow = values[headerRowIdx];
  const headerMap = buildHeaderMap(headerRow);
  const col = {
    tanggalKirim: findColByKeywords(headerMap, ['tanggal', 'kirim']),
    estimasiBiaya: findColByKeywords(headerMap, ['estimasi', 'biaya']),
    // Header "Adhoc" DUPLIKAT di sheet ini (2 kolom sama-sama bernama
    // "Adhoc" — satu berisi status "Reguler"/"Cancel", satu lagi selalu
    // kosong/tidak dipakai). buildHeaderMap cuma simpan kemunculan
    // TERAKHIR untuk nama yang sama, padahal kolom status ada di
    // kemunculan PERTAMA — jadi dicari langsung dari headerRow mentah,
    // bukan lewat headerMap/findColByKeywords, supaya ambil kolom yang
    // benar (dikonfirmasi dari data: kolom status persis setelah
    // "Gramasi").
    status: headerRow.findIndex(function (h) { return String(h).trim().toLowerCase() === 'adhoc'; })
  };
  if (col.tanggalKirim === -1) throw new Error('Kolom "Tanggal Kirim" tidak ditemukan di sheet "' + SPK_SHEET_NAME + '".');
  if (col.estimasiBiaya === -1) throw new Error('Kolom "Estimasi Biaya" tidak ditemukan di sheet "' + SPK_SHEET_NAME + '".');

  // Pengiriman berstatus "Cancel" DIKECUALIKAN dari total Biaya
  // Distribusi (dikonfirmasi user).
  const isCancel_ = function (v) { return String(v || '').trim().toLowerCase().indexOf('cancel') !== -1; };

  const rows = [];
  values.slice(headerRowIdx + 1).forEach(function (row) {
    const tanggalKirim = getByCol(row, col.tanggalKirim);
    if (!tanggalKirim) return;
    if (col.status !== -1 && isCancel_(getByCol(row, col.status))) return;
    const estimasiBiaya = Number(getByCol(row, col.estimasiBiaya)) || 0;
    rows.push({ TanggalKirim: tanggalKirim, EstimasiBiaya: estimasiBiaya });
  });
  return rows;
}

/**
 * SUM(Estimasi Biaya) seluruh SPK Pengiriman dengan Tanggal Kirim di
 * rentang startDate..endDate (inklusif, via isDateInPeriod — SAMA
 * helper tanggal yang sudah dipakai di seluruh Kinerja Bulanan).
 * Baris dengan Estimasi Biaya = 0 otomatis tidak menambah total
 * (ikut disertakan di SUM, tapi kontribusinya nol).
 */
function getBiayaDistribusiSpkTotal_(startDate, endDate) {
  const rows = readSpkPengirimanRows_();
  return rows
    .filter(function (r) { return isDateInPeriod(r.TanggalKirim, startDate, endDate); })
    .reduce(function (s, r) { return s + r.EstimasiBiaya; }, 0);
}

/**
 * Sama seperti readEgrJualRowsForDate_ tapi untuk Pembelian ETF, dari
 * sheet Input_Beli_EGR. Bentuk objeknya sama seperti baris
 * PembelianTransaksi (NomorPO, AfterDiskon, Qty, TotalHarga, Seller).
 */
function readEgrBeliRowsForDate_(dateStr) {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.BELI);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return [];
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    nomorPO: findColByKeywords(headerMap, ['nomor', 'purchase']),
    seller: findColByKeywords(headerMap, ['seller']),
    qty: findColByKeywords(headerMap, ['qty']),
    afterDiskon: findColByKeywords(headerMap, ['after', 'diskon']),
    harga: findColByKeywords(headerMap, ['harga'], ['after', 'total', 'simulasi']),
    totalHarga: findColByKeywords(headerMap, ['total', 'harga'])
  };
  if (col.tanggal === -1 || col.totalHarga === -1) return [];

  const target = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
  const rows = [];
  values.slice(headerRowIdx + 1).forEach(row => {
    const tanggal = getByCol(row, col.tanggal);
    if (!tanggal || extractDateOnly(tanggal) !== target) return;
    const totalHarga = Number(getByCol(row, col.totalHarga)) || 0;
    if (totalHarga <= 0) return;
    const afterDiskonRaw = col.afterDiskon !== -1 ? Number(getByCol(row, col.afterDiskon)) : 0;
    const harga = col.harga !== -1 ? (Number(getByCol(row, col.harga)) || 0) : 0;
    rows.push({
      NomorPO: col.nomorPO !== -1 ? getByCol(row, col.nomorPO) : '-',
      AfterDiskon: afterDiskonRaw || harga,
      Qty: col.qty !== -1 ? (Number(getByCol(row, col.qty)) || 0) : 0,
      TotalHarga: totalHarga,
      Seller: col.seller !== -1 ? getByCol(row, col.seller) : '-',
      Kategori: 'ETF'
    });
  });
  return rows;
}

/**
 * Total Trading (dari PenjualanTransaksi/PembelianTransaksi, seperti
 * sekarang) + ETF (dari sheet Input_Jual_EGR/Input_Beli_EGR, read-only)
 * untuk kartu KPI Penjualan/Pembelian/Margin/Jumlah Transaksi di Dashboard.
 * Dipakai KHUSUS untuk tampilan Dashboard — tidak mengubah getSalesSummaryAll/
 * getPembelianSummaryAll/calculateProfitMarginOverall itu sendiri, supaya
 * Laporan Harian/Bulanan dan halaman lain yang masih memakai fungsi-fungsi
 * itu tidak ikut berubah.
 */
/**
 * Set NomorPO yang saat ini valid di F.2.Input_Trans_Beli (sumber Trading
 * Pembelian) — dipakai untuk memastikan Trading Pembelian di Dashboard
 * cuma menghitung baris yang memang masih match dengan sumbernya sekarang
 * (baris PembelianTransaksi yang PO-nya sudah tidak ada di F.2 dianggap
 * bukan Trading lagi, termasuk yang ternyata untuk keperluan ETF).
 */
function getValidPembelianSourcePOs_() {
  const sourceSs = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sourceSheet = sourceSs.getSheetByName('F.2.Input_Trans_Beli');
  if (!sourceSheet) return null;

  const values = sourceSheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['tanggal', 'purchase', 'seller', 'qty', 'harga']);
  if (headerRowIdx === -1) return null;
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    nomorPO: findColByKeywords(headerMap, ['nomor', 'purchase']),
    qty: findColByKeywords(headerMap, ['qty'])
  };

  const validPOs = {};
  values.slice(headerRowIdx + 1).forEach(row => {
    const tanggal = getByCol(row, col.tanggal);
    const nomorPO = String(getByCol(row, col.nomorPO) || '').trim();
    const qty = Number(getByCol(row, col.qty)) || 0;
    if (tanggal && nomorPO && qty > 0) validPOs[nomorPO] = true;
  });
  return validPOs;
}

function getDashboardKategoriBreakdown(startDate, endDate) {
  const isEgr_ = v => String(v || '').trim().toUpperCase() === 'EGR';
  // Baris tanpa TotalHarga valid (misal penanda "end" sisa import lama)
  // bukan transaksi asli — dikecualikan juga.
  const isValidRow_ = p => Number(p.TotalHarga || 0) > 0;

  // Beberapa baris ETF pernah ke-sync ke PenjualanTransaksi (dari saat
  // sebelum Input_Jual_EGR dipisah) — dikecualikan dari Trading supaya
  // tidak dobel dengan angka ETF (dari Input_Jual_EGR di bawah).
  // isDateInPeriod(...) menerapkan Global Period Filter Dashboard — tanpa
  // startDate/endDate (mis. dipanggil dari halaman lain) perilakunya tetap
  // seperti semula (seluruh waktu).
  const penjualanTrading = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => isValidRow_(p) && !isEgr_(p.Pembeli) && isDateInPeriod(p.Tanggal, startDate, endDate));
  const tradingPenjualan = {
    nominal: penjualanTrading.reduce((s, p) => s + Number(p.TotalHarga || 0), 0),
    gram: penjualanTrading.reduce((s, p) => s + Number(p.Qty || 0), 0),
    count: penjualanTrading.length
  };

  // Pembelian: Trading = baris yang Nomor PO-nya masih match dengan
  // F.2.Input_Trans_Beli saat ini (bukan berdasar nama Seller — beberapa
  // pembelian ETF pakai Seller yang sama dengan Trading biasa, jadi
  // pembedanya adalah apakah baris itu masih tercatat sebagai Trading
  // di sumbernya sekarang).
  const validPOs = getValidPembelianSourcePOs_();
  const pembelianTrading = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL).filter(p => {
    if (!isValidRow_(p) || !isDateInPeriod(p.Tanggal, startDate, endDate)) return false;
    if (!validPOs) return true; // gagal baca sumber -> jangan filter apa-apa (aman)
    return !!validPOs[String(p.NomorPO || '').trim()];
  });
  const tradingPembelian = {
    nominal: pembelianTrading.reduce((s, p) => s + Number(p.TotalHarga || 0), 0),
    gram: pembelianTrading.reduce((s, p) => s + Number(p.Qty || 0), 0),
    count: pembelianTrading.length
  };

  const etfPenjualan = readEgrSheetTotalRange_(EGR_SHEET.JUAL, startDate, endDate);
  const etfPembelian = readEgrSheetTotalRange_(EGR_SHEET.BELI, startDate, endDate);
  // Redemption EGR ditampilkan TERPISAH — sengaja TIDAK dijumlahkan ke
  // penjualan/pembelian/total/margin sampai klasifikasi arus kasnya pasti.
  const etfRedemption = readEgrRedemTotalRangeSafe_(startDate, endDate);

  return {
    redemption: { etf: etfRedemption },
    penjualan: {
      trading: tradingPenjualan,
      etf: etfPenjualan,
      total: {
        nominal: tradingPenjualan.nominal + etfPenjualan.nominal,
        gram: tradingPenjualan.gram + etfPenjualan.gram,
        count: tradingPenjualan.count + etfPenjualan.count
      }
    },
    pembelian: {
      trading: tradingPembelian,
      etf: etfPembelian,
      total: {
        nominal: tradingPembelian.nominal + etfPembelian.nominal,
        gram: tradingPembelian.gram + etfPembelian.gram,
        count: tradingPembelian.count + etfPembelian.count
      }
    }
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
 * Nilai harian sederhana (angka saja, untuk sparkline mini-chart
 * di kartu KPI Dashboard) — dari transaksi pertama tercatat s/d
 * hari ini (bukan rolling window).
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
  const sentMap = getDealEmailSentMap_('PenjualanTransaksi');
  return sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => extractDateOnly(p.Tanggal) === target)
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt))
    .map(p => Object.assign(p, { DealEmailSent: !!sentMap[p.ID] }));
}

function getPenjualanTransaksiRecent(days) {
  days = days || 30;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  const sentMap = getDealEmailSentMap_('PenjualanTransaksi');
  return sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => parseDateValue(p.Tanggal) >= cutoff.getTime())
    .sort((a, b) => parseDateValue(b.Tanggal) - parseDateValue(a.Tanggal))
    .map(p => Object.assign(p, { DealEmailSent: !!sentMap[p.ID] }));
}

// ------------------------------------------------------------
// JADWAL PICKUP (Penjualan) — panel "Jadwal Pickup" menggantikan
// "Penjualan Hari Ini" di halaman Penjualan. Disimpan di 4 kolom BARU
// di sheet PenjualanTransaksi (ditambahkan idempoten oleh
// ensurePenjualanPickupColumns_ — kolom & data existing TIDAK
// diubah/dipindah, cuma ditambah di akhir, persis pola yang sudah ada
// utk kolom lain): MetodePengambilan, JamPengambilan, PICPengambil,
// StatusPickup. TanggalPengambilan & Catatan pakai kolom yang SUDAH
// ADA sebelumnya (tidak bikin kolom tanggal/catatan baru).
// ------------------------------------------------------------
const PICKUP_SLOTS = ['09:00', '10:00', '11:00', '13:00', '14:00', '15:00', '16:00'];

function pickupSlotLabel_(jam) {
  if (!jam || PICKUP_SLOTS.indexOf(jam) === -1) return '';
  const startH = parseInt(jam.split(':')[0], 10);
  const endH = startH + 1;
  const pad = n => String(n).padStart(2, '0');
  return pad(startH) + '.00–' + pad(endH) + '.00';
}

function ensurePenjualanPickupColumns_() {
  const sh = getSheet(SHEET_NAMES.PENJUALAN_FULL);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const needed = ['MetodePengambilan', 'JamPengambilan', 'PICPengambil', 'StatusPickup'];
  const missing = needed.filter(h => headers.indexOf(h) === -1);
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    formatHeader(sh, lastCol + missing.length);
  }

  // Kolom JamPengambilan HARUS format teks polos ("@") — kalau tidak,
  // Google Sheets otomatis mengonversi string "09:00" yang ditulis via
  // setValue() jadi nilai waktu/Date, sehingga tidak lagi match dgn
  // PICKUP_SLOTS ("09:00" dst) di manapun dibandingkan. Cukup dijalankan
  // sekali (ditandai via Script Properties) supaya tidak boros quota
  // Sheets API di setiap page load.
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('JAM_PENGAMBILAN_TEXT_FORMATTED')) {
    const headersNow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const jamColIdx = headersNow.indexOf('JamPengambilan');
    if (jamColIdx !== -1) {
      const numRows = Math.max(sh.getMaxRows() - 1, 1);
      sh.getRange(2, jamColIdx + 1, numRows, 1).setNumberFormat('@');
      props.setProperty('JAM_PENGAMBILAN_TEXT_FORMATTED', '1');
    }
  }
}

/**
 * Normalisasi nilai JamPengambilan jadi "HH:mm" polos, tahan terhadap
 * baris lama yang keburu ke-convert Google Sheets jadi Date/waktu
 * sebelum kolom diformat sebagai teks (lihat ensurePenjualanPickupColumns_).
 * sheetToObjects() sudah men-stringify Date jadi "yyyy-MM-dd'T'HH:mm:ss"
 * sebelum sampai sini, jadi cek pola waktu setelah "T" juga.
 */
function extractJamOnly_(value) {
  const str = String(value == null ? '' : value).trim();
  if (!str) return '';
  let m = str.match(/^(\d{2}):(\d{2})/);
  if (m) return m[1] + ':' + m[2];
  m = str.match(/T(\d{2}):(\d{2})/);
  if (m) return m[1] + ':' + m[2];
  return '';
}

/**
 * Info transaksi read-only utk modal "Atur/Edit Jadwal Pickup" —
 * dipanggil saat modal dibuka, supaya user tidak perlu input ulang
 * data transaksi yang sudah ada.
 */
function getPickupTransaksiInfo(id) {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  const row = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL).find(r => String(r.ID) === String(id));
  if (!row) throw new Error('Transaksi tidak ditemukan.');
  return {
    id: row.ID,
    pembeli: row.Pembeli,
    nomorInvoice: row.NomorInvoice,
    qty: Number(row.Qty) || 0,
    merk: row.Merk || '',
    tanggal: extractDateOnly(row.Tanggal),
    statusBayar: row.Status || '',
    metodePengambilan: row.MetodePengambilan || '',
    jamPengambilan: extractJamOnly_(row.JamPengambilan),
    picPengambil: row.PICPengambil || '',
    tanggalPengambilan: extractDateOnly(row.TanggalPengambilan) || '',
    catatan: (row.Catatan === 'Diimpor dari spreadsheet lama' ? '' : (row.Catatan || '')),
    statusPickup: row.StatusPickup || ''
  };
}

/**
 * Simpan jadwal pickup 1 transaksi. metode: 'Ambil Sendiri' | 'Brinks'
 * | 'Hold'. Tanggal & jam wajib utk Ambil Sendiri/Brinks, tidak wajib
 * utk Hold. Slot jam adalah pengelompokan waktu, BUKAN reservasi
 * eksklusif — banyak buyer/transaksi (Ambil Sendiri maupun Brinks)
 * boleh berbagi slot tanggal+jam yang sama, jadi tidak ada validasi
 * bentrok di sini.
 */
function simpanJadwalPickup(id, data) {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  data = data || {};
  const metode = String(data.metode || '').trim();
  if (['Ambil Sendiri', 'Brinks', 'Hold'].indexOf(metode) === -1) {
    throw new Error('Metode pengambilan tidak valid.');
  }

  let tanggal = String(data.tanggal || '').trim();
  let jam = String(data.jam || '').trim();
  const pic = String(data.pic || '').trim();
  const catatan = String(data.catatan || '').trim();

  if (metode === 'Hold') {
    tanggal = '';
    jam = '';
  } else {
    if (!tanggal) throw new Error('Tanggal pengambilan wajib diisi.');
    if (!jam || PICKUP_SLOTS.indexOf(jam) === -1) throw new Error('Jam pengambilan wajib dipilih.');
  }

  const statusPickup = metode === 'Hold' ? 'Hold - Belum Ada Jadwal' : 'Terjadwal';

  updateRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', id, {
    MetodePengambilan: metode === 'Hold' ? '' : metode,
    JamPengambilan: jam,
    PICPengambil: pic,
    TanggalPengambilan: tanggal,
    Catatan: catatan,
    StatusPickup: statusPickup
  });
  logAudit('UPDATE', 'PenjualanTransaksi', id, 'Atur jadwal pickup — ' + statusPickup + (jam ? ' (' + tanggal + ' ' + jam + ')' : ''));
  return { ok: true, statusPickup: statusPickup };
}

/**
 * Simpan jadwal pickup ke BEBERAPA invoice sekaligus — dipakai fitur
 * "Atur Jadwal Massal". Satu pengaturan (metode/tanggal/jam/pic/catatan)
 * diterapkan ke tiap ID dengan memanggil simpanJadwalPickup() SATU PER
 * SATU (bukan transaksi gabungan baru, nominal/gramasi asli tidak
 * disentuh). Kalau salah satu invoice gagal (mis. metode/tanggal/jam
 * tidak valid), invoice lain tetap lanjut diproses — hasil per-invoice
 * dikembalikan supaya client bisa menampilkan mana yang gagal.
 */
function simpanJadwalPickupMassal(ids, data) {
  requireRole('Viewer');
  ids = ids || [];
  if (!ids.length) throw new Error('Tidak ada transaksi yang dipilih.');
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const ok = [];
  const failed = [];
  ids.forEach(id => {
    const row = rows.find(r => String(r.ID) === String(id));
    const nomorInvoice = row ? row.NomorInvoice : String(id);
    try {
      simpanJadwalPickup(id, data);
      ok.push({ id: id, nomorInvoice: nomorInvoice });
    } catch (e) {
      failed.push({ id: id, nomorInvoice: nomorInvoice, error: e.message || String(e) });
    }
  });
  return { ok: ok, failed: failed };
}

/** Tandai pickup selesai — StatusPickup='Selesai', hilang dari Jadwal Pickup & Pending. */
function selesaikanPickup(id) {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  updateRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', id, { StatusPickup: 'Selesai' });
  logAudit('UPDATE', 'PenjualanTransaksi', id, 'Pickup selesai');
  return { ok: true };
}

/**
 * Data untuk modal "Pickup Massal" — SELURUH invoice yang belum selesai
 * pickup (StatusPickup != 'Selesai'), TERMASUK yang belum pernah punya
 * jadwal sama sekali. Tidak dibatasi 30 hari seperti "Barang Belum
 * Dilakukan Pickup" di panel Jadwal Pickup, karena tujuannya melihat
 * SELURUH pekerjaan pickup yang tersisa, bukan hanya yang terbaru.
 */
function getPenjualanBelumSelesaiPickup() {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  return sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(function (r) { return String(r.ID || '').trim() && (r.StatusPickup || '') !== 'Selesai'; })
    .map(function (r) {
      const tglAmbil = extractDateOnly(r.TanggalPengambilan) || '';
      const jamAmbil = extractJamOnly_(r.JamPengambilan);
      return {
        id: r.ID, nomorInvoice: r.NomorInvoice, pembeli: r.Pembeli, merk: r.Merk || '',
        qty: Number(r.Qty) || 0, tanggal: extractDateOnly(r.Tanggal),
        tanggalPengambilan: tglAmbil, jamPengambilan: jamAmbil, statusPickup: r.StatusPickup || ''
      };
    });
}

/**
 * Selesaikan pickup untuk BEBERAPA invoice sekaligus (fitur "Pickup
 * Massal") — TIDAK perlu memilih tanggal/hari/slot jam, StatusPickup
 * langsung ditandai 'Selesai' dengan timestamp submit saat ini (jadwal
 * tanggal/jam/metode yang mungkin sudah ada TIDAK diubah/dihapus, hanya
 * StatusPickup). ID di-dedupe & invoice yang sudah 'Selesai' dilewati
 * (idempotent) supaya aman kalau tombol submit ter-klik dobel / request
 * terkirim ulang — satu invoice tidak akan diproses dua kali dalam satu
 * pemanggilan maupun lintas pemanggilan. Audit trail per invoice memakai
 * logAudit yang sama seperti selesaikanPickup() satuan (mencatat user
 * yang login lewat getCurrentUser()).
 */
function selesaikanPickupMassal(ids) {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  ids = Array.from(new Set((ids || []).map(String).filter(function (s) { return s.trim(); })));
  if (!ids.length) throw new Error('Tidak ada invoice yang dipilih.');

  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const done = [];
  const skipped = [];
  let totalGram = 0;

  ids.forEach(function (id) {
    const row = rows.find(function (r) { return String(r.ID) === id; });
    if (!row) { skipped.push({ id: id, reason: 'Invoice tidak ditemukan' }); return; }
    if ((row.StatusPickup || '') === 'Selesai') { skipped.push({ id: id, nomorInvoice: row.NomorInvoice, reason: 'Sudah selesai' }); return; }
    updateRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', id, { StatusPickup: 'Selesai' });
    logAudit('UPDATE', 'PenjualanTransaksi', id, 'Pickup selesai (Pickup Massal) — ' + (row.NomorInvoice || id));
    totalGram += Number(row.Qty) || 0;
    done.push({ id: id, nomorInvoice: row.NomorInvoice });
  });

  return { doneCount: done.length, done: done, skippedCount: skipped.length, skipped: skipped, totalGram: totalGram };
}

/**
 * Preview Upload Jadwal Excel (fitur "Upload Jadwal Excel") — TIDAK
 * menyimpan apa pun. rows = [{nomorInvoice, tanggal, jam}] hasil parsing
 * client-side. Tiap baris dicocokkan ke NomorInvoice PenjualanTransaksi
 * (case-insensitive, spasi berlebih diabaikan — reuse rkNormNamaKey_ yang
 * sama dipakai fitur Alias Buyer) dan diberi status:
 *  - DITEMUKAN            : invoice ada, belum ada jadwal (Terjadwal/Dalam Proses)
 *  - SUDAH_ADA_JADWAL     : invoice ada, StatusPickup sudah Terjadwal/Dalam Proses
 *  - TIDAK_DITEMUKAN      : No. Invoice tidak ada di data penjualan
 *  - DATA_TIDAK_VALID     : format Tanggal Pickup / Jam Pickup tidak valid
 */
function previewImportJadwalPickup(rows) {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  const dbRows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const byInvoice = {};
  dbRows.forEach(function (r) {
    const key = rkNormNamaKey_(r.NomorInvoice);
    if (key && !byInvoice[key]) byInvoice[key] = r;
  });

  return (rows || []).map(function (r) {
    const nomorInvoiceExcel = String(r.nomorInvoice || '').trim();
    const tanggal = extractDateOnly(r.tanggal);
    const jam = extractJamOnly_(r.jam) || String(r.jam || '').trim();
    const row = byInvoice[rkNormNamaKey_(nomorInvoiceExcel)];

    if (!row) {
      return { id: '', nomorInvoiceExcel: nomorInvoiceExcel, tanggal: tanggal, jam: jam, pembeli: '', merk: '', qty: 0, jadwalLama: '', status: 'TIDAK_DITEMUKAN' };
    }
    if (!tanggal || PICKUP_SLOTS.indexOf(jam) === -1) {
      return { id: row.ID, nomorInvoiceExcel: row.NomorInvoice, tanggal: tanggal, jam: jam, pembeli: row.Pembeli, merk: row.Merk || '', qty: Number(row.Qty) || 0, jadwalLama: '', status: 'DATA_TIDAK_VALID' };
    }
    const sudahAda = ['Terjadwal', 'Dalam Proses'].indexOf(row.StatusPickup || '') !== -1;
    return {
      id: row.ID, nomorInvoiceExcel: row.NomorInvoice, tanggal: tanggal, jam: jam,
      pembeli: row.Pembeli, merk: row.Merk || '', qty: Number(row.Qty) || 0,
      jadwalLama: sudahAda ? (extractDateOnly(row.TanggalPengambilan) + ' ' + extractJamOnly_(row.JamPengambilan)) : '',
      status: sudahAda ? 'SUDAH_ADA_JADWAL' : 'DITEMUKAN'
    };
  });
}

/**
 * Import Jadwal Pickup massal dari Excel — rows = [{id, tanggal, jam}]
 * hasil pencocokan previewImportJadwalPickup() yang sudah difilter di
 * client (hanya DITEMUKAN, atau SUDAH_ADA_JADWAL dengan pilihan
 * "Update"). Metode/PIC/Catatan TIDAK diminta dari file — kalau invoice
 * sudah punya MetodePengambilan dipertahankan, kalau belum default ke
 * 'Ambil Sendiri', supaya user tidak perlu mengisi ulang data yang sudah
 * ada. Disimpan lewat simpanJadwalPickup() satu per satu — mekanisme &
 * StatusPickup yang dihasilkan SAMA PERSIS dengan Atur Jadwal manual.
 */
function importJadwalPickupMassal(rows) {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  rows = rows || [];
  if (!rows.length) throw new Error('Tidak ada data untuk diimpor.');

  const dbRows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const done = [];
  const failed = [];
  rows.forEach(function (r) {
    const row = dbRows.find(function (x) { return String(x.ID) === String(r.id); });
    if (!row) { failed.push({ id: r.id, error: 'Invoice tidak ditemukan' }); return; }
    const tanggal = extractDateOnly(r.tanggal);
    const jam = String(r.jam || '').trim();
    if (!tanggal) { failed.push({ id: r.id, nomorInvoice: row.NomorInvoice, error: 'Tanggal tidak valid' }); return; }
    if (PICKUP_SLOTS.indexOf(jam) === -1) { failed.push({ id: r.id, nomorInvoice: row.NomorInvoice, error: 'Jam tidak valid' }); return; }
    const metode = ['Ambil Sendiri', 'Brinks'].indexOf(row.MetodePengambilan) !== -1 ? row.MetodePengambilan : 'Ambil Sendiri';
    try {
      simpanJadwalPickup(row.ID, {
        metode: metode, tanggal: tanggal, jam: jam,
        pic: row.PICPengambil || '', catatan: (row.Catatan === 'Diimpor dari spreadsheet lama' ? '' : (row.Catatan || ''))
      });
      done.push({ id: row.ID, nomorInvoice: row.NomorInvoice });
    } catch (e) {
      failed.push({ id: row.ID, nomorInvoice: row.NomorInvoice, error: e.message || String(e) });
    }
  });

  return { doneCount: done.length, done: done, failedCount: failed.length, failed: failed };
}

/**
 * Operasi satu-kali (per instruksi user, 31 Agustus 2026): tandai SEMUA
 * transaksi yang saat ini muncul di "Barang Belum Dipickup" sebagai
 * Selesai Diambil, KECUALI 25 invoice di bawah yang dikonfirmasi masih
 * benar-benar menunggu pickup. Logika filter "belum dipickup" identik
 * dengan getJadwalPickup().pending (30 hari terakhir, StatusPickup
 * bukan Selesai/Terjadwal/Dalam Proses). Hanya StatusPickup yang
 * diubah — nominal, gramasi, jadwal, dan data lain TIDAK disentuh.
 */
function pjlSelesaikanBelumPickupKecuali31Agu() {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  const KECUALI = [
    '00678/INV-00017/2026', '00679/INV-00017/2026', '00684/INV-00017/2026',
    '00685/INV-00017/2026', '00686/INV-00017/2026', '00690/INV-00017/2026',
    '00691/INV-00017/2026', '00692/INV-00017/2026', '00693/INV-00017/2026',
    '00695/INV-00017/2026', '00698/INV-00017/2026', '00703/INV-00017/2026',
    '00704/INV-00017/2026', '00705/INV-00017/2026', '00706/INV-00017/2026',
    '00707/INV-00017/2026', '00708/INV-00017/2026', '00710/INV-00017/2026',
    '00711/INV-00017/2026', '00712/INV-00017/2026', '00713/INV-00017/2026',
    '00714/INV-00017/2026', '00715/INV-00017/2026', '00717/INV-00017/2026',
    '00719/INV-00017/2026'
  ];
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL).filter(r => String(r.ID || '').trim());
  const cutoffTime = (function () { const d = new Date(); d.setDate(d.getDate() - 30); return d.getTime(); })();
  const pending = rows.filter(p => {
    const st = p.StatusPickup || '';
    if (st === 'Selesai' || st === 'Terjadwal' || st === 'Dalam Proses') return false;
    return parseDateValue(p.Tanggal) >= cutoffTime;
  });
  const done = [];
  const dikecualikan = [];
  pending.forEach(p => {
    const inv = String(p.NomorInvoice || '').trim();
    if (KECUALI.indexOf(inv) !== -1) { dikecualikan.push(inv); return; }
    updateRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', p.ID, { StatusPickup: 'Selesai' });
    logAudit('UPDATE', 'PenjualanTransaksi', p.ID, 'Pickup selesai (bulk, kecuali daftar 31 Agu 2026) — ' + inv);
    done.push(inv);
  });
  return { doneCount: done.length, done: done, dikecualikanCount: dikecualikan.length, dikecualikan: dikecualikan };
}

/**
 * Koreksi satu-kali (per instruksi user, 31 Agu 2026): TanggalPengambilan
 * 3 invoice PT King Halim Jewelry (sebelumnya kesimpan 31/08/2026)
 * dikembalikan ke 28/08/2026 — tanggal pickup yang sebenarnya. Hanya
 * TanggalPengambilan yang diubah; StatusPickup ('Selesai') dan data
 * lain tidak disentuh.
 */
function pjlUbahJadwalKingHalim28Agu() {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  const TARGETS = ['00691/INV-00017/2026', '00692/INV-00017/2026', '00693/INV-00017/2026'];
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  let done = 0;
  const notFound = [];
  TARGETS.forEach(inv => {
    const row = rows.find(r => String(r.NomorInvoice || '').trim() === inv);
    if (!row) { notFound.push(inv); return; }
    updateRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', row.ID, { TanggalPengambilan: '2026-08-28' });
    logAudit('UPDATE', 'PenjualanTransaksi', row.ID, 'Koreksi TanggalPengambilan ke 28 Agustus 2026 (King Halim, per instruksi user) — ' + inv);
    done++;
  });
  return { done: done, notFound: notFound };
}

/**
 * Perbaikan satu-kali: sejumlah invoice sempat ikut ke-tandai
 * 'Selesai' oleh migrasi tanggal 18 Agustus 2026, padahal
 * pengirimannya belum diproses. Kembalikan ke status kosong supaya
 * transaksi ini muncul lagi di "Barang Belum Dilakukan Pickup" dan
 * bisa dipilih metode pengirimannya.
 */
function bukaKembaliPickupByInvoice(nomorInvoiceList) {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  const targets = (nomorInvoiceList || []).map(s => String(s).trim());
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  let done = 0;
  const notFound = [];
  targets.forEach(nomorInvoice => {
    const row = rows.find(r => String(r.NomorInvoice || '').trim() === nomorInvoice);
    if (!row) { notFound.push(nomorInvoice); return; }
    updateRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', row.ID, { StatusPickup: '' });
    logAudit('UPDATE', 'PenjualanTransaksi', row.ID, 'Buka kembali pickup (koreksi migrasi) — ' + nomorInvoice);
    done++;
  });
  return { ok: true, done: done, notFound: notFound };
}

/**
 * Perbaikan satu-kali: invoice yang sebenarnya SUDAH punya jadwal
 * pickup (MetodePengambilan/TanggalPengambilan/JamPengambilan sudah
 * terisi) tapi StatusPickup-nya kosong, sehingga dobel muncul — baik
 * di slot jadwal maupun di "Barang Belum Dilakukan Pickup". Kembalikan
 * StatusPickup ke 'Terjadwal' saja; TIDAK menyentuh data jadwal yang
 * sudah ada.
 */
function tandaiTerjadwalByInvoice(nomorInvoiceList) {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  const targets = (nomorInvoiceList || []).map(s => String(s).trim());
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  let done = 0;
  const notFound = [];
  targets.forEach(nomorInvoice => {
    const row = rows.find(r => String(r.NomorInvoice || '').trim() === nomorInvoice);
    if (!row) { notFound.push(nomorInvoice); return; }
    updateRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', row.ID, { StatusPickup: 'Terjadwal' });
    logAudit('UPDATE', 'PenjualanTransaksi', row.ID, 'Tandai Terjadwal (koreksi dobel-listing) — ' + nomorInvoice);
    done++;
  });
  return { ok: true, done: done, notFound: notFound };
}

/** Batalkan jadwal pickup — masuk lagi ke "Barang Belum Dilakukan Pickup". */
function batalkanPickup(id) {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  updateRowById(SHEET_NAMES.PENJUALAN_FULL, 'ID', id, { StatusPickup: 'Dibatalkan' });
  logAudit('UPDATE', 'PenjualanTransaksi', id, 'Pickup dibatalkan');
  return { ok: true };
}

/**
 * Data Jadwal Pickup utk 1 tanggal (berdasarkan TanggalPengambilan,
 * BUKAN tanggal transaksi) + daftar "Barang Belum Dilakukan Pickup"
 * (transaksi 30 hari terakhir yang belum Selesai/Terjadwal/Dalam
 * Proses — termasuk Hold, Dibatalkan, atau belum pernah diatur sama
 * sekali).
 */
function getJadwalPickup(tanggal) {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  const target = tanggal ? extractDateOnly(tanggal) : formatDateOnly(new Date());

  const all = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL).filter(p => String(p.ID || '').trim());
  const slotMap = {};
  PICKUP_SLOTS.forEach(s => { slotMap[s] = []; });

  all.forEach(p => {
    // Transaksi Selesai TETAP tampil di slot jadwal aslinya (badge "✓
    // Selesai Diambil" ditambahkan di sisi client) — hanya Dibatalkan
    // yang disingkirkan dari tampilan slot.
    if (p.StatusPickup === 'Dibatalkan') return;
    if (extractDateOnly(p.TanggalPengambilan) !== target) return;
    const jam = extractJamOnly_(p.JamPengambilan);
    if (PICKUP_SLOTS.indexOf(jam) === -1) return;
    slotMap[jam].push({
      id: p.ID, pembeli: p.Pembeli, nomorInvoice: p.NomorInvoice, qty: Number(p.Qty) || 0,
      merk: p.Merk || '', metode: p.MetodePengambilan || '', statusPickup: p.StatusPickup || ''
    });
  });

  const slots = PICKUP_SLOTS.map(jam => ({ jam: jam, label: pickupSlotLabel_(jam), items: slotMap[jam] }));

  const cutoffTime = (function () { const d = new Date(); d.setDate(d.getDate() - 30); return d.getTime(); })();
  const pending = all.filter(p => {
    const st = p.StatusPickup || '';
    if (st === 'Selesai' || st === 'Terjadwal' || st === 'Dalam Proses') return false;
    return parseDateValue(p.Tanggal) >= cutoffTime;
  }).map(p => ({
    id: p.ID, pembeli: p.Pembeli, nomorInvoice: p.NomorInvoice, qty: Number(p.Qty) || 0,
    merk: p.Merk || '', tanggal: extractDateOnly(p.Tanggal), statusPickup: p.StatusPickup || 'Hold - Belum Ada Jadwal'
  })).sort((a, b) => a.tanggal < b.tanggal ? -1 : (a.tanggal > b.tanggal ? 1 : 0));

  return { tanggal: target, slots: slots, pending: pending };
}

/**
 * Scorecard operasional utk halaman Penjualan: "Pickup Hari Ini" &
 * "Barang Belum Dipickup". "Pickup Hari Ini" dihitung dari
 * TanggalPengambilan (bukan tanggal transaksi) = hari ini, MENCAKUP baik
 * yang masih Menunggu Pickup maupun yang sudah Selesai Diambil (dengan
 * rincian total/selesai/menunggu) — hanya Dibatalkan yang tidak dihitung.
 * "Barang Belum Dipickup" TIDAK PERNAH memasukkan transaksi Selesai.
 */
function getPickupScorecards() {
  requireRole('Viewer');
  ensurePenjualanPickupColumns_();
  const todayStr = formatDateOnly(new Date());
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL).filter(r => String(r.ID || '').trim());

  // "Pickup Hari Ini" — SEMUA transaksi dijadwalkan (TanggalPengambilan)
  // hari ini, baik Menunggu maupun Selesai Diambil. Dibatalkan tidak dihitung.
  const hariIniAll = rows.filter(r => {
    const st = r.StatusPickup || '';
    return st !== 'Dibatalkan' && extractDateOnly(r.TanggalPengambilan) === todayStr;
  });
  const hariIniSelesai = hariIniAll.filter(r => (r.StatusPickup || '') === 'Selesai');
  const hariIniMenunggu = hariIniAll.filter(r => (r.StatusPickup || '') !== 'Selesai');

  // "Barang Belum Dipickup" — HARUS sama persis dengan daftar "Barang
  // Belum Dilakukan Pickup" di panel Jadwal Pickup (getJadwalPickup:
  // 30 hari terakhir, StatusPickup bukan Selesai/Terjadwal/Dalam
  // Proses — StatusPickup AKTUAL adalah satu-satunya penentu, BUKAN
  // tanggal jadwal). Sebelumnya digabung (union) dengan hariIniMenunggu,
  // yang hanya mengecualikan 'Selesai' — TIDAK mengecualikan
  // 'Terjadwal'/'Dalam Proses' — sehingga transaksi yang SUDAH punya
  // jadwal pickup hari ini (bukan outstanding) ikut kehitung di sini,
  // mencampur konsep "Jadwal Pickup" dengan "Belum Dipickup". Union
  // itu dibuang; pakai pendingList apa adanya supaya konsisten dengan
  // daftar yang tampil di UI.
  const cutoffTime = (function () { const d = new Date(); d.setDate(d.getDate() - 30); return d.getTime(); })();
  const belumDipickup = rows.filter(r => {
    const st = r.StatusPickup || '';
    if (st === 'Selesai' || st === 'Terjadwal' || st === 'Dalam Proses') return false;
    return parseDateValue(r.Tanggal) >= cutoffTime;
  });

  function summarize(list) {
    const gram = list.reduce((sum, r) => sum + (Number(r.Qty) || 0), 0);
    const buyerSet = {};
    list.forEach(r => { buyerSet[String(r.Pembeli || '').trim()] = true; });
    return { gram: gram, buyerCount: Object.keys(buyerSet).length, invoiceCount: list.length };
  }

  return {
    tanggal: todayStr,
    hariIni: {
      total: summarize(hariIniAll),
      selesai: summarize(hariIniSelesai),
      menunggu: summarize(hariIniMenunggu)
    },
    belumDipickup: summarize(belumDipickup)
  };
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
  const sentMap = getDealEmailSentMap_('PembelianTransaksi');
  return sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL)
    .filter(p => extractDateOnly(p.Tanggal) === target)
    .sort((a, b) => new Date(b.CreatedAt) - new Date(a.CreatedAt))
    .map(p => Object.assign(p, { DealEmailSent: !!sentMap[p.ID] }));
}

function getPembelianTransaksiRecent(days) {
  days = days || 30;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - days);
  const sentMap = getDealEmailSentMap_('PembelianTransaksi');
  return sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL)
    .filter(p => parseDateValue(p.Tanggal) >= cutoff.getTime())
    .sort((a, b) => parseDateValue(b.Tanggal) - parseDateValue(a.Tanggal))
    .map(p => Object.assign(p, { DealEmailSent: !!sentMap[p.ID] }));
}

// ------------------------------------------------------------
// KINERJA & PROFITABILITAS BULANAN
// Template/source of truth: sheet "Kinerja Bulanan" (diaudit dari file
// "Laporan_Bulanan_Trading.xlsx" yang diberikan user — workbook Kertas
// Kerja Monitoring Bulanan Bullion Trading). Struktur (urutan No,
// nama indikator, satuan, formula ANALISIS) di-reuse APA ADANYA dari
// hasil audit, TIDAK dibuat ulang berdasarkan asumsi.
//
// Keputusan hasil audit (dikonfirmasi user):
// 1. Nomor indikator template asli meloncat (1,2,3,4,8,9,10,11,12,13,16
//    — gap di 5,6,7,14,15). User: "typo kayanya" -> gap dibiarkan apa
//    adanya, TIDAK diisi/direkonstruksi.
// 2. C21/C22 template asli ("Laba kotor penjualan"/"Margin laba kotor")
//    formulanya #REF! (referensi ke kolom HPP/Modal Penjualan yang
//    sudah terhapus dari sheet asli). User: "kosongin dulu aja" ->
//    di sini DIKOSONGKAN (bukan 0, bukan #REF!) sampai ada sumber HPP.
// 3. C25 ("Perputaran persediaan") formulanya juga #REF! (referensi ke
//    "Stok Awal" yang tidak ada baris inputnya). User: "kosongin dulu
//    aja" -> sama, dikosongkan.
// 4. "Pembelian wholesale – nilai": user konfirmasi SEBELUM PAJAK ->
//    pakai TotalHarga (pre PajakWapu), konsisten dengan "Penjualan –
//    nilai (sebelum pajak)".
// 5. User setuju bikin sheet manual baru (KinerjaBulananManual) untuk
//    4 indikator yang belum ada sumber datanya di database (Buyback
//    retail, Biaya distribusi, Stok akhir, Keluhan nasabah).
// 6. Susulan: user konfirmasi Penjualan & Pembelian DIGABUNG Trading +
//    EGR (bukan Trading-only lagi). Reuse getDashboardKategoriBreakdown()
//    APA ADANYA — fungsi SAMA yang sudah dipakai kartu KPI Dashboard
//    untuk gabungan ini, termasuk logic dedup-nya (baris PenjualanTransaksi
//    lama yang Pembeli='EGR' dikecualikan dari Trading; Pembelian Trading
//    difilter berdasar NomorPO yang masih valid di F.2.Input_Trans_Beli)
//    — TIDAK ditulis ulang di sini supaya tidak berisiko salah dedup.
// 7. Susulan: "Biaya distribusi ditanggung Perusahaan" TIDAK manual
//    lagi — user beri sumbernya: sheet "SPK" di spreadsheet
//    PRICE_DISCOVERY_SOURCE_ID (mulai kolom L), SUM(Estimasi Biaya)
//    di-filter Tanggal Kirim = bulan/YTD periode (readSpkPengirimanRows_/
//    getBiayaDistribusiSpkTotal_), status "Cancel" dikecualikan. Kolom
//    BiayaDistribusi di sheet KinerjaBulananManual TIDAK dihapus (jaga
//    histori lama), cuma sudah tidak dibaca/ditulis aplikasi lagi.
// 8. Susulan: "Stok akhir – gram" TIDAK manual lagi — user: "gold in
//    transit itu piutang ya, samain aja kalo piutang egr dan trading" /
//    "Stok akhir – gram itu piutang, samain aja". Sumbernya Gold in
//    Transit (Piutang Supplier) GABUNGAN Trading+EGR, reuse
//    getPiutangSupplierDashboardCombined(tanggalPosisi).totalGramasi
//    APA ADANYA — fungsi yang sama dipakai card Dashboard tab
//    Gabungan. Posisi = AKHIR bulan terpilih (saldo, bukan arus/flow).
//    Kolom StokAkhirGram di KinerjaBulananManual TIDAK dihapus, cuma
//    sudah tidak dibaca/ditulis lagi.
// ------------------------------------------------------------
const KINERJA_BULANAN_INDIKATOR = [
  { no: 1, key: 'penjualanGram', label: 'Penjualan – gram', satuan: 'g', source: 'auto', sourceLabel: 'Data PenjualanTransaksi + EGR', numfmt: '#,##0.000' },
  { no: 2, key: 'penjualanNilai', label: 'Penjualan – nilai (Rp, sebelum pajak)', satuan: 'Rp', source: 'auto', sourceLabel: 'Data PenjualanTransaksi + EGR', numfmt: '#,##0' },
  { no: 3, key: 'penjualanTrx', label: 'Penjualan – jumlah transaksi', satuan: 'trx', source: 'auto', sourceLabel: 'Data PenjualanTransaksi + EGR', numfmt: '#,##0' },
  { no: 4, key: 'buyerAktif', label: 'Jumlah buyer aktif', satuan: 'pihak', source: 'auto', sourceLabel: 'Data PenjualanTransaksi + EGR', numfmt: '#,##0' },
  { no: 8, key: 'pembelianGram', label: 'Pembelian wholesale – gram', satuan: 'g', source: 'auto', sourceLabel: 'Data PembelianTransaksi + EGR', numfmt: '#,##0.000' },
  { no: 9, key: 'pembelianNilai', label: 'Pembelian wholesale – nilai (sebelum pajak)', satuan: 'Rp', source: 'auto', sourceLabel: 'Data PembelianTransaksi + EGR', numfmt: '#,##0' },
  { no: 10, key: 'buybackGram', label: 'Buyback retail – gram', satuan: 'g', source: 'manual', sourceLabel: 'Input manual', numfmt: '#,##0.000' },
  { no: 11, key: 'buybackNilai', label: 'Buyback retail – nilai', satuan: 'Rp', source: 'manual', sourceLabel: 'Input manual', numfmt: '#,##0' },
  { no: 12, key: 'biayaDistribusi', label: 'Biaya distribusi ditanggung Perusahaan', satuan: 'Rp', source: 'auto', sourceLabel: 'Data SPK Pengiriman', numfmt: '#,##0' },
  { no: 13, key: 'stokAkhir', label: 'Stok akhir – gram', satuan: 'g', source: 'auto', sourceLabel: 'Data Gold in Transit / Piutang (Trading + EGR)', numfmt: '#,##0.000' },
  { no: 16, key: 'keluhanNasabah', label: 'Keluhan nasabah diterima', satuan: 'kasus', source: 'manual', sourceLabel: 'Input manual', numfmt: '#,##0' }
];

const KINERJA_BULANAN_ANALISIS = [
  { key: 'hargaJualRataRata', label: 'Harga jual rata-rata', satuan: 'Rp/g', numfmt: '#,##0' },
  { key: 'hargaBeliRataRata', label: 'Harga beli wholesale rata-rata', satuan: 'Rp/g', numfmt: '#,##0' },
  {
    key: 'labaKotorPenjualan', label: 'Laba kotor penjualan', satuan: 'Rp', numfmt: '#,##0', pending: true,
    pendingNote: 'Template asli #REF! (referensi HPP/Modal Penjualan sudah terhapus). Dikosongkan sampai ada sumber data HPP — sesuai instruksi user, bukan diisi asumsi.'
  },
  {
    key: 'marginLabaKotor', label: 'Margin laba kotor', satuan: '%', numfmt: '0.00%', pending: true,
    pendingNote: 'Sama seperti Laba kotor penjualan — menunggu sumber data HPP/Modal Penjualan.'
  },
  { key: 'spreadRealisasi', label: 'Spread realisasi jual vs beli wholesale', satuan: '%', numfmt: '0.00%' },
  {
    key: 'selisihMutasiStok', label: 'Selisih mutasi stok tak terjelaskan (stok akhir − stok awal − pembelian − buyback + penjualan )', satuan: 'g', numfmt: '#,##0.000', pending: true,
    pendingNote: 'Template asli juga tidak punya formula di sel ini — butuh indikator "Stok Awal" yang belum ada input-nya.'
  },
  {
    key: 'perputaranPersediaan', label: 'Perputaran persediaan (penjualan gram ÷ rata-rata stok)', satuan: 'kali', numfmt: '0.00"x"', pending: true,
    pendingNote: 'Template asli #REF! (referensi "Stok Awal" sudah terhapus). Dikosongkan sampai ada input Stok Awal.'
  }
];

const KINERJA_BULANAN_MANUAL_HEADERS = ['Periode', 'BuybackGram', 'BuybackNilai', 'BiayaDistribusi', 'StokAkhirGram', 'KeluhanNasabah', 'UpdatedAt', 'UpdatedBy'];

function getOrCreateKinerjaBulananManualSheet_() {
  const db = getDb();
  let sh = db.getSheetByName(SHEET_NAMES.KINERJA_BULANAN_MANUAL);
  if (sh) return sh;
  sh = db.insertSheet(SHEET_NAMES.KINERJA_BULANAN_MANUAL);
  sh.getRange(1, 1, 1, KINERJA_BULANAN_MANUAL_HEADERS.length).setValues([KINERJA_BULANAN_MANUAL_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, KINERJA_BULANAN_MANUAL_HEADERS.length);
  return sh;
}

function getKinerjaBulananManual_(periode) {
  getOrCreateKinerjaBulananManualSheet_();
  const row = sheetToObjects(SHEET_NAMES.KINERJA_BULANAN_MANUAL).find(r => r.Periode === periode);
  return {
    buybackGram: row ? Number(row.BuybackGram || 0) : 0,
    buybackNilai: row ? Number(row.BuybackNilai || 0) : 0,
    keluhanNasabah: row ? Number(row.KeluhanNasabah || 0) : 0,
    updatedAt: row ? row.UpdatedAt : null,
    updatedBy: row ? row.UpdatedBy : null
  };
}

/**
 * Jumlah indikator manual dari Januari s/d `uptoMonth` di tahun `year`
 * (dipakai kolom YTD). HANYA untuk indikator yang sifatnya arus/flow
 * per bulan (Buyback, Keluhan nasabah) — dijumlah antar bulan. "Stok
 * akhir" TIDAK ikut dijumlah di sini karena itu saldo/posisi (bukan
 * arus) — YTD-nya = nilai bulan terpilih itu sendiri, ditangani
 * terpisah di getKinerjaBulananData(). "Biaya distribusi" TIDAK lagi
 * di sini sejak jadi AUTO (sumber SPK Pengiriman) — lihat
 * getBiayaDistribusiSpkTotal_().
 */
function getKinerjaBulananManualRange_(year, uptoMonth) {
  getOrCreateKinerjaBulananManualSheet_();
  const rows = sheetToObjects(SHEET_NAMES.KINERJA_BULANAN_MANUAL).filter(function (r) {
    const parts = String(r.Periode || '').split('-');
    const y = Number(parts[0]), m = Number(parts[1]);
    return y === year && m >= 1 && m <= uptoMonth;
  });
  const sum = { buybackGram: 0, buybackNilai: 0, keluhanNasabah: 0 };
  rows.forEach(function (r) {
    sum.buybackGram += Number(r.BuybackGram || 0);
    sum.buybackNilai += Number(r.BuybackNilai || 0);
    sum.keluhanNasabah += Number(r.KeluhanNasabah || 0);
  });
  return sum;
}

/**
 * data = { buybackGram, buybackNilai, keluhanNasabah }
 * Upsert by Periode ('YYYY-MM') — satu baris per bulan. Kolom
 * BiayaDistribusi & StokAkhirGram di sheet TIDAK disentuh lagi
 * (biarkan histori lama apa adanya) — keduanya sudah AUTO (SPK
 * Pengiriman & Gold in Transit gabungan), bukan dari form manual ini
 * lagi. updateRowById hanya menulis key yang ADA di payload, jadi
 * kolom itu otomatis tidak ke-overwrite.
 */
function saveKinerjaBulananManual(periode, data) {
  requireRole('Supervisor');
  const sh = getOrCreateKinerjaBulananManualSheet_();
  const user = getCurrentUser();
  const existing = sheetToObjects(SHEET_NAMES.KINERJA_BULANAN_MANUAL).find(r => r.Periode === periode);
  const payload = {
    BuybackGram: Number(data.buybackGram || 0),
    BuybackNilai: Number(data.buybackNilai || 0),
    KeluhanNasabah: Number(data.keluhanNasabah || 0),
    UpdatedAt: new Date(),
    UpdatedBy: user.Email
  };
  if (existing) {
    updateRowById(SHEET_NAMES.KINERJA_BULANAN_MANUAL, 'Periode', periode, payload);
  } else {
    sh.appendRow([periode, payload.BuybackGram, payload.BuybackNilai, '', '', payload.KeluhanNasabah, payload.UpdatedAt, payload.UpdatedBy]);
  }
  logAudit('UPSERT', 'KinerjaBulananManual', periode, JSON.stringify(data));
  return getKinerjaBulananManual_(periode);
}

// Jumlah buyer aktif: distinct Pembeli dari PenjualanTransaksi Trading
// (exclude baris EGR-lama yang ke-sync, SAMA filter dengan
// getDashboardKategoriBreakdown) UNION distinct Pembeli dari
// Input_Jual_EGR pada rentang tanggal yang diberikan. Dipakai untuk
// kolom bulan maupun kolom YTD (rentang tanggalnya beda, logic sama).
function kbDistinctBuyerCount_(startDate, endDate, startDateStr, endDateStr) {
  const isEgrBuyer_ = v => String(v || '').trim().toUpperCase() === 'EGR';
  const buyerSet_ = {};
  sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(p => Number(p.TotalHarga || 0) > 0 && !isEgrBuyer_(p.Pembeli) && isDateInPeriod(p.Tanggal, startDate, endDate))
    .forEach(p => { if (p.Pembeli) buyerSet_[p.Pembeli] = true; });
  readEgrJualRowsForRange_(startDateStr, endDateStr).forEach(r => { if (r.Pembeli) buyerSet_[r.Pembeli] = true; });
  return Object.keys(buyerSet_).length;
}

// Harga rata-rata & spread dari sepasang gram/nilai — dipakai baik utk
// kolom bulan maupun kolom YTD (rumus SAMA, cuma input beda rentang).
function kbHitungAnalisis_(penjualanGram, penjualanNilai, pembelianGram, pembelianNilai) {
  const hargaJualRataRata = penjualanGram > 0 ? penjualanNilai / penjualanGram : null;
  const hargaBeliRataRata = pembelianGram > 0 ? pembelianNilai / pembelianGram : null;
  const spreadRealisasi = (hargaJualRataRata !== null && hargaBeliRataRata !== null && hargaJualRataRata !== 0)
    ? (hargaJualRataRata - hargaBeliRataRata) / hargaJualRataRata : null;
  return {
    hargaJualRataRata: hargaJualRataRata,
    hargaBeliRataRata: hargaBeliRataRata,
    labaKotorPenjualan: null,
    marginLabaKotor: null,
    spreadRealisasi: spreadRealisasi,
    selisihMutasiStok: null,
    perputaranPersediaan: null
  };
}

/**
 * Data lengkap halaman Kinerja & Profitabilitas Bulanan untuk satu
 * periode (year, month 1-12), SEKALIGUS kolom YTD (Januari s/d bulan
 * terpilih, tahun yang sama) — sesuai catatan asli template ('"YTD" =
 * Januari s.d. bulan periode', ditemukan waktu audit). Indikator
 * Penjualan/Pembelian AUTO = GABUNGAN Trading + EGR, reuse
 * getDashboardKategoriBreakdown() (fungsi KPI Dashboard) apa adanya —
 * TIDAK ada logic gabung/dedup baru ditulis di sini. Indikator MANUAL
 * dari KinerjaBulananManual — untuk YTD, indikator ARUS (Buyback,
 * Biaya distribusi, Keluhan nasabah) DIJUMLAH antar bulan Jan..bulan
 * terpilih, sedangkan "Stok akhir" (SALDO/posisi, bukan arus) YTD-nya
 * = nilai bulan terpilih itu sendiri (tidak dijumlah). ANALISIS
 * dihitung dari kedua itu, mengikuti formula template PERSIS (2 baris
 * "pending" dikosongkan sesuai keputusan user, lihat komentar di atas
 * KINERJA_BULANAN_ANALISIS) — untuk kedua kolom, bulan maupun YTD.
 */
function getKinerjaBulananData(year, month) {
  requireRole('Viewer');
  year = Number(year);
  month = Number(month);
  const periode = year + '-' + String(month).padStart(2, '0');
  const monthStart = new Date(year, month - 1, 1);
  const monthEnd = new Date(year, month, 0);
  const monthStartStr = formatDateOnly(monthStart);
  const monthEndStr = formatDateOnly(monthEnd);
  const yearStart = new Date(year, 0, 1);
  const yearStartStr = formatDateOnly(yearStart);

  // --- Kolom bulan ---
  const kategori = getDashboardKategoriBreakdown(monthStartStr, monthEndStr);
  const penjualanGram = kategori.penjualan.total.gram;
  const penjualanNilai = kategori.penjualan.total.nominal;
  const penjualanTrx = kategori.penjualan.total.count;
  const pembelianGram = kategori.pembelian.total.gram;
  const pembelianNilai = kategori.pembelian.total.nominal;
  const buyerAktif = kbDistinctBuyerCount_(monthStart, monthEnd, monthStartStr, monthEndStr);

  // --- Kolom YTD (Januari s/d bulan terpilih, tahun yang sama) ---
  const kategoriYtd = getDashboardKategoriBreakdown(yearStartStr, monthEndStr);
  const penjualanGramYtd = kategoriYtd.penjualan.total.gram;
  const penjualanNilaiYtd = kategoriYtd.penjualan.total.nominal;
  const penjualanTrxYtd = kategoriYtd.penjualan.total.count;
  const pembelianGramYtd = kategoriYtd.pembelian.total.gram;
  const pembelianNilaiYtd = kategoriYtd.pembelian.total.nominal;
  const buyerAktifYtd = kbDistinctBuyerCount_(yearStart, monthEnd, yearStartStr, monthEndStr);

  // Biaya distribusi: AUTO dari SPK Pengiriman — SUM(Estimasi Biaya)
  // ter-filter Tanggal Kirim, sheet "SPK" (PRICE_DISCOVERY_SOURCE_ID).
  const biayaDistribusi = getBiayaDistribusiSpkTotal_(monthStart, monthEnd);
  const biayaDistribusiYtd = getBiayaDistribusiSpkTotal_(yearStart, monthEnd);

  // Stok akhir – gram: AUTO dari Gold in Transit (= Piutang Supplier)
  // GABUNGAN Trading + EGR, reuse getPiutangSupplierDashboardCombined()
  // APA ADANYA — fungsi yang sama dipakai card "Piutang per Supplier"
  // tab Gabungan di Dashboard (dan Menu Piutang Supplier Trading/EGR
  // untuk komponennya masing-masing) — TIDAK ada logic gabung baru
  // ditulis di sini. "Stok akhir" = posisi Gold in Transit PER AKHIR
  // bulan terpilih (bukan dijumlah), dikonfirmasi user: "gold in
  // transit itu piutang ya, samain aja kalo piutang egr dan trading" /
  // "Stok akhir – gram itu piutang, samain aja".
  const stokAkhir = getPiutangSupplierDashboardCombined(monthEndStr).totalGramasi;

  const manual = getKinerjaBulananManual_(periode);
  const manualYtdFlow = getKinerjaBulananManualRange_(year, month);

  const values = {
    penjualanGram: penjualanGram, penjualanNilai: penjualanNilai, penjualanTrx: penjualanTrx, buyerAktif: buyerAktif,
    pembelianGram: pembelianGram, pembelianNilai: pembelianNilai,
    buybackGram: manual.buybackGram, buybackNilai: manual.buybackNilai,
    biayaDistribusi: biayaDistribusi, stokAkhir: stokAkhir, keluhanNasabah: manual.keluhanNasabah
  };
  const valuesYtd = {
    penjualanGram: penjualanGramYtd, penjualanNilai: penjualanNilaiYtd, penjualanTrx: penjualanTrxYtd, buyerAktif: buyerAktifYtd,
    pembelianGram: pembelianGramYtd, pembelianNilai: pembelianNilaiYtd,
    buybackGram: manualYtdFlow.buybackGram, buybackNilai: manualYtdFlow.buybackNilai,
    biayaDistribusi: biayaDistribusiYtd,
    stokAkhir: stokAkhir, // saldo/posisi, BUKAN dijumlah -> = nilai bulan terpilih
    keluhanNasabah: manualYtdFlow.keluhanNasabah
  };

  const analisisValues = kbHitungAnalisis_(penjualanGram, penjualanNilai, pembelianGram, pembelianNilai);
  const analisisValuesYtd = kbHitungAnalisis_(penjualanGramYtd, penjualanNilaiYtd, pembelianGramYtd, pembelianNilaiYtd);

  return {
    tahun: year,
    bulan: month,
    periode: periode,
    periodeLabel: NAMA_BULAN_ID[month - 1] + ' ' + year,
    ytdLabel: 'Jan–' + NAMA_BULAN_ID[month - 1].slice(0, 3) + ' ' + year,
    indikator: KINERJA_BULANAN_INDIKATOR.map(ind => Object.assign({}, ind, { value: values[ind.key], valueYtd: valuesYtd[ind.key] })),
    analisis: KINERJA_BULANAN_ANALISIS.map(a => Object.assign({}, a, { value: analisisValues[a.key], valueYtd: analisisValuesYtd[a.key] })),
    manualUpdatedAt: manual.updatedAt,
    manualUpdatedBy: manual.updatedBy
  };
}

function kbSetCell_(sh, a1, value, opts) {
  const range = sh.getRange(a1);
  if (value !== undefined) range.setValue(value);
  if (opts) {
    if (opts.bold !== undefined) range.setFontWeight(opts.bold ? 'bold' : 'normal');
    if (opts.size) range.setFontSize(opts.size);
    if (opts.color) range.setFontColor(opts.color);
    if (opts.bg) range.setBackground(opts.bg);
    if (opts.numfmt) range.setNumberFormat(opts.numfmt);
    if (opts.align) range.setHorizontalAlignment(opts.align);
  }
  return range;
}

/**
 * Menulis SELURUH tampilan sheet "Kinerja Bulanan" (dibersihkan &
 * dibuat ulang tiap kali — pola SAMA seperti prepareMonthlyReportSheet)
 * berdasarkan format/warna/merge hasil audit template asli. Dipakai
 * HANYA saat export Excel (halaman web tidak butuh sheet ini, cukup
 * getKinerjaBulananData()). Formula ANALISIS ditulis sebagai FORMULA
 * (bukan angka hasil), merujuk sel INPUT di sheet yang sama — TETAP
 * hidup/recalculate kalau dibuka & diedit manual di Excel/Sheets.
 */
function buildKinerjaBulananSheet_(sh, data) {
  sh.clear();
  sh.clearNotes();
  // sh.clear() TIDAK menghapus frozen rows/columns maupun merged cells
  // sisa generate sebelumnya (gotcha Apps Script) — kalau dibiarkan,
  // merge baru di bawah bisa bentrok dgn sisa freeze/merge lama dan
  // melempar "You can't merge frozen and non-frozen columns.". Reset
  // eksplisit dulu supaya sheet ini SELALU mulai dari kondisi bersih
  // tiap generate, apapun struktur kolom hasil generate sebelumnya.
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  const maxRows = sh.getMaxRows(), maxCols = sh.getMaxColumns();
  if (maxRows > 0 && maxCols > 0) sh.getRange(1, 1, maxRows, maxCols).breakApart();

  // Kolom: A=No, B=Indikator, C=<Bulan>, D=YTD, E=Satuan (YTD ditambah
  // sebagai kolom baru sesuai catatan template asli — '"YTD" = Januari
  // s.d. bulan periode' — yang tadinya belum diimplementasi).
  sh.setColumnWidth(1, 36);
  sh.setColumnWidth(2, 380);
  sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 150);

  const GREEN = '#1D5C44';
  const GREEN_SOFT = '#DCE9E2';
  const NOTE_GRAY = '#58645D';
  const YELLOW = '#FFFF00';
  const BLUE = '#0000FF';
  const CREAM = '#FFF2CC';

  kbSetCell_(sh, 'A1', 'Kinerja & Profitabilitas Bulanan', { bold: true, size: 15, color: GREEN });
  kbSetCell_(sh, 'A2', 'Baris kuning = input manual (Buyback, Keluhan nasabah). Baris putih = otomatis: Penjualan/Pembelian gabungan Trading + EGR, Biaya distribusi dari SPK Pengiriman, Stok akhir dari Gold in Transit/Piutang gabungan Trading + EGR — diisi aplikasi Gold Trading Management System, bukan ditulis manual di sini. Kolom YTD = akumulasi Januari s.d. bulan periode (Stok akhir tidak diakumulasi — itu saldo/posisi, bukan arus).', { size: 9, color: NOTE_GRAY });

  const monthAbbrev = NAMA_BULAN_ID[data.bulan - 1].slice(0, 3);
  kbSetCell_(sh, 'C4', new Date(data.tahun, data.bulan - 1, 1), { size: 8, color: '#808080', numfmt: 'mmm-yy' });

  sh.getRange('A5:E5').setValues([['No', 'Indikator', monthAbbrev, 'YTD', 'Satuan']]);
  sh.getRange('A5:E5').setFontWeight('bold').setFontColor('#FFFFFF').setBackground(GREEN).setHorizontalAlignment('center');

  sh.getRange('B6:E6').merge();
  kbSetCell_(sh, 'B6', 'INPUT', { bold: true, bg: GREEN_SOFT });

  let r = 7;
  const inputRowByKey = {};
  data.indikator.forEach(function (ind) {
    inputRowByKey[ind.key] = r;
    sh.getRange(r, 1).setValue(ind.no).setHorizontalAlignment('center');
    sh.getRange(r, 2).setValue(ind.label);
    const cCell = sh.getRange(r, 3);
    if (ind.value !== null && ind.value !== undefined) cCell.setValue(ind.value);
    cCell.setNumberFormat(ind.numfmt).setHorizontalAlignment('center');
    const dCell = sh.getRange(r, 4);
    if (ind.valueYtd !== null && ind.valueYtd !== undefined) dCell.setValue(ind.valueYtd);
    dCell.setNumberFormat(ind.numfmt).setHorizontalAlignment('center');
    if (ind.source === 'manual') {
      cCell.setBackground(YELLOW).setFontColor(BLUE);
      dCell.setBackground(YELLOW).setFontColor(BLUE);
    }
    sh.getRange(r, 5).setValue(ind.satuan).setHorizontalAlignment('center');
    r++;
  });

  const rPenjualanGram = inputRowByKey.penjualanGram, rPenjualanNilai = inputRowByKey.penjualanNilai;
  const rPembelianGram = inputRowByKey.pembelianGram, rPembelianNilai = inputRowByKey.pembelianNilai;

  const analisisHeaderRow = r;
  sh.getRange(analisisHeaderRow, 2, 1, 4).merge();
  kbSetCell_(sh, 'B' + analisisHeaderRow, 'ANALISIS (otomatis)', { bold: true, bg: GREEN_SOFT });
  r++;

  // Formula IFERROR ditulis 2x (kolom C=bulan, D=YTD) — supaya kolom
  // YTD juga LIVE recalculate (bukan angka statis) kalau file dibuka &
  // diedit manual, bukan cuma kolom bulan.
  data.analisis.forEach(function (a) {
    sh.getRange(r, 2).setValue(a.label).setWrap(true);
    const cCell = sh.getRange(r, 3);
    const dCell = sh.getRange(r, 4);
    if (a.pending) {
      sh.getRange(r, 1, 1, 5).setNote(a.pendingNote || 'Data belum tersedia.');
    } else if (a.key === 'hargaJualRataRata') {
      cCell.setFormula('=IFERROR(C' + rPenjualanNilai + '/C' + rPenjualanGram + ',"")');
      dCell.setFormula('=IFERROR(D' + rPenjualanNilai + '/D' + rPenjualanGram + ',"")');
    } else if (a.key === 'hargaBeliRataRata') {
      cCell.setFormula('=IFERROR(C' + rPembelianNilai + '/C' + rPembelianGram + ',"")');
      dCell.setFormula('=IFERROR(D' + rPembelianNilai + '/D' + rPembelianGram + ',"")');
    } else if (a.key === 'spreadRealisasi') {
      cCell.setFormula('=IFERROR((C' + rPenjualanNilai + '/C' + rPenjualanGram + '-C' + rPembelianNilai + '/C' + rPembelianGram + ')/(C' + rPenjualanNilai + '/C' + rPenjualanGram + '),"")');
      dCell.setFormula('=IFERROR((D' + rPenjualanNilai + '/D' + rPenjualanGram + '-D' + rPembelianNilai + '/D' + rPembelianGram + ')/(D' + rPenjualanNilai + '/D' + rPenjualanGram + '),"")');
    }
    cCell.setNumberFormat(a.numfmt).setHorizontalAlignment('center');
    dCell.setNumberFormat(a.numfmt).setHorizontalAlignment('center');
    sh.getRange(r, 5).setValue(a.satuan).setHorizontalAlignment('center');
    r++;
  });

  r++; // baris kosong pemisah, sama seperti template asli
  kbSetCell_(sh, 'B' + r, 'Pengesahan', { bold: true, size: 11, color: GREEN });
  r++;
  sh.getRange(r, 2, 1, 3).setValues([['Peran', 'Nama', 'Tanggal & paraf']]);
  sh.getRange(r, 2, 1, 3).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(GREEN).setHorizontalAlignment('center');
  r++;
  [
    ['Disusun – Kepala Departemen Operasional Perdagangan', 'M. REZA FEBRIAN'],
    ['Diperiksa – Deputy Bisnis Bulion', 'DARMA SATRIA'],
    ['Diketahui – Kepala Divisi Bisnis Bulion', 'KADEK EVA SUPUTRA']
  ].forEach(function (p) {
    sh.getRange(r, 2).setValue(p[0]).setFontWeight('bold');
    sh.getRange(r, 3).setValue(p[1]).setFontColor(BLUE).setBackground(CREAM);
    r++;
  });

  sh.setFrozenRows(5);
}

/**
 * COPY TEMPLATE -> ISI DATA SESUAI PERIODE -> PERTAHANKAN FORMULA ->
 * PERTAHANKAN FORMAT -> DOWNLOAD XLSX. Sheet "Kinerja Bulanan" ditulis
 * ulang di spreadsheet DB ini (sama pola dgn prepareMonthlyReportSheet:
 * dibersihkan & dibuat ulang tiap generate), lalu di-copy ke SATU
 * spreadsheet sementara (supaya file yang di-download HANYA berisi
 * sheet ini — TIDAK ikut membocorkan sheet lain di DB seperti Users/
 * AuditTrail/Outstanding), di-export sbg xlsx via Drive export endpoint
 * (pola SAMA PERSIS dgn generateNodinDueDiligenceDocx/Laporan Bulanan
 * PPTX yang sudah ada), lalu file sementara dihapus.
 */
function exportKinerjaBulananExcel(year, month) {
  requireRole('Viewer');
  const data = getKinerjaBulananData(year, month);
  const db = getDb();
  let sh = db.getSheetByName(SHEET_NAMES.KINERJA_BULANAN_TEMPLATE);
  if (!sh) sh = db.insertSheet(SHEET_NAMES.KINERJA_BULANAN_TEMPLATE);
  buildKinerjaBulananSheet_(sh, data);
  SpreadsheetApp.flush();

  const tempSs = SpreadsheetApp.create('TMP_Kinerja_Bulanan_' + data.periode + '_' + new Date().getTime());
  const tempId = tempSs.getId();
  try {
    const copied = sh.copyTo(tempSs);
    copied.setName('Kinerja Bulanan');
    const defaultSheet = tempSs.getSheetByName('Sheet1');
    if (defaultSheet) tempSs.deleteSheet(defaultSheet);
    SpreadsheetApp.flush();

    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tempId + '/export?format=xlsx';
    const resp = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error('Gagal export ke xlsx (HTTP ' + resp.getResponseCode() + ')');
    }
    const base64 = Utilities.base64Encode(resp.getBlob().getBytes());
    return {
      base64: base64,
      filename: 'Kinerja_Bulanan_' + data.periode + '.xlsx'
    };
  } finally {
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}

// ------------------------------------------------------------
// MONITORING DELIVERY PO (tab kedua menu Pembelian) — pemenuhan PO
// supplier (Order vs Delivery vs Remaining), TERPISAH TOTAL dari
// PembelianTransaksi (yang mencatat transaksi harga/nominal aktual) dan
// dari sheet Rekap (tidak pernah ditulis). Sheet "MonitoringDeliveryPO"
// dibuat lazy/idempotent (bukan lewat setupDatabase manual) krn
// Execution API tidak tersedia di sesi kerja ini — sheet baru otomatis
// terbentuk saat tab Monitoring Delivery PO pertama kali diakses oleh
// user sungguhan. Nomor PO = unique key (upsert, tidak pernah duplikat).
// Remaining/Status/Aging PO SELALU dihitung ulang saat dibaca (tidak
// pernah disimpan statis) supaya tidak pernah basi.
// ------------------------------------------------------------
const MONITORING_PO_CSV_HEADERS = ['No', 'Tanggal PO', 'Nomor PO', 'Nama Supplier', 'Total Order', 'Delivery', 'Remaining', 'Jumlah Hari'];

function ensureMonitoringPoSheet_() {
  const ss = getDb();
  let sh = ss.getSheetByName(SHEET_NAMES.MONITORING_PO);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAMES.MONITORING_PO);
    const headers = ['NomorPO', 'TanggalPO', 'Supplier', 'TotalOrder', 'Delivery', 'CreatedAt', 'UpdatedAt'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    formatHeader(sh, headers.length);
  }
  return sh;
}

// "Jumlah Hari" pada file sumber adalah snapshot umur PO PADA SAAT file
// dibuat — akan basi begitu tanggal berganti. Aplikasi menghitung ULANG
// "Aging PO" setiap kali dibaca (hari ini - Tanggal PO), bukan memakai
// angka statis dari file (lihat catatan user: label diganti "Aging PO").
function monitoringPoAgingBucket_(hari) {
  if (hari <= 7) return '0-7';
  if (hari <= 14) return '8-14';
  if (hari <= 30) return '15-30';
  return '>30';
}

function getMonitoringPoData() {
  requireRole('Viewer');
  ensureMonitoringPoSheet_();
  const rows = sheetToObjects(SHEET_NAMES.MONITORING_PO).filter(r => String(r.NomorPO || '').trim());
  const todayTs = new Date(formatDateOnly(new Date()) + 'T00:00:00').getTime();

  const items = rows.map(r => {
    const totalOrder = Number(r.TotalOrder) || 0;
    const delivery = Number(r.Delivery) || 0;
    const remaining = Math.max(0, totalOrder - delivery);
    const tanggalPO = extractDateOnly(r.TanggalPO);
    const agingHari = tanggalPO ? Math.max(0, Math.round((todayTs - new Date(tanggalPO + 'T00:00:00').getTime()) / 86400000)) : 0;
    let status = 'Belum Delivery';
    if (remaining === 0) status = 'Selesai';
    else if (delivery > 0) status = 'Sebagian';
    return {
      nomorPO: r.NomorPO, tanggalPO: tanggalPO, supplier: r.Supplier || '',
      totalOrder: totalOrder, delivery: delivery, remaining: remaining,
      agingHari: agingHari, agingBucket: monitoringPoAgingBucket_(agingHari), status: status
    };
  });

  const supplierList = Array.from(new Set(items.map(r => r.supplier).filter(Boolean))).sort();
  return { items: items, supplierList: supplierList };
}

function validateMonitoringPoCsvHeader_(headerRow) {
  const norm = function (s) { return String(s || '').trim().toLowerCase(); };
  const got = (headerRow || []).map(norm);
  const want = MONITORING_PO_CSV_HEADERS.map(norm);
  for (let i = 0; i < want.length; i++) {
    if (got[i] !== want[i]) {
      return { valid: false, message: 'Kolom ke-' + (i + 1) + ' harus "' + MONITORING_PO_CSV_HEADERS[i] + '", ditemukan "' + (headerRow[i] || '(kosong)') + '".' };
    }
  }
  return { valid: true };
}

/**
 * rows = array of { tanggalPO, nomorPO, supplier, totalOrder, delivery }
 * (sudah diparse & angka format titik-ribuan sudah dikonversi di CLIENT
 * — lihat pjlParseAngkaTitik_ di JavaScript.html). Dipakai baik utk
 * preview (dryRun=true, TIDAK menulis apa pun) maupun commit
 * (dryRun=false) supaya logic pencocokan Nomor PO SAMA PERSIS antara
 * preview & hasil akhir. Upsert murni: baris lama yg tidak ada di file
 * baru TIDAK PERNAH dihapus.
 */
function processMonitoringPoImport_(rows, dryRun) {
  ensureMonitoringPoSheet_();
  const sh = getSheet(SHEET_NAMES.MONITORING_PO);
  const existing = sheetToObjects(SHEET_NAMES.MONITORING_PO);
  const existingMap = {};
  existing.forEach(function (r) { existingMap[String(r.NomorPO || '').trim()] = r; });

  const baru = [], diperbarui = [], dilewati = [], gagal = [];
  const seen = {};
  const now = new Date();

  (rows || []).forEach(function (row, idx) {
    const baris = idx + 2; // +1 header, +1 1-based
    const nomorPO = String(row.nomorPO || '').trim();
    if (!nomorPO) { gagal.push({ baris: baris, alasan: 'Nomor PO kosong' }); return; }
    if (seen[nomorPO]) { gagal.push({ baris: baris, nomorPO: nomorPO, alasan: 'Nomor PO duplikat di dalam file yang sama' }); return; }
    seen[nomorPO] = true;

    const tanggalPO = extractDateOnly(row.tanggalPO);
    if (!tanggalPO) { gagal.push({ baris: baris, nomorPO: nomorPO, alasan: 'Tanggal PO tidak valid' }); return; }
    const supplier = String(row.supplier || '').trim();
    if (!supplier) { gagal.push({ baris: baris, nomorPO: nomorPO, alasan: 'Nama Supplier kosong' }); return; }
    const totalOrder = Number(row.totalOrder) || 0;
    const delivery = Number(row.delivery) || 0;
    if (totalOrder <= 0) { gagal.push({ baris: baris, nomorPO: nomorPO, alasan: 'Total Order harus lebih dari 0' }); return; }
    if (delivery < 0) { gagal.push({ baris: baris, nomorPO: nomorPO, alasan: 'Delivery tidak boleh negatif' }); return; }

    const ex = existingMap[nomorPO];
    if (ex) {
      const sama = extractDateOnly(ex.TanggalPO) === tanggalPO && String(ex.Supplier || '').trim() === supplier &&
        Number(ex.TotalOrder) === totalOrder && Number(ex.Delivery) === delivery;
      if (sama) {
        dilewati.push({ nomorPO: nomorPO, alasan: 'Data sama, tidak ada perubahan' });
      } else {
        diperbarui.push({ nomorPO: nomorPO, tanggalPO: tanggalPO, supplier: supplier, totalOrder: totalOrder, delivery: delivery });
        if (!dryRun) {
          updateRowById(SHEET_NAMES.MONITORING_PO, 'NomorPO', nomorPO, {
            TanggalPO: tanggalPO, Supplier: supplier, TotalOrder: totalOrder, Delivery: delivery, UpdatedAt: now
          });
        }
      }
    } else {
      baru.push({ nomorPO: nomorPO, tanggalPO: tanggalPO, supplier: supplier, totalOrder: totalOrder, delivery: delivery });
      if (!dryRun) {
        sh.appendRow([nomorPO, tanggalPO, supplier, totalOrder, delivery, now, now]);
      }
    }
  });

  return { baru: baru, diperbarui: diperbarui, dilewati: dilewati, gagal: gagal };
}

function previewMonitoringPoImport(headerRow, rows) {
  requireRole('Viewer');
  const headerCheck = validateMonitoringPoCsvHeader_(headerRow);
  if (!headerCheck.valid) throw new Error(headerCheck.message);
  return processMonitoringPoImport_(rows, true);
}

function commitMonitoringPoImport(headerRow, rows) {
  requireRole('Viewer');
  const headerCheck = validateMonitoringPoCsvHeader_(headerRow);
  if (!headerCheck.valid) throw new Error(headerCheck.message);
  const result = processMonitoringPoImport_(rows, false);
  logAudit('IMPORT', SHEET_NAMES.MONITORING_PO, 'BULK',
    'Import CSV Monitoring Delivery PO — baru:' + result.baru.length + ' diperbarui:' + result.diperbarui.length +
    ' dilewati:' + result.dilewati.length + ' gagal:' + result.gagal.length);
  return result;
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

// ------------------------------------------------------------
// NAMA SHEET EGR — TERPUSAT. Sheet EGR di spreadsheet sumber
// (PRICE_DISCOVERY_SOURCE_ID) di-rename: Input_Jual_EGR -> Input_Subs_EGR
// (penjualan/subscription) dan Input_Beli_EGR -> Input_Restocking_EGR
// (pembelian/restocking). Semua fungsi EGR membuka sheet lewat
// getEgrSheet_ supaya nama cukup diatur di sini. Nama lama tetap dicoba
// sebagai fallback (kompatibilitas kalau nama dikembalikan).
// ------------------------------------------------------------
const EGR_SHEET = {
  JUAL: 'Input_Subs_EGR',
  BELI: 'Input_Restocking_EGR',
  // Redemption/buyback EGR dari nasabah (emas kembali ke inventory).
  REDEM: 'Input_Redem_EGR'
};
const EGR_SHEET_FALLBACK_ = {
  'Input_Subs_EGR': ['Input_Jual_EGR'],
  'Input_Restocking_EGR': ['Input_Beli_EGR']
};

/**
 * Buka sheet EGR berdasarkan nama utama (EGR_SHEET.JUAL / EGR_SHEET.BELI),
 * lalu coba nama lama kalau nama utama tidak ada. Mengembalikan null kalau
 * semua kandidat tidak ditemukan (perilaku sama seperti getSheetByName).
 */
function getEgrSheet_(ss, sheetName) {
  const candidates = [sheetName].concat(EGR_SHEET_FALLBACK_[sheetName] || []);
  for (let i = 0; i < candidates.length; i++) {
    const sh = ss.getSheetByName(candidates[i]);
    if (sh) return sh;
  }
  return null;
}
const PRICE_DISCOVERY_SHEET_NAME = 'Price V3';

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

  // OPTIMASI LOADING: baca SATU KALI blok baris header/goldInTransit/uangMuka
  // x seluruh kolom supplier via getValues(), lalu index ke array di memory
  // — sebelumnya 24 pemanggilan getRange(...).getValue() terpisah (8
  // supplier x 3 sel) per eksekusi fungsi ini. Sel & nilai yang dibaca
  // PERSIS SAMA, hanya jumlah round-trip ke Sheets service yang berkurang.
  const piutangRowStart = REKAP_PIUTANG_ROW.header;
  const piutangRowEnd = REKAP_PIUTANG_ROW.uangMuka;
  const piutangColStart = Math.min.apply(null, REKAP_PIUTANG_SUPPLIER_COLS.map(function (c) { return Math.min(c.labelCol, c.valueCol); }));
  const piutangColEnd = Math.max.apply(null, REKAP_PIUTANG_SUPPLIER_COLS.map(function (c) { return Math.max(c.labelCol, c.valueCol); }));
  const piutangBlock = sheet.getRange(piutangRowStart, piutangColStart, piutangRowEnd - piutangRowStart + 1, piutangColEnd - piutangColStart + 1).getValues();
  function piutangCell_(row, col) { return piutangBlock[row - piutangRowStart][col - piutangColStart]; }

  const rows = REKAP_PIUTANG_SUPPLIER_COLS.map(function (colSet) {
    const namaSupplier = piutangCell_(REKAP_PIUTANG_ROW.header, colSet.labelCol);
    const gramasi = Number(piutangCell_(REKAP_PIUTANG_ROW.goldInTransit, colSet.valueCol)) || 0;
    const valuasi = Number(piutangCell_(REKAP_PIUTANG_ROW.uangMuka, colSet.valueCol)) || 0;
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

// Indeks kolom (0-based) di sheet Price Discovery V3 — diverifikasi
// ULANG langsung dari baris header sheet (bukan tebak-tebak), karena
// mapping lama di atas sudah TIDAK SINKRON: sheet sekarang punya 2
// kolom baru yang disisipkan setelah mapping lama dibuat — "AMMAN"
// (setelah IGS) dan "WARIS" (setelah LOTUS) — yang menggeser semua
// kolom di sebelah kanannya. Akibatnya kolom Z yang dulu "Keputusan
// Harga" sekarang berisi kolom LAIN bernama "Average" (kolom bantu
// internal, bukan harga keputusan), sementara "Keputusan Harga" yang
// asli sudah pindah ke kolom AA. Ini menyebabkan chart "Pergerakan
// Harga Supplier" (menu Supplier Price Trend Analysis) salah membaca
// nilai pada tanggal yang punya sesi Harga Ke- kedua tanpa keputusan
// resmi (kolom Average tetap menghasilkan angka dari placeholder,
// padahal kolom Keputusan Harga kosong) — contoh kasus: 11 & 17
// September 2026.
// Header aktual (hasil verifikasi, 0-based): 0=(kosong) 1=Tanggal
// 2=Harga Ke- 3=Pukul 4=HRTA 5=HRTA+disc 6=IGS 7=AMMAN 8=SJL 9=IDN
// 10=MKB/STRGLD 11=LOTUS 12=WARIS 13=APEPI 14=Spot Goldprice
// 15=XAUIDRK 16=LBMA H-1 17=Kitco H-1 18=Kurs Jual 19=Kurs Beli
// 20=Rp LBMA 21=Rp Kitco 22=Average Price DN 23=Average Price Spot
// 24=Average Price Fixing 25=Average 26=Keputusan Harga 27=Buyback
// 28=Deposito 29=Sentra Buyback 30=Buyback Harta 31=Selisih Spot
// 32=Catatan.
// Kalau sheet sumber berubah struktur lagi di masa depan (kolom
// disisipkan/dihapus), index di bawah ini perlu diverifikasi ulang
// dari baris header — BUKAN diasumsikan tetap sama.
const PD_COL = {
  tanggal: 1, hargaKe: 2, pukul: 3,
  hrta: 4, hrtaDisc: 5, igs: 6, sjl: 8, idn: 9, mkb: 10, lotus: 11, apepi: 13,
  spotGoldprice: 14, xauidrk: 15, lbma: 16, kitco: 17,
  kursJual: 18, kursBeli: 19, rpLbma: 20, rpKitco: 21,
  keputusanHarga: 26, buyback: 27, sentraBuyback: 29
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
  const targetDateStr = formatDateOnly(targetDate);

  // --- Sumber Data: LIVE dari Price Discovery V3 ---
  const sumberData = getSumberDataLive(dateStr);
  const cheapestSupplier = getCheapestFromSumberData(sumberData);

  // --- Piutang: HARUS sama persis dengan Menu Piutang Supplier Trading
  // (getGoldInTransitData) & card "Piutang per Supplier" Dashboard Utama
  // tab Trading, jadi reuse getPiutangSupplierDashboard(dateStr) — SAMA
  // fungsi yang sudah dipakai Dashboard, sumbernya Gold in Transit per PO
  // (readF3BarangMasuk_/goldInTransitPerPO_), posisi PER TANGGAL LAPORAN
  // (dateStr) — BUKAN lagi getPiutangLive() (snapshot statis sheet
  // "Rekap" tanpa filter tanggal) yang sebelumnya dipakai di sini dan
  // menyebabkan angka Piutang Laporan Harian berbeda dari Dashboard
  // Utama / Menu Piutang Supplier. Bentuk hasilnya (rows[].supplierNama/
  // gramasi/valuasi/gramasiPercent/valuasiPercent, totalGramasi/
  // totalValuasi) identik dengan sebelumnya, jadi frontend tidak berubah.
  // TIDAK menyentuh logic Menu Piutang Supplier / Dashboard itu sendiri.
  const piutang = getPiutangSupplierDashboard(dateStr);
  // DEBUG SEMENTARA — hapus setelah verifikasi selesai. Buka Extensions >
  // Apps Script > Executions untuk memastikan baris ini benar-benar
  // tereksekusi dan sumbernya getPiutangSupplierDashboard(), bukan
  // getPiutangLive().
  Logger.log('[DEBUG getLaporanHarianData] piutang source = getPiutangSupplierDashboard(dateStr=' + dateStr + ') | posisi=' + piutang.posisi +
    ' | rows=' + JSON.stringify((piutang.rows || []).map(r => ({ supplier: r.supplierNama, gramasi: r.gramasi, valuasi: r.valuasi }))) +
    ' | totalGramasi=' + piutang.totalGramasi + ' | totalValuasi=' + piutang.totalValuasi);

  // --- Penetapan Harga: LIVE dari Price Discovery V3 (Keputusan Harga & Buyback).
  // Tidak lagi menghitung/menampilkan Margin Rata-rata atau Profit Margin
  // (Trading/ETF) di Laporan Harian, sesuai revisi terbaru.
  const keputusan = getKeputusanHargaLive(dateStr);

  // --- Penjualan & Pembelian Trading pada tanggal ini ---
  const penjualanTradingRows = getPenjualanTransaksiToday(dateStr).map(r => { r.Kategori = 'Trading'; return r; });
  const pembelianTradingRows = getPembelianTransaksiToday(dateStr).map(r => { r.Kategori = 'Trading'; return r; });
  const penjualanTradingSummary = calculatePenjualanSummaryToday(dateStr);
  const pembelianTradingSummary = calculatePembelianSummaryToday(dateStr);

  // --- Penjualan & Pembelian ETF pada tanggal ini (read-only dari
  // Input_Jual_EGR/Input_Beli_EGR) — baris detail + total ---
  const penjualanEtfRows = readEgrJualRowsForDate_(targetDateStr);
  const pembelianEtfRows = readEgrBeliRowsForDate_(targetDateStr);
  const etfPenjualanTotal = readEgrSheetTotal_(EGR_SHEET.JUAL, targetDateStr);
  const etfPembelianTotal = readEgrSheetTotal_(EGR_SHEET.BELI, targetDateStr);

  // Angka0_ memastikan tiap komponen (nominal/gram/count) SELALU Number
  // dan tidak pernah kosong/NaN/null — kalau tidak ada transaksi, hasilnya 0.
  function angka0_(v) {
    const n = Number(v);
    return isNaN(n) ? 0 : n;
  }
  // catNormal_ menyeragamkan bentuk objek kategori jadi {nominal, gram, count}
  // dari SUMBER MANA PUN (fungsi ringkasan Trading Penjualan/Pembelian punya
  // nama field count yang beda: jumlahInvoice vs jumlahTransaksi — makanya
  // field count-nya harus disebut eksplisit di sini, bukan ditebak/disamakan).
  function catNormal_(nominal, gram, count) {
    return { nominal: angka0_(nominal), gram: angka0_(gram), count: angka0_(count) };
  }
  function gabung(trading, etf) {
    return catNormal_(trading.nominal + etf.nominal, trading.gram + etf.gram, trading.count + etf.count);
  }

  const penjualanTradingCat = catNormal_(penjualanTradingSummary.totalNominal, penjualanTradingSummary.totalGram, penjualanTradingSummary.jumlahInvoice);
  const pembelianTradingCat = catNormal_(pembelianTradingSummary.totalNominal, pembelianTradingSummary.totalGram, pembelianTradingSummary.jumlahTransaksi);
  const penjualanEtfCat = catNormal_(etfPenjualanTotal.nominal, etfPenjualanTotal.gram, etfPenjualanTotal.count);
  const pembelianEtfCat = catNormal_(etfPembelianTotal.nominal, etfPembelianTotal.gram, etfPembelianTotal.count);

  return {
    tanggalLabel: tanggalLabel,
    sumberData: sumberData,
    cheapestSupplier: cheapestSupplier,
    piutang: piutang,
    penetapanHarga: {
      saleVals: keputusan.saleVals,
      buybackVals: keputusan.buybackVals
    },
    penjualan: {
      rows: penjualanTradingRows.concat(penjualanEtfRows),
      trading: penjualanTradingCat,
      etf: penjualanEtfCat,
      total: gabung(penjualanTradingCat, penjualanEtfCat)
    },
    pembelian: {
      rows: pembelianTradingRows.concat(pembelianEtfRows),
      trading: pembelianTradingCat,
      etf: pembelianEtfCat,
      total: gabung(pembelianTradingCat, pembelianEtfCat)
    }
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
  lines.push('Penjualan (Trading + ETF): ' + (data.penjualan.total.gram / 1000).toLocaleString('id-ID', { maximumFractionDigits: 2 }) +
    ' kg senilai ' + formatRupiahServer(data.penjualan.total.nominal) + ' (' + data.penjualan.total.count + ' transaksi).');
  lines.push('- Trading: ' + formatRupiahServer(data.penjualan.trading.nominal) + ' (' + (data.penjualan.trading.gram / 1000).toLocaleString('id-ID', { maximumFractionDigits: 2 }) + ' kg, ' + data.penjualan.trading.count + ' trx)');
  lines.push('- ETF: ' + formatRupiahServer(data.penjualan.etf.nominal) + ' (' + (data.penjualan.etf.gram / 1000).toLocaleString('id-ID', { maximumFractionDigits: 2 }) + ' kg, ' + data.penjualan.etf.count + ' trx)');
  lines.push('');
  lines.push('Pembelian (Trading + ETF): ' + (data.pembelian.total.gram / 1000).toLocaleString('id-ID', { maximumFractionDigits: 2 }) +
    ' kg senilai ' + formatRupiahServer(data.pembelian.total.nominal) + ' (' + data.pembelian.total.count + ' transaksi).');
  lines.push('- Trading: ' + formatRupiahServer(data.pembelian.trading.nominal) + ' (' + (data.pembelian.trading.gram / 1000).toLocaleString('id-ID', { maximumFractionDigits: 2 }) + ' kg, ' + data.pembelian.trading.count + ' trx)');
  lines.push('- ETF: ' + formatRupiahServer(data.pembelian.etf.nominal) + ' (' + (data.pembelian.etf.gram / 1000).toLocaleString('id-ID', { maximumFractionDigits: 2 }) + ' kg, ' + data.pembelian.etf.count + ' trx)');
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

// Sentra Buyback — data historis BENCHMARK pasar (bukan supplier),
// sengaja dipisah dari SUPPLIER_PRICE_SOURCES supaya TIDAK PERNAH ikut
// terhitung di jumlahSupplier/termurah/paling stabil/paling fluktuatif/
// tren naik-turun/market spread pada getSupplierPriceKpiAndInsights
// (semua itu hanya mengiterasi supplierList, tidak pernah benchmarkList).
const SUPPLIER_PRICE_BENCHMARK_SOURCES = [
  { key: 'Sentra Buyback', col: PD_COL.sentraBuyback }
];

/**
 * Baca SEKALI seluruh Price Discovery V3 (bukan per-tanggal) lalu
 * susun jadi time-series per supplier untuk rentang tanggal
 * tertentu — jauh lebih cepat daripada memanggil getSumberDataLive
 * berkali-kali dalam loop.
 * benchmarkSeries/benchmarkList (Sentra Buyback) memakai MEKANISME
 * IDENTIK dengan series/supplierList — sumber sheet sama, filter
 * tanggal sama, aturan "harga terakhir/Harga Ke- tertinggi per hari"
 * sama, dan v>0 sebagai validasi (0/kosong/teks/'-' otomatis diabaikan
 * karena Number(...) hasilnya bukan >0) — hanya arraynya dipisah biar
 * tidak pernah ketercampur ke perhitungan supplier di atas.
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
    SUPPLIER_PRICE_BENCHMARK_SOURCES.forEach(s => {
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
  const benchmarkSeries = {};
  SUPPLIER_PRICE_BENCHMARK_SOURCES.forEach(s => {
    benchmarkSeries[s.key] = dates.map(d => (byDate[d][s.key] ? byDate[d][s.key].value : null));
  });

  return {
    dates: dates, series: series, supplierList: SUPPLIER_PRICE_SOURCES.map(s => s.key),
    benchmarkSeries: benchmarkSeries, benchmarkList: SUPPLIER_PRICE_BENCHMARK_SOURCES.map(s => s.key)
  };
}

/**
 * Bungkus time-series harian jadi Mingguan/Bulanan (ambil nilai
 * terakhir yang tersedia dalam tiap bucket). benchmarkSeries (Sentra
 * Buyback) dibungkus dengan mekanisme bucket YANG SAMA PERSIS dengan
 * series supplier — supaya konsisten mengikuti pilihan frekuensi
 * Harian/Mingguan/Bulanan seperti data historis supplier lainnya.
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
  const benchmarkSeries = {};
  (raw.benchmarkList || []).forEach(key => { benchmarkSeries[key] = buckets.map(k => raw.benchmarkSeries[key][bucketToIdx[k]]); });

  return { dates: buckets, series: series, supplierList: raw.supplierList, benchmarkSeries: benchmarkSeries, benchmarkList: raw.benchmarkList };
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

  // --- Rata-Rata Sentra Buyback (benchmark, BUKAN supplier) — dihitung
  // terpisah dari insights[] di atas, jadi tidak pernah menyentuh
  // jumlahSupplier/termurah/stabil/fluktuatif/tren/market spread. ---
  const benchmarkAverages = {};
  (data.benchmarkList || []).forEach(function (key) {
    const vals = (data.benchmarkSeries[key] || []).filter(function (v) { return v != null; });
    benchmarkAverages[key] = vals.length
      ? { avg: vals.reduce(function (a, b) { return a + b; }, 0) / vals.length, count: vals.length }
      : null;
  });
  const sentraBuybackAvg = benchmarkAverages['Sentra Buyback'] ? benchmarkAverages['Sentra Buyback'].avg : null;

  // --- Harga Rata-Rata per Supplier — SEMUA supplier di supplierList
  // KECUALI "Penetapan Harga" (itu bukan supplier, hanya keputusan
  // harga internal). Beda dari insights[] di atas yang otomatis
  // membuang supplier tanpa data sama sekali — di sini supplier tanpa
  // data TETAP ditampilkan (avg:null) supaya frontend bisa menulis
  // "Tidak ada data", bukan menghilangkannya diam-diam.
  const supplierAverages = data.supplierList
    .filter(function (key) { return key !== 'Penetapan Harga'; })
    .map(function (key) {
      const vals = data.series[key].filter(function (v) { return v != null; });
      const avg = vals.length ? (vals.reduce(function (a, b) { return a + b; }, 0) / vals.length) : null;
      const selisih = (avg != null && sentraBuybackAvg != null) ? (avg - sentraBuybackAvg) : null;
      return { key: key, avg: avg, count: vals.length, selisihVsSentraBuyback: selisih };
    });
  supplierAverages.sort(function (a, b) {
    if (a.avg == null && b.avg == null) return 0;
    if (a.avg == null) return 1;
    if (b.avg == null) return -1;
    return a.avg - b.avg;
  });
  const supplierWithData = supplierAverages.filter(function (r) { return r.avg != null; });
  if (supplierWithData.length) {
    supplierWithData[0].isTermurah = true;
    supplierWithData[supplierWithData.length - 1].isTertinggi = true;
  }

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
    trenTurun: trenTurun.slice(0, 3),
    supplierAverages: supplierAverages,
    benchmarkAverages: benchmarkAverages
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
        extractDateOnly(existing.Tanggal) !== extractDateOnly(tanggal) ||
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
          Tanggal: tanggal, HargaPerGram: hargaPerGram, Qty: qty, TotalHarga: totalHarga, PphPasal22: pph, TotalHargaAfterTax: totalAfterTax,
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

  // Kunci pencocokan baris existing = Nomor PO LENGKAP + Tanggal LENGKAP
  // (bukan Nomor PO saja). Nomor PO ternyata bisa dipakai ulang di sheet
  // sumber untuk transaksi yang benar-benar berbeda pada tanggal lain —
  // kalau cuma dicocokkan dari Nomor PO, baris database lama (dengan
  // Seller/Qty/Harga transaksi SEBELUMNYA) bisa "nyangkut" dianggap sama
  // dengan transaksi baru yang kebetulan pakai nomor PO yang sama, padahal
  // datanya sudah beda. Dengan PO+Tanggal, dua transaksi berbeda pada
  // tanggal berbeda selalu dianggap baris terpisah, bukan baris yang sama.
  function poDateKey_(nomorPO, tanggal) {
    return String(nomorPO || '').trim() + '|' + extractDateOnly(tanggal);
  }

  const existingRows = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL);
  const existingByPoDate = {};
  existingRows.forEach(r => { existingByPoDate[poDateKey_(r.NomorPO, r.Tanggal)] = r; });

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

    const existing = existingByPoDate[poDateKey_(nomorPO, tanggal)];
    if (existing) {
      const bolehAutoUpdate = autoSyncCatatan.indexOf(String(existing.Catatan || '')) !== -1;
      if (!bolehAutoUpdate) return;

      const berubah =
        extractDateOnly(existing.Tanggal) !== extractDateOnly(tanggal) ||
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
          Tanggal: tanggal, Harga: harga, AfterDiskon: afterDiskon, Qty: qty,
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
    existingByPoDate[poDateKey_(nomorPO, tanggal)] = { NomorPO: nomorPO, Tanggal: tanggal }; // cegah baris sama dobel di run yang sama
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
 * Konfigurasi per jenis transaksi (penjualan/pembelian) untuk
 * Confirmation Deal Email — supaya satu set fungsi builder/kirim bisa
 * dipakai untuk keduanya tanpa duplikasi (cuma beda nama field &
 * label). "pihak" = Buyer untuk penjualan, Seller untuk pembelian.
 */
const DEAL_JENIS_CONFIG_ = {
  penjualan: {
    sheetName: 'PenjualanTransaksi', moduleName: 'PenjualanTransaksi',
    noLabel: 'Nomor Invoice', noField: 'NomorInvoice',
    pihakLabel: 'Buyer', pihakField: 'Pembeli',
    hargaField: 'HargaPerGram', nilaiField: 'TotalHargaAfterTax',
    hasMerk: true, showLokasiPengambilan: true,
    judulTransaksi: 'penjualan', subjectPrefix: 'Confirmation Deal — Invoice '
  },
  pembelian: {
    sheetName: 'PembelianTransaksi', moduleName: 'PembelianTransaksi',
    noLabel: 'Nomor PO', noField: 'NomorPO',
    pihakLabel: 'Seller', pihakField: 'Seller',
    hargaField: 'AfterDiskon', nilaiField: 'TotalHargaIncludePajak',
    hasMerk: false, showLokasiPengambilan: false,
    judulTransaksi: 'pembelian', subjectPrefix: 'Confirmation Deal — PO '
  }
};

function getDealJenisConfig_(jenis) {
  const cfg = DEAL_JENIS_CONFIG_[jenis];
  if (!cfg) throw new Error('Jenis transaksi Confirmation Deal tidak dikenali: ' + jenis);
  return cfg;
}

/**
 * Bangun TEKS POLOS Confirmation Deal untuk satu baris transaksi —
 * dipakai sebagai plainBody (fallback klien email yang tidak bisa
 * render HTML). jenis = 'penjualan' atau 'pembelian'.
 */
function buildConfirmationDealBody_(row, jenis) {
  const cfg = getDealJenisConfig_(jenis);
  const tanggalLabel = formatTanggalDealLabel_(row.Tanggal);
  const pihakNama = row[cfg.pihakField] || '-';

  let detail =
    '- ' + cfg.noLabel + ' : ' + (row[cfg.noField] || '-') + '\n' +
    '- Tanggal Deal : ' + tanggalLabel + '\n' +
    '- ' + cfg.pihakLabel + ' : ' + pihakNama + '\n';
  if (cfg.hasMerk) detail += '- Produk (Merk) : ' + (row.Merk || '-') + '\n';
  detail +=
    '- Berat (Gramasi) : ' + Number(row.Qty || 0).toLocaleString('id-ID', { maximumFractionDigits: 2 }) + ' gram\n' +
    '- Harga Deal : ' + formatRupiahServer(row[cfg.hargaField]) + '\n' +
    '- Nilai Transaksi : ' + formatRupiahServer(row[cfg.nilaiField]) + '\n';
  if (cfg.showLokasiPengambilan) detail += '- Lokasi Pengambilan : ' + CONFIRMATION_DEAL_COMPANY.lokasiPengambilanDefault + '\n';

  return 'Yth. Bapak/Ibu ' + pihakNama + ',\n\n' +
    'Terima kasih atas kepercayaan yang telah diberikan kepada PT Pegadaian.\n\n' +
    'Dengan ini kami mengonfirmasi bahwa transaksi ' + cfg.judulTransaksi + ' emas telah disepakati (deal) dengan rincian sebagai berikut:\n\n' +
    'Detail Transaksi\n' + detail + '\n' +
    'Segala perubahan harga setelah dokumen ini diterbitkan hanya dapat dilakukan berdasarkan persetujuan tertulis antara PT Pegadaian dan ' + cfg.pihakLabel + '.\n\n' +
    'Dokumen ini menjadi dasar penerbitan invoice dan pelaksanaan proses transaksi sesuai ketentuan yang berlaku.\n\n' +
    'Apabila terdapat pertanyaan atau memerlukan informasi lebih lanjut, silakan menghubungi PIC PT Pegadaian.\n\n' +
    'Hormat kami,\n' +
    'PT Pegadaian - Divisi Bisnis Bulion\n\n' +
    '---\n' +
    'Email ini dibuat secara otomatis oleh sistem. Mohon tidak membalas email ini.';
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
 * baris yang sama dengan versi teks polos, supaya kedua versi selalu
 * konsisten. jenis = 'penjualan' atau 'pembelian'.
 */
function buildConfirmationDealHtml_(row, jenis) {
  const cfg = getDealJenisConfig_(jenis);
  const c = CONFIRMATION_DEAL_COMPANY;
  const tanggalLabel = formatTanggalDealLabel_(row.Tanggal);
  const gramasiLabel = Number(row.Qty || 0).toLocaleString('id-ID', { maximumFractionDigits: 2 }) + ' gram';
  const tahun = new Date().getFullYear();

  let rows =
    confirmationDealHtmlRow_(cfg.noLabel, row[cfg.noField] || '-') +
    confirmationDealHtmlRow_('Tanggal Deal', tanggalLabel) +
    confirmationDealHtmlRow_(cfg.pihakLabel, row[cfg.pihakField] || '-');
  if (cfg.hasMerk) rows += confirmationDealHtmlRow_('Produk (Merk)', row.Merk || '-');
  rows +=
    confirmationDealHtmlRow_('Berat (Gramasi)', gramasiLabel) +
    confirmationDealHtmlRow_('Harga Deal', formatRupiahServer(row[cfg.hargaField])) +
    confirmationDealHtmlRow_('Nilai Transaksi', '<b>' + formatRupiahServer(row[cfg.nilaiField]) + '</b>');
  if (cfg.showLokasiPengambilan) rows += confirmationDealHtmlRow_('Lokasi Pengambilan', c.lokasiPengambilanDefault);

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

      '<p style="font-size:13.5px;line-height:1.6;color:#333;">Dokumen ini merupakan konfirmasi resmi atas kesepakatan harga transaksi jual beli emas antara PT Pegadaian dan ' + cfg.pihakLabel + '. Seluruh informasi di bawah ini telah disepakati oleh kedua belah pihak dan menjadi dasar pelaksanaan transaksi.</p>' +

      '<div style="font-size:14.5px;font-weight:800;color:#046A38;margin:22px 0 8px;">DETAIL TRANSAKSI</div>' +
      '<table width="100%" style="border-collapse:collapse;border:1px solid #E7EAEE;">' + rows + '</table>' +

      '<div style="font-size:14.5px;font-weight:800;color:#046A38;margin:22px 0 8px;">KONFIRMASI</div>' +
      '<p style="font-size:13.5px;line-height:1.6;color:#333;">Dengan ini kami menyatakan bahwa rincian harga transaksi sebagaimana tercantum pada dokumen ini telah diperiksa dan disepakati oleh kedua belah pihak.</p>' +
      '<p style="font-size:13.5px;line-height:1.6;color:#333;">Segala perubahan harga setelah dokumen ini diterbitkan hanya dapat dilakukan berdasarkan persetujuan tertulis antara PT Pegadaian dan ' + cfg.pihakLabel + '.</p>' +
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
 * mengembalikan nomor referensi (invoice/PO), nama pihak (buyer/
 * seller), dan email default (hasil lookup Master_Customer, HANYA
 * untuk jenis='penjualan' — belum ada sumber email otomatis untuk
 * seller/supplier) supaya field "To" bisa terisi otomatis tapi tetap
 * bisa diedit/ditambah oleh user sebelum kirim. jenis = 'penjualan'
 * atau 'pembelian'.
 */
function getConfirmationDealPrefill(jenis, transaksiId) {
  requireRole('Viewer');
  const cfg = getDealJenisConfig_(jenis);
  const rows = sheetToObjects(cfg.sheetName);
  const row = rows.find(r => r.ID === transaksiId);
  if (!row) throw new Error('Transaksi ' + cfg.moduleName + ' dengan ID "' + transaksiId + '" tidak ditemukan.');
  const pihak = row[cfg.pihakField] || '-';
  return {
    noRef: row[cfg.noField] || transaksiId,
    pihak: pihak,
    emailDefault: jenis === 'penjualan' ? (getBuyerEmailByName_(pihak) || '') : ''
  };
}

/**
 * Dipicu dari modal "Kirim Confirmation Deal", untuk transaksi
 * Penjualan maupun Pembelian (jenis = 'penjualan'/'pembelian').
 * toOverride & ccOverride boleh berisi lebih dari satu alamat email
 * dipisah koma (mis. "a@x.com, b@x.com") — GmailApp.sendEmail
 * menerima format ini langsung tanpa perlu diparse manual. Kalau
 * toOverride kosong DAN jenis='penjualan', email diambil otomatis
 * dari Master_Customer (kolom Email PIC) berdasarkan nama Pembeli;
 * untuk jenis='pembelian' toOverride wajib diisi manual (belum ada
 * sumber Email PIC Seller/Supplier).
 */
function sendConfirmationDealEmail(jenis, transaksiId, toOverride, ccOverride) {
  requireRole('Viewer');
  const cfg = getDealJenisConfig_(jenis);
  const rows = sheetToObjects(cfg.sheetName);
  const row = rows.find(r => r.ID === transaksiId);
  if (!row) throw new Error('Transaksi ' + cfg.moduleName + ' dengan ID "' + transaksiId + '" tidak ditemukan.');
  const pihakNama = row[cfg.pihakField] || '-';

  const manualTo = String(toOverride || '').trim();
  let to = manualTo;
  if (!to && jenis === 'penjualan') {
    to = getBuyerEmailByName_(pihakNama);
  }
  if (!to) {
    throw new Error(
      jenis === 'penjualan'
        ? ('Email PIC untuk buyer "' + pihakNama + '" tidak ditemukan di Master_Customer. Lengkapi dulu kolom Email PIC di Master_Customer, atau isi email tujuan manual di form kirim.')
        : ('Email tujuan untuk seller "' + pihakNama + '" belum diisi. Isi email tujuan manual di form kirim (belum ada sumber otomatis untuk Seller/Supplier).')
    );
  }
  const cc = String(ccOverride || '').trim();

  const subject = cfg.subjectPrefix + (row[cfg.noField] || transaksiId);
  const plainBody = buildConfirmationDealBody_(row, jenis);
  const htmlBody = buildConfirmationDealHtml_(row, jenis);
  const logoBlob = Utilities.newBlob(Utilities.base64Decode(CONFIRMATION_DEAL_LOGO_BASE64), 'image/jpeg', 'logo_pegadaian.jpg');

  const sendOptions = {
    htmlBody: htmlBody,
    inlineImages: { logo_pegadaian: logoBlob },
    from: CONFIRMATION_DEAL_SENDER_EMAIL,
    name: CONFIRMATION_DEAL_SENDER_NAME
  };
  if (cc) sendOptions.cc = cc;

  GmailApp.sendEmail(to, subject, plainBody, sendOptions);
  logAudit('SEND_EMAIL', cfg.moduleName, transaksiId, 'Confirmation Deal dikirim ke ' + to + (cc ? ' (CC: ' + cc + ')' : '') + ' dari ' + CONFIRMATION_DEAL_SENDER_EMAIL);
  return { sent: true, to: to, cc: cc };
}

// ------------------------------------------------------------
// KALENDER HUT & ULANG TAHUN
// ------------------------------------------------------------
// Modul baru & berdiri sendiri: sheet Master_HUT + HUT_EmailLog dibuat
// OTOMATIS di database aplikasi (getDb()) kalau belum ada, lewat
// getOrCreate...Sheet_() di bawah — TIDAK PERNAH menghapus/menimpa
// sheet lain yang sudah ada (Master_Customer, PenjualanTransaksi,
// PembelianTransaksi, dst sama sekali tidak disentuh oleh modul ini).
const HUT_KATEGORI_PERUSAHAAN = 'Perusahaan/Institusi';
const HUT_KATEGORI_INDIVIDU = 'Individu';
const HUT_JENIS_SUPPLIER = 'Supplier';
const HUT_JENIS_BUYER = 'Buyer';
const HUT_MASTER_HEADERS = ['ID', 'Nomor', 'Nama', 'Kategori', 'Jenis Perusahaan', 'Tanggal Berdiri/Lahir', 'Email PIC', 'Status Aktif', 'Keterangan'];
// Data awal sesuai permintaan — hanya dipakai untuk mengisi sheet SAAT
// PERTAMA KALI dibuat; tidak pernah menimpa data yang sudah ada/diedit.
// Jenis Perusahaan sengaja dikosongkan (bukan sekadar tebakan) — bisa
// diisi belakangan lewat fitur Edit.
const HUT_MASTER_SEED = [
  ['HUT-001', 1, 'Universitas Muhammadiyah Sumatera Utara', HUT_KATEGORI_PERUSAHAAN, '', new Date(1957, 1, 27), '', 'Aktif', ''],
  ['HUT-002', 2, 'PT Hartadinata Abadi Tbk', HUT_KATEGORI_PERUSAHAAN, '', new Date(2004, 2, 29), '', 'Aktif', ''],
  ['HUT-003', 3, 'PT Pegadaian', HUT_KATEGORI_PERUSAHAAN, '', new Date(1901, 3, 1), '', 'Aktif', ''],
  ['HUT-004', 4, 'PT Elang Mulia Abadi Sempurna', HUT_KATEGORI_PERUSAHAAN, '', new Date(2019, 3, 10), '', 'Aktif', ''],
  ['HUT-005', 5, 'PT Maha Karya Baru', HUT_KATEGORI_PERUSAHAAN, '', new Date(2015, 3, 22), '', 'Aktif', ''],
  ['HUT-006', 6, 'PT ANTAM Tbk', HUT_KATEGORI_PERUSAHAAN, '', new Date(1968, 6, 5), '', 'Aktif', ''],
  ['HUT-007', 7, 'Elvi Rofiqotul Hidayah', HUT_KATEGORI_INDIVIDU, '', new Date(1971, 6, 31), '', 'Aktif', ''],
  ['HUT-008', 8, 'PT Pegadaian Galeri 24', HUT_KATEGORI_PERUSAHAAN, '', new Date(2018, 7, 3), '', 'Aktif', ''],
  ['HUT-009', 9, 'PT Triliun Anugrah Nusantara', HUT_KATEGORI_PERUSAHAAN, '', new Date(2024, 8, 2), '', 'Aktif', ''],
  ['HUT-010', 10, 'PT Karya Utama Putra Mandiri', HUT_KATEGORI_PERUSAHAAN, '', new Date(2001, 10, 10), '', 'Aktif', ''],
  ['HUT-011', 11, 'PT Laku Emas Indonesia', HUT_KATEGORI_PERUSAHAAN, '', new Date(2018, 6, 17), '', 'Aktif', '']
];

/**
 * Migrasi ringan utk sheet Master_HUT yg SUDAH ADA dari sebelum kolom
 * ID/Jenis Perusahaan ditambahkan — tambah kolom yg belum ada TANPA
 * menyentuh data lain, dan backfill ID utk baris lama yg belum punya ID
 * (Jenis Perusahaan baris lama sengaja dibiarkan kosong, bukan ditebak).
 */
function ensureMasterHutSchema_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerIdx = {};
  headerRow.forEach(function (h, i) { if (h) headerIdx[String(h).trim()] = i + 1; });

  function ensureCol(name) {
    if (headerIdx[name]) return headerIdx[name];
    const newCol = sh.getLastColumn() + 1;
    sh.getRange(1, newCol).setValue(name).setFontWeight('bold');
    headerIdx[name] = newCol;
    return newCol;
  }

  const hadId = !!headerIdx['ID'];
  const idCol = ensureCol('ID');
  ensureCol('Jenis Perusahaan');

  if (!hadId) {
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const idValues = sh.getRange(2, idCol, lastRow - 1, 1).getValues();
      let nextNum = 1;
      for (let i = 0; i < idValues.length; i++) {
        if (!idValues[i][0]) {
          idValues[i][0] = 'HUT-' + String(nextNum).padStart(3, '0');
          nextNum++;
        }
      }
      sh.getRange(2, idCol, idValues.length, 1).setValues(idValues);
    }
  }
}

function getOrCreateMasterHutSheet_() {
  const db = getDb();
  let sh = db.getSheetByName(SHEET_NAMES.MASTER_HUT);
  if (sh) {
    ensureMasterHutSchema_(sh);
    return sh;
  }
  sh = db.insertSheet(SHEET_NAMES.MASTER_HUT);
  sh.getRange(1, 1, 1, HUT_MASTER_HEADERS.length).setValues([HUT_MASTER_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, HUT_MASTER_SEED.length, HUT_MASTER_HEADERS.length).setValues(HUT_MASTER_SEED);
  sh.autoResizeColumns(1, HUT_MASTER_HEADERS.length);
  return sh;
}

const HUT_LOG_HEADERS = ['ID', 'Timestamp', 'Nama', 'Kategori', 'TahunPeringatan', 'TipeNotifikasi', 'TanggalHUT', 'Usia', 'EmailTujuan', 'Status', 'Keterangan'];

function getOrCreateHutEmailLogSheet_() {
  const db = getDb();
  let sh = db.getSheetByName(SHEET_NAMES.HUT_EMAIL_LOG);
  if (sh) return sh;
  sh = db.insertSheet(SHEET_NAMES.HUT_EMAIL_LOG);
  sh.getRange(1, 1, 1, HUT_LOG_HEADERS.length).setValues([HUT_LOG_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, HUT_LOG_HEADERS.length);
  return sh;
}

/**
 * Baca semua baris Master_HUT (sheet dibuat+diisi data awal otomatis
 * kalau belum ada) jadi array of objects header-based.
 */
function getMasterHutRows_() {
  getOrCreateMasterHutSheet_();
  return sheetToObjects(SHEET_NAMES.MASTER_HUT);
}

/**
 * Diagnostik read-only — kembalikan header & isi mentah Master_HUT apa
 * adanya di sheet (dipakai sekali utk menelusuri kasus "berhasil
 * ditambahkan tapi tidak muncul di tabel Rundown HUT" — dihapus lagi
 * setelah masalahnya beres).
 */
function debugMasterHutRaw() {
  requireRole('Viewer');
  const sh = getOrCreateMasterHutSheet_();
  const values = sh.getDataRange().getValues();
  return { headers: values[0], rows: values.slice(1) };
}

/**
 * Validasi field form Tambah/Edit Perusahaan (HUT) — dipakai bareng oleh
 * addMasterHut & updateMasterHut supaya aturannya konsisten persis sama.
 * excludeId diisi ID yang sedang diedit (biar tidak dianggap duplikat
 * nama dirinya sendiri).
 */
function validateMasterHutInput_(data, excludeId) {
  const nama = String((data && data.nama) || '').trim();
  if (!nama) throw new Error('Nama perusahaan wajib diisi.');

  const tanggalBerdiri = data && data.tanggalBerdiri;
  if (!tanggalBerdiri) throw new Error('Tanggal berdiri wajib diisi.');
  const tglOnly = extractDateOnly(tanggalBerdiri);
  if (!tglOnly) throw new Error('Tanggal berdiri tidak valid.');
  if (tglOnly > formatDateOnly(new Date())) throw new Error('Tanggal berdiri tidak boleh melebihi tanggal hari ini.');

  const kategori = data && data.kategori;
  if (kategori !== HUT_KATEGORI_PERUSAHAAN && kategori !== HUT_KATEGORI_INDIVIDU) {
    throw new Error('Kategori wajib dipilih.');
  }

  const jenisPerusahaan = data && data.jenisPerusahaan;
  if (jenisPerusahaan !== HUT_JENIS_SUPPLIER && jenisPerusahaan !== HUT_JENIS_BUYER) {
    throw new Error('Jenis Perusahaan wajib dipilih.');
  }

  const namaKey = nama.toLowerCase().replace(/\s+/g, ' ');
  const existing = sheetToObjects(SHEET_NAMES.MASTER_HUT);
  const dup = existing.some(function (r) {
    if (excludeId && String(r['ID']) === String(excludeId)) return false;
    return String(r['Nama'] || '').trim().toLowerCase().replace(/\s+/g, ' ') === namaKey;
  });
  if (dup) throw new Error('Nama perusahaan "' + nama + '" sudah terdaftar.');

  return { nama: nama, tglOnly: tglOnly, kategori: kategori, jenisPerusahaan: jenisPerusahaan };
}

/**
 * Tambah perusahaan/individu baru ke Master_HUT — dipakai form "Tambah
 * Perusahaan" pada halaman Rundown HUT.
 */
function addMasterHut(data) {
  requireRole('Viewer');
  const sh = getOrCreateMasterHutSheet_();
  const v = validateMasterHutInput_(data, null);

  const existingCount = sheetToObjects(SHEET_NAMES.MASTER_HUT).length;
  const id = getNextId(SHEET_NAMES.MASTER_HUT, 'ID', 'HUT');

  // Tulis berdasarkan NAMA header sheet yang sebenarnya (bukan posisi
  // kolom tetap) — sheet lama bisa saja punya urutan kolom berbeda dari
  // HUT_MASTER_HEADERS krn migrasi ensureMasterHutSchema_ menambah kolom
  // di ujung kanan, bukan menyisipkan di posisi "kanonik". Kalau ditulis
  // secara posisional, nilainya bisa masuk ke kolom yang salah — baris
  // "berhasil" tersimpan tapi datanya (mis. Tanggal Berdiri/Lahir) rusak
  // sehingga tidak muncul di tabel.
  const headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rowObj = {
    'ID': id,
    'Nomor': existingCount + 1,
    'Nama': v.nama,
    'Kategori': v.kategori,
    'Jenis Perusahaan': v.jenisPerusahaan,
    'Tanggal Berdiri/Lahir': v.tglOnly,
    'Email PIC': String((data && data.emailPic) || '').trim(),
    'Status Aktif': String((data && data.statusAktif) || 'Aktif').trim() || 'Aktif',
    'Keterangan': String((data && data.keterangan) || '').trim()
  };
  const rowValues = headerRow.map(function (h) { return rowObj[h] !== undefined ? rowObj[h] : ''; });
  sh.appendRow(rowValues);

  logAudit('CREATE', 'MasterHut', id, v.nama);
  return id;
}

/**
 * Perbarui perusahaan/individu Master_HUT berdasarkan ID unik (BUKAN
 * nomor urut/nama) — dipakai form "Edit Perusahaan" pada halaman Rundown
 * HUT. Hanya baris dengan ID ini yang tersentuh; baris lain tidak
 * berubah sama sekali.
 */
function updateMasterHut(id, data) {
  requireRole('Supervisor');
  if (!id) throw new Error('ID perusahaan tidak valid.');
  getOrCreateMasterHutSheet_();
  const v = validateMasterHutInput_(data, id);

  updateRowById(SHEET_NAMES.MASTER_HUT, 'ID', id, {
    'Nama': v.nama,
    'Kategori': v.kategori,
    'Jenis Perusahaan': v.jenisPerusahaan,
    'Tanggal Berdiri/Lahir': v.tglOnly,
    'Email PIC': String((data && data.emailPic) || '').trim(),
    'Status Aktif': String((data && data.statusAktif) || 'Aktif').trim() || 'Aktif',
    'Keterangan': String((data && data.keterangan) || '').trim()
  });
  logAudit('UPDATE', 'MasterHut', id, v.nama);
  return true;
}

/**
 * Hapus satu baris Master_HUT berdasarkan ID unik.
 */
function deleteMasterHut(id) {
  requireRole('Admin');
  if (!id) throw new Error('ID perusahaan tidak valid.');
  getOrCreateMasterHutSheet_();
  deleteRowById(SHEET_NAMES.MASTER_HUT, 'ID', id);
  logAudit('DELETE', 'MasterHut', id, '');
  return true;
}

/**
 * Hitung tanggal HUT/ulang tahun BERIKUTNYA dari suatu tanggal
 * berdiri/lahir, relatif terhadap "today" (default sekarang, zona
 * Asia/Jakarta — appsscript.json timeZone). Kalau tanggal tahun ini
 * sudah lewat, otomatis pakai tahun depan. Usia dihitung otomatis
 * dari selisih tahun asal vs tahun target — TIDAK PERNAH statis/hardcode.
 */
function hitungHutBerikutnya_(tanggalBerdiri, today) {
  const asal = new Date(extractDateOnly(tanggalBerdiri) + 'T00:00:00');
  const now = today || new Date();
  const todayOnly = new Date(extractDateOnly(now) + 'T00:00:00');

  let tahunTarget = todayOnly.getFullYear();
  let berikutnya = new Date(tahunTarget, asal.getMonth(), asal.getDate());
  if (berikutnya.getTime() < todayOnly.getTime()) {
    tahunTarget += 1;
    berikutnya = new Date(tahunTarget, asal.getMonth(), asal.getDate());
  }

  const daysUntil = Math.round((berikutnya.getTime() - todayOnly.getTime()) / 86400000);
  const usia = tahunTarget - asal.getFullYear();

  return { tanggalAsal: asal, tanggalBerikutnya: berikutnya, tahunTarget: tahunTarget, daysUntil: daysUntil, usia: usia };
}

function hutCountdownLabel_(daysUntil) {
  return daysUntil <= 0 ? 'HARI INI' : ('H-' + daysUntil);
}

function hutColorClass_(daysUntil) {
  if (daysUntil <= 1) return 'red';
  if (daysUntil <= 7) return 'yellow';
  return 'blue';
}

/**
 * Widget Agenda HUT & Ulang Tahun di Dashboard — entri Aktif dalam
 * jendela 30 hari ke depan (H-30), diurutkan dari yang paling dekat.
 */
function getHutAgenda() {
  requireRole('Viewer');
  const rows = getMasterHutRows_();
  const now = new Date();
  return rows
    .filter(r => String(r['Status Aktif'] || '').trim().toLowerCase() !== 'nonaktif' && String(r['Nama'] || '').trim())
    .filter(r => !!extractDateOnly(r['Tanggal Berdiri/Lahir'])) // lewati baris dgn tanggal berdiri kosong/tidak valid
    .map(r => {
      const hitung = hitungHutBerikutnya_(r['Tanggal Berdiri/Lahir'], now);
      return {
        nama: r['Nama'],
        kategori: r['Kategori'] || HUT_KATEGORI_PERUSAHAAN,
        tanggalBerikutnya: formatDateOnly(hitung.tanggalBerikutnya),
        tanggalLabel: hitung.tanggalBerikutnya.getDate() + ' ' + NAMA_BULAN_ID[hitung.tanggalBerikutnya.getMonth()] + ' ' + hitung.tahunTarget,
        usia: hitung.usia,
        daysUntil: hitung.daysUntil,
        countdownLabel: hutCountdownLabel_(hitung.daysUntil),
        colorClass: hutColorClass_(hitung.daysUntil)
      };
    })
    .filter(a => a.daysUntil <= 30)
    .sort((a, b) => a.daysUntil - b.daysUntil);
}

/**
 * Rundown HUT satu tahun penuh (semua entri Aktif, tidak dibatasi
 * H-30) untuk halaman detail Rundown HUT — tabel + timeline bulanan.
 */
function getHutRundown(year) {
  requireRole('Viewer');
  const rows = getMasterHutRows_();
  const targetYear = year ? parseInt(year, 10) : new Date().getFullYear();

  const list = rows
    .filter(r => String(r['Status Aktif'] || '').trim().toLowerCase() !== 'nonaktif' && String(r['Nama'] || '').trim())
    .map(r => {
      // Tanggal Berdiri/Lahir kosong atau tidak bisa diparse -> lewati
      // baris ini (jangan diproses jadi NaN/Invalid Date yang bisa
      // merusak sort/tampilan seluruh Rundown HUT).
      const tglIso = extractDateOnly(r['Tanggal Berdiri/Lahir']);
      if (!tglIso) return null;
      const asal = new Date(tglIso + 'T00:00:00');
      if (isNaN(asal.getTime())) return null;

      const tanggalTahunIni = new Date(targetYear, asal.getMonth(), asal.getDate());
      const hitung = hitungHutBerikutnya_(r['Tanggal Berdiri/Lahir']);
      return {
        // ID unik (BUKAN nomor urut) — dipakai form Edit & tombol Hapus
        // supaya selalu menunjuk baris yang benar walau urutan tabel
        // berubah-ubah mengikuti HUT terdekat.
        id: r['ID'],
        nama: r['Nama'],
        kategori: r['Kategori'] || HUT_KATEGORI_PERUSAHAAN,
        jenisPerusahaan: r['Jenis Perusahaan'] || '',
        // Field mentah (bukan hasil format tampilan) — dipakai utk isi
        // otomatis form Edit.
        tanggalBerdiriIso: tglIso,
        emailPic: r['Email PIC'] || '',
        statusAktif: r['Status Aktif'] || 'Aktif',
        keterangan: r['Keterangan'] || '',
        bulan: asal.getMonth(),
        tanggal: tanggalTahunIni.getDate(),
        // Tanggal & tahun ASLI pendirian/kelahiran (mis. PT Pegadaian tetap
        // 1 April 1901) — bukan tahun rundown yang sedang dipilih.
        tanggalLengkap: asal.getDate() + ' ' + NAMA_BULAN_ID[asal.getMonth()] + ' ' + asal.getFullYear(),
        // Usia & hitung mundur selalu mengacu ke HUT BERIKUTNYA dari hari
        // ini (tidak pernah minus; otomatis pindah ke tahun depan kalau
        // tanggal tahun ini sudah lewat) — tidak tergantung tahun rundown.
        usia: hitung.usia,
        countdownLabel: hutCountdownLabel_(hitung.daysUntil),
        colorClass: hutColorClass_(hitung.daysUntil)
      };
    })
    .filter(Boolean)
    .sort((a, b) => (a.bulan - b.bulan) || (a.tanggal - b.tanggal));

  const byMonth = NAMA_BULAN_ID.map((label, idx) => ({ bulan: idx, label: label, items: list.filter(i => i.bulan === idx) }));

  return { year: targetYear, rows: list, byMonth: byMonth };
}

// ---- Settings: daftar email penerima notifikasi HUT ----
const HUT_SETTINGS_KEY_EMAILS = 'HUT_NOTIFICATION_EMAILS';

function getOrCreateSettingsSheet_() {
  const db = getDb();
  let sh = db.getSheetByName(SHEET_NAMES.SETTINGS);
  if (sh) return sh;
  sh = db.insertSheet(SHEET_NAMES.SETTINGS);
  sh.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Description']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

function getSettingValue_(key, defaultValue) {
  const sh = getOrCreateSettingsSheet_();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) return values[i][1];
  }
  return defaultValue !== undefined ? defaultValue : '';
}

function setSettingValue_(key, value, description) {
  const sh = getOrCreateSettingsSheet_();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      if (description) sh.getRange(i + 1, 3).setValue(description);
      return;
    }
  }
  sh.appendRow([key, value, description || '']);
}

/**
 * Daftar email penerima notifikasi HUT (Settings, Key =
 * HUT_NOTIFICATION_EMAILS, dipisah koma). Dipakai untuk mengisi form
 * di halaman Settings maupun oleh pengiriman notifikasi harian.
 */
function getHutNotificationEmails() {
  requireRole('Viewer');
  const raw = String(getSettingValue_(HUT_SETTINGS_KEY_EMAILS, '') || '');
  return raw.split(',').map(s => s.trim()).filter(Boolean);
}

function saveHutNotificationEmails(emailsArray) {
  requireRole('Admin');
  const emails = (emailsArray || []).map(e => String(e || '').trim()).filter(Boolean);
  const invalid = emails.filter(e => !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e));
  if (invalid.length) throw new Error('Format email tidak valid: ' + invalid.join(', '));
  setSettingValue_(HUT_SETTINGS_KEY_EMAILS, emails.join(','), 'Daftar email penerima notifikasi HUT & Ulang Tahun (H-7 & H-1), dipisah koma');
  logAudit('UPDATE', 'HutSettings', '-', 'Email penerima notifikasi HUT diperbarui: ' + emails.join(', '));
  return { emails: emails };
}

// ---- Pengiriman notifikasi email H-7 / H-1 ----

/**
 * Susun & kirim satu email pengingat HUT sesuai format yang diminta.
 */
function kirimHutReminderEmail_(entry, recipients) {
  const jenisPeringatan = entry.kategori === HUT_KATEGORI_INDIVIDU ? 'Ulang Tahun' : 'HUT';
  const subject = 'Pengingat HUT H-' + entry.daysUntil + ' – ' + entry.nama;
  const body = 'Yth. Tim,\n\n' +
    'Kami informasikan bahwa ' + entry.nama + ' akan memperingati ' + jenisPeringatan + ' pada ' + entry.tanggalLabel + ', atau H-' + entry.daysUntil + ' dari hari ini.\n\n' +
    'Pada tanggal tersebut, ' + entry.nama + ' akan berusia ' + entry.usia + ' tahun.\n\n' +
    'Mohon dapat dipersiapkan tindak lanjut dan ucapan yang diperlukan.\n\n' +
    'Terima kasih.';
  GmailApp.sendEmail(recipients.join(','), subject, body);
}

/**
 * Fungsi harian (dipasang lewat trigger installHutNotificationTrigger,
 * jalan otomatis tiap hari ~08:00 WIB) — cek entri Master_HUT yang
 * tepat H-7 atau H-1 dari hari ini, kirim email ke daftar penerima di
 * Settings, dan catat log ke HUT_EmailLog supaya tidak dobel-kirim
 * untuk nama+tahun peringatan+tipe notifikasi (H-7/H-1) yang sama.
 */
function sendHutNotificationsDaily_() {
  const logSheet = getOrCreateHutEmailLogSheet_();
  const alreadySent = {};
  sheetToObjects(SHEET_NAMES.HUT_EMAIL_LOG).forEach(l => {
    if (String(l.Status).trim() === 'Berhasil') {
      alreadySent[String(l.Nama).trim() + '|' + l.TahunPeringatan + '|' + l.TipeNotifikasi] = true;
    }
  });

  const recipients = String(getSettingValue_(HUT_SETTINGS_KEY_EMAILS, '') || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!recipients.length) return { sent: 0, skipped: 0, note: 'Belum ada email penerima di Settings — notifikasi tidak dikirim.' };

  const rows = getMasterHutRows_();
  const now = new Date();
  let sent = 0, skipped = 0;

  rows.forEach(r => {
    const nama = String(r['Nama'] || '').trim();
    if (!nama || String(r['Status Aktif'] || '').trim().toLowerCase() === 'nonaktif') return;

    const hitung = hitungHutBerikutnya_(r['Tanggal Berdiri/Lahir'], now);
    if (hitung.daysUntil !== 7 && hitung.daysUntil !== 1) return;

    const kategori = r['Kategori'] || HUT_KATEGORI_PERUSAHAAN;
    const tipeNotifikasi = 'H-' + hitung.daysUntil;
    const dedupKey = nama + '|' + hitung.tahunTarget + '|' + tipeNotifikasi;
    if (alreadySent[dedupKey]) { skipped++; return; }

    const tanggalLabel = hitung.tanggalBerikutnya.getDate() + ' ' + NAMA_BULAN_ID[hitung.tanggalBerikutnya.getMonth()] + ' ' + hitung.tahunTarget;
    const entry = { nama: nama, kategori: kategori, daysUntil: hitung.daysUntil, usia: hitung.usia, tanggalLabel: tanggalLabel };

    let status = 'Berhasil', keterangan = '';
    try {
      kirimHutReminderEmail_(entry, recipients);
      sent++;
    } catch (err) {
      status = 'Gagal';
      keterangan = String(err.message || err);
    }

    logSheet.appendRow([
      getNextId(SHEET_NAMES.HUT_EMAIL_LOG, 'ID', 'HUTLOG'),
      new Date(), nama, kategori, hitung.tahunTarget, tipeNotifikasi,
      formatDateOnly(hitung.tanggalBerikutnya), hitung.usia, recipients.join(','), status, keterangan
    ]);
  });

  return { sent: sent, skipped: skipped, totalAktif: rows.length };
}

/**
 * Tombol "Kirim Sekarang (Test)" di Settings — jalankan pengiriman
 * notifikasi HUT secara manual kapan saja (memakai logic & dedup yang
 * sama persis dengan trigger harian), untuk verifikasi tanpa menunggu
 * jam 08.00.
 */
function sendHutNotificationsNow() {
  requireRole('Admin');
  return sendHutNotificationsDaily_();
}

/**
 * Kirim SATU email uji coba (format sama persis dengan notifikasi
 * asli, ditandai "[UJI COBA]" di subjek) ke daftar penerima Settings,
 * memakai entri Master_HUT yang tanggal HUT-nya paling dekat saat ini
 * — TIDAK terikat aturan H-7/H-1 dan TIDAK dicatat ke HUT_EmailLog
 * sebagai pengiriman asli (supaya tidak mengganggu dedup notifikasi
 * sungguhan). Dipakai untuk memverifikasi pengiriman GmailApp & daftar
 * penerima benar-benar berfungsi, tanpa harus menunggu tanggal H-7/H-1.
 */
function sendHutTestEmail() {
  requireRole('Admin');
  const recipients = String(getSettingValue_(HUT_SETTINGS_KEY_EMAILS, '') || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!recipients.length) throw new Error('Isi & simpan dulu minimal satu email penerima sebelum kirim uji coba.');

  const rows = getMasterHutRows_().filter(r => String(r['Nama'] || '').trim());
  if (!rows.length) throw new Error('Belum ada data di Master_HUT.');

  const now = new Date();
  let closest = null;
  rows.forEach(r => {
    const hitung = hitungHutBerikutnya_(r['Tanggal Berdiri/Lahir'], now);
    if (!closest || hitung.daysUntil < closest.hitung.daysUntil) closest = { row: r, hitung: hitung };
  });

  const nama = closest.row['Nama'];
  const kategori = closest.row['Kategori'] || HUT_KATEGORI_PERUSAHAAN;
  const tanggalLabel = closest.hitung.tanggalBerikutnya.getDate() + ' ' + NAMA_BULAN_ID[closest.hitung.tanggalBerikutnya.getMonth()] + ' ' + closest.hitung.tahunTarget;
  const daysUntil = closest.hitung.daysUntil;
  const jenisPeringatan = kategori === HUT_KATEGORI_INDIVIDU ? 'Ulang Tahun' : 'HUT';

  const subject = '[UJI COBA] Pengingat HUT H-' + daysUntil + ' – ' + nama;
  const body = 'Yth. Tim,\n\n' +
    '(Email ini adalah UJI COBA untuk memverifikasi pengiriman notifikasi HUT — bukan pengingat H-7/H-1 sungguhan.)\n\n' +
    'Kami informasikan bahwa ' + nama + ' akan memperingati ' + jenisPeringatan + ' pada ' + tanggalLabel + ', atau H-' + daysUntil + ' dari hari ini.\n\n' +
    'Pada tanggal tersebut, ' + nama + ' akan berusia ' + closest.hitung.usia + ' tahun.\n\n' +
    'Mohon dapat dipersiapkan tindak lanjut dan ucapan yang diperlukan.\n\n' +
    'Terima kasih.';

  GmailApp.sendEmail(recipients.join(','), subject, body);
  logAudit('SEND_EMAIL', 'HutTest', '-', 'Email uji coba HUT dikirim ke ' + recipients.join(', ') + ' (contoh: ' + nama + ', H-' + daysUntil + ')');
  return { to: recipients, nama: nama, daysUntil: daysUntil };
}

// TEMP DIAGNOSTIK — investigasi kenapa "Kirim Email Uji Coba" dilaporkan
// tidak sampai ke inbox. Cek recipients tersimpan lalu tangkap pesan
// error persis kalau GmailApp.sendEmail gagal. (MailApp.getRemainingDailyQuota
// SENGAJA tidak dipakai lagi — versi sebelumnya gagal duluan di situ
// dengan error scope "script.send_mail" milik MailApp, BUKAN scope milik
// GmailApp yang sebenarnya dipakai untuk kirim email fitur ini, jadi
// menyesatkan diagnosanya.) (dihapus setelah selesai)
function _debugHutTestEmail_() {
  const result = {};
  try {
    result.recipientsRaw = getSettingValue_(HUT_SETTINGS_KEY_EMAILS, '');
    const recipients = String(result.recipientsRaw || '').split(',').map(s => s.trim()).filter(Boolean);
    result.recipients = recipients;
    try { result.effectiveUser = Session.getEffectiveUser().getEmail(); } catch (e2) { result.effectiveUserError = String(e2 && e2.message || e2); }
    if (!recipients.length) { result.error = 'Belum ada email penerima tersimpan di Settings (cek apakah Simpan Email Penerima berhasil).'; return result; }

    const rows = getMasterHutRows_().filter(r => String(r['Nama'] || '').trim());
    result.masterHutRowCount = rows.length;

    const subject = '[UJI COBA DEBUG] Test HUT Email - ' + new Date().toISOString();
    const body = 'Ini email uji coba debug dari sistem Kalender HUT & Ulang Tahun. Kalau email ini sampai, pengiriman GmailApp berfungsi normal.';
    GmailApp.sendEmail(recipients.join(','), subject, body);
    result.sendSuccess = true;
  } catch (err) {
    result.sendSuccess = false;
    result.errorMessage = String(err && err.message || err);
    result.errorStack = err && err.stack ? String(err.stack) : null;
  }
  return result;
}

/**
 * Log pengiriman notifikasi HUT terbaru (untuk ditampilkan di
 * Settings — status berhasil/gagal & waktu pengiriman).
 */
function getHutEmailLog(limit) {
  requireRole('Admin');
  getOrCreateHutEmailLogSheet_();
  const rows = sheetToObjects(SHEET_NAMES.HUT_EMAIL_LOG);
  rows.sort((a, b) => new Date(b.Timestamp || 0) - new Date(a.Timestamp || 0));
  return rows.slice(0, limit || 30);
}

function installHutNotificationTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendHutNotificationsDaily_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendHutNotificationsDaily_')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
}

function isHutTriggerInstalled_() {
  return ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sendHutNotificationsDaily_');
}

/**
 * JALANKAN dari menu Settings ("Pasang Trigger Notifikasi Otomatis")
 * — trigger harian akan mengeksekusi sendHutNotificationsDaily_()
 * sekitar jam 08.00 WIB (zona waktu project = Asia/Jakarta, sesuai
 * appsscript.json). Aman dijalankan berkali-kali (trigger lama
 * dihapus dulu supaya tidak dobel-dobel jalan).
 */
function installHutNotificationTrigger() {
  requireRole('Admin');
  installHutNotificationTrigger_();
  return { installed: true };
}

function uninstallHutNotificationTrigger() {
  requireRole('Admin');
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendHutNotificationsDaily_') { ScriptApp.deleteTrigger(t); removed++; }
  });
  return { removed: removed };
}

function getHutTriggerStatus() {
  requireRole('Admin');
  return { installed: isHutTriggerInstalled_() };
}

// ------------------------------------------------------------
// LAPORAN HARIAN ETF (EGR)
// ------------------------------------------------------------
// Modul baru, berdiri sendiri dari Laporan Harian (Trading) — TIDAK
// mencampur data Trading, TIDAK menyentuh sheet/fungsi existing apapun.
//
// PENTING (temuan investigasi): Top Up/Restock EGR dan Saldo berjalan
// TERNYATA sudah dicatat langsung di sheet Input_Jual_EGR itu sendiri,
// sebagai 2 kolom tambahan pada baris-baris transaksi penjualan ("Top Up
// EGR" dan "Saldo") — BUKAN dari sheet Input_Beli_EGR (yang ternyata
// mencatat hal lain: pembelian EGR dari supplier/PO, tidak sinkron
// dengan Top Up yang dimaksud di sini). Semua dibaca read-only.
//
// Aturan (sesuai koreksi terakhir):
// - Saldo EGR Tersedia (tanggal X) = Saldo baris TERAKHIR sebelum
//   tanggal X (0 kalau belum ada data sama sekali) + total kolom
//   "Top Up EGR" pada tanggal X.
// - Saldo Tersisa (tanggal X) = nilai kolom "Saldo" pada baris TERAKHIR
//   tanggal X (dibaca langsung dari sheet — sumber kebenaran resmi),
//   fallback ke Saldo EGR Tersedia - Penjualan hari ini kalau kolom
//   Saldo kosong.
// - Restock/Top Up ditampilkan terpisah sbg rincian tapi nilainya SAMA
//   dengan yang sudah dihitung di Saldo EGR Tersedia (tidak dihitung dua
//   kali / tidak ditambahkan lagi ke Saldo Tersisa).

/**
 * Baca "ledger" EGR dari Input_Jual_EGR: saldo baris terakhir SEBELUM
 * tanggal target, total Top Up EGR & Penjualan PADA tanggal target, dan
 * saldo baris terakhir tanggal target (kalau ada). Header/kolom dicari
 * via findHeaderRowIndex/buildHeaderMap/findColByKeywords (bukan index
 * tetap) — aman meski urutan kolom sumber berubah.
 */
function readEgrLedger_(targetDateStr) {
  const empty = { saldoSebelum: 0, topUpHariIni: 0, penjualanHariIni: 0, penjualanNilaiHariIni: 0, saldoAkhirHariIni: null, topUpRows: [] };
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (!sheet) return empty;

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return empty;

  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    qty: findColByKeywords(headerMap, ['qty']),
    topUp: findColByKeywords(headerMap, ['top', 'up']),
    saldo: findColByKeywords(headerMap, ['saldo']),
    referensi: findColByKeywords(headerMap, ['nomor']),
    totalAfterTax: findColByKeywords(headerMap, ['after', 'tax']),
    totalHarga: findColByKeywords(headerMap, ['total', 'harga'], ['after', 'tax'])
  };
  if (col.tanggal === -1) return empty;

  let saldoSebelum = 0;
  let topUpHariIni = 0;
  let penjualanHariIni = 0;
  let penjualanNilaiHariIni = 0;
  let saldoAkhirHariIni = null;
  const topUpRows = [];

  values.slice(headerRowIdx + 1).forEach(row => {
    const tgl = extractDateOnly(getByCol(row, col.tanggal));
    if (!tgl) return;
    if (tgl < targetDateStr) {
      if (col.saldo !== -1) {
        const s = getByCol(row, col.saldo);
        if (s !== '' && s !== null) saldoSebelum = Number(s) || 0;
      }
    } else if (tgl === targetDateStr) {
      if (col.qty !== -1) penjualanHariIni += Number(getByCol(row, col.qty)) || 0;
      let nilai = col.totalAfterTax !== -1 ? Number(getByCol(row, col.totalAfterTax)) || 0 : 0;
      if (!nilai && col.totalHarga !== -1) nilai = Number(getByCol(row, col.totalHarga)) || 0;
      penjualanNilaiHariIni += nilai;
      const topUpVal = col.topUp !== -1 ? Number(getByCol(row, col.topUp)) || 0 : 0;
      if (topUpVal > 0) {
        topUpHariIni += topUpVal;
        topUpRows.push({
          Referensi: col.referensi !== -1 ? getByCol(row, col.referensi) : '-',
          Gram: topUpVal,
          SaldoSetelah: col.saldo !== -1 ? (Number(getByCol(row, col.saldo)) || 0) : null
        });
      }
      if (col.saldo !== -1) {
        const s = getByCol(row, col.saldo);
        if (s !== '' && s !== null) saldoAkhirHariIni = Number(s) || 0;
      }
    }
  });

  return {
    saldoSebelum: saldoSebelum, topUpHariIni: topUpHariIni, penjualanHariIni: penjualanHariIni,
    penjualanNilaiHariIni: penjualanNilaiHariIni, saldoAkhirHariIni: saldoAkhirHariIni, topUpRows: topUpRows
  };
}

/**
 * Inventory Utama Kumulatif = akumulasi (Total Pembelian EGR - Total
 * Penjualan EGR + Total Redemption EGR) SEMUA HARI sejak data paling awal s.d. SEBELUM tanggal
 * target (tidak termasuk tanggal target sendiri). Dipakai sbg "Inventory
 * Sebelumnya" utk Inventory Utama Hari Ini = Inventory Sebelumnya + Net
 * Hari Ini. Murni akumulasi (bukan lookup per-tanggal), jadi otomatis
 * "lompat" ke tanggal aktif terakhir kalau ada hari tanpa transaksi
 * (mis. akhir pekan) — tanpa perlu logic khusus terpisah.
 */
function getEgrInventoryKumulatifSebelum_(targetDateStr) {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);

  function cumGramSebelum(sheetName) {
    const sheet = getEgrSheet_(ss, sheetName);
    if (!sheet) return 0;
    const values = sheet.getDataRange().getValues();
    const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
    if (headerRowIdx === -1) return 0;
    const headerMap = buildHeaderMap(values[headerRowIdx]);
    const tanggalCol = findColByKeywords(headerMap, ['tanggal']);
    const qtyCol = findColByKeywords(headerMap, ['qty']);
    const totalCol = findColByKeywords(headerMap, ['total', 'harga']);
    if (tanggalCol === -1 || qtyCol === -1 || totalCol === -1) return 0;

    let sum = 0;
    values.slice(headerRowIdx + 1).forEach(row => {
      const tgl = extractDateOnly(getByCol(row, tanggalCol));
      if (!tgl || tgl >= targetDateStr) return;
      const val = Number(getByCol(row, totalCol)) || 0;
      if (val <= 0) return;
      sum += Number(getByCol(row, qtyCol)) || 0;
    });
    return sum;
  }

  // Restocking menambah, Subscription mengurangi, Redemption menambah kembali.
  return cumGramSebelum(EGR_SHEET.BELI) - cumGramSebelum(EGR_SHEET.JUAL) + cumGramSebelum(EGR_SHEET.REDEM);
}

/**
 * Total gram Redemption EGR (Input_Redem_EGR) dari baris paling awal
 * s.d. SEBELUM tanggal target. Kriteria baris sama dengan inventory
 * kumulatif (Total Harga > 0). 0 kalau sheet/header/kolom tidak ada.
 */
function getEgrRedemGramSebelum_(targetDateStr) {
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.REDEM);
  if (!sheet) return 0;
  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return 0;
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const tanggalCol = findColByKeywords(headerMap, ['tanggal']);
  const qtyCol = findColByKeywords(headerMap, ['qty']);
  const totalCol = findColByKeywords(headerMap, ['total', 'harga']);
  if (tanggalCol === -1 || qtyCol === -1 || totalCol === -1) return 0;
  let sum = 0;
  values.slice(headerRowIdx + 1).forEach(row => {
    const tgl = extractDateOnly(getByCol(row, tanggalCol));
    if (!tgl || tgl >= targetDateStr) return;
    if ((Number(getByCol(row, totalCol)) || 0) <= 0) return;
    sum += Number(getByCol(row, qtyCol)) || 0;
  });
  return sum;
}

/**
 * Data Laporan Harian EGR — lihat catatan aturan Saldo EGR di atas.
 * Redemption (Input_Redem_EGR) = emas kembali ke saldo/inventory:
 * - Saldo EGR Tersedia = Saldo ledger Subs sebelum X + Redemption
 *   kumulatif sebelum X + Top Up X + Redemption X.
 * - Saldo Tersisa = Saldo ledger Subs akhir X (fallback: tersedia ledger
 *   - penjualan) + Redemption kumulatif s.d. X.
 * (Kolom Saldo di Input_Subs_EGR diperlakukan sbg ledger subscription
 * saja; redemption ditambahkan terpisah dari sheet-nya sendiri.)
 */
function getLaporanHarianEtfData(dateStr) {
  requireRole('Viewer');
  const targetDate = dateStr ? new Date(extractDateOnly(dateStr) + 'T00:00:00') : new Date();
  const tanggalLabel = targetDate.getDate() + ' ' + NAMA_BULAN_ID[targetDate.getMonth()] + ' ' + targetDate.getFullYear();
  const targetDateStr = formatDateOnly(targetDate);

  const ledger = readEgrLedger_(targetDateStr);

  // Redemption EGR (Input_Redem_EGR) — ditampilkan terpisah; emas kembali
  // ke saldo & inventory. Net Subscription = Penjualan - Redemption (gram).
  const redemRows = readEgrRedemRowsForDate_(targetDateStr);
  const redemTotalGram = redemRows.reduce((sum, r) => sum + (Number(r.Qty) || 0), 0);
  const redemTotalNominal = redemRows.reduce((sum, r) => sum + (Number(r.TotalHarga) || 0), 0);
  const netSubscriptionGram = ledger.penjualanHariIni - redemTotalGram;
  const redemSebelumGram = getEgrRedemGramSebelum_(targetDateStr);

  const saldoTersediaLedger = ledger.saldoSebelum + ledger.topUpHariIni;
  const saldoAkhirLedger = ledger.saldoAkhirHariIni !== null ? ledger.saldoAkhirHariIni : (saldoTersediaLedger - ledger.penjualanHariIni);
  const saldoTersedia = saldoTersediaLedger + redemSebelumGram + redemTotalGram;
  const saldoAkhir = saldoAkhirLedger + redemSebelumGram + redemTotalGram;

  // Detail transaksi Penjualan (Input_Jual_EGR) & Pembelian (Input_Beli_EGR)
  // pada tanggal ini — daftar baris, bukan sekadar total. Read-only,
  // fungsi baca sudah ada sebelumnya (dipakai juga di Laporan Harian Trading
  // utk menggabung kategori ETF), tidak memengaruhi perhitungan Saldo di atas.
  const penjualanRows = readEgrJualRowsForDate_(targetDateStr);
  const pembelianRows = readEgrBeliRowsForDate_(targetDateStr);

  // Restock Pembelian EGR = HANYA dari Input_Beli_EGR (bukan kolom Top Up
  // EGR di Input_Jual_EGR) — dipakai jg utk ringkasan Inventory Utama di
  // bawahnya. Berdiri sendiri, TIDAK sama dengan Saldo Tersisa.
  const restockTotalGram = pembelianRows.reduce((sum, r) => sum + (Number(r.Qty) || 0), 0);
  const restockTotalNominal = pembelianRows.reduce((sum, r) => sum + (Number(r.TotalHarga) || 0), 0);

  // Inventory Utama Kumulatif: Net Hari Ini = Pembelian - Penjualan +
  // Redemption hari ini; Inventory Utama Hari Ini = Inventory Sebelumnya (akumulasi s.d.
  // sebelum tanggal ini) + Net Hari Ini — BUKAN cuma Net Hari Ini sendiri.
  const netEgrHariIni = restockTotalGram - ledger.penjualanHariIni + redemTotalGram;
  const inventorySebelumnya = getEgrInventoryKumulatifSebelum_(targetDateStr);
  const inventoryUtamaKumulatif = inventorySebelumnya + netEgrHariIni;

  const tanggalSebelum = new Date(targetDate);
  tanggalSebelum.setDate(tanggalSebelum.getDate() - 1);
  const tanggalSebelumLabel = tanggalSebelum.getDate() + ' ' + NAMA_BULAN_ID[tanggalSebelum.getMonth()] + ' ' + tanggalSebelum.getFullYear();

  return {
    tanggalLabel: tanggalLabel,
    tanggalSebelumLabel: tanggalSebelumLabel,
    targetDateStr: targetDateStr,
    saldoTersedia: saldoTersedia,
    penjualanEtfGram: ledger.penjualanHariIni,
    penjualanEtfNilai: ledger.penjualanNilaiHariIni,
    saldoAkhir: saldoAkhir,
    restock: { totalGram: restockTotalGram, totalNominal: restockTotalNominal },
    redemption: { totalGram: redemTotalGram, totalNominal: redemTotalNominal, rows: redemRows },
    netSubscriptionGram: netSubscriptionGram,
    inventorySebelumnya: inventorySebelumnya,
    netEgrHariIni: netEgrHariIni,
    inventoryUtamaKumulatif: inventoryUtamaKumulatif,
    penjualan: { rows: penjualanRows },
    pembelian: { rows: pembelianRows }
  };
}

// ------------------------------------------------------------
// PERBANDINGAN PENJUALAN EGR (vs tanggal transaksi terakhir sebelumnya)
// ------------------------------------------------------------
// Modul baru, berdiri sendiri — read-only dari Input_Jual_EGR saja, tidak
// menyentuh sheet/fungsi existing apapun. "Penjualan Sebelumnya" adalah
// tanggal TRANSAKSI TERAKHIR sebelum tanggal laporan (bukan sekadar H-1
// kalender) — kalau tgl laporan Senin & Jumat adalah transaksi terakhir
// sebelumnya, otomatis dibandingkan dgn Jumat.
// Kolom paling KANAN yang header-nya mengandung kata tertentu — beda dari
// findColByKeywords (yg mengembalikan match pertama/paling kiri secara
// nama unik) karena di sini kita butuh posisi kolom TERAKHIR di baris
// header mentah, walau ada beberapa kolom dgn kata kunci yg sama.
function findLastColContaining_(headerRow, keyword) {
  let found = -1;
  headerRow.forEach(function (h, i) {
    if (String(h).trim().toLowerCase().indexOf(keyword) !== -1) found = i;
  });
  return found;
}

// Parse angka dari sel yang mungkin berupa ANGKA murni ATAU TEKS berformat
// rupiah manual (mis. "Rp2.440.905", "Rp 76.950.000") — beberapa kolom di
// Input_Jual_EGR (terutama Buyback) ternyata diisi sbg teks dengan prefix
// "Rp", bukan angka murni, sehingga Number(...) biasa gagal parse (NaN)
// dan nilainya jadi terlihat "tidak ada" padahal datanya ada.
function parseAngkaSel_(raw) {
  if (raw === '' || raw === null || raw === undefined) return null;
  if (typeof raw === 'number') return raw;
  const cleaned = String(raw).replace(/[^0-9,.-]/g, '');
  if (!cleaned) return null;
  const normalized = cleaned.replace(/\./g, '').replace(',', '.');
  const num = Number(normalized);
  return isNaN(num) ? null : num;
}

// Hitung Spread Rupiah & Persentase Spread dari satu pasang Harga
// Jual/Buyback (salah satu/keduanya boleh null kalau data belum ada).
function computeEgrSpread_(hargaJual, hargaBuyback) {
  if (hargaJual === null || hargaJual === undefined) return { spread: null, spreadPersen: null };
  const buyback = (hargaBuyback === null || hargaBuyback === undefined) ? 0 : hargaBuyback;
  const spread = hargaJual - buyback;
  const spreadPersen = hargaJual ? (spread / hargaJual) * 100 : null;
  return { spread: spread, spreadPersen: spreadPersen };
}

const EGR_HARGA_POSISI_LABEL = ['Harga I', 'Harga II', 'Harga III', 'Harga IV', 'Harga V', 'Harga VI', 'Harga VII', 'Harga VIII'];

function getEgrPerbandinganPenjualanData(dateStr) {
  requireRole('Viewer');
  const targetDate = dateStr ? new Date(extractDateOnly(dateStr) + 'T00:00:00') : new Date();
  const tanggalLabel = targetDate.getDate() + ' ' + NAMA_BULAN_ID[targetDate.getMonth()] + ' ' + targetDate.getFullYear();
  const targetDateStr = formatDateOnly(targetDate);

  const kosongAgg = { qty: 0, nilai: 0, hargaJual: null, buybackTerakhir: null };
  const kosongPosisi = { hargaJual: null, hargaBuyback: null, spread: null, spreadPersen: null };
  const emptyPrevDate = new Date(targetDate);
  emptyPrevDate.setDate(emptyPrevDate.getDate() - 1);
  const emptyPrevDateStr = formatDateOnly(emptyPrevDate);
  const empty = {
    tanggalLabel: tanggalLabel, targetDateStr: targetDateStr,
    hariIni: { qty: 0, nilai: 0, hargaJual: null },
    sebelumnya: { qty: 0, nilai: 0 },
    tanggalSebelumnyaStr: emptyPrevDateStr, tanggalSebelumnyaLabel: formatTanggalPanjangIndo_(emptyPrevDateStr),
    adaDataSebelumnya: true, sebelumnyaAdaTransaksi: false,
    hargaBuyback: 0, spread: null, spreadPersen: null,
    rekapEgrNilai: 0, rekapBuybackGram: 0, rekapBuybackNilai: 0,
    hargaPosisi: { h1: kosongPosisi, list: [] }
  };

  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (!sheet) return empty;

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return empty;
  const headerRow = values[headerRowIdx];
  const headerMap = buildHeaderMap(headerRow);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    qty: findColByKeywords(headerMap, ['qty']),
    // Exclude 'buyback' juga — supaya kolom "Harga Buyback" (kalau ada)
    // TIDAK ketuker/kepilih sebagai kolom Harga Jual (Harga per Gram).
    hargaPerGram: findColByKeywords(headerMap, ['harga'], ['total', 'after', 'buyback']),
    totalAfterTax: findColByKeywords(headerMap, ['after', 'tax']),
    totalHarga: findColByKeywords(headerMap, ['total', 'harga'], ['after', 'tax']),
    // Kolom Buyback PALING KANAN yang header-nya mengandung "buyback" —
    // bukan dari Input_Beli_EGR.
    buyback: findLastColContaining_(headerRow, 'buyback')
  };
  if (col.tanggal === -1) return empty;

  // Agregat per tanggal: Qty & Nilai Transaksi (Total harga after tax,
  // fallback Total Harga — sama persis dgn cara "Penjualan Hari Ini"
  // dihitung di Laporan Harian EGR supaya konsisten antar-halaman), Harga
  // per Gram TERAKHIR yg terisi & nilai Buyback TERAKHIR yg terisi (urutan
  // baris sesuai sheet, KEDUANYA TIDAK dijumlah/dirata-rata/dihitung —
  // langsung ambil apa adanya dari kolomnya). Khusus tanggal laporan, JUGA
  // deteksi SETIAP posisi harga (pasangan Jual+Buyback) yang benar-benar
  // berbeda dari posisi sebelumnya, berurutan sesuai baris sheet (dipakai
  // sbg proxy urutan waktu transaksi, krn sheet tidak punya kolom jam).
  const perTanggal = {};
  const posisiHariIni = [];
  let curJualPos = null, curBuybackPos = null;
  values.slice(headerRowIdx + 1).forEach(function (row) {
    const tgl = extractDateOnly(getByCol(row, col.tanggal));
    if (!tgl) return;
    if (!perTanggal[tgl]) perTanggal[tgl] = { qty: 0, nilai: 0, hargaJual: null, buybackTerakhir: null };

    const qty = col.qty !== -1 ? (parseAngkaSel_(getByCol(row, col.qty)) || 0) : 0;
    perTanggal[tgl].qty += qty;

    let nilai = col.totalAfterTax !== -1 ? (parseAngkaSel_(getByCol(row, col.totalAfterTax)) || 0) : 0;
    if (!nilai && col.totalHarga !== -1) nilai = parseAngkaSel_(getByCol(row, col.totalHarga)) || 0;
    perTanggal[tgl].nilai += nilai;

    const h = col.hargaPerGram !== -1 ? parseAngkaSel_(getByCol(row, col.hargaPerGram)) : null;
    if (h !== null && h !== 0) perTanggal[tgl].hargaJual = h;
    const b = col.buyback !== -1 ? parseAngkaSel_(getByCol(row, col.buyback)) : null;
    if (b !== null && b !== 0) perTanggal[tgl].buybackTerakhir = b;

    if (tgl === targetDateStr) {
      if (h !== null && h !== 0) curJualPos = h;
      if (b !== null && b !== 0) curBuybackPos = b;
      if (curJualPos !== null || curBuybackPos !== null) {
        const last = posisiHariIni[posisiHariIni.length - 1];
        if (!last || last.hargaJual !== curJualPos || last.hargaBuyback !== curBuybackPos) {
          posisiHariIni.push({ hargaJual: curJualPos, hargaBuyback: curBuybackPos });
        }
      }
    }
  });

  const hariIniAgg = perTanggal[targetDateStr] || kosongAgg;

  // H-1 = SATU HARI KALENDER sebelum tanggal laporan, SELALU — BUKAN
  // "tanggal transaksi terakhir sebelum ini". Tetap dipakai/ditampilkan
  // walau tidak ada transaksi pada tanggal itu (nilainya 0, bukan
  // dilewati/dicari mundur ke tanggal lain yg ada datanya).
  const prevDate = new Date(targetDate);
  prevDate.setDate(prevDate.getDate() - 1);
  const tanggalSebelumnyaStr = formatDateOnly(prevDate);
  const sebelumnyaAgg = perTanggal[tanggalSebelumnyaStr] || kosongAgg;
  const sebelumnyaAdaTransaksi = !!perTanggal[tanggalSebelumnyaStr];

  // --- Bagian A: kolom harga dinamis (H-1, Harga I, Harga II, ...) ---
  const h1Spread = computeEgrSpread_(sebelumnyaAgg.hargaJual, sebelumnyaAgg.buybackTerakhir);
  const hargaPosisiList = posisiHariIni.map(function (p, idx) {
    const sp = computeEgrSpread_(p.hargaJual, p.hargaBuyback);
    return {
      label: EGR_HARGA_POSISI_LABEL[idx] || ('Harga ' + (idx + 1)),
      hargaJual: p.hargaJual, hargaBuyback: p.hargaBuyback,
      spread: sp.spread, spreadPersen: sp.spreadPersen
    };
  });

  // "Harga hari ini" yang dipakai di tempat lain (mis. narasi Summary)
  // SELALU Harga I — perubahan harga berikutnya di hari yg sama (Harga
  // II, III, dst) hanya ditampilkan di tabel Perbandingan Harga, tidak
  // menggantikan harga acuan hari ini.
  const hargaIJual = hargaPosisiList.length ? hargaPosisiList[0].hargaJual : null;
  const hargaIBuyback = hargaPosisiList.length ? (hargaPosisiList[0].hargaBuyback || 0) : 0;
  const hargaISpread = computeEgrSpread_(hargaIJual, hargaIBuyback);

  return {
    tanggalLabel: tanggalLabel,
    targetDateStr: targetDateStr,
    hariIni: { qty: hariIniAgg.qty, nilai: hariIniAgg.nilai, hargaJual: hargaIJual },
    sebelumnya: { qty: sebelumnyaAgg.qty, nilai: sebelumnyaAgg.nilai },
    tanggalSebelumnyaStr: tanggalSebelumnyaStr,
    tanggalSebelumnyaLabel: formatTanggalPanjangIndo_(tanggalSebelumnyaStr),
    adaDataSebelumnya: true, // H-1 (tanggal kalender) SELALU ada & ditampilkan, walau nilainya 0
    sebelumnyaAdaTransaksi: sebelumnyaAdaTransaksi, // dipakai narasi "Baru"/"Tidak ada perubahan" saat H-1 nol
    hargaBuyback: hargaIBuyback,
    spread: hargaISpread.spread,
    spreadPersen: hargaISpread.spreadPersen,
    rekapEgrNilai: hariIniAgg.nilai,
    // Transaksi Buyback: BELUM ADA transaksi Buyback sungguhan sampai saat
    // ini — SELALU 0, TIDAK dihitung dari kolom Buyback (kolom itu cuma
    // dipakai sbg Harga Buyback di Bagian A, dua data yg berbeda).
    rekapBuybackGram: 0,
    rekapBuybackNilai: 0,
    // Kolom harga dinamis Bagian A — h1 = harga terakhir tanggal
    // sebelumnya, list = setiap posisi harga berbeda pada tanggal
    // laporan berurutan sesuai baris sheet (Harga I, II, III, ...).
    hargaPosisi: {
      h1: { hargaJual: sebelumnyaAgg.hargaJual, hargaBuyback: sebelumnyaAgg.buybackTerakhir, spread: h1Spread.spread, spreadPersen: h1Spread.spreadPersen },
      list: hargaPosisiList
    }
  };
}

// ------------------------------------------------------------
// LAPORAN REKAP EGR (rentang tanggal, rekap per Buyer)
// ------------------------------------------------------------
// Modul baru, berdiri sendiri — read-only dari Input_Jual_EGR &
// Input_Beli_EGR, tidak menyentuh sheet/fungsi existing apapun (termasuk
// tidak menyentuh logic Laporan Harian EGR/Trading di atas).
/**
 * Saldo EGR / Total Transaksi kumulatif SAMPAI tanggal akhir laporan
 * (inklusif, TIDAK dibatasi tanggal mulai) — dipakai utk scorecard utama
 * Laporan Rekap EGR (Saldo EGR / Total Transaksi / Saldo Tersisa).
 * Sumber: Input_Jual_EGR saja — kolom Top Up EGR utk Saldo, Qty utk
 * Total Transaksi, Harga per Gram TERAKHIR (baris terakhir yg terisi,
 * bukan dijumlah/dirata-rata) utk valuasi Saldo EGR.
 */
function getEgrSaldoKumulatifSampai_(endDateStr) {
  const result = { saldoGram: 0, totalTransaksiGram: 0, totalTransaksiNilai: 0, hargaTerakhir: null };
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (!sheet) return result;

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return result;
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    qty: findColByKeywords(headerMap, ['qty']),
    topUp: findColByKeywords(headerMap, ['top', 'up']),
    hargaPerGram: findColByKeywords(headerMap, ['harga'], ['total', 'after']),
    totalAfterTax: findColByKeywords(headerMap, ['after', 'tax']),
    totalHarga: findColByKeywords(headerMap, ['total', 'harga'], ['after', 'tax'])
  };
  if (col.tanggal === -1) return result;

  values.slice(headerRowIdx + 1).forEach(function (row) {
    const tgl = extractDateOnly(getByCol(row, col.tanggal));
    if (!tgl || tgl > endDateStr) return;

    if (col.topUp !== -1) result.saldoGram += Number(getByCol(row, col.topUp)) || 0;
    if (col.qty !== -1) result.totalTransaksiGram += Number(getByCol(row, col.qty)) || 0;

    let nilai = col.totalAfterTax !== -1 ? Number(getByCol(row, col.totalAfterTax)) || 0 : 0;
    if (!nilai && col.totalHarga !== -1) nilai = Number(getByCol(row, col.totalHarga)) || 0;
    result.totalTransaksiNilai += nilai;

    if (col.hargaPerGram !== -1) {
      const h = getByCol(row, col.hargaPerGram);
      if (h !== '' && h !== null && !isNaN(Number(h)) && Number(h) !== 0) result.hargaTerakhir = Number(h);
    }
  });

  return result;
}

/**
 * Rekap EGR untuk satu rentang tanggal [startDateStr, endDateStr]
 * (inklusif): Penjualan dikelompokkan per Buyer (nama digabung meski
 * beda kapital/spasi), + scorecard utama Saldo EGR/Total
 * Transaksi/Saldo Tersisa (KUMULATIF sampai endDateStr, lihat
 * getEgrSaldoKumulatifSampai_ — tidak dibatasi tanggal mulai, karena
 * merepresentasikan posisi saldo per tanggal akhir, bukan delta rentang).
 * buyerFilter opsional — kalau diisi, tabel Buyer dipersempit ke buyer
 * itu saja (scorecard utama tidak ikut difilter Buyer).
 */
function getEgrRekapData(startDateStr, endDateStr, buyerFilter) {
  requireRole('Viewer');
  const start = extractDateOnly(startDateStr);
  const end = extractDateOnly(endDateStr);
  if (!start || !end) throw new Error('Tanggal Mulai dan Tanggal Akhir wajib diisi.');
  const buyerKeyFilter = buyerFilter ? String(buyerFilter).trim().toLowerCase().replace(/\s+/g, ' ') : '';

  // --- Penjualan EGR (Input_Jual_EGR) dalam rentang tanggal, dikelompokkan per Buyer ---
  const buyerMap = {};
  let totalPenjualanGram = 0, totalPenjualanNilai = 0, totalPenjualanCount = 0;

  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const jualSheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (jualSheet) {
    const values = jualSheet.getDataRange().getValues();
    const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
    if (headerRowIdx !== -1) {
      const headerMap = buildHeaderMap(values[headerRowIdx]);
      const col = {
        tanggal: findColByKeywords(headerMap, ['tanggal']),
        pembeli: findColByKeywords(headerMap, ['pembeli']),
        qty: findColByKeywords(headerMap, ['qty']),
        totalAfterTax: findColByKeywords(headerMap, ['after', 'tax']),
        totalHarga: findColByKeywords(headerMap, ['total', 'harga'], ['after', 'tax'])
      };
      values.slice(headerRowIdx + 1).forEach(row => {
        const tgl = extractDateOnly(getByCol(row, col.tanggal));
        if (!tgl || tgl < start || tgl > end) return;
        // Qty (Gram) = penjualan sesungguhnya pada baris ini; kolom Top
        // Up EGR (kalau ada di baris yang sama) TIDAK ikut dihitung di sini.
        const qty = col.qty !== -1 ? Number(getByCol(row, col.qty)) || 0 : 0;
        if (qty <= 0) return; // abaikan baris kosong/tidak valid

        const namaAsli = col.pembeli !== -1 ? String(getByCol(row, col.pembeli) || '').trim() : '';
        if (!namaAsli) return;

        const namaKey = namaAsli.toLowerCase().replace(/\s+/g, ' ');
        if (buyerKeyFilter && namaKey !== buyerKeyFilter) return;

        let nilai = col.totalAfterTax !== -1 ? Number(getByCol(row, col.totalAfterTax)) || 0 : 0;
        if (!nilai && col.totalHarga !== -1) nilai = Number(getByCol(row, col.totalHarga)) || 0;

        if (!buyerMap[namaKey]) buyerMap[namaKey] = { nama: namaAsli, count: 0, gram: 0, nilai: 0 };
        buyerMap[namaKey].count++;
        buyerMap[namaKey].gram += qty;
        buyerMap[namaKey].nilai += nilai;

        totalPenjualanGram += qty;
        totalPenjualanNilai += nilai;
        totalPenjualanCount++;
      });
    }
  }

  const buyerRows = Object.keys(buyerMap).map(k => buyerMap[k]).sort((a, b) => b.gram - a.gram);

  const kumulatif = getEgrSaldoKumulatifSampai_(end);
  const saldoEgrGram = kumulatif.saldoGram;
  const saldoEgrNilai = kumulatif.hargaTerakhir !== null ? saldoEgrGram * kumulatif.hargaTerakhir : 0;
  const saldoTersisaGram = saldoEgrGram - kumulatif.totalTransaksiGram;

  return {
    startDateStr: start,
    endDateStr: end,
    penjualan: { totalGram: totalPenjualanGram, totalNilai: totalPenjualanNilai, totalCount: totalPenjualanCount },
    saldoEgr: { gram: saldoEgrGram, nilai: saldoEgrNilai },
    totalTransaksi: { gram: kumulatif.totalTransaksiGram, nilai: kumulatif.totalTransaksiNilai },
    saldoTersisaGram: saldoTersisaGram,
    buyerRows: buyerRows
  };
}

// ------------------------------------------------------------
// LAPORAN BUYER (Trading & EGR) — rekap per Buyer, filter Bulan +
// Minggu ke- (berdasarkan tanggal transaksi, tahun berjalan). Data
// transaksi asli TIDAK diubah/digabung — penggabungan per Buyer hanya
// terjadi di hasil laporan ini.
// ------------------------------------------------------------
function laporanBuyerMingguKe_(day) {
  return Math.min(5, Math.ceil(day / 7));
}

// tglStr wajib format yyyy-MM-dd (hasil extractDateOnly). monthNum 1-12,
// weekNum 0 (Semua Minggu) atau 1-5. Tahun SELALU tahun berjalan (Asia/Jakarta).
function laporanBuyerCocokFilter_(tglStr, monthNum, weekNum) {
  if (!tglStr) return false;
  const p = tglStr.split('-');
  const y = Number(p[0]), m = Number(p[1]), d = Number(p[2]);
  const nowYear = Number(formatDateOnly(new Date()).split('-')[0]);
  if (y !== nowYear || m !== monthNum) return false;
  if (weekNum && laporanBuyerMingguKe_(d) !== weekNum) return false;
  return true;
}

/**
 * Sama seperti laporanBuyerCocokFilter_, tapi tahun & bulan eksplisit
 * (bukan selalu "tahun berjalan") — dipakai untuk bandingkan periode
 * terpilih dengan bulan sebelumnya (yang bisa jatuh di tahun sebelumnya
 * kalau bulan terpilih Januari).
 */
function laporanBuyerCocokFilterYM_(tglStr, year, monthNum, weekNum) {
  if (!tglStr) return false;
  const p = tglStr.split('-');
  const y = Number(p[0]), m = Number(p[1]), d = Number(p[2]);
  if (y !== year || m !== monthNum) return false;
  if (weekNum && laporanBuyerMingguKe_(d) !== weekNum) return false;
  return true;
}

/**
 * Cocok untuk opsi "Semua Bulan (Januari–Saat Ini)" — tglStr (yyyy-MM-dd)
 * berada di antara 1 Januari tahun berjalan s.d. hari ini (inklusif).
 * String yyyy-MM-dd aman dibandingkan leksikografis (urutannya sama
 * dengan urutan tanggal asli).
 */
function laporanBuyerCocokFilterSemuaBulan_(tglStr, todayStr, jan1Str) {
  if (!tglStr) return false;
  return tglStr >= jan1Str && tglStr <= todayStr;
}

// ------------------------------------------------------------
// "Download Excel" DETAIL — Laporan Buyer Trading & Laporan Supplier
// Bulanan. Builder & export-runner DIBAGI (reuse) antara kedua
// laporan supaya styling, struktur kolom, baris TOTAL, dan usability
// (freeze/filter/autofit/number format) konsisten di satu tempat,
// bukan dua implementasi terpisah. Pola export-ke-xlsx (buat
// spreadsheet sementara -> export via Drive export endpoint -> hapus
// spreadsheet sementara) SAMA PERSIS dengan exportKinerjaBulananExcel
// yang sudah ada — tidak ada pendekatan baru.
// ------------------------------------------------------------

/**
 * Tulis sheet "Detail Laporan Buyer/Supplier" — 1 baris = 1 transaksi
 * (bukan rekap/summary). opts: { judul, periodeLabel, namaKolomEntitas,
 * rows: [{nama, tanggal('yyyy-MM-dd'), noTransaksi, gram(gram asli,
 * BUKAN kg), hargaPerGram, totalHarga}] }.
 */
function buildDetailLaporanSheet_(sh, opts) {
  sh.clear();
  sh.clearNotes();
  // sh.clear() TIDAK menghapus frozen rows/kolom maupun merge sisa
  // generate sebelumnya — reset eksplisit dulu (pola sama seperti
  // buildKinerjaBulananSheet_) supaya tidak error "can't merge frozen
  // and non-frozen columns" kalau fungsi ini dipanggil berulang.
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  const maxRows0 = sh.getMaxRows(), maxCols0 = sh.getMaxColumns();
  if (maxRows0 > 0 && maxCols0 > 0) sh.getRange(1, 1, maxRows0, maxCols0).breakApart();

  const GREEN_DARK = '#1D5C44';
  const GREEN_SOFT = '#E8F3EC';
  const GOLD = '#C9A227';
  const WHITE = '#FFFFFF';
  const COLS = 7; // No, Nama, Tanggal, No. Transaksi, Gramasi (Kg), Harga (Rp/gr), Total Rupiah
  const rows = opts.rows || [];

  sh.setColumnWidths(1, COLS, 100);
  sh.setColumnWidth(1, 45);

  // --- Judul ---
  sh.getRange(1, 1, 1, COLS).merge();
  kbSetCell_(sh, 'A1', opts.judul, { bold: true, size: 16, color: WHITE });
  sh.getRange(1, 1, 1, COLS).setBackground(GREEN_DARK).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);

  // --- Periode (mengikuti filter dashboard yang aktif) ---
  sh.getRange(2, 1, 1, COLS).merge();
  kbSetCell_(sh, 'A2', 'Periode Laporan: ' + (opts.periodeLabel || '-'), { bold: true, size: 10, color: GREEN_DARK });
  sh.getRange(2, 1, 1, COLS).setBackground(GREEN_SOFT).setHorizontalAlignment('center');

  // --- Header tabel (baris 4, baris 3 dikosongkan sbg spacer) ---
  const headerRow = 4;
  sh.getRange(headerRow, 1, 1, COLS).setValues([['No', opts.namaKolomEntitas, 'Tanggal', 'No. Transaksi', 'Gramasi (Kg)', 'Harga (Rp/gr)', 'Total Rupiah']]);
  sh.getRange(headerRow, 1, 1, COLS)
    .setFontWeight('bold').setFontColor(WHITE).setBackground(GREEN_DARK)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#0F3A29', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(headerRow, 26);

  // --- Isi tabel: zebra rows, border tipis ---
  let r = headerRow + 1;
  rows.forEach(function (row, idx) {
    const bg = idx % 2 === 0 ? WHITE : GREEN_SOFT;
    sh.getRange(r, 1).setValue(idx + 1);
    sh.getRange(r, 2).setValue(row.nama);
    const tglParts = String(row.tanggal || '').split('-');
    if (tglParts.length === 3) {
      sh.getRange(r, 3).setValue(new Date(Number(tglParts[0]), Number(tglParts[1]) - 1, Number(tglParts[2]))).setNumberFormat('dd mmm yyyy');
    } else {
      sh.getRange(r, 3).setValue('-');
    }
    sh.getRange(r, 4).setValue(String(row.noTransaksi || '-')).setNumberFormat('@');
    sh.getRange(r, 5).setValue((Number(row.gram) || 0) / 1000).setNumberFormat('#,##0.00" Kg"');
    sh.getRange(r, 6).setValue(Number(row.hargaPerGram) || 0).setNumberFormat('"Rp"#,##0');
    sh.getRange(r, 7).setValue(Number(row.totalHarga) || 0).setNumberFormat('"Rp"#,##0');
    sh.getRange(r, 1, 1, COLS).setBackground(bg)
      .setBorder(true, true, true, true, false, false, '#D9E6DF', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 1).setHorizontalAlignment('center');
    sh.getRange(r, 3).setHorizontalAlignment('center');
    sh.getRange(r, 5, 1, 3).setHorizontalAlignment('right');
    r++;
  });

  if (!rows.length) {
    sh.getRange(r, 1, 1, COLS).merge();
    kbSetCell_(sh, 'A' + r, 'Tidak ada transaksi pada periode/filter ini.', { color: '#6B7684', align: 'center' });
    r++;
  }

  // --- Baris TOTAL ---
  const totalGramKg = rows.reduce(function (s, x) { return s + (Number(x.gram) || 0); }, 0) / 1000;
  const totalNilai = rows.reduce(function (s, x) { return s + (Number(x.totalHarga) || 0); }, 0);
  const totalRow = r;
  sh.getRange(totalRow, 1, 1, 4).merge();
  kbSetCell_(sh, 'A' + totalRow, 'TOTAL — ' + rows.length.toLocaleString('id-ID') + ' transaksi', { bold: true, color: WHITE });
  sh.getRange(totalRow, 5).setValue(totalGramKg).setNumberFormat('#,##0.00" Kg"');
  sh.getRange(totalRow, 7).setValue(totalNilai).setNumberFormat('"Rp"#,##0');
  sh.getRange(totalRow, 1, 1, COLS)
    .setBackground(GREEN_DARK).setFontColor(WHITE).setFontWeight('bold')
    .setBorder(true, true, true, true, false, false, GOLD, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(totalRow, 5, 1, 3).setHorizontalAlignment('right');
  sh.setRowHeight(totalRow, 26);

  // --- Usability: freeze header, autofilter, autofit kolom ---
  sh.setFrozenRows(headerRow);
  sh.getRange(headerRow, 1, (totalRow - headerRow), COLS).createFilter();
  for (let c = 1; c <= COLS; c++) sh.autoResizeColumn(c);
  if (sh.getColumnWidth(2) < 220) sh.setColumnWidth(2, 220);
  if (sh.getColumnWidth(4) < 170) sh.setColumnWidth(4, 170);
}

/**
 * Bungkus buildDetailLaporanSheet_ ke dalam SATU spreadsheet
 * sementara (supaya file yang di-download HANYA berisi sheet detail
 * ini, tidak ikut sheet lain) lalu export sebagai xlsx via Drive
 * export endpoint — pola SAMA PERSIS dengan exportKinerjaBulananExcel
 * yang sudah ada. Spreadsheet sementara selalu dihapus (finally),
 * TIDAK PERNAH menyentuh sheet/database aplikasi manapun.
 */
function exportDetailLaporanToXlsx_(sheetTitle, buildOpts, filenameBase) {
  const tempSs = SpreadsheetApp.create('TMP_' + filenameBase + '_' + new Date().getTime());
  const tempId = tempSs.getId();
  try {
    const sh = tempSs.getSheets()[0];
    sh.setName(sheetTitle);
    buildDetailLaporanSheet_(sh, buildOpts);
    SpreadsheetApp.flush();

    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tempId + '/export?format=xlsx';
    const resp = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error('Gagal export ke xlsx (HTTP ' + resp.getResponseCode() + ')');
    }
    const base64 = Utilities.base64Encode(resp.getBlob().getBytes());
    return { base64: base64, filename: filenameBase + '.xlsx' };
  } finally {
    DriveApp.getFileById(tempId).setTrashed(true);
  }
}

/**
 * Laporan Buyer — Trading. Baca PenjualanTransaksi (sheet aplikasi ini
 * sendiri), kelompokkan per Pembeli (nama digabung meski beda
 * kapital/spasi) untuk Bulan (+ opsional Minggu ke-) yang dipilih, tahun
 * berjalan. Gramasi (Qty) dalam Gram — dikonversi ke Kg di client.
 *
 * buyerFilter (opsional): kalau diisi, buyerRows cuma berisi 1 buyer itu
 * — tapi `totals` & `%volume` tiap baris tetap dihitung dari SELURUH
 * buyer periode ini (supaya KPI & persentase tidak berubah arti waktu
 * difilter 1 buyer). `prevTotals` = total periode bulan sebelumnya
 * (Minggu yang sama kalau Minggu dipilih) untuk perbandingan di KPI.
 *
 * monthNum juga menerima string 'all' — opsi "Semua Bulan
 * (Januari–Saat Ini)": gabungkan SEMUA transaksi tahun berjalan dari
 * 1 Januari s.d. hari ini jadi satu periode (Minggu diabaikan, karena
 * tidak ada satu bulan spesifik). Buyer tetap dihitung unik lewat
 * buyerMap yang sama (1 buyer = 1 entri walau muncul di banyak bulan).
 * Tidak ada "periode sebelumnya" yang relevan untuk rentang kumulatif
 * ini, jadi prevTotals dikembalikan kosong (0) & prevPeriodeLabel ''.
 */
function getLaporanBuyerTrading(monthNum, weekNum, buyerFilter) {
  requireRole('Viewer');
  const isSemuaBulan = String(monthNum) === 'all';
  monthNum = isSemuaBulan ? 0 : Number(monthNum);
  weekNum = (!isSemuaBulan && weekNum) ? Number(weekNum) : 0;
  if (!isSemuaBulan && !monthNum) throw new Error('Bulan wajib dipilih.');
  const buyerFilterKey = String(buyerFilter || '').trim().toLowerCase().replace(/\s+/g, ' ');

  const todayStr = formatDateOnly(new Date());
  const nowYear = Number(todayStr.split('-')[0]);
  const jan1Str = nowYear + '-01-01';
  let prevYear = nowYear, prevMonth = monthNum - 1;
  if (prevMonth < 1) { prevMonth = 12; prevYear = nowYear - 1; }

  const buyerMap = {};
  let prevTransaksi = 0, prevGram = 0, prevNilai = 0;
  const prevBuyerSet = {};

  sheetToObjects(SHEET_NAMES.PENJUALAN_FULL).forEach(function (p) {
    const qty = Number(p.Qty) || 0;
    if (qty <= 0) return;
    const namaAsli = String(p.Pembeli || '').trim();
    if (!namaAsli) return;
    const tgl = extractDateOnly(p.Tanggal);
    const nilai = Number(p.TotalHarga) || 0;
    const key = namaAsli.toLowerCase().replace(/\s+/g, ' ');

    const cocokPeriodeAktif = isSemuaBulan
      ? laporanBuyerCocokFilterSemuaBulan_(tgl, todayStr, jan1Str)
      : laporanBuyerCocokFilterYM_(tgl, nowYear, monthNum, weekNum);

    if (cocokPeriodeAktif) {
      if (!buyerMap[key]) buyerMap[key] = { nama: namaAsli, gram: 0, nilai: 0, jumlahTransaksi: 0 };
      buyerMap[key].gram += qty;
      buyerMap[key].nilai += nilai;
      buyerMap[key].jumlahTransaksi += 1;
    }
    if (!isSemuaBulan && laporanBuyerCocokFilterYM_(tgl, prevYear, prevMonth, weekNum)) {
      prevTransaksi += 1; prevGram += qty; prevNilai += nilai; prevBuyerSet[key] = true;
    }
  });

  const allRows = Object.keys(buyerMap).map(function (k) { return buyerMap[k]; }).sort(function (a, b) { return b.gram - a.gram; });
  const totals = {
    transaksi: allRows.reduce(function (s, r) { return s + r.jumlahTransaksi; }, 0),
    buyer: allRows.length,
    gram: allRows.reduce(function (s, r) { return s + r.gram; }, 0),
    nilai: allRows.reduce(function (s, r) { return s + r.nilai; }, 0)
  };
  allRows.forEach(function (r) { r.persenVolume = totals.gram > 0 ? (r.gram / totals.gram * 100) : 0; });

  const buyerRows = buyerFilterKey ? allRows.filter(function (r) { return r.nama.toLowerCase().replace(/\s+/g, ' ') === buyerFilterKey; }) : allRows;

  return {
    buyerRows: buyerRows,
    totals: totals,
    prevTotals: { transaksi: prevTransaksi, buyer: Object.keys(prevBuyerSet).length, gram: prevGram, nilai: prevNilai },
    prevPeriodeLabel: isSemuaBulan ? '' : (NAMA_BULAN_ID[prevMonth - 1] + ' ' + prevYear)
  };
}

/**
 * Baris transaksi individual Trading (PenjualanTransaksi), dipakai
 * BERSAMA oleh modal "Detail Buyer" (getLaporanBuyerTradingDetail, 1
 * buyer) dan export "Download Excel" (exportLaporanBuyerTradingDetailExcel,
 * BISA banyak buyer sekaligus) — supaya angkanya SELALU sama antara
 * dashboard, pop-up, dan Excel (satu sumber filter, tidak ada logic
 * duplikat). nameFilterSet: Set nama (lowercase, spasi dirapikan) yang
 * diizinkan; null/undefined = semua buyer. monthNum juga menerima
 * 'all' (Semua Bulan Januari–Saat Ini), sama seperti getLaporanBuyerTrading.
 */
function getLaporanBuyerTradingDetailRows_(monthNum, weekNum, nameFilterSet) {
  const isSemuaBulan = String(monthNum) === 'all';
  monthNum = isSemuaBulan ? 0 : Number(monthNum);
  weekNum = (!isSemuaBulan && weekNum) ? Number(weekNum) : 0;
  if (!isSemuaBulan && !monthNum) return [];

  const todayStr = formatDateOnly(new Date());
  const nowYear = Number(todayStr.split('-')[0]);
  const jan1Str = nowYear + '-01-01';

  return sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(function (p) {
      const qty = Number(p.Qty) || 0;
      if (qty <= 0) return false;
      const namaAsli = String(p.Pembeli || '').trim();
      if (!namaAsli) return false;
      if (nameFilterSet && !nameFilterSet.has(namaAsli.toLowerCase().replace(/\s+/g, ' '))) return false;
      const tgl = extractDateOnly(p.Tanggal);
      return isSemuaBulan
        ? laporanBuyerCocokFilterSemuaBulan_(tgl, todayStr, jan1Str)
        : laporanBuyerCocokFilterYM_(tgl, nowYear, monthNum, weekNum);
    })
    .map(function (p) {
      return {
        nama: String(p.Pembeli || '').trim(),
        tanggal: extractDateOnly(p.Tanggal),
        nomorInvoice: String(p.NomorInvoice || '-'),
        qty: Number(p.Qty) || 0,
        hargaPerGram: Number(p.HargaPerGram) || 0,
        totalHarga: Number(p.TotalHarga) || 0
      };
    })
    .sort(function (a, b) {
      if (a.nama !== b.nama) return a.nama < b.nama ? -1 : 1;
      return a.tanggal < b.tanggal ? -1 : (a.tanggal > b.tanggal ? 1 : 0);
    });
}

/**
 * Detail transaksi individual Trading utk 1 buyer (dipakai modal
 * "Detail Buyer" di halaman) — TIDAK PERNAH digabung, 1 baris
 * PenjualanTransaksi = 1 baris hasil. Sekarang murni wrapper tipis di
 * atas getLaporanBuyerTradingDetailRows_ (lihat komentar di atasnya).
 */
function getLaporanBuyerTradingDetail(monthNum, weekNum, buyerNama) {
  requireRole('Viewer');
  const targetKey = String(buyerNama || '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (!targetKey) return { rows: [] };
  const rows = getLaporanBuyerTradingDetailRows_(monthNum, weekNum, new Set([targetKey]));
  return { rows: rows };
}

/**
 * "Download Excel" Laporan Buyer Trading — SEBELUMNYA hanya export
 * rekap/summary per buyer. Diganti agar berisi SELURUH detail
 * transaksi (1 transaksi = 1 baris) dari SEMUA buyer yang sedang
 * masuk ke laporan (sesuai filter Bulan/Minggu/Buyer yang aktif),
 * persis seperti data yang dipakai modal "Detail Buyer" — buyerRows
 * dari getLaporanBuyerTrading() (fungsi yang SUDAH dipakai tabel
 * rekap, TIDAK dihitung ulang) menentukan buyer mana yang masuk,
 * lalu getLaporanBuyerTradingDetailRows_ mengambil transaksinya
 * sekaligus (bukan loop per-buyer) supaya angka pasti sama dgn
 * dashboard & tidak membuat logic agregasi baru.
 */
function exportLaporanBuyerTradingDetailExcel(monthNum, weekNum, buyerFilter, periodeLabel) {
  requireRole('Viewer');
  const rekap = getLaporanBuyerTrading(monthNum, weekNum, buyerFilter);
  const nameSet = new Set(rekap.buyerRows.map(function (r) { return r.nama.toLowerCase().replace(/\s+/g, ' '); }));
  const detailRows = getLaporanBuyerTradingDetailRows_(monthNum, weekNum, nameSet);
  const rows = detailRows.map(function (r) {
    return { nama: r.nama, tanggal: r.tanggal, noTransaksi: r.nomorInvoice, gram: r.qty, hargaPerGram: r.hargaPerGram, totalHarga: r.totalHarga };
  });
  const filenameBase = 'Detail_Laporan_Buyer_' + laporanBuyerPeriodeFileSlug_(monthNum, weekNum);
  return exportDetailLaporanToXlsx_('Detail Laporan Buyer', {
    judul: 'DETAIL LAPORAN BUYER TRADING',
    periodeLabel: periodeLabel,
    namaKolomEntitas: 'Nama Buyer',
    rows: rows
  }, filenameBase);
}

function laporanBuyerPeriodeFileSlug_(monthNum, weekNum) {
  const tahun = Number(formatDateOnly(new Date()).split('-')[0]);
  if (String(monthNum) === 'all') {
    const bulanIniSingkat = NAMA_BULAN_ID[new Date().getMonth()].slice(0, 3);
    return 'Jan-' + bulanIniSingkat + '_' + tahun;
  }
  const bulanNama = NAMA_BULAN_ID[Number(monthNum) - 1] || '';
  const mingguSuffix = Number(weekNum) ? ('-M' + Number(weekNum)) : '';
  return bulanNama + mingguSuffix + '_' + tahun;
}

/**
 * Tren gramasi Trading (Kg) — mode 'bulanan' (Januari-Desember tahun
 * berjalan) atau 'mingguan' (Minggu 1-5 dari monthNum yang dipilih,
 * tahun berjalan). Sumber sama, PenjualanTransaksi, cuma diagregasi per
 * bulan/minggu (tidak mengubah data asli). buyerFilter (opsional) —
 * kalau diisi, SELURUH agregasi (array tren + perbandingan MtM/WtW)
 * cuma menghitung transaksi buyer itu. weekNum (opsional, mode
 * mingguan saja) — "minggu aktif" utk WtW; kosong = otomatis pakai
 * minggu terakhir yg ada datanya di bulan tsb.
 * Include jg "compare" (MtM/bulanan atau WtW/mingguan, sudah
 * memperhitungkan pergantian tahun & lintas-bulan utk Minggu 1) dan
 * "prevBoundaryGram/Label" (dipakai tooltip titik pertama chart, yg
 * pembandingnya di luar array gramKg — Des tahun lalu utk Januari,
 * minggu terakhir bulan sebelumnya utk Minggu 1).
 *
 * monthNum juga menerima 'all' (Semua Bulan Januari–Saat Ini) — mode
 * dipaksa 'bulanan' (array gramKg tetap Jan-Des tahun berjalan seperti
 * biasa, sudah mencakup seluruh transaksi s.d. bulan berjalan karena
 * bulan-bulan setelah hari ini otomatis 0/belum ada data), dan bulan
 * berjalan dipakai sebagai acuan kartu perbandingan MtM (tidak ada 1
 * bulan spesifik yang relevan untuk rentang kumulatif).
 */
function getLaporanBuyerTradingTrend(monthNum, mode, buyerFilter, weekNum) {
  requireRole('Viewer');
  const isSemuaBulan = String(monthNum) === 'all';
  mode = (!isSemuaBulan && mode === 'mingguan') ? 'mingguan' : 'bulanan';
  monthNum = isSemuaBulan ? (new Date().getMonth() + 1) : (Number(monthNum) || (new Date().getMonth() + 1));
  weekNum = (!isSemuaBulan && weekNum) ? Number(weekNum) : 0;
  const nowYear = Number(formatDateOnly(new Date()).split('-')[0]);
  const all = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const buyerFilterKey = String(buyerFilter || '').trim().toLowerCase().replace(/\s+/g, ' ');

  function cocokBuyer(p) {
    if (!buyerFilterKey) return true;
    const namaAsli = String(p.Pembeli || '').trim().toLowerCase().replace(/\s+/g, ' ');
    return namaAsli === buyerFilterKey;
  }
  function sumFor(year, month, week) {
    let sum = 0;
    all.forEach(function (p) {
      const qty = Number(p.Qty) || 0;
      if (qty <= 0 || !cocokBuyer(p)) return;
      const tgl = extractDateOnly(p.Tanggal);
      if (!laporanBuyerCocokFilterYM_(tgl, year, month, week)) return;
      sum += qty;
    });
    return sum;
  }

  let result;
  if (mode === 'bulanan') {
    const gram = new Array(12).fill(0);
    const hasData = new Array(12).fill(false);
    all.forEach(function (p) {
      const qty = Number(p.Qty) || 0;
      if (qty <= 0 || !cocokBuyer(p)) return;
      const tgl = extractDateOnly(p.Tanggal);
      if (!tgl) return;
      const parts = tgl.split('-');
      const y = Number(parts[0]), m = Number(parts[1]);
      if (y !== nowYear) return;
      gram[m - 1] += qty; hasData[m - 1] = true;
    });

    let prevYear = nowYear, prevMonth = monthNum - 1;
    if (prevMonth < 1) { prevMonth = 12; prevYear = nowYear - 1; }
    const aktifGram = sumFor(nowYear, monthNum, 0);
    const prevGram = sumFor(prevYear, prevMonth, 0);

    result = {
      mode: 'bulanan', labels: NAMA_BULAN_ID, gramKg: gram.map(function (g) { return g / 1000; }), hasData: hasData,
      bulanAktifNum: monthNum, tahunAktif: nowYear,
      prevBoundaryGram: prevGram / 1000, prevBoundaryLabel: NAMA_BULAN_ID[prevMonth - 1] + ' ' + prevYear,
      compare: {
        periodType: 'MtM', aktifLabel: NAMA_BULAN_ID[monthNum - 1] + ' ' + nowYear,
        prevLabel: NAMA_BULAN_ID[prevMonth - 1] + ' ' + prevYear,
        aktifGram: aktifGram, prevGram: prevGram
      }
    };
  } else {
    const gram = [0, 0, 0, 0, 0];
    const hasData = [false, false, false, false, false];
    all.forEach(function (p) {
      const qty = Number(p.Qty) || 0;
      if (qty <= 0 || !cocokBuyer(p)) return;
      const tgl = extractDateOnly(p.Tanggal);
      if (!tgl) return;
      const parts = tgl.split('-');
      const y = Number(parts[0]), m = Number(parts[1]), d = Number(parts[2]);
      if (y !== nowYear || m !== monthNum) return;
      const w = laporanBuyerMingguKe_(d);
      gram[w - 1] += qty; hasData[w - 1] = true;
    });

    let wAktif = weekNum;
    if (!wAktif) {
      for (let w = 5; w >= 1; w--) { if (hasData[w - 1]) { wAktif = w; break; } }
      if (!wAktif) wAktif = 1;
    }

    let prevYear = nowYear, prevMonth = monthNum, prevWeek = wAktif - 1, prevLabel;
    if (prevWeek < 1) {
      prevMonth = monthNum - 1;
      if (prevMonth < 1) { prevMonth = 12; prevYear = nowYear - 1; }
      const lastDayPrevMonth = new Date(prevYear, prevMonth, 0).getDate();
      prevWeek = laporanBuyerMingguKe_(lastDayPrevMonth);
      prevLabel = 'Minggu ' + prevWeek + ' — ' + NAMA_BULAN_ID[prevMonth - 1] + ' ' + prevYear;
    } else {
      prevLabel = 'Minggu ' + prevWeek;
    }
    const aktifGram = sumFor(nowYear, monthNum, wAktif);
    const prevGram = sumFor(prevYear, prevMonth, prevWeek);

    // Boundary Minggu 1 (titik pertama chart) — beda dari "prev" di atas
    // kalau wAktif bukan 1 (compare card & boundary titik-1 chart adalah
    // dua hal yg beda konteksnya, meski rumusnya sama saat wAktif = 1).
    const lastDayThisPrevMonth = new Date(prevYear, prevMonth, 0).getDate();
    let boundaryMonth = monthNum - 1, boundaryYear = nowYear;
    if (boundaryMonth < 1) { boundaryMonth = 12; boundaryYear = nowYear - 1; }
    const boundaryLastDay = new Date(boundaryYear, boundaryMonth, 0).getDate();
    const boundaryWeek = laporanBuyerMingguKe_(boundaryLastDay);
    const boundaryGram = sumFor(boundaryYear, boundaryMonth, boundaryWeek);

    result = {
      mode: 'mingguan', labels: ['Minggu 1', 'Minggu 2', 'Minggu 3', 'Minggu 4', 'Minggu 5'],
      gramKg: gram.map(function (g) { return g / 1000; }), hasData: hasData,
      bulanAktifNum: monthNum, tahunAktif: nowYear,
      prevBoundaryGram: boundaryGram / 1000, prevBoundaryLabel: 'Minggu ' + boundaryWeek + ' — ' + NAMA_BULAN_ID[boundaryMonth - 1] + ' ' + boundaryYear,
      compare: {
        periodType: 'WtW', aktifLabel: 'Minggu ' + wAktif + ' — ' + NAMA_BULAN_ID[monthNum - 1] + ' ' + nowYear,
        prevLabel: prevLabel, aktifGram: aktifGram, prevGram: prevGram, weekNumAktif: wAktif
      }
    };
  }

  const c = result.compare;
  c.selisihGram = c.aktifGram - c.prevGram;
  c.persen = c.prevGram > 0 ? ((c.aktifGram - c.prevGram) / c.prevGram) * 100 : null;
  c.isBaru = c.prevGram === 0 && c.aktifGram > 0;
  c.isZero = c.prevGram === 0 && c.aktifGram === 0;
  return result;
}

// ------------------------------------------------------------
// LAPORAN SUPPLIER BULANAN — rekap transaksi pembelian TRADING per
// supplier (sama pola dengan Laporan Buyer Trading, tapi sumbernya
// PembelianTransaksi). Filter Tahun+Bulan (bukan Minggu — sesuai
// permintaan). "Trading" dibedakan dari ETF/EGR pakai pola yang SUDAH
// ADA (getValidPembelianSourcePOs_ — cek Nomor PO masih match
// F.2.Input_Trans_Beli saat ini), sama seperti getDashboardKategoriBreakdown.
// Data transaksi asli TIDAK diubah — penggabungan per supplier hanya
// terjadi di hasil laporan ini.
// ------------------------------------------------------------

/**
 * Ambil baris PembelianTransaksi yang valid & masih Trading (bukan
 * ETF/EGR), reuse getValidPembelianSourcePOs_ yang sudah ada (dipakai
 * juga oleh getDashboardKategoriBreakdown).
 */
function lsbAmbilPembelianTrading_() {
  const validPOs = getValidPembelianSourcePOs_();
  return sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL).filter(function (p) {
    if (Number(p.TotalHarga || 0) <= 0) return false;
    if (!validPOs) return true; // gagal baca sumber -> jangan filter apa-apa (aman)
    return !!validPOs[String(p.NomorPO || '').trim()];
  });
}

/**
 * bulan=0 (atau kosong) berarti "Semua Bulan" -- cocok cuma
 * berdasarkan Tahun, tidak mempersempit ke bulan tertentu.
 */
function lsbCocokFilter_(tglStr, tahun, bulan) {
  if (!tglStr) return false;
  const p = tglStr.split('-');
  if (Number(p[0]) !== tahun) return false;
  return !bulan || Number(p[1]) === bulan;
}

/**
 * Laporan Supplier Bulanan — rekap per Seller (supplier), untuk
 * Tahun+Bulan yang dipilih (bulan=0 = Semua Bulan/satu tahun penuh).
 * supplierFilter (opsional) mempersempit supplierRows ke 1 supplier
 * saja, tapi `totals` tetap dari SELURUH supplier periode itu (sama
 * seperti pola Laporan Buyer Trading).
 * searchTerm (opsional) mencari nama supplier ATAU Nomor PO/Invoice
 * transaksi individual milik supplier tsb.
 * Average price pakai metode yang SUDAH ADA di getTopSuppliers:
 * rata-rata AfterDiskon per transaksi (bukan nilai/gram).
 */
function getLaporanSupplierBulanan(tahun, bulan, supplierFilter, searchTerm) {
  requireRole('Viewer');
  tahun = Number(tahun);
  bulan = Number(bulan) || 0;
  if (!tahun) throw new Error('Tahun wajib dipilih.');
  const supplierFilterKey = String(supplierFilter || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const search = String(searchTerm || '').trim().toLowerCase();

  const supplierMap = {};
  lsbAmbilPembelianTrading_().forEach(function (p) {
    const tgl = extractDateOnly(p.Tanggal);
    if (!lsbCocokFilter_(tgl, tahun, bulan)) return;
    const qty = Number(p.Qty) || 0;
    if (qty <= 0) return;
    const namaAsli = String(p.Seller || '').trim();
    if (!namaAsli) return;
    const key = namaAsli.toLowerCase().replace(/\s+/g, ' ');

    if (!supplierMap[key]) {
      supplierMap[key] = { nama: namaAsli, gram: 0, nilai: 0, jumlahTransaksi: 0, sumHargaTransaksi: 0, nomorList: [] };
    }
    const s = supplierMap[key];
    s.gram += qty;
    s.nilai += Number(p.TotalHarga) || 0;
    s.jumlahTransaksi += 1;
    s.sumHargaTransaksi += Number(p.AfterDiskon) || 0;
    s.nomorList.push(String(p.NomorPO || '') + ' ' + String(p.InvoiceRef || ''));
  });

  const allRows = Object.keys(supplierMap).map(function (k) { return supplierMap[k]; })
    .map(function (s) {
      return {
        nama: s.nama,
        gram: s.gram,
        nilai: s.nilai,
        jumlahTransaksi: s.jumlahTransaksi,
        avgPrice: s.jumlahTransaksi ? (s.sumHargaTransaksi / s.jumlahTransaksi) : 0,
        _nomorGabungan: s.nomorList.join(' | ').toLowerCase()
      };
    })
    .sort(function (a, b) { return b.nilai - a.nilai; }); // urut nominal terbesar, sesuai instruksi

  const totals = {
    supplierAktif: allRows.length,
    transaksi: allRows.reduce(function (s, r) { return s + r.jumlahTransaksi; }, 0),
    gram: allRows.reduce(function (s, r) { return s + r.gram; }, 0),
    nilai: allRows.reduce(function (s, r) { return s + r.nilai; }, 0)
  };
  const totalSumHarga = allRows.reduce(function (s, r) { return s + (r.avgPrice * r.jumlahTransaksi); }, 0);
  totals.avgPrice = totals.transaksi ? (totalSumHarga / totals.transaksi) : 0;

  allRows.forEach(function (r) { r.persenKontribusi = totals.nilai > 0 ? (r.nilai / totals.nilai * 100) : 0; });

  // Top 5 berdasarkan GRAMASI (beda dari urutan tabel utama yang
  // berdasarkan nominal) — SELALU dari seluruh supplier periode ini,
  // tidak ikut menyempit walau filter Supplier/Cari sedang aktif (sama
  // seperti pola Top 5 Buyer di Laporan Buyer Trading).
  const top5ByGramasi = allRows.slice().sort(function (a, b) { return b.gram - a.gram; }).slice(0, 5)
    .map(function (r) { return { nama: r.nama, gram: r.gram, nilai: r.nilai, jumlahTransaksi: r.jumlahTransaksi, avgPrice: r.avgPrice, persenKontribusi: r.persenKontribusi }; });

  let supplierRows = supplierFilterKey
    ? allRows.filter(function (r) { return r.nama.toLowerCase().replace(/\s+/g, ' ') === supplierFilterKey; })
    : allRows;

  if (search) {
    supplierRows = supplierRows.filter(function (r) {
      return r.nama.toLowerCase().indexOf(search) !== -1 || r._nomorGabungan.indexOf(search) !== -1;
    });
  }
  supplierRows.forEach(function (r) { delete r._nomorGabungan; });

  return { supplierRows: supplierRows, totals: totals, tahun: tahun, bulan: bulan, top5ByGramasi: top5ByGramasi };
}

/**
 * Tren gramasi pembelian Trading (Kg) per bulan (Januari-Desember)
 * untuk 1 tahun penuh — dipakai chart "Tren Gramasi Pembelian Trading"
 * di Laporan Supplier Bulanan. supplierFilter (opsional) mempersempit
 * ke 1 supplier saja; kalau kosong, total seluruh supplier. Sumbernya
 * SAMA seperti getLaporanSupplierBulanan (PembelianTransaksi, hanya
 * baris Trading — lsbAmbilPembelianTrading_), jadi tidak ada transaksi
 * penjualan/EGR yang ikut masuk.
 */
function getLaporanSupplierBulananTrend(tahun, supplierFilter) {
  requireRole('Viewer');
  tahun = Number(tahun);
  if (!tahun) throw new Error('Tahun wajib dipilih.');
  const supplierFilterKey = String(supplierFilter || '').trim().toLowerCase().replace(/\s+/g, ' ');

  const gram = new Array(12).fill(0);
  const nilai = new Array(12).fill(0);
  const jumlahTransaksi = new Array(12).fill(0);

  lsbAmbilPembelianTrading_().forEach(function (p) {
    const qty = Number(p.Qty) || 0;
    if (qty <= 0) return;
    const namaAsli = String(p.Seller || '').trim();
    if (!namaAsli) return;
    if (supplierFilterKey && namaAsli.toLowerCase().replace(/\s+/g, ' ') !== supplierFilterKey) return;
    const tgl = extractDateOnly(p.Tanggal);
    if (!tgl) return;
    const parts = tgl.split('-');
    const y = Number(parts[0]), m = Number(parts[1]);
    if (y !== tahun) return;

    gram[m - 1] += qty;
    nilai[m - 1] += Number(p.TotalHarga) || 0;
    jumlahTransaksi[m - 1] += 1;
  });

  return {
    labels: NAMA_BULAN_ID,
    gramKg: gram.map(function (g) { return g / 1000; }),
    nilai: nilai,
    jumlahTransaksi: jumlahTransaksi
  };
}

/**
 * Baris transaksi individual pembelian Trading (PembelianTransaksi),
 * dipakai BERSAMA oleh modal "Detail Supplier"
 * (getLaporanSupplierBulananDetail, 1 supplier) dan export "Download
 * Excel" (exportLaporanSupplierBulananDetailExcel, BISA banyak
 * supplier sekaligus) — satu sumber filter, tidak ada logic duplikat.
 * nameFilterSet: Set nama (lowercase, spasi dirapikan) yang
 * diizinkan; null/undefined = semua supplier. bulan=0 = Semua Bulan
 * (satu tahun penuh), sama seperti getLaporanSupplierBulanan.
 */
function getLaporanSupplierBulananDetailRows_(tahun, bulan, nameFilterSet) {
  tahun = Number(tahun);
  bulan = Number(bulan) || 0;
  if (!tahun) return [];

  return lsbAmbilPembelianTrading_()
    .filter(function (p) {
      const qty = Number(p.Qty) || 0;
      if (qty <= 0) return false;
      const namaAsli = String(p.Seller || '').trim();
      if (!namaAsli) return false;
      if (nameFilterSet && !nameFilterSet.has(namaAsli.toLowerCase().replace(/\s+/g, ' '))) return false;
      const tgl = extractDateOnly(p.Tanggal);
      return lsbCocokFilter_(tgl, tahun, bulan);
    })
    .map(function (p) {
      return {
        nama: String(p.Seller || '').trim(),
        tanggal: extractDateOnly(p.Tanggal),
        nomorPO: String(p.NomorPO || '-'),
        invoiceRef: String(p.InvoiceRef || ''),
        qty: Number(p.Qty) || 0,
        hargaPerGram: Number(p.AfterDiskon) || 0,
        totalHarga: Number(p.TotalHarga) || 0,
        catatan: String(p.Catatan || '')
      };
    })
    .sort(function (a, b) {
      if (a.nama !== b.nama) return a.nama < b.nama ? -1 : 1;
      return a.tanggal < b.tanggal ? -1 : (a.tanggal > b.tanggal ? 1 : 0);
    });
}

/**
 * Detail transaksi individual pembelian Trading utk 1 supplier
 * (dipakai modal "Detail Supplier" di halaman) — TIDAK PERNAH
 * digabung, 1 baris PembelianTransaksi = 1 baris hasil. Sekarang
 * murni wrapper tipis di atas getLaporanSupplierBulananDetailRows_
 * (lihat komentar di atasnya); field hasil TETAP SAMA persis seperti
 * sebelum refactor supaya modal yang sudah ada tidak perlu diubah.
 */
function getLaporanSupplierBulananDetail(tahun, bulan, supplierNama) {
  requireRole('Viewer');
  const targetKey = String(supplierNama || '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (!Number(tahun) || !targetKey) return { rows: [] };
  const rows = getLaporanSupplierBulananDetailRows_(tahun, bulan, new Set([targetKey]))
    .map(function (r) {
      return {
        tanggal: r.tanggal, nomorPO: r.nomorPO, invoiceRef: r.invoiceRef, seller: r.nama,
        qty: r.qty, hargaPerGram: r.hargaPerGram, totalHarga: r.totalHarga, catatan: r.catatan
      };
    });
  return { rows: rows };
}

/**
 * "Download Excel" Laporan Supplier Bulanan — menu ini SEBELUMNYA
 * tidak punya tombol Excel sama sekali (hanya PDF/Gambar/Copy Text).
 * Ditambahkan agar berisi SELURUH detail transaksi (1 transaksi = 1
 * baris) dari SEMUA supplier yang sedang masuk ke laporan (sesuai
 * filter Tahun/Bulan/Supplier/Cari yang aktif), persis seperti data
 * yang dipakai modal "Detail Supplier" — supplierRows dari
 * getLaporanSupplierBulanan() (fungsi yang SUDAH dipakai tabel rekap,
 * termasuk logic filter "Cari" yang TIDAK diduplikasi di sini)
 * menentukan supplier mana yang masuk, lalu
 * getLaporanSupplierBulananDetailRows_ mengambil transaksinya
 * sekaligus (bukan loop per-supplier).
 */
function exportLaporanSupplierBulananDetailExcel(tahun, bulan, supplierFilter, searchTerm, periodeLabel) {
  requireRole('Viewer');
  const rekap = getLaporanSupplierBulanan(tahun, bulan, supplierFilter, searchTerm);
  const nameSet = new Set(rekap.supplierRows.map(function (r) { return r.nama.toLowerCase().replace(/\s+/g, ' '); }));
  const detailRows = getLaporanSupplierBulananDetailRows_(tahun, bulan, nameSet);
  const rows = detailRows.map(function (r) {
    const noTransaksi = (r.nomorPO || '-') + (r.invoiceRef ? ' / ' + r.invoiceRef : '');
    return { nama: r.nama, tanggal: r.tanggal, noTransaksi: noTransaksi, gram: r.qty, hargaPerGram: r.hargaPerGram, totalHarga: r.totalHarga };
  });
  const filenameBase = 'Detail_Laporan_Supplier_' + laporanSupplierPeriodeFileSlug_(tahun, bulan);
  return exportDetailLaporanToXlsx_('Detail Laporan Supplier', {
    judul: 'DETAIL LAPORAN SUPPLIER TRADING',
    periodeLabel: periodeLabel,
    namaKolomEntitas: 'Nama Supplier',
    rows: rows
  }, filenameBase);
}

function laporanSupplierPeriodeFileSlug_(tahun, bulan) {
  bulan = Number(bulan) || 0;
  if (!bulan) return 'Jan-Des_' + tahun;
  const bulanNama = NAMA_BULAN_ID[bulan - 1] || '';
  return bulanNama + '_' + tahun;
}

/**
 * Laporan Buyer — EGR. Baca Input_Jual_EGR (spreadsheet EGR eksternal,
 * PRICE_DISCOVERY_SOURCE_ID) dengan pola deteksi kolom dinamis yang sama
 * dengan getEgrRekapData, kelompokkan per Pembeli untuk Bulan (+
 * opsional Minggu ke-) yang dipilih, tahun berjalan.
 */
function getLaporanBuyerEgr(monthNum, weekNum, buyerFilter) {
  requireRole('Viewer');
  monthNum = Number(monthNum);
  weekNum = weekNum ? Number(weekNum) : 0;
  if (!monthNum) throw new Error('Bulan wajib dipilih.');
  const buyerFilterKey = String(buyerFilter || '').trim().toLowerCase().replace(/\s+/g, ' ');

  const nowYear = Number(formatDateOnly(new Date()).split('-')[0]);
  let prevYear = nowYear, prevMonth = monthNum - 1;
  if (prevMonth < 1) { prevMonth = 12; prevYear = nowYear - 1; }

  const buyerMap = {};
  let prevTransaksi = 0, prevGram = 0, prevNilai = 0;
  const prevBuyerSet = {};

  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const jualSheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (jualSheet) {
    const values = jualSheet.getDataRange().getValues();
    const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
    if (headerRowIdx !== -1) {
      const headerMap = buildHeaderMap(values[headerRowIdx]);
      const col = {
        tanggal: findColByKeywords(headerMap, ['tanggal']),
        pembeli: findColByKeywords(headerMap, ['pembeli']),
        qty: findColByKeywords(headerMap, ['qty']),
        totalAfterTax: findColByKeywords(headerMap, ['after', 'tax']),
        totalHarga: findColByKeywords(headerMap, ['total', 'harga'], ['after', 'tax'])
      };
      values.slice(headerRowIdx + 1).forEach(function (row) {
        const qty = col.qty !== -1 ? Number(getByCol(row, col.qty)) || 0 : 0;
        if (qty <= 0) return;
        const namaAsli = col.pembeli !== -1 ? String(getByCol(row, col.pembeli) || '').trim() : '';
        if (!namaAsli) return;
        const tgl = extractDateOnly(getByCol(row, col.tanggal));
        const key = namaAsli.toLowerCase().replace(/\s+/g, ' ');

        let nilai = col.totalAfterTax !== -1 ? Number(getByCol(row, col.totalAfterTax)) || 0 : 0;
        if (!nilai && col.totalHarga !== -1) nilai = Number(getByCol(row, col.totalHarga)) || 0;

        if (laporanBuyerCocokFilterYM_(tgl, nowYear, monthNum, weekNum)) {
          if (!buyerMap[key]) buyerMap[key] = { nama: namaAsli, gram: 0, nilai: 0, jumlahTransaksi: 0 };
          buyerMap[key].gram += qty;
          buyerMap[key].nilai += nilai;
          buyerMap[key].jumlahTransaksi += 1;
        }
        if (laporanBuyerCocokFilterYM_(tgl, prevYear, prevMonth, weekNum)) {
          prevTransaksi += 1; prevGram += qty; prevNilai += nilai; prevBuyerSet[key] = true;
        }
      });
    }
  }

  const allRows = Object.keys(buyerMap).map(function (k) { return buyerMap[k]; }).sort(function (a, b) { return b.gram - a.gram; });
  const totals = {
    transaksi: allRows.reduce(function (s, r) { return s + r.jumlahTransaksi; }, 0),
    buyer: allRows.length,
    gram: allRows.reduce(function (s, r) { return s + r.gram; }, 0),
    nilai: allRows.reduce(function (s, r) { return s + r.nilai; }, 0)
  };
  allRows.forEach(function (r) { r.persenVolume = totals.gram > 0 ? (r.gram / totals.gram * 100) : 0; });

  const buyerRows = buyerFilterKey ? allRows.filter(function (r) { return r.nama.toLowerCase().replace(/\s+/g, ' ') === buyerFilterKey; }) : allRows;

  return {
    buyerRows: buyerRows,
    totals: totals,
    prevTotals: { transaksi: prevTransaksi, buyer: Object.keys(prevBuyerSet).length, gram: prevGram, nilai: prevNilai },
    prevPeriodeLabel: NAMA_BULAN_ID[prevMonth - 1] + ' ' + prevYear
  };
}

/**
 * Detail transaksi individual EGR utk 1 buyer, dari Input_Jual_EGR,
 * mengikuti filter Bulan+Minggu yang sama seperti getLaporanBuyerEgr.
 * TIDAK PERNAH digabung — 1 baris sheet sumber = 1 baris hasil di sini.
 * Pola deteksi kolom (termasuk NomorInvoice & HargaPerGram) sama
 * seperti readEgrJualRowsForDate_.
 */
function getLaporanBuyerEgrDetail(monthNum, weekNum, buyerNama) {
  requireRole('Viewer');
  monthNum = Number(monthNum);
  weekNum = weekNum ? Number(weekNum) : 0;
  const targetKey = String(buyerNama || '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (!monthNum || !targetKey) return { rows: [] };

  const nowYear = Number(formatDateOnly(new Date()).split('-')[0]);
  const rows = [];

  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const jualSheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (jualSheet) {
    const values = jualSheet.getDataRange().getValues();
    const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
    if (headerRowIdx !== -1) {
      const headerMap = buildHeaderMap(values[headerRowIdx]);
      const col = {
        tanggal: findColByKeywords(headerMap, ['tanggal']),
        // Header di Input_Subs_EGR = "Nomor TC" (nomor transaksi EGR);
        // "Nomor Invoice" tetap dicoba duluan utk kompatibilitas sheet lama.
        invoice: findColByKeywordsMulti_(headerMap, [['nomor', 'invoice'], ['nomor', 'tc']]),
        hargaPerGram: findColByKeywords(headerMap, ['harga'], ['total', 'after']),
        qty: findColByKeywords(headerMap, ['qty']),
        totalAfterTax: findColByKeywords(headerMap, ['after', 'tax']),
        totalHarga: findColByKeywords(headerMap, ['total', 'harga'], ['after', 'tax']),
        pembeli: findColByKeywords(headerMap, ['pembeli'])
      };
      values.slice(headerRowIdx + 1).forEach(function (row) {
        const qty = col.qty !== -1 ? Number(getByCol(row, col.qty)) || 0 : 0;
        if (qty <= 0) return;
        const namaAsli = col.pembeli !== -1 ? String(getByCol(row, col.pembeli) || '').trim() : '';
        if (!namaAsli || namaAsli.toLowerCase().replace(/\s+/g, ' ') !== targetKey) return;
        const tgl = extractDateOnly(getByCol(row, col.tanggal));
        if (!laporanBuyerCocokFilterYM_(tgl, nowYear, monthNum, weekNum)) return;

        let nilai = col.totalAfterTax !== -1 ? Number(getByCol(row, col.totalAfterTax)) || 0 : 0;
        if (!nilai && col.totalHarga !== -1) nilai = Number(getByCol(row, col.totalHarga)) || 0;

        rows.push({
          tanggal: tgl,
          nomorInvoice: col.invoice !== -1 ? String(getByCol(row, col.invoice) || '-') : '-',
          qty: qty,
          hargaPerGram: col.hargaPerGram !== -1 ? (Number(getByCol(row, col.hargaPerGram)) || 0) : 0,
          totalHarga: nilai
        });
      });
    }
  }

  rows.sort(function (a, b) { return a.tanggal < b.tanggal ? -1 : (a.tanggal > b.tanggal ? 1 : 0); });
  return { rows: rows };
}

/**
 * Tren gramasi EGR (Kg) dari Input_Jual_EGR — mode 'bulanan'
 * (Januari-Desember tahun berjalan) atau 'mingguan' (Minggu 1-5 dari
 * monthNum yang dipilih, tahun berjalan). Sama seperti
 * getLaporanBuyerTradingTrend, cuma sumbernya Input_Jual_EGR.
 */
function getLaporanBuyerEgrTrend(monthNum, mode) {
  requireRole('Viewer');
  mode = (mode === 'mingguan') ? 'mingguan' : 'bulanan';
  monthNum = Number(monthNum) || (new Date().getMonth() + 1);
  const nowYear = Number(formatDateOnly(new Date()).split('-')[0]);

  const gramBulanan = new Array(12).fill(0);
  const hasDataBulanan = new Array(12).fill(false);
  const gramMingguan = [0, 0, 0, 0, 0];
  const hasDataMingguan = [false, false, false, false, false];

  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const jualSheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (jualSheet) {
    const values = jualSheet.getDataRange().getValues();
    const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
    if (headerRowIdx !== -1) {
      const headerMap = buildHeaderMap(values[headerRowIdx]);
      const col = {
        tanggal: findColByKeywords(headerMap, ['tanggal']),
        qty: findColByKeywords(headerMap, ['qty'])
      };
      values.slice(headerRowIdx + 1).forEach(function (row) {
        const qty = col.qty !== -1 ? Number(getByCol(row, col.qty)) || 0 : 0;
        if (qty <= 0) return;
        const tgl = extractDateOnly(getByCol(row, col.tanggal));
        if (!tgl) return;
        const parts = tgl.split('-');
        const y = Number(parts[0]), m = Number(parts[1]), d = Number(parts[2]);
        if (y !== nowYear) return;
        gramBulanan[m - 1] += qty; hasDataBulanan[m - 1] = true;
        if (m === monthNum) {
          const w = laporanBuyerMingguKe_(d);
          gramMingguan[w - 1] += qty; hasDataMingguan[w - 1] = true;
        }
      });
    }
  }

  if (mode === 'bulanan') {
    return { mode: 'bulanan', labels: NAMA_BULAN_ID, gramKg: gramBulanan.map(function (g) { return g / 1000; }), hasData: hasDataBulanan };
  }
  return { mode: 'mingguan', labels: ['Minggu 1', 'Minggu 2', 'Minggu 3', 'Minggu 4', 'Minggu 5'], gramKg: gramMingguan.map(function (g) { return g / 1000; }), hasData: hasDataMingguan };
}

/**
 * Summary Teks — Laporan Rekap EGR. SELALU baca langsung dari
 * Input_Jual_EGR berdasarkan tanggal laporan (= Tanggal Akhir yang
 * dipilih di filter), TIDAK PERNAH memakai snapshot sesi 12.00/16.00
 * (fitur snapshot itu KHUSUS Laporan Harian EGR, berdiri sendiri, tidak
 * disentuh oleh fungsi ini sama sekali). Tidak ada pembatasan jam —
 * kalau tanggal akhir = hari ini, dipakai seluruh transaksi yang sudah
 * ada di sheet saat fungsi ini dipanggil; kalau tanggal lampau, dipakai
 * seluruh transaksi pada tanggal itu.
 */
function getEgrRekapSummaryData(endDateStr) {
  requireRole('Viewer');
  const tgl = extractDateOnly(endDateStr) || formatDateOnly(new Date());
  const tanggalLabel = formatTanggalPanjangIndo_(tgl);

  const result = {
    tanggalStr: tgl, tanggalLabel: tanggalLabel,
    hargaJualTerakhir: null, hargaBuybackTerakhir: null,
    totalGramHariIni: 0, buyerRows: [], totalGramKumulatif: 0
  };

  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (!sheet) return result;

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return result;
  const headerRow = values[headerRowIdx];
  const headerMap = buildHeaderMap(headerRow);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    qty: findColByKeywords(headerMap, ['qty']),
    pembeli: findColByKeywords(headerMap, ['pembeli']),
    hargaPerGram: findColByKeywords(headerMap, ['harga'], ['total', 'after', 'buyback']),
    buyback: findLastColContaining_(headerRow, 'buyback')
  };
  if (col.tanggal === -1) return result;

  const buyerMap = {};
  values.slice(headerRowIdx + 1).forEach(function (row) {
    const tglRow = extractDateOnly(getByCol(row, col.tanggal));
    if (!tglRow) return;

    const qty = col.qty !== -1 ? (parseAngkaSel_(getByCol(row, col.qty)) || 0) : 0;
    if (qty <= 0) return; // abaikan baris kosong/subtotal/total

    // Total kumulatif: SELURUH Qty (Gram) dari awal data s.d. tanggal
    // laporan (inklusif) — bukan hanya tanggal laporan itu sendiri.
    if (tglRow <= tgl) result.totalGramKumulatif += qty;

    if (tglRow !== tgl) return;

    result.totalGramHariIni += qty;

    const namaAsli = col.pembeli !== -1 ? String(getByCol(row, col.pembeli) || '').trim() : '';
    if (namaAsli) {
      const key = namaAsli.toLowerCase().replace(/\s+/g, ' ');
      if (!buyerMap[key]) buyerMap[key] = { nama: namaAsli, gram: 0 };
      buyerMap[key].gram += qty;
    }

    // Harga Jual/Buyback TERAKHIR pada tanggal laporan (baris terakhir yg
    // terisi sesuai urutan sheet, bukan dijumlah/dirata-rata).
    const h = col.hargaPerGram !== -1 ? parseAngkaSel_(getByCol(row, col.hargaPerGram)) : null;
    if (h !== null && h !== 0) result.hargaJualTerakhir = h;
    const b = col.buyback !== -1 ? parseAngkaSel_(getByCol(row, col.buyback)) : null;
    if (b !== null && b !== 0) result.hargaBuybackTerakhir = b;
  });

  result.buyerRows = Object.keys(buyerMap).map(function (k) { return buyerMap[k]; })
    .sort(function (a, b) { return b.gram - a.gram; });

  return result;
}

/**
 * Transaksi Hari Ini — kartu ringkas paling atas Laporan Rekap EGR.
 * SELALU membaca tanggal hari ini (Session.getScriptTimeZone() =
 * Asia/Jakarta, sesuai appsscript.json), independen dari filter rentang
 * tanggal laporan (start/end/buyer) di getEgrRekapData di atas. Sumber:
 * Input_Jual_EGR saja — Qty (Gram) & Total harga after tax (fallback
 * Total Harga), baris kosong/subtotal/total diabaikan (qty <= 0). TIDAK
 * menyentuh Input_Beli_EGR, Top Up EGR, maupun Harga Buyback.
 */
// reportDateStr opsional = Tanggal Akhir filter halaman ("Hari Ini" harus
// mengikuti tanggal laporan, BUKAN tanggal server/browser — itu penyebab
// card ini sering tampil 0 padahal Summary Teks & Rekap Buyer di halaman
// yang sama sudah benar membaca data tanggal tsb). Sumber sheet, kolom, dan
// filtering SAMA PERSIS dengan yang sudah dipakai getEgrRekapSummaryData/
// getEgrRekapData — tidak ada logic/data baru, hanya parameter tanggalnya.
function getEgrTransaksiHariIni(reportDateStr) {
  requireRole('Viewer');
  const todayStr = reportDateStr ? extractDateOnly(reportDateStr) : formatDateOnly(new Date());
  const result = { tanggalStr: todayStr, gram: 0, nilai: 0 };

  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (!sheet) return result;

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return result;
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    qty: findColByKeywords(headerMap, ['qty']),
    totalAfterTax: findColByKeywords(headerMap, ['after', 'tax']),
    totalHarga: findColByKeywords(headerMap, ['total', 'harga'], ['after', 'tax'])
  };
  if (col.tanggal === -1) return result;

  values.slice(headerRowIdx + 1).forEach(function (row) {
    const tgl = extractDateOnly(getByCol(row, col.tanggal));
    if (tgl !== todayStr) return;

    const qty = col.qty !== -1 ? (parseAngkaSel_(getByCol(row, col.qty)) || 0) : 0;
    if (qty <= 0) return; // abaikan baris kosong/subtotal/total

    let nilai = col.totalAfterTax !== -1 ? (parseAngkaSel_(getByCol(row, col.totalAfterTax)) || 0) : 0;
    if (!nilai && col.totalHarga !== -1) nilai = parseAngkaSel_(getByCol(row, col.totalHarga)) || 0;

    result.gram += qty;
    result.nilai += nilai;
  });

  return result;
}

// Stretch target 120% Realisasi Nominal Transaksi 2026 (angka resmi dari
// user, BUKAN hasil kali TARGET_TAHUNAN x1.2 — dua target ini independen).
const RK_TARGET_120_NOMINAL = 35000000000000; // Rp35 Triliun

/**
 * Total gram & nilai transaksi EGR (Input_Jual_EGR) dalam rentang tanggal
 * [startStr, endStr] — logic & kolom SAMA PERSIS dengan blok "Penjualan
 * EGR" di getEgrRekapData (qty & Total harga after tax, fallback Total
 * Harga; baris kosong/subtotal/qty<=0 diabaikan), hanya diekstrak jadi
 * helper agar bisa dipakai ulang untuk periode YTD di
 * getRealisasiNominal2026 tanpa duplikasi logic.
 */
function getEgrYtdTotals_(startStr, endStr) {
  let gram = 0, nilai = 0;
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const jualSheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (!jualSheet) return { gram: gram, nilai: nilai };

  const values = jualSheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return { gram: gram, nilai: nilai };
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    qty: findColByKeywords(headerMap, ['qty']),
    totalAfterTax: findColByKeywords(headerMap, ['after', 'tax']),
    totalHarga: findColByKeywords(headerMap, ['total', 'harga'], ['after', 'tax'])
  };
  if (col.tanggal === -1) return { gram: gram, nilai: nilai };

  values.slice(headerRowIdx + 1).forEach(function (row) {
    const tgl = extractDateOnly(getByCol(row, col.tanggal));
    if (!tgl || tgl < startStr || tgl > endStr) return;

    const qty = col.qty !== -1 ? (parseAngkaSel_(getByCol(row, col.qty)) || 0) : 0;
    if (qty <= 0) return; // abaikan baris kosong/subtotal/total

    let n = col.totalAfterTax !== -1 ? (parseAngkaSel_(getByCol(row, col.totalAfterTax)) || 0) : 0;
    if (!n && col.totalHarga !== -1) n = parseAngkaSel_(getByCol(row, col.totalHarga)) || 0;

    gram += qty;
    nilai += n;
  });

  return { gram: gram, nilai: nilai };
}

/**
 * Realisasi Nominal Transaksi Perdagangan 2026 (YTD, 1 Jan s.d. Tanggal
 * Akhir filter halaman) — GABUNGAN Trading (PenjualanTransaksi +
 * PembelianTransaksi, TotalHarga & Qty) DAN EGR (Input_Jual_EGR, via
 * getEgrYtdTotals_), TIDAK dipengaruhi filter Buyer EGR di halaman ini
 * per instruksi user. Sumber Trading berasal dari sheet aplikasi ini,
 * sumber EGR dari spreadsheet eksternal PRICE_DISCOVERY_SOURCE_ID — dua
 * sumber data yang sepenuhnya terpisah, sehingga menjumlahkannya tidak
 * menimbulkan double counting. Target 100% = TARGET_TAHUNAN existing
 * (Rp32,2 T), target 120% = RK_TARGET_120_NOMINAL (Rp35 T, stretch
 * target bisnis, bukan x1,2).
 */
function getRealisasiNominal2026(endDateStr) {
  requireRole('Viewer');
  const end = extractDateOnly(endDateStr) || formatDateOnly(new Date());
  const yearStart = end.slice(0, 4) + '-01-01';

  const pjYtd = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL)
    .filter(function (p) { const tgl = extractDateOnly(p.Tanggal); return tgl >= yearStart && tgl <= end; });
  const pbYtd = sheetToObjects(SHEET_NAMES.PEMBELIAN_FULL)
    .filter(function (p) { const tgl = extractDateOnly(p.Tanggal); return tgl >= yearStart && tgl <= end; });

  const tradingNominal = pjYtd.reduce(function (s, p) { return s + Number(p.TotalHarga || 0); }, 0)
    + pbYtd.reduce(function (s, p) { return s + Number(p.TotalHarga || 0); }, 0);
  const tradingGramasi = pjYtd.reduce(function (s, p) { return s + Number(p.Qty || 0); }, 0)
    + pbYtd.reduce(function (s, p) { return s + Number(p.Qty || 0); }, 0);

  const egrYtd = getEgrYtdTotals_(yearStart, end);

  const realisasiNominal = tradingNominal + egrYtd.nilai;
  const realisasiGramasi = tradingGramasi + egrYtd.gram;

  const target100 = TARGET_TAHUNAN;
  const target120 = RK_TARGET_120_NOMINAL;
  const persen = target100 > 0 ? (realisasiNominal / target100) * 100 : 0;

  return {
    endDateStr: end,
    realisasiNominal: realisasiNominal,
    realisasiGramasi: realisasiGramasi,
    target100: target100,
    target120: target120,
    persenRealisasi: persen
  };
}

// ------------------------------------------------------------
// DOWNLOAD PENJUALAN HARI INI (Menu Penjualan) — dari sheet
// "F.1.Input_Trans_Jual" (spreadsheet EGR eksternal,
// PRICE_DISCOVERY_SOURCE_ID), BUKAN dari sheet PenjualanTransaksi/Invoice
// aplikasi ini.
// ------------------------------------------------------------
// Coba beberapa kandidat kata kunci berurutan — kembalikan indeks kolom
// pertama yang cocok. Dipakai krn header sheet "F.1.Input_Trans_Jual"
// bisa beda-beda penamaan (mis. "Buyer" vs "Pembeli", "Gramasi" vs "Qty").
function findColByKeywordsMulti_(headerMap, keywordSets) {
  for (let i = 0; i < keywordSets.length; i++) {
    const idx = findColByKeywords(headerMap, keywordSets[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Download Penjualan Hari Ini — dibaca LANGSUNG dari sheet
 * "F.1.Input_Trans_Jual", TIDAK PERNAH digabung/dijumlah berdasarkan
 * Buyer/Merk/Tanggal: 1 baris transaksi di "F.1.Input_Trans_Jual" = 1
 * baris laporan, urutan mengikuti urutan baris asli di sheet. Hanya
 * transaksi bertanggal HARI INI (Asia/Jakarta) yang diambil. Gramasi
 * dikonversi ke Kg (1.000 gram = 1 Kg). Read-only — tidak pernah
 * mengubah data "F.1.Input_Trans_Jual".
 */
function getPenjualanHariIniF1TransJual() {
  requireRole('Viewer');
  const todayStr = formatDateOnly(new Date());

  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = ss.getSheetByName('F.1.Input_Trans_Jual');
  if (!sheet) throw new Error('Sheet "F.1.Input_Trans_Jual" tidak ditemukan di spreadsheet sumber.');

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['tanggal', 'merk']);
  if (headerRowIdx === -1) throw new Error('Baris header (kolom Tanggal & Merk) tidak ditemukan di sheet "F.1.Input_Trans_Jual".');
  const headerRow = values[headerRowIdx];
  const headerMap = buildHeaderMap(headerRow);

  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    buyer: findColByKeywordsMulti_(headerMap, [['buyer'], ['pembeli'], ['nama']]),
    merk: findColByKeywords(headerMap, ['merk']),
    gramasi: findColByKeywordsMulti_(headerMap, [['gramasi'], ['qty'], ['gram']])
  };

  const missing = [];
  if (col.tanggal === -1) missing.push('Tanggal');
  if (col.buyer === -1) missing.push('Buyer/Pembeli/Nama');
  if (col.merk === -1) missing.push('Merk');
  if (col.gramasi === -1) missing.push('Gramasi/Qty');
  if (missing.length) {
    throw new Error('Kolom berikut tidak ditemukan di sheet "F.1.Input_Trans_Jual": ' + missing.join(', ') +
      '. Header yang terdeteksi: ' + headerRow.filter(function (h) { return String(h).trim(); }).join(', '));
  }

  const rows = [];
  values.slice(headerRowIdx + 1).forEach(function (row) {
    const tglRow = extractDateOnly(getByCol(row, col.tanggal));
    if (tglRow !== todayStr) return;

    const buyer = String(getByCol(row, col.buyer) || '').trim();
    const merk = String(getByCol(row, col.merk) || '').trim();
    const gramasiRaw = parseAngkaSel_(getByCol(row, col.gramasi));
    if (!buyer || gramasiRaw === null || gramasiRaw <= 0) return; // abaikan baris kosong/subtotal/total

    // SATU baris sheet = SATU baris laporan — tidak pernah digabung
    // dengan baris lain walau Buyer/Merk/Tanggal sama persis.
    rows.push({ buyer: buyer, merk: merk, gramasiKg: gramasiRaw / 1000 });
  });

  const jadwalPickupHariIni = getJadwalPickup(todayStr);

  return {
    tanggalStr: todayStr, tanggalLabel: formatTanggalPanjangIndo_(todayStr), rows: rows,
    jadwalPickup: jadwalPickupHariIni
  };
}

// ------------------------------------------------------------
// REKAP TRANSAKSI EGR PER SESI (12:00 / 16:00 WIB) — Laporan Harian EGR
// ------------------------------------------------------------
// Input_Jual_EGR TIDAK punya kolom jam, jadi jam transaksi TIDAK PERNAH
// ditebak dari urutan baris. Posisi kumulatif per sesi diambil OTOMATIS
// oleh time-driven trigger Apps Script (lihat egrSesiSnapshot1200/1600 &
// installEgrSesiSnapshotTriggers di Setup.gs, JAM/timezone Asia/Jakarta —
// SAMA dgn Session.getScriptTimeZone() krn appsscript.json sudah diset
// "timeZone": "Asia/Jakarta", jadi TIDAK ADA perhitungan tanggal/jam yg
// jatuh ke UTC di modul ini). TIDAK ADA tombol simpan manual di halaman
// pengguna. Kalau trigger BELUM terpasang/telat/gagal, getEgrSesiSnapshot
// di bawah otomatis membuat snapshot PEMULIHAN saat laporan dibuka (lihat
// ensureEgrSesiPemulihanJikaPerlu_) — jadi tabel tidak pernah nyangkut di
// "Snapshot tidak tersedia" selamanya kalau datanya sebenarnya bisa
// dihitung. Buyback selalu 0 (belum ada transaksi Buyback sungguhan).
const EGR_SESI_SNAPSHOT_HEADERS = [
  'ID', 'Tanggal', 'Sesi', 'KunciUnik', 'JualGramKumulatif', 'JualNilaiKumulatif',
  'BuybackGramKumulatif', 'BuybackNilaiKumulatif', 'HargaJualTerakhir',
  'HargaBuybackTerakhir', 'WaktuEksekusi', 'Status', 'Keterangan'
];

// Kunci unik "yyyy-MM-dd_HH.00" (mis. "2026-08-12_12.00") — dipakai SAMA
// PERSIS oleh fungsi penyimpan (runEgrSesiSnapshot_) & fungsi pembaca
// (getEgrSesiSnapshot) supaya tidak pernah ada ambiguitas format.
function egrSesiKunciUnik_(tgl, sesi) {
  return tgl + '_' + (sesi === '12:00' ? '12.00' : '16.00');
}

// Header-name -> nomor kolom (1-based). Dipakai (bukan indeks posisi
// tetap) supaya migrasi skema dari versi sebelumnya tidak merusak
// baca/tulis data baru — kolom lama yg sudah tidak dipakai (mis. SavedAt/
// SavedBy dari versi manual pertama) dibiarkan ada, TIDAK dihapus.
function egrSesiSnapshotHeaderMap_(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headerRow.forEach(function (h, i) { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

function ensureEgrSesiSnapshotSchema_(sh) {
  const map = egrSesiSnapshotHeaderMap_(sh);
  EGR_SESI_SNAPSHOT_HEADERS.forEach(function (name) {
    if (!map[name]) {
      const newCol = sh.getLastColumn() + 1;
      sh.getRange(1, newCol).setValue(name).setFontWeight('bold');
      map[name] = newCol;
    }
  });
}

// Storage PERSISTEN (Google Sheet di database aplikasi, bukan cache/
// properties sementara) — dibaca kembali oleh web app lewat
// getEgrSesiSnapshot(), bisa dibuka manual lewat Apps Script editor utk
// verifikasi ("Lokasi penyimpanan" di log, lihat runEgrSesiSnapshot_).
function getOrCreateEgrSesiSnapshotSheet_() {
  const db = getDb();
  let sh = db.getSheetByName(SHEET_NAMES.EGR_SESI_SNAPSHOT);
  if (sh) {
    ensureEgrSesiSnapshotSchema_(sh);
    return sh;
  }
  sh = db.insertSheet(SHEET_NAMES.EGR_SESI_SNAPSHOT);
  sh.getRange(1, 1, 1, EGR_SESI_SNAPSHOT_HEADERS.length).setValues([EGR_SESI_SNAPSHOT_HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

/**
 * Posisi kumulatif transaksi EGR (Input_Jual_EGR) utk satu tanggal, TANPA
 * batasan jam (sheet sumber tidak punya kolom jam) — dipakai sbg "posisi
 * yang tersedia saat ini" ketika snapshot sesi dijalankan (trigger
 * otomatis, pemulihan otomatis saat halaman dibuka, maupun pemulihan
 * manual admin). TIDAK PERNAH merekonstruksi posisi historis dari
 * urutan baris — HANYA membaca sheet SUMBER (read-only), tidak pernah
 * mengubah struktur/isi Input_Jual_EGR maupun Input_Beli_EGR.
 */
function getEgrSesiPosisiKumulatif_(dateStr) {
  const result = { jualGram: 0, jualNilai: 0, hargaJualTerakhir: null, hargaBuybackTerakhir: null };
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = getEgrSheet_(ss, EGR_SHEET.JUAL);
  if (!sheet) return result;
  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['total', 'harga']);
  if (headerRowIdx === -1) return result;
  const headerRow = values[headerRowIdx];
  const headerMap = buildHeaderMap(headerRow);
  const col = {
    tanggal: findColByKeywords(headerMap, ['tanggal']),
    qty: findColByKeywords(headerMap, ['qty']),
    hargaPerGram: findColByKeywords(headerMap, ['harga'], ['total', 'after', 'buyback']),
    totalAfterTax: findColByKeywords(headerMap, ['after', 'tax']),
    totalHarga: findColByKeywords(headerMap, ['total', 'harga'], ['after', 'tax']),
    buyback: findLastColContaining_(headerRow, 'buyback')
  };
  if (col.tanggal === -1) return result;

  values.slice(headerRowIdx + 1).forEach(function (row) {
    const tgl = extractDateOnly(getByCol(row, col.tanggal));
    if (tgl !== dateStr) return;
    const qty = col.qty !== -1 ? (parseAngkaSel_(getByCol(row, col.qty)) || 0) : 0;
    if (qty <= 0) return;
    result.jualGram += qty;

    let nilai = col.totalAfterTax !== -1 ? (parseAngkaSel_(getByCol(row, col.totalAfterTax)) || 0) : 0;
    if (!nilai && col.totalHarga !== -1) nilai = parseAngkaSel_(getByCol(row, col.totalHarga)) || 0;
    result.jualNilai += nilai;

    const h = col.hargaPerGram !== -1 ? parseAngkaSel_(getByCol(row, col.hargaPerGram)) : null;
    if (h !== null && h !== 0) result.hargaJualTerakhir = h;
    const b = col.buyback !== -1 ? parseAngkaSel_(getByCol(row, col.buyback)) : null;
    if (b !== null && b !== 0) result.hargaBuybackTerakhir = b;
  });
  return result;
}

/**
 * Inti eksekusi snapshot sesi "12:00"/"16:00" — dipanggil oleh (a)
 * trigger otomatis (tanpa dateStrOverride -> tanggal HARI INI, Status
 * "Otomatis"), (b) pemulihan otomatis saat halaman dibuka (dgn
 * dateStrOverride = hari ini, Status "Pemulihan"), atau (c) pemulihan
 * manual admin (dgn dateStrOverride = tanggal manapun, Status
 * "Pemulihan"). Upsert berdasarkan KunciUnik (tanggal+sesi) — dijalankan
 * berkali-kali pada sesi yang sama akan MENIMPA record yang sama, TIDAK
 * PERNAH membuat duplikasi. LockService mencegah dua eksekusi bersamaan
 * menyimpan data ganda. Setiap langkah (mulai, selesai, error) dicatat
 * ke Logger DAN Audit Trail — tidak ada try/catch kosong.
 */
function runEgrSesiSnapshot_(sesi, dateStrOverride) {
  if (sesi !== '12:00' && sesi !== '16:00') throw new Error('Sesi tidak valid: ' + sesi);

  const tz = Session.getScriptTimeZone();
  const isPemulihan = !!dateStrOverride;
  const statusLabel = isPemulihan ? 'Pemulihan' : 'Otomatis';

  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(30000);
  if (!gotLock) {
    Logger.log('[EgrSesiSnapshot] LOCK GAGAL — sesi=' + sesi + ' status=' + statusLabel + ' — snapshot dilewati (kemungkinan eksekusi lain sedang berjalan).');
    logAudit('SNAPSHOT_SKIP', 'EgrSesiSnapshot', sesi, 'Gagal mendapatkan lock, snapshot dilewati.');
    return { success: false, skipped: true, reason: 'lock' };
  }

  let tgl = null;
  try {
    tgl = isPemulihan ? extractDateOnly(dateStrOverride) : formatDateOnly(new Date());
    const now = new Date();
    Logger.log('[EgrSesiSnapshot] MULAI — sesi=' + sesi + ' tanggal=' + tgl + ' timezone=' + tz +
      ' waktuEksekusiAktual=' + Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss') + ' status=' + statusLabel);

    const posisi = getEgrSesiPosisiKumulatif_(tgl);
    const kunciUnik = egrSesiKunciUnik_(tgl, sesi);

    const sh = getOrCreateEgrSesiSnapshotSheet_();
    const map = egrSesiSnapshotHeaderMap_(sh);
    const values = sh.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][map['KunciUnik'] - 1]) === kunciUnik) { rowIdx = i; break; }
    }

    const rowValues = {
      Tanggal: tgl,
      Sesi: sesi,
      KunciUnik: kunciUnik,
      JualGramKumulatif: posisi.jualGram,
      JualNilaiKumulatif: posisi.jualNilai,
      BuybackGramKumulatif: 0,
      BuybackNilaiKumulatif: 0,
      HargaJualTerakhir: posisi.hargaJualTerakhir,
      HargaBuybackTerakhir: posisi.hargaBuybackTerakhir,
      WaktuEksekusi: now,
      Status: statusLabel,
      Keterangan: isPemulihan ? 'Posisi saat fungsi pemulihan dijalankan, bukan posisi historis tepat jam sesi.' : 'Snapshot otomatis dari time-driven trigger.'
    };

    if (rowIdx === -1) {
      const id = getNextId(SHEET_NAMES.EGR_SESI_SNAPSHOT, 'ID', 'SESI');
      const newRow = new Array(sh.getLastColumn()).fill('');
      newRow[map['ID'] - 1] = id;
      Object.keys(rowValues).forEach(function (k) { newRow[map[k] - 1] = rowValues[k]; });
      sh.appendRow(newRow);
    } else {
      Object.keys(rowValues).forEach(function (k) {
        sh.getRange(rowIdx + 1, map[k]).setValue(rowValues[k]);
      });
    }

    Logger.log('[EgrSesiSnapshot] SELESAI — kunciUnik=' + kunciUnik + ' jualGram=' + posisi.jualGram +
      ' jualNilai=' + posisi.jualNilai + ' hargaJualTerakhir=' + posisi.hargaJualTerakhir +
      ' hargaBuybackTerakhir=' + posisi.hargaBuybackTerakhir + ' status=' + statusLabel +
      ' lokasiPenyimpanan=' + sh.getParent().getId() + '!' + sh.getName());
    logAudit('SNAPSHOT_' + statusLabel.toUpperCase(), 'EgrSesiSnapshot', kunciUnik,
      'JualGram=' + posisi.jualGram + ' JualNilai=' + posisi.jualNilai + ' WaktuEksekusi=' + Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'));

    return {
      success: true, tanggalStr: tgl, sesi: sesi, status: statusLabel,
      jualGram: posisi.jualGram, jualNilai: posisi.jualNilai,
      savedAtLabel: Utilities.formatDate(now, tz, 'HH:mm')
    };
  } catch (err) {
    Logger.log('[EgrSesiSnapshot] ERROR — sesi=' + sesi + ' tanggal=' + (tgl || '-') + ' status=' + statusLabel + ' pesan=' + (err && err.message ? err.message : err));
    logAudit('SNAPSHOT_FAILED', 'EgrSesiSnapshot', (tgl || '-') + ' ' + sesi, String(err && err.stack ? err.stack : err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// Target trigger — nama fungsi ini yang didaftarkan oleh
// installEgrSesiSnapshotTriggers() di Setup.gs. Tidak menerima argumen
// (event trigger diabaikan), selalu snapshot tanggal HARI INI (Asia/
// Jakarta), Status "Otomatis".
function egrSesiSnapshot1200() { runEgrSesiSnapshot_('12:00'); }
function egrSesiSnapshot1600() { runEgrSesiSnapshot_('16:00'); }

/**
 * Pemulihan manual kalau trigger otomatis gagal berjalan — ADMIN SAJA,
 * TIDAK ditampilkan sbg tombol di halaman pengguna. Perhitungan SAMA
 * PERSIS dgn trigger (total transaksi tanggal itu SAAT FUNGSI INI
 * DIJALANKAN) — tidak merekonstruksi posisi historis dari urutan baris.
 * Status selalu "Pemulihan".
 */
function adminRerunEgrSesiSnapshot(dateStr, sesi) {
  requireRole('Admin');
  const tgl = extractDateOnly(dateStr);
  if (!tgl) throw new Error('Tanggal wajib diisi.');
  if (sesi !== '12:00' && sesi !== '16:00') throw new Error('Sesi tidak valid.');
  return runEgrSesiSnapshot_(sesi, tgl);
}

/**
 * FUNGSI PEMULIHAN OTOMATIS — dipanggil setiap kali getEgrSesiSnapshot()
 * dibuka utk tanggal HARI INI. Kalau waktu sekarang (Asia/Jakarta) sudah
 * lewat jam sesi tapi snapshot sesi itu belum ada, buat snapshot
 * "Pemulihan" pakai posisi yang tersedia SAAT INI — supaya tabel Rekap
 * Sesi tidak nyangkut selamanya di "Snapshot tidak tersedia" hanya krn
 * trigger belum terpasang/telat/gagal. TIDAK berlaku utk tanggal selain
 * hari ini (tidak merekonstruksi posisi historis).
 */
function ensureEgrSesiPemulihanJikaPerlu_(tgl) {
  // debug dikembalikan (BUKAN dilempar/dibuang) supaya getEgrSesiSnapshot
  // bisa membalikkannya ke frontend — satu-satunya cara menelusuri
  // eksekusi server di lingkungan ini krn Stackdriver/Logger tidak bisa
  // diakses dari sesi kerja ini.
  const debug = { tgl: tgl };
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayStr = formatDateOnly(now);
  debug.todayStr = todayStr;
  debug.sameDay = tgl === todayStr;
  if (tgl !== todayStr) return debug; // hanya berlaku utk laporan tanggal hari ini

  const nowHour = Number(Utilities.formatDate(now, tz, 'H'));
  const nowMinute = Number(Utilities.formatDate(now, tz, 'm'));
  const nowMinutesOfDay = nowHour * 60 + nowMinute;
  debug.nowHour = nowHour;
  debug.nowMinute = nowMinute;

  const sh = getOrCreateEgrSesiSnapshotSheet_();
  const map = egrSesiSnapshotHeaderMap_(sh);
  debug.kunciUnikColIdx = map['KunciUnik']; // undefined -> bukti kolom KunciUnik hilang/tidak ketemu
  debug.sheetLastCol = sh.getLastColumn();
  debug.headerRowLive = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const values = sh.getDataRange().getValues();
  debug.totalDataRows = values.length - 1;
  const existing = { '12:00': false, '16:00': false };
  for (let i = 1; i < values.length; i++) {
    const kunci = String(values[i][map['KunciUnik'] - 1]);
    if (kunci === egrSesiKunciUnik_(tgl, '12:00')) existing['12:00'] = true;
    if (kunci === egrSesiKunciUnik_(tgl, '16:00')) existing['16:00'] = true;
  }
  debug.existingSebelum = existing;

  // Masing-masing sesi dibungkus try/catch SENDIRI-SENDIRI — supaya kalau
  // pemulihan sesi 12:00 gagal, sesi 16:00 tetap dicoba (tidak ikut gagal
  // gara-gara error di sesi lain). Error tetap dilempar ulang ke
  // pemanggil (getEgrSesiSnapshot) SETELAH kedua sesi dicoba, supaya
  // tetap tercatat di sana juga — tidak ada try/catch kosong.
  let firstError = null;
  debug.cutoff1200Lewat = nowMinutesOfDay >= 12 * 60;
  debug.cutoff1600Lewat = nowMinutesOfDay >= 16 * 60;
  if (nowMinutesOfDay >= 12 * 60 && !existing['12:00']) {
    Logger.log('[EgrSesiSnapshot] Pemulihan otomatis dipicu utk sesi 12:00 (tanggal=' + tgl + ', waktu sekarang sudah lewat 12.00 WIB tapi snapshot belum ada).');
    try {
      debug.hasil1200 = runEgrSesiSnapshot_('12:00', tgl);
    } catch (err) {
      Logger.log('[EgrSesiSnapshot] Pemulihan sesi 12:00 GAGAL — ' + (err && err.message ? err.message : err));
      debug.error1200 = String(err && err.message ? err.message : err);
      firstError = firstError || err;
    }
  }
  if (nowMinutesOfDay >= 16 * 60 && !existing['16:00']) {
    Logger.log('[EgrSesiSnapshot] Pemulihan otomatis dipicu utk sesi 16:00 (tanggal=' + tgl + ', waktu sekarang sudah lewat 16.00 WIB tapi snapshot belum ada).');
    try {
      debug.hasil1600 = runEgrSesiSnapshot_('16:00', tgl);
    } catch (err) {
      Logger.log('[EgrSesiSnapshot] Pemulihan sesi 16:00 GAGAL — ' + (err && err.message ? err.message : err));
      debug.error1600 = String(err && err.message ? err.message : err);
      firstError = firstError || err;
    }
  }
  debug.firstError = firstError ? String(firstError.message || firstError) : null;
  return debug;
}

/**
 * Ambil snapshot sesi 12:00 & 16:00 utk satu tanggal — null kalau sesi
 * itu belum pernah dijalankan (BUKAN 0, supaya UI bisa membedakan "belum
 * ada snapshot" vs "sudah snapshot dan nilainya nol"). Utk tanggal HARI
 * INI, otomatis memicu pemulihan (lihat ensureEgrSesiPemulihanJikaPerlu_)
 * SEBELUM membaca, supaya hasilnya selalu up-to-date.
 */
function getEgrSesiSnapshot(dateStr) {
  requireRole('Viewer');
  const tgl = extractDateOnly(dateStr);
  const result = { tanggalStr: tgl, s1200: null, s1600: null };
  if (!tgl) return result;

  // Pemulihan otomatis TIDAK BOLEH menjatuhkan seluruh pembacaan kalau ia
  // gagal (mis. error transient saat menulis) — kalau dibiarkan throw,
  // seluruh getEgrSesiSnapshot() gagal dan tabel di halaman nyangkut di
  // "Memuat..." selamanya walau data snapshot yang SUDAH ADA sebenarnya
  // masih bisa dibaca. Error tetap DICATAT (Logger + Audit Trail + balik
  // ke frontend lewat result.recoveryError) & debug lengkapnya dibalikkan
  // lewat result.debugPemulihan (tidak disembunyikan — lingkungan kerja
  // ini tidak bisa akses Stackdriver/Logger sama sekali).
  try {
    result.debugPemulihan = ensureEgrSesiPemulihanJikaPerlu_(tgl);
    if (result.debugPemulihan && result.debugPemulihan.firstError) {
      result.recoveryError = result.debugPemulihan.firstError;
      logAudit('SNAPSHOT_RECOVERY_FAILED', 'EgrSesiSnapshot', tgl, result.debugPemulihan.firstError);
    }
  } catch (err) {
    const pesan = String(err && err.message ? err.message : err);
    Logger.log('[EgrSesiSnapshot] Pemulihan otomatis GAGAL saat getEgrSesiSnapshot dibuka — tanggal=' + tgl + ' pesan=' + pesan);
    logAudit('SNAPSHOT_RECOVERY_FAILED', 'EgrSesiSnapshot', tgl, String(err && err.stack ? err.stack : err));
    result.recoveryError = pesan;
  }

  const sh = getOrCreateEgrSesiSnapshotSheet_();
  const map = egrSesiSnapshotHeaderMap_(sh);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const kunci = String(row[map['KunciUnik'] - 1]);
    if (kunci !== egrSesiKunciUnik_(tgl, '12:00') && kunci !== egrSesiKunciUnik_(tgl, '16:00')) continue;
    const sesi = String(row[map['Sesi'] - 1]);
    const waktuEksekusi = row[map['WaktuEksekusi'] - 1];
    const entry = {
      jualGram: Number(row[map['JualGramKumulatif'] - 1]) || 0,
      jualNilai: Number(row[map['JualNilaiKumulatif'] - 1]) || 0,
      buybackGram: Number(row[map['BuybackGramKumulatif'] - 1]) || 0,
      buybackNilai: Number(row[map['BuybackNilaiKumulatif'] - 1]) || 0,
      hargaJualTerakhir: row[map['HargaJualTerakhir'] - 1] || null,
      hargaBuybackTerakhir: row[map['HargaBuybackTerakhir'] - 1] || null,
      status: row[map['Status'] - 1] || 'Otomatis',
      savedAtLabel: waktuEksekusi instanceof Date ? Utilities.formatDate(waktuEksekusi, Session.getScriptTimeZone(), 'HH:mm') : String(waktuEksekusi || '')
    };
    if (sesi === '12:00') result.s1200 = entry;
    else if (sesi === '16:00') result.s1600 = entry;
  }
  return result;
}

/**
 * Diagnostik ADMIN — jalankan manual dari editor Apps Script utk
 * memverifikasi trigger snapshot benar-benar terpasang (bukan cuma
 * diasumsikan). Mencetak ke Logger, tidak mengubah data apa pun.
 */
function logEgrSesiTriggerStatus() {
  const targetFunctions = ['egrSesiSnapshot1200', 'egrSesiSnapshot1600'];
  const triggers = ScriptApp.getProjectTriggers();
  targetFunctions.forEach(function (fnName) {
    const found = triggers.filter(function (t) { return t.getHandlerFunction() === fnName; });
    if (found.length === 0) {
      Logger.log('[EgrSesiSnapshot] Trigger TIDAK DITEMUKAN utk fungsi "' + fnName + '" — jalankan installEgrSesiSnapshotTriggers() di Setup.gs.');
    } else {
      found.forEach(function (t) {
        Logger.log('[EgrSesiSnapshot] Trigger DITEMUKAN — fungsi=' + fnName + ' tipe=' + t.getEventType() + ' sumber=' + t.getTriggerSource());
      });
    }
  });
  Logger.log('[EgrSesiSnapshot] Timezone project (Session.getScriptTimeZone()) = ' + Session.getScriptTimeZone());
}

// TEMP DIAGNOSTIK — verifikasi hasil getLaporanHarianEtfData (versi
// ledger Top Up EGR/Saldo dari Input_Jual_EGR) untuk satu tanggal.
// Read-only. (dihapus setelah selesai)
function _debugEgrSaldo_(dateStr) {
  const result = {};
  try {
    const targetDateStr = dateStr ? extractDateOnly(dateStr) : formatDateOnly(new Date());
    result.targetDateStr = targetDateStr;
    result.ledger = readEgrLedger_(targetDateStr);
    result.laporan = getLaporanHarianEtfData(dateStr);
  } catch (err) {
    result.error = String(err && err.message || err);
    result.stack = err && err.stack ? String(err.stack) : null;
  }
  return result;
}

// TEMP DIAGNOSTIC — investigasi bug Jadwal Pickup/Barang Belum Dipickup
// (per laporan user, 31 Agu 2026). Read-only, TIDAK mengubah data
// apapun. Menampilkan SEMUA field mentah dari PenjualanTransaksi utk
// invoice yang dilaporkan salah tampil, baik dicari dengan nomor
// "INV-00017" (yang tersimpan di app) maupun "WO-00017" (yang
// disebut user sebagai nomor di source), supaya bisa dipastikan
// benar-benar nomor mana yang ada di sheet. (dihapus setelah selesai)
function _debugPjlPickupCheck_() {
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const nomorTarget = ['00691', '00692', '00693', '00710', '00714', '00686'];
  const cocok = rows.filter(function (r) {
    const inv = String(r.NomorInvoice || '').trim();
    return nomorTarget.some(function (n) { return inv.indexOf(n) === 0; });
  }).map(function (r) {
    return {
      ID: r.ID,
      NomorInvoice: r.NomorInvoice,
      Pembeli: r.Pembeli,
      Status: r.Status,
      StatusPickup: r.StatusPickup,
      Tanggal: String(r.Tanggal),
      TanggalPengambilan: String(r.TanggalPengambilan),
      JamPengambilan: r.JamPengambilan,
      MetodePengambilan: r.MetodePengambilan,
      EstimasiPenjemputan: String(r.EstimasiPenjemputan),
      Catatan: r.Catatan
    };
  });
  return { totalRowsPenjualanTransaksi: rows.length, cocok: cocok };
}

// TEMP DIAGNOSTIC — cek sheet SUMBER F.1.Input_Trans_Jual (bukan
// PenjualanTransaksi app ini): header lengkap + baris mentah utk
// invoice yang dilaporkan salah, supaya bisa dipastikan (a) apakah
// ada kolom status-pickup di sumber yang belum ikut di-sync, dan
// (b) format nomor transaksi ASLI di sumber (INV vs WO). Read-only.
// (dihapus setelah selesai)
function _debugF1SourceCheck_() {
  const sourceSs = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sourceSheet = sourceSs.getSheetByName('F.1.Input_Trans_Jual');
  if (!sourceSheet) return { error: 'Sheet F.1.Input_Trans_Jual tidak ditemukan' };
  const values = sourceSheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['tanggal', 'invoice', 'pembeli', 'qty', 'harga']);
  if (headerRowIdx === -1) return { error: 'Header tidak terdeteksi' };
  const header = values[headerRowIdx];
  const nomorTarget = ['00691', '00692', '00693', '00710', '00714', '00686'];
  const dataRows = values.slice(headerRowIdx + 1);
  const cocok = dataRows.filter(function (row) {
    return row.some(function (cell) {
      return nomorTarget.some(function (n) { return String(cell || '').indexOf(n) === 0; });
    });
  }).map(function (row) {
    const obj = {};
    header.forEach(function (h, i) { obj[String(h || ('col' + i))] = row[i] instanceof Date ? String(row[i]) : row[i]; });
    return obj;
  });
  return { headerRowIdx: headerRowIdx, header: header, totalDataRows: dataRows.length, cocok: cocok };
}

// ------------------------------------------------------------
// KPI PERDAGANGAN EMAS
// Kebijakan sumber data (per instruksi user, revisi terbaru):
//   Target=J50, Sisa Target=J51, Sisa Hari Kerja=J52, Target Gramasi
//   Harian=J53 (gram), Kebutuhan Transaksi/Hari=K53, Pencapaian=J56,
//   Average Price=O5 — SEMUA dibaca LANGSUNG dari sheet "Rekap"
//   (spreadsheet EGR eksternal, PRICE_DISCOVERY_SOURCE_ID), TIDAK
//   dihitung ulang di aplikasi. Realisasi Transaksi=I15 (Nominal Trx,
//   total tahun berjalan) & grafik bulanan=I3:I14 — juga langsung dari
//   sheet. Formula aplikasi (kpiHitungTingkat_ dkk) HANYA dipakai
//   sebagai fallback kalau cell sumber di atas kosong/tidak valid;
//   kalau fallback kepakai, dicatat via Logger.log (Extensions > Apps
//   Script > Executions) — tidak pernah ditampilkan ke user. Sheet
//   Rekap sendiri TIDAK PERNAH ditulis/diubah oleh fungsi ini.
//   Prognosa 110%/120% dan Proyeksi Akhir Tahun TIDAK punya cell
//   sumber di sheet sama sekali (konsep baru di luar sheet), jadi
//   keduanya tetap dihitung di aplikasi — itu bukan "menghitung ulang
//   nilai yang sudah ada", karena memang belum pernah ada di sheet.
// Target 100/110/120% adalah angka resmi yang diberikan user; khusus
// KPI_TARGET_100 hanya dipakai sebagai FALLBACK kalau J50 kosong.
// ------------------------------------------------------------
const KPI_TARGET_100 = 32207565800000;
const KPI_TARGET_110 = KPI_TARGET_100 * 1.1;
const KPI_TARGET_120 = KPI_TARGET_100 * 1.2;

function kpiValidNumber_(v) {
  if (v === '' || v === null || v === undefined) return false;
  const n = Number(v);
  return !isNaN(n) && isFinite(n);
}

/**
 * Jumlah hari kerja (Senin-Jumat, sama seperti konvensi
 * getMarginTrendHarian — skip Sabtu/Minggu) di antara dua tanggal
 * (inklusif kedua ujungnya).
 */
function hitungHariKerjaAntara_(startDate, endDate) {
  let count = 0;
  const d = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  while (d <= end) {
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) count++;
    d.setDate(d.getDate() + 1);
  }
  return count;
}

/**
 * Hitung satu tingkat target: sisa, pencapaian, kebutuhan per hari,
 * kebutuhan gramasi per hari. Dipakai untuk Prognosa 110%/120% (tidak
 * ada cell sumbernya di sheet, jadi murni dihitung) dan sebagai
 * fallback tingkat 100% kalau cell J50-J56/K53 kosong/invalid. Sisa
 * TIDAK PERNAH negatif.
 */
function kpiHitungTingkat_(targetNominal, realisasi, sisaHariKerja, avgPrice) {
  const tercapai = realisasi >= targetNominal;
  const sisa = tercapai ? 0 : (targetNominal - realisasi);
  const surplus = tercapai ? (realisasi - targetNominal) : 0;
  const pencapaianPersen = targetNominal > 0 ? (realisasi / targetNominal) * 100 : 0;
  const kebutuhanPerHari = (!tercapai && sisaHariKerja > 0) ? (sisa / sisaHariKerja) : 0;
  const kebutuhanGramPerHari = (kebutuhanPerHari > 0 && avgPrice > 0) ? (kebutuhanPerHari / avgPrice) : 0;
  return {
    targetNominal: targetNominal,
    realisasi: realisasi,
    tercapai: tercapai,
    sisa: sisa,
    surplus: surplus,
    pencapaianPersen: pencapaianPersen,
    kebutuhanPerHari: kebutuhanPerHari,
    kebutuhanGramPerHari: kebutuhanGramPerHari
  };
}

function getKpiPerdaganganEmas() {
  requireRole('Viewer');
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sh = ss.getSheetByName('Rekap');
  if (!sh) throw new Error('Sheet "Rekap" tidak ditemukan.');

  // OPTIMASI LOADING: baca SATU KALI blok baris 3-56 kolom A-O (mencakup
  // seluruh range/sel yang dipakai fungsi ini: A3:A14, I3:I14, I15, E15,
  // J51, J52, J53, J56, K53, O5), lalu index ke array di memory —
  // sebelumnya 10 pemanggilan getRange(...).getValues()/getValue()
  // terpisah per eksekusi. Sel & nilai yang dibaca PERSIS SAMA, hanya
  // jumlah round-trip ke Sheets service yang berkurang.
  const KPI_ROW_MIN = 3, KPI_ROW_MAX = 56, KPI_COL_MIN = 1, KPI_COL_MAX = 15; // kolom A..O
  const kpiBlock = sh.getRange(KPI_ROW_MIN, KPI_COL_MIN, KPI_ROW_MAX - KPI_ROW_MIN + 1, KPI_COL_MAX - KPI_COL_MIN + 1).getValues();
  function kpiCell_(row, col) { return kpiBlock[row - KPI_ROW_MIN][col - KPI_COL_MIN]; }
  const KPI_COL_A = 1, KPI_COL_E = 5, KPI_COL_I = 9, KPI_COL_J = 10, KPI_COL_K = 11, KPI_COL_O = 15;

  const bulanLabels = [];
  const nominalTrxBulananTrading = [];
  for (let r = 3; r <= 14; r++) {
    bulanLabels.push(String(kpiCell_(r, KPI_COL_A) || '').trim());
    nominalTrxBulananTrading.push(Number(kpiCell_(r, KPI_COL_I)) || 0);
  }

  // --- EGR digabung ke Realisasi Transaksi ---
  // Blok bulanan EGR ada terpisah di baris 18-22 (mulai Agustus, sejak
  // transaksi EGR berjalan), Total EGR di baris 23. Sumbernya sheet
  // BEDA (Input_Jual_EGR/Input_Beli_EGR) dari blok Trading di atas
  // (F.1.Input_Trans_Jual/F.2.Input_Trans_Beli), jadi tidak ada risiko
  // double counting antar keduanya. Sebelumnya Realisasi Transaksi KPI
  // Perdagangan Emas HANYA memakai I15 (Total Trading) — EGR belum
  // ikut terhitung. Di-map per label bulan (bukan asumsi posisi baris)
  // supaya tetap benar walau baris EGR yang terisi berubah.
  const KPI_EGR_ROW_MIN = 18, KPI_EGR_ROW_MAX = 22, KPI_EGR_TOTAL_ROW = 23;
  const egrNominalByBulan = {};
  for (let r = KPI_EGR_ROW_MIN; r <= KPI_EGR_ROW_MAX; r++) {
    const label = String(kpiCell_(r, KPI_COL_A) || '').trim();
    if (label) egrNominalByBulan[label] = Number(kpiCell_(r, KPI_COL_I)) || 0;
  }
  const nominalTrxBulananEgr = bulanLabels.map(function (label) {
    return egrNominalByBulan[label] || 0;
  });
  const nominalTrxBulanan = nominalTrxBulananTrading.map(function (v, i) {
    return v + nominalTrxBulananEgr[i];
  });

  const realisasiTrading = Number(kpiCell_(15, KPI_COL_I)) || 0;
  const realisasiEgr = nominalTrxBulananEgr.reduce(function (a, b) { return a + b; }, 0);
  const realisasi = realisasiTrading + realisasiEgr;
  Logger.log('[KPI Perdagangan Emas] Realisasi Trading = Rp' + realisasiTrading +
    ' | Realisasi EGR = Rp' + realisasiEgr + ' | Total Realisasi (Trading+EGR) = Rp' + realisasi);

  const totalGramasiTrading = Number(kpiCell_(15, KPI_COL_E)) || 0;
  const totalGramasiEgr = Number(kpiCell_(KPI_EGR_TOTAL_ROW, KPI_COL_E)) || 0;
  const totalGramasi = totalGramasiTrading + totalGramasiEgr;

  // --- Target 100% — di-hardcode di aplikasi (KPI_TARGET_100), TIDAK
  // lagi dibaca dari J50, per instruksi user. Prognosa 110%/120%
  // (KPI_TARGET_110/120) otomatis mengikuti angka ini (x1.1/x1.2).
  const target = KPI_TARGET_100;

  // --- Sisa Hari Kerja ---
  // SEBELUMNYA dibaca dari Rekap!J52 — tapi J52 BUKAN sel "Sisa Hari
  // Kerja". Di sheet, label "hari" (baris mini-tabel Target/Sisa/hari/
  // kilogram/Penjualan (Rp)/Penjualan (Kg)/Pencapaian, kolom I-K) ada
  // di baris 51 dengan nilai 65 (masuk akal: ±65 hari kerja ke 31
  // Des), sedangkan baris 52 labelnya "kilogram" — nilai besaran
  // gramasi target, BUKAN jumlah hari. Seluruh mini-tabel ini rupanya
  // sudah bergeser 1 baris dari referensi lama di kode (J50/J51/J52/
  // J53/K53/J56 — lihat komentar di atas fungsi ini), kemungkinan
  // karena ada baris yang pernah disisipkan di sheet. Itulah sebab
  // "Sisa Hari Kerja" tampil ribuan hari (6.600/8.700 dst) — yang
  // ditampilkan sebenarnya nilai gramasi dari baris "kilogram".
  //
  // Diperbaiki: TIDAK lagi membaca cell sheet manapun untuk Sisa Hari
  // Kerja (supaya tidak rawan salah lagi kalau sheet bergeser lagi ke
  // depannya) — selalu dihitung langsung: dari hari ini sampai 31
  // Desember tahun berjalan, hanya Senin-Jumat (Sabtu/Minggu tidak
  // dihitung). Belum ada data hari libur nasional/cuti bersama di
  // aplikasi ini untuk dikurangi.
  const todayKerja = new Date();
  const akhirTahunKerja = new Date(todayKerja.getFullYear(), 11, 31);
  const sisaHariKerja = hitungHariKerjaAntara_(todayKerja, akhirTahunKerja);

  const tercapai = realisasi >= target;

  // --- Sisa Target Transaksi ---
  // Dihitung langsung dari target (KPI_TARGET_100, hardcode) dikurangi
  // realisasi — TIDAK lagi dibaca dari Rekap!J51, karena cell itu masih
  // mengikuti Target lama di sheet (J50) yang sudah tidak dipakai sejak
  // Target 100% di-hardcode di aplikasi, sehingga nilainya bisa tidak
  // sinkron dengan target yang sebenarnya ditampilkan.
  let sisaTarget;
  if (tercapai) {
    sisaTarget = 0;
  } else {
    sisaTarget = Math.max(0, target - realisasi);
  }
  const surplus = tercapai ? Math.max(0, realisasi - target) : 0;

  // --- Pencapaian % ---
  // Dihitung LANGSUNG dari realisasi/target (KPI_TARGET_100, hardcode di
  // app) — TIDAK lagi dibaca dari Rekap!J56, alasannya SAMA PERSIS seperti
  // Sisa Target Transaksi di atas: J56 masih mengikuti Target lama di
  // sheet (J50) sehingga bisa tidak sinkron dengan target yang sebenarnya
  // ditampilkan. J56 juga tidak bisa dipakai untuk menebak format sel
  // (desimal 0-1 vs sudah persen) secara aman: heuristik lama
  // "Math.abs(x) <= 1 -> x *= 100" SALAH tepat di atas 100% — begitu
  // realisasi melewati target, rasio desimalnya (mis. 1,0058 = 100,58%)
  // sudah > 1 sehingga TIDAK dikali 100 dan malah tampil "1,01%". Target
  // sudah tercapai TIDAK BOLEH membuat persentase seolah reset — jadi
  // dihitung langsung, tidak lagi bergantung pada tebakan format sel.
  const pencapaianPersen = target > 0 ? (realisasi / target) * 100 : 0;

  // --- Average Price (O5) ---
  const rawAvgPrice = kpiCell_(5, KPI_COL_O);
  let avgPrice;
  if (kpiValidNumber_(rawAvgPrice) && Number(rawAvgPrice) > 0) {
    avgPrice = Number(rawAvgPrice);
  } else {
    avgPrice = totalGramasi > 0 ? (realisasi / totalGramasi) : 0;
    Logger.log('[KPI fallback] Average Price (Rekap!O5) kosong/invalid, dihitung ulang dari Nominal Trx / Jumlah Gramasi = ' + avgPrice);
  }

  // --- Kebutuhan Transaksi / Hari ---
  // TIDAK lagi dibaca dari Rekap!K53 — sama seperti J51/J56 di atas:
  // formula K53 di sheet ("=roundup((J50-I15)/2/J52;-2)") memakai
  // Target LAMA (J50) dan Realisasi Trading-only (I15), jadi begitu
  // Realisasi digabung Trading+EGR di atas, K53 tidak lagi sinkron.
  // Dihitung langsung dari sisaTarget/sisaHariKerja yang SUDAH memakai
  // target & realisasi gabungan.
  const kebutuhanHarian = tercapai ? 0 : (sisaHariKerja > 0 ? (sisaTarget / sisaHariKerja) : 0);

  // --- Target Gramasi Harian, dalam GRAM ---
  // TIDAK lagi dibaca dari Rekap!J53, alasan sama seperti Kebutuhan
  // Transaksi/Hari di atas.
  const gramasiHarian = tercapai ? 0 : (avgPrice > 0 ? (kebutuhanHarian / avgPrice) : 0);

  const tingkat100 = {
    targetNominal: target,
    realisasi: realisasi,
    tercapai: tercapai,
    sisa: sisaTarget,
    surplus: surplus,
    pencapaianPersen: pencapaianPersen,
    kebutuhanPerHari: kebutuhanHarian,
    kebutuhanGramPerHari: gramasiHarian
  };
  // Prognosa 110%/120% TIDAK ADA cell sumbernya di sheet Rekap — murni
  // dihitung, pakai realisasi/sisaHariKerja/avgPrice yang sudah
  // (kalau perlu) di-fallback di atas.
  const tingkat110 = kpiHitungTingkat_(KPI_TARGET_110, realisasi, sisaHariKerja, avgPrice);
  const tingkat120 = kpiHitungTingkat_(KPI_TARGET_120, realisasi, sisaHariKerja, avgPrice);

  // --- Grafik bulanan: realisasi (Nominal Trx) per bulan = Trading +
  // EGR (lihat nominalTrxBulanan di atas), target per bulan (Target
  // dibagi rata 12 bulan — sheet tidak punya breakdown target
  // bulanan), kumulatif realisasi & target. Total kumulatif di bulan
  // berjalan = realisasi (Trading+EGR gabungan) di atas.
  const targetBulanan = target / 12;
  let kumRealisasi = 0, kumTarget = 0;
  const realisasiKumulatif = [];
  const targetKumulatif = [];
  nominalTrxBulanan.forEach(function (v) {
    kumRealisasi += v;
    kumTarget += targetBulanan;
    realisasiKumulatif.push(kumRealisasi);
    targetKumulatif.push(kumTarget);
  });

  // --- Proyeksi akhir tahun — TIDAK ADA cell sumbernya di sheet,
  // murni dihitung dari rata-rata realisasi harian (hari kerja yang
  // SUDAH berjalan tahun ini, 1 Jan s/d hari ini, konvensi hari kerja
  // sama seperti getMarginTrendHarian: skip Sabtu/Minggu).
  const today2 = new Date();
  const awalTahun = new Date(today2.getFullYear(), 0, 1);
  const hariKerjaBerjalan = hitungHariKerjaAntara_(awalTahun, today2);
  const rataRataHarian = hariKerjaBerjalan > 0 ? (realisasi / hariKerjaBerjalan) : 0;
  const rataRataBulanan = rataRataHarian * 21; // ~21 hari kerja/bulan, murni estimasi tampilan
  const proyeksiAkhirTahun = realisasi + (rataRataHarian * sisaHariKerja);
  const proyeksiPersen100 = target > 0 ? (proyeksiAkhirTahun / target) * 100 : 0;

  let proyeksiTingkatTertinggi = 'Di Bawah Target';
  if (proyeksiAkhirTahun >= KPI_TARGET_120) proyeksiTingkatTertinggi = '120%';
  else if (proyeksiAkhirTahun >= KPI_TARGET_110) proyeksiTingkatTertinggi = '110%';
  else if (proyeksiAkhirTahun >= target) proyeksiTingkatTertinggi = '100%';

  return {
    realisasi: realisasi,
    realisasiTrading: realisasiTrading,
    realisasiEgr: realisasiEgr,
    sisaHariKerja: sisaHariKerja,
    avgPriceTransaksi: avgPrice,
    tingkat100: tingkat100,
    tingkat110: tingkat110,
    tingkat120: tingkat120,
    grafikBulanan: {
      labels: bulanLabels,
      realisasi: nominalTrxBulanan,
      realisasiTrading: nominalTrxBulananTrading,
      realisasiEgr: nominalTrxBulananEgr,
      target: nominalTrxBulanan.map(function () { return targetBulanan; }),
      realisasiKumulatif: realisasiKumulatif,
      targetKumulatif: targetKumulatif
    },
    proyeksi: {
      hariKerjaBerjalan: hariKerjaBerjalan,
      rataRataHarian: rataRataHarian,
      rataRataBulanan: rataRataBulanan,
      nominalAkhirTahun: proyeksiAkhirTahun,
      persenTerhadapTarget100: proyeksiPersen100,
      tingkatTertinggiBerpotensi: proyeksiTingkatTertinggi,
      selisihTarget100: proyeksiAkhirTahun - target,
      selisihTarget110: proyeksiAkhirTahun - KPI_TARGET_110,
      selisihTarget120: proyeksiAkhirTahun - KPI_TARGET_120
    }
  };
}

// ------------------------------------------------------------
// TEMP AUDIT — investigasi struktur & formula sheet "Print_ V3"
// (dan sheet-sheet yang direferensikannya) untuk perencanaan menu
// "Penetapan Harga". READ-ONLY, tidak menulis apa pun. Dihapus
// setelah audit selesai.
// ------------------------------------------------------------
function _debugRkInvoiceRaw_(nomorInvoice) {
  const target = String(nomorInvoice || '').trim();
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL).filter(function (r) {
    return String(r.NomorInvoice || '').trim() === target;
  });
  return { nomorInvoice: target, rowCount: rows.length, rows: rows };
}

function _debugRkInvoiceCheck_(buyerSubstr) {
  const needle = String(buyerSubstr || '').toUpperCase();
  const invoices = rkGetInvoiceAggregates_();
  const allocByInvoice = rkGetAllocationsByInvoice_();
  const matchedInvoices = Object.keys(invoices)
    .map(function (k) { return invoices[k]; })
    .filter(function (inv) { return String(inv.buyer || '').toUpperCase().indexOf(needle) > -1; })
    .map(function (inv) {
      const allocs = allocByInvoice[inv.nomorInvoice] || [];
      const total = allocs.reduce(function (s, a) { return s + Number(a.AllocatedAmount || 0); }, 0);
      return {
        nomorInvoice: inv.nomorInvoice, buyer: inv.buyer, tanggal: inv.tanggal, nilai: inv.nilai,
        totalAlokasi: total, outstanding: inv.nilai - total, allocs: allocs
      };
    });

  const mutasiRows = sheetToObjects(RK_SHEETS.MUTASI).filter(function (m) {
    return String(m.NamaPengirim || '').toUpperCase().indexOf(needle) > -1 ||
      String(m.Remark || '').toUpperCase().indexOf(needle) > -1;
  }).map(function (m) {
    return {
      fingerprint: m.Fingerprint, tanggal: m.Tanggal, jam: m.Jam, remark: m.Remark,
      namaPengirim: m.NamaPengirim, credit: m.Credit, status: m.Status,
      potentialInvoices: m.PotentialInvoices, alasan: m.Alasan
    };
  });

  return { buyerSearched: buyerSubstr, invoices: matchedInvoices, mutasi: mutasiRows };
}

function _debugRkPaymentColumn_() {
  const sh = getSheet(SHEET_NAMES.PENJUALAN_FULL);
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerMap = headers.map(function (h, i) { return colLetterLocal_(i + 1) + ' (col ' + (i + 1) + ') = "' + h + '"'; });

  // cari beberapa baris yang kolom PaymentTime-nya terisi (kalau field bernama PaymentTime ada)
  const paymentColIdx = headers.indexOf('PaymentTime'); // 0-based
  let sampleFilled = [];
  if (paymentColIdx > -1 && lastRow > 1) {
    const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (let r = 0; r < values.length && sampleFilled.length < 5; r++) {
      if (values[r][paymentColIdx]) {
        const rowObj = {};
        headers.forEach(function (h, i) { rowObj[colLetterLocal_(i + 1)] = values[r][i]; });
        sampleFilled.push(rowObj);
      }
    }
  }
  return {
    sheetName: SHEET_NAMES.PENJUALAN_FULL, lastCol: lastCol, lastRow: lastRow,
    headerMap: headerMap, columnS_header: headers[18] || null, // S = col 19 = index 18
    columnM_header: headers[12] || null, // M = col 13 = index 12 (PaymentTime per Setup.js)
    sampleRowsWithPaymentFilled: sampleFilled
  };
}
function colLetterLocal_(c) {
  let s = '';
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

function colLetter_(c) {
  let s = '';
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

/**
 * Dump PENUH (values + formulas) sebuah range sebagai array of
 * {cell, value, formula} per baris — supaya gampang dibaca manusia
 * lewat JSON, tanpa perlu hitung index kolom manual.
 */
function dumpSheetGrid_(sheet, startRow, numRows, startCol, numCols) {
  const values = sheet.getRange(startRow, startCol, numRows, numCols).getValues();
  const formulas = sheet.getRange(startRow, startCol, numRows, numCols).getFormulas();
  const rows = [];
  for (let r = 0; r < numRows; r++) {
    const rowCells = {};
    let rowHasData = false;
    for (let c = 0; c < numCols; c++) {
      const v = values[r][c];
      const f = formulas[r][c];
      if (v !== '' || f) {
        rowCells[colLetter_(startCol + c)] = f ? { v: v, f: f } : v;
        rowHasData = true;
      }
    }
    if (rowHasData) rows.push({ row: startRow + r, cells: rowCells });
  }
  return rows;
}

function auditPrintV3Diagnostic() {
  const ss = SpreadsheetApp.openById('13JQCulmX_lCgBYpsabH3gmz6QSLDVKVhhCNesDJmm8o');
  const result = {};

  const printV3 = ss.getSheetByName('Print_ V3');
  if (printV3) {
    result.printV3_full = dumpSheetGrid_(printV3, 1, Math.min(printV3.getLastRow(), 165), 1, Math.min(printV3.getLastColumn(), 12));
  } else {
    result.printV3Error = 'tidak ditemukan';
  }

  const priceV3 = ss.getSheetByName('Price V3');
  if (priceV3) {
    result.priceV3_dims = { rows: priceV3.getLastRow(), cols: priceV3.getLastColumn() };
    // Baris header (asumsi baris 1) + 2 baris data contoh, semua kolom.
    result.priceV3_header = dumpSheetGrid_(priceV3, 1, 1, 1, priceV3.getLastColumn());
    result.priceV3_sample = dumpSheetGrid_(priceV3, 2, 3, 1, priceV3.getLastColumn());
  } else {
    result.priceV3Error = 'tidak ditemukan';
  }

  const priceAntam = ss.getSheetByName('Price ANTAM');
  if (priceAntam) {
    result.priceAntam_dims = { rows: priceAntam.getLastRow(), cols: priceAntam.getLastColumn() };
    result.priceAntam_header = dumpSheetGrid_(priceAntam, 1, 1, 1, priceAntam.getLastColumn());
    result.priceAntam_sample = dumpSheetGrid_(priceAntam, 2, 3, 1, priceAntam.getLastColumn());
  } else {
    result.priceAntamError = 'tidak ditemukan';
  }

  const antam = ss.getSheetByName('Price ANTAM');
  if (antam) {
    result.antam_dataRows = dumpSheetGrid_(antam, 4, 8, 1, antam.getLastColumn());
  }

  const npd = ss.getSheetByName('New Price Discovery');
  if (npd) {
    result.npd_dims = { rows: npd.getLastRow(), cols: npd.getLastColumn() };
    result.npd_rows1to5 = dumpSheetGrid_(npd, 1, 5, 1, Math.min(npd.getLastColumn(), 15));
  } else {
    result.npdError = 'tidak ditemukan';
  }

  const f3 = ss.getSheetByName('F.3.Input_Barang Masuk');
  if (f3) {
    result.f3_dims = { rows: f3.getLastRow(), cols: f3.getLastColumn() };
    // Header ada di baris 11 (per komentar existing getBarangTibaTrend/readF3BarangMasuk_).
    result.f3_headerRow11 = dumpSheetGrid_(f3, 11, 1, 1, f3.getLastColumn());
    result.f3_sampleRows12to14 = dumpSheetGrid_(f3, 12, 3, 1, f3.getLastColumn());
    // Kolom W (23) & AE (31) — dipakai formula Piutang di Print_ V3.
    result.f3_colW_col23_header = f3.getRange(11, 23).getValue();
    result.f3_colAE_col31_header = f3.getRange(11, 31).getValue();
  } else {
    result.f3Error = 'tidak ditemukan';
  }

  return result;
}

// ==============================================================
// REKONSILIASI PEMBAYARAN
// Modul terisolasi: mencocokkan mutasi rekening bank (upload
// CSV/XLSX QLola/BRI) dengan invoice Trading (PenjualanTransaksi).
// TIDAK PERNAH mengubah PenjualanTransaksi / Buyer / sheet lain
// yang sudah ada. Semua state modul ini hidup di 4 sheet baru:
// Bank_Mutasi, Bank_Mutasi_Match, Master_Alias_Buyer, Upload_History.
// ==============================================================

const RK_SHEETS = {
  MUTASI: 'Bank_Mutasi',
  MATCH: 'Bank_Mutasi_Match',
  ALIAS: 'Master_Alias_Buyer',
  UPLOAD_HISTORY: 'Upload_History'
};

const RK_MUTASI_HEADERS = ['ID', 'Fingerprint', 'Tanggal', 'Jam', 'Remark', 'Debit', 'Credit', 'TellerID', 'NamaPengirim', 'Status', 'PotentialInvoices', 'Alasan', 'SourceFile', 'UploadedAt', 'UploadedBy'];
const RK_MATCH_HEADERS = ['ID', 'Fingerprint', 'NomorInvoice', 'AllocatedAmount', 'MatchMethod', 'MatchedAt', 'MatchedBy', 'PreviousStatus', 'NewStatus', 'Notes'];
const RK_ALIAS_HEADERS = ['ID', 'BuyerNama', 'Alias', 'CreatedAt', 'CreatedBy'];
const RK_UPLOAD_HISTORY_HEADERS = ['ID', 'TanggalUpload', 'JamUpload', 'NamaFile', 'PeriodeMulai', 'PeriodeSelesai', 'TotalBaris', 'TransaksiBaru', 'SudahAda', 'GagalDibaca', 'UploadedBy'];

const RK_EPS = 1; // toleransi rupiah untuk pembulatan

function rkEnsureSheet_(name, headers) {
  const ss = getDb();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  // Beberapa panggilan google.script.run bisa jalan hampir bersamaan (mis.
  // upload + refresh dashboard) dan sama-sama melihat sheet belum ada —
  // pakai lock supaya cuma satu yang benar-benar membuatnya.
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    sh = ss.getSheetByName(name);
    if (!sh) {
      try {
        sh = ss.insertSheet(name);
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        formatHeader(sh, headers.length);
      } catch (e) {
        sh = ss.getSheetByName(name); // sheet sudah sempat dibuat proses lain
        if (!sh) throw e;
      }
    }
  } finally {
    lock.releaseLock();
  }
  return sh;
}
function rkMutasiSheet_() { return rkEnsureSheet_(RK_SHEETS.MUTASI, RK_MUTASI_HEADERS); }
function rkMatchSheet_() { return rkEnsureSheet_(RK_SHEETS.MATCH, RK_MATCH_HEADERS); }
function rkAliasSheet_() { return rkEnsureSheet_(RK_SHEETS.ALIAS, RK_ALIAS_HEADERS); }
function rkUploadHistorySheet_() { return rkEnsureSheet_(RK_SHEETS.UPLOAD_HISTORY, RK_UPLOAD_HISTORY_HEADERS); }

function rkFormatTgl_(d) {
  return Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd');
}

function rkParseTanggal_(raw) {
  if (!raw) return '';
  if (raw instanceof Date) return rkFormatTgl_(raw);
  const s = String(raw).trim();
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/); // DD/MM/YYYY
  if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/); // YYYY-MM-DD
  if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})$/); // DD/MM/YY (mis. mutasi BRI QLola: 09/09/26)
  if (m) return (2000 + Number(m[3])) + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  const dt = new Date(s);
  if (!isNaN(dt.getTime())) return rkFormatTgl_(dt);
  return '';
}

function rkParseJam_(raw) {
  if (!raw) return '00:00:00';
  if (raw instanceof Date) return Utilities.formatDate(raw, 'Asia/Jakarta', 'HH:mm:ss');
  let s = String(raw).trim();
  const m = s.match(/^(\d{1,2})[:.](\d{1,2})(?:[:.](\d{1,2}))?/);
  if (m) return m[1].padStart(2, '0') + ':' + m[2].padStart(2, '0') + ':' + (m[3] || '00').padStart(2, '0');
  return '00:00:00';
}

function rkParseNumber_(raw) {
  if (raw === null || raw === undefined || raw === '') return 0;
  if (typeof raw === 'number') return raw;
  let s = String(raw).trim().replace(/[^\d,.\-]/g, '');
  if (!s) return 0;
  // Deteksi konvensi berdasarkan pemisah TERAKHIR yang muncul — supaya cocok
  // baik format Indonesia (1.000.000,00) maupun internasional/bank BRI QLola
  // (1,000,000.00) tanpa perlu asumsi tetap.
  const lastComma = s.lastIndexOf(',');
  const lastDot = s.lastIndexOf('.');
  if (lastComma > -1 && lastDot > -1) {
    if (lastDot > lastComma) s = s.replace(/,/g, ''); // koma = ribuan, titik = desimal
    else s = s.replace(/\./g, '').replace(',', '.'); // titik = ribuan, koma = desimal
  } else if (lastComma > -1) {
    const decimals = s.length - lastComma - 1;
    s = (decimals === 1 || decimals === 2) ? s.replace(',', '.') : s.replace(/,/g, '');
  } else if (lastDot > -1) {
    const decimals = s.length - lastDot - 1;
    if (decimals !== 1 && decimals !== 2) s = s.replace(/\./g, '');
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function rkFingerprint_(tanggal, jam, remark, debit, credit, tellerId) {
  const raw = [tanggal, jam, String(remark || '').trim(), Math.round(debit), Math.round(credit), String(tellerId || '').trim()].join('|');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8);
  return digest.map(function (b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function rkExtractNamaPengirim_(remark) {
  if (!remark) return '';
  let s = String(remark).trim();

  // Pola BI-FAST/BFST TANPA '#' (beda dari "BI-FAST#NAMA#..." di bawah):
  // "BFST<digit referensi><NAMA>:<KODE BANK> <sisa kode...>", mis.
  // "BFST033501003627305PASSION ABAD:CENAIDJA 20260813CENAIDJA010O..."
  // -> "PASSION ABAD". Branch terpisah dengan early-return supaya logic
  // existing di bawah (RTGS#/TRSF#/_TO_/ESB:/dst) sama sekali tidak
  // tersentuh untuk remark yang bukan format ini.
  const bfstMatch = s.match(/^BFST\d+([A-Za-z .,&'\-]+?):[A-Z0-9]+(?:\s|$)/i);
  if (bfstMatch) {
    return bfstMatch[1].replace(/\s+/g, ' ').trim().toUpperCase();
  }

  // Pola "..._TO_<NAMA>,..." (IFT_TO_, RTG_TO_, dst — transfer internal antar
  // rekening dengan referensi nomor PO di depan, format umum BRI QLola).
  const toMatch = s.match(/_TO_([^,]+)/i);
  if (toMatch) {
    s = toMatch[1];
  } else {
    // Pola "RTGS#<NAMA>#...", "TRSF#<NAMA>#...", dst.
    s = s.replace(/^(RTGS|TRSF|TRANSFER|BI-?FAST|OVERBOOKING|KLIRING|SKN|TRF)\s*#\s*/i, '');
    s = s.split('#')[0];
  }

  s = s.split(/\s*ESB\s*:/i)[0];
  s = s.replace(/\b\d{6,}\b/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  return s.toUpperCase();
}

function rkNormalizeName_(s) {
  return String(s || '')
    .toUpperCase()
    .replace(/\bPT\b|\bCV\b|\bTBK\b|\bPERSERO\b/g, '')
    .replace(/[^A-Z0-9 ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function rkGetExistingFingerprints_(sh) {
  const set = {};
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return set;
  const col = RK_MUTASI_HEADERS.indexOf('Fingerprint') + 1;
  const values = sh.getRange(2, col, lastRow - 1, 1).getValues();
  values.forEach(function (r) { if (r[0]) set[r[0]] = true; });
  return set;
}

// ---------- Master Alias Buyer ----------

/**
 * Daftar nama buyer/customer RESMI dari sheet "Master_Customer" (kolom
 * "Nama Customer"), spreadsheet eksternal PRICE_DISCOVERY_SOURCE_ID —
 * dipakai sebagai sumber dropdown "Nama Buyer (resmi)" saat Tambah Alias
 * Buyer (Master Alias Buyer, Rekonsiliasi Pembayaran). Read-only, tidak
 * pernah mengubah sheet Master_Customer.
 */
function getMasterCustomerNames() {
  requireRole('Viewer');
  const ss = SpreadsheetApp.openById(PRICE_DISCOVERY_SOURCE_ID);
  const sheet = ss.getSheetByName('Master_Customer');
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const headerRowIdx = findHeaderRowIndex(values, ['nama', 'customer']);
  if (headerRowIdx === -1) return [];
  const headerMap = buildHeaderMap(values[headerRowIdx]);
  const col = findColByKeywords(headerMap, ['nama', 'customer']);
  if (col === -1) return [];

  const names = [];
  const seen = {};
  values.slice(headerRowIdx + 1).forEach(function (row) {
    const nama = String(getByCol(row, col) || '').trim();
    if (!nama || seen[nama]) return;
    seen[nama] = true;
    names.push(nama);
  });
  names.sort(function (a, b) { return a.localeCompare(b, 'id'); });
  return names;
}

function getMasterAliasBuyer() {
  requireRole('Viewer');
  rkAliasSheet_();
  return sheetToObjects(RK_SHEETS.ALIAS);
}

function addAliasBuyer(buyerNama, alias) {
  const user = requireRole('Supervisor');
  buyerNama = String(buyerNama || '').trim();
  alias = String(alias || '').trim();
  if (!buyerNama || !alias) throw new Error('Nama buyer dan alias wajib diisi.');
  const sh = rkAliasSheet_();
  const id = 'ALIAS-' + new Date().getTime();
  sh.appendRow([id, buyerNama, alias, new Date(), user.Email]);
  logAudit('CREATE', 'Rekonsiliasi.Alias', id, buyerNama + ' -> ' + alias);
  return { success: true, id: id };
}

function deleteAliasBuyer(id) {
  requireRole('Supervisor');
  deleteRowById(RK_SHEETS.ALIAS, 'ID', id);
  logAudit('DELETE', 'Rekonsiliasi.Alias', id, '');
  return { success: true };
}

function rkNormNamaKey_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

// Cocokkan nama buyer dari file Excel dengan Master Buyer resmi
// (Master_Customer — sumber yang sama dipakai dropdown "Nama Buyer
// (resmi)" di form Tambah Alias Buyer, lihat getMasterCustomerNames),
// case-insensitive & spasi berlebih diabaikan. Kembalikan ejaan resmi
// Master_Customer kalau cocok, atau null.
function rkCariBuyerMasterResmi_(namaInput, masterNames) {
  const norm = rkNormNamaKey_(namaInput);
  if (!norm) return null;
  for (let i = 0; i < masterNames.length; i++) {
    if (rkNormNamaKey_(masterNames[i]) === norm) return masterNames[i];
  }
  return null;
}

/**
 * Preview Upload Alias Buyer massal dari Excel — TIDAK menyimpan apa pun.
 * rows = [{buyerNama, alias}] (baris kosong sudah difilter di client).
 * Tiap baris dicocokkan terhadap Master Buyer resmi (Master_Customer) &
 * alias yang sudah ada (Master_Alias_Buyer, sheet/mekanisme yang sama
 * dengan addAliasBuyer) untuk menentukan status VALID / DUPLIKAT /
 * BUYER TIDAK DITEMUKAN.
 */
function previewUploadAliasBuyer(rows) {
  requireRole('Supervisor');
  const masterNames = getMasterCustomerNames();
  rkAliasSheet_();
  const existing = sheetToObjects(RK_SHEETS.ALIAS);
  const existingKeySet = {};
  existing.forEach(function (r) {
    existingKeySet[rkNormNamaKey_(r.BuyerNama) + '||' + rkNormNamaKey_(r.Alias)] = true;
  });

  const seenInBatch = {};
  return (rows || []).map(function (r) {
    const buyerInput = String(r.buyerNama || '').trim();
    const alias = String(r.alias || '').trim();
    const buyerMaster = rkCariBuyerMasterResmi_(buyerInput, masterNames);
    let status;
    if (!buyerMaster) {
      status = 'BUYER TIDAK DITEMUKAN';
    } else {
      const key = rkNormNamaKey_(buyerMaster) + '||' + rkNormNamaKey_(alias);
      if (existingKeySet[key] || seenInBatch[key]) {
        status = 'DUPLIKAT';
      } else {
        status = 'VALID';
        seenInBatch[key] = true;
      }
    }
    return { buyerNamaExcel: buyerInput, alias: alias, buyerMaster: buyerMaster || '', status: status };
  });
}

/**
 * Simpan HANYA baris berstatus VALID hasil previewUploadAliasBuyer (client
 * mengirim buyerNama = ejaan resmi Master_Customer, BUKAN nama dari
 * Excel). Duplikat dicek ulang di sini secara defensif, lalu di-append ke
 * Master_Alias_Buyer dengan mekanisme yang sama seperti addAliasBuyer
 * (ID unik per baris supaya aman dipanggil dalam satu batch/loop).
 */
function saveUploadAliasBuyer(rows) {
  const user = requireRole('Supervisor');
  const sh = rkAliasSheet_();
  const existing = sheetToObjects(RK_SHEETS.ALIAS);
  const existingKeySet = {};
  existing.forEach(function (r) {
    existingKeySet[rkNormNamaKey_(r.BuyerNama) + '||' + rkNormNamaKey_(r.Alias)] = true;
  });

  let inserted = 0, skipped = 0;
  const now = new Date().getTime();
  (rows || []).forEach(function (r, idx) {
    const buyerNama = String(r.buyerNama || '').trim();
    const alias = String(r.alias || '').trim();
    if (!buyerNama || !alias) { skipped++; return; }
    const key = rkNormNamaKey_(buyerNama) + '||' + rkNormNamaKey_(alias);
    if (existingKeySet[key]) { skipped++; return; }
    existingKeySet[key] = true;
    const id = 'ALIAS-' + now + '-' + idx;
    sh.appendRow([id, buyerNama, alias, new Date(), user.Email]);
    logAudit('CREATE', 'Rekonsiliasi.Alias', id, buyerNama + ' -> ' + alias);
    inserted++;
  });

  return { inserted: inserted, skipped: skipped };
}

function rkBuildAliasMap_() {
  const rows = sheetToObjects(RK_SHEETS.ALIAS);
  const map = {};
  rows.forEach(function (r) {
    const buyerKey = rkNormalizeName_(r.BuyerNama);
    map[rkNormalizeName_(r.BuyerNama)] = buyerKey;
    map[rkNormalizeName_(r.Alias)] = buyerKey;
  });
  return map;
}

// ---------- Invoice aggregation (baca PenjualanTransaksi, TIDAK menulis) ----------

// Cut-off go-live Rekonsiliasi Pembayaran: invoice dengan Tanggal Invoice
// SEBELUM tanggal ini dianggap data historis yang sudah selesai/lunas dan
// TIDAK PERNAH masuk ke modul ini (tidak direkonsiliasi, tidak jadi
// kandidat, kolom Payment-nya tidak disentuh). Mulai tanggal ini dan
// seterusnya (dinamis, bukan daftar nomor invoice yang di-hardcode),
// invoice baru otomatis ikut mekanisme normal begitu muncul di
// PenjualanTransaksi.
const RK_CUTOFF_DATE = '2026-09-09';

// Koreksi nilai invoice HANYA untuk modul Rekonsiliasi Pembayaran — TIDAK
// mengubah data di sheet PenjualanTransaksi sama sekali. Dipakai saat kolom
// PphPasal22/TotalHargaAfterTax pada satu baris tertentu tidak konsisten
// dengan pola baris lain milik buyer yang sama (mis. invoice lain buyer ini
// tidak kena PPh 22, tapi baris ini kena — kemungkinan salah input di
// sumber). 'TotalHarga' berarti abaikan PphPasal22/TotalHargaAfterTax untuk
// invoice tsb dan pakai TotalHarga saja sebagai nilai invoice.
// 00755/INV-00017/2026 (PT Jatramas Utama Jaya, 09/09/2026): PphPasal22
// Rp11.665.000 tidak konsisten dgn invoice Jatramas lain (00735, 00743) yang
// PPh-nya kosong — dikoreksi khusus atas konfirmasi user.
const RK_INVOICE_VALUE_OVERRIDE = {
  '00755/INV-00017/2026': 'TotalHarga', // PT Jatramas Utama Jaya — konfirmasi user 09/09/2026
  '00754/INV-00017/2026': 'TotalHarga', // PT Triliun Anugrah Nusantara — konfirmasi user 09/09/2026
  '00756/INV-00017/2026': 'TotalHarga'  // PT Triliun Anugrah Nusantara — konfirmasi user 09/09/2026
};

// RULE MUTLAK: Payment terisi = invoice SELESAI/CLOSED, Payment kosong =
// OUTSTANDING. Rekonsiliasi hanya boleh melihat invoice yang (a) tanggal
// invoice-nya >= RK_CUTOFF_DATE, dan (b) Payment-nya masih kosong — invoice
// historis (sebelum cut-off) atau yang sudah punya nilai Payment (manual
// maupun hasil writeback sebelumnya) dibuang total dari sini, jadi otomatis
// tidak pernah muncul lagi sebagai kandidat/potential match/kandidat
// verifikasi di semua fungsi lain yang memakai peta ini.
//
// testModeDate (opsional, 'yyyy-MM-dd') — HANYA dipakai oleh
// previewRekonsiliasiTestMode() (TEST MODE sementara 13 Agustus 2026):
// invoice yang tanggalnya PERSIS sama dengan testModeDate tetap masuk
// walaupun Payment sudah terisi atau tanggalnya sebelum RK_CUTOFF_DATE.
// Semua pemanggil lain (tanpa argumen) TIDAK terpengaruh sama sekali —
// rule production di atas tetap berlaku 100% seperti sebelumnya.
function rkGetInvoiceAggregates_(testModeDate) {
  const rows = sheetToObjects(SHEET_NAMES.PENJUALAN_FULL);
  const closedInvoices = {};
  rows.forEach(function (r) {
    const inv = String(r.NomorInvoice || '').trim();
    if (inv && String(r.Payment || '').trim()) closedInvoices[inv] = true;
  });
  const map = {};
  rows.forEach(function (r) {
    const inv = String(r.NomorInvoice || '').trim();
    if (!inv) return;
    const tanggal = extractDateOnly(r.Tanggal);
    const isTestModeRow = !!(testModeDate && tanggal === testModeDate);
    if (!isTestModeRow) {
      if (closedInvoices[inv]) return; // Payment sudah terisi -> SELESAI, skip total
      if (tanggal < RK_CUTOFF_DATE) return; // sebelum cut-off -> historical/closed, skip total
    }
    const nilai = RK_INVOICE_VALUE_OVERRIDE[inv] === 'TotalHarga'
      ? Number(r.TotalHarga || 0)
      : Number(r.TotalHargaAfterTax || r.TotalHarga || 0);
    if (!map[inv]) map[inv] = { nomorInvoice: inv, tanggal: tanggal, buyer: r.Pembeli, nilai: 0, testMode: isTestModeRow };
    map[inv].nilai += nilai;
  });
  return map;
}

// Reset hasil match/potential-candidate SEBELUMNYA (dari upload test) untuk
// invoice pada tanggal tertentu — TIDAK menghapus invoice, TIDAK mengubah
// Payment, TIDAK menghapus data Trading, TIDAK mengubah struktur sheet.
// Hanya membersihkan: (1) baris Bank_Mutasi_Match yang menunjuk ke invoice
// target, (2) referensi invoice target di kolom PotentialInvoices pada
// Bank_Mutasi, lalu mengembalikan Status transaksi bank yang jadi "polos"
// (tidak lagi punya match/potential apa pun) ke BELUM_DIPROSES supaya siap
// diproses ulang bersih oleh upload berikutnya.
function rkResetHasilUntukTanggal_(tanggal) {
  const invoices = rkGetInvoiceAggregates_();
  const targetInvoices = {};
  Object.keys(invoices).forEach(function (k) {
    if (invoices[k].tanggal === tanggal) targetInvoices[invoices[k].nomorInvoice] = true;
  });
  const targetList = Object.keys(targetInvoices);
  if (!targetList.length) return { tanggal: tanggal, targetInvoices: [], matchRowsDeleted: 0, mutasiRowsUpdated: 0 };

  // 1) Hapus baris Bank_Mutasi_Match yang NomorInvoice-nya termasuk target
  const matchSh = rkMatchSheet_();
  const matchValues = matchSh.getDataRange().getValues();
  const matchHeaders = matchValues[0];
  const invColIdx = matchHeaders.indexOf('NomorInvoice');
  const fpColIdx = matchHeaders.indexOf('Fingerprint');
  const affectedFingerprints = {};
  const rowsToDelete = [];
  for (let r = 1; r < matchValues.length; r++) {
    const invNo = String(matchValues[r][invColIdx] || '').trim();
    if (targetInvoices[invNo]) {
      rowsToDelete.push(r + 1); // nomor baris 1-based di sheet
      const fp = matchValues[r][fpColIdx];
      if (fp) affectedFingerprints[fp] = true;
    }
  }
  rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (rowNum) { matchSh.deleteRow(rowNum); });

  // 2) Bersihkan PotentialInvoices di Bank_Mutasi yang menunjuk target, dan
  //    kembalikan Status transaksi yang jadi polos ke BELUM_DIPROSES
  const mutasiSh = rkMutasiSheet_();
  const mutasiValues = mutasiSh.getDataRange().getValues();
  const mutasiHeaders = mutasiValues[0];
  const fpCol = mutasiHeaders.indexOf('Fingerprint');
  const statusCol = mutasiHeaders.indexOf('Status');
  const potCol = mutasiHeaders.indexOf('PotentialInvoices');
  const alasanCol = mutasiHeaders.indexOf('Alasan');
  const creditCol = mutasiHeaders.indexOf('Credit');
  let mutasiRowsUpdated = 0;

  for (let r = 1; r < mutasiValues.length; r++) {
    const fp = mutasiValues[r][fpCol];
    const potRaw = String(mutasiValues[r][potCol] || '');
    let potCodes = potRaw ? potRaw.split(',').filter(Boolean) : [];
    const hadTargetPotential = potCodes.some(function (c) { return targetInvoices[c]; });
    const wasAffectedByMatch = !!affectedFingerprints[fp];
    if (!wasAffectedByMatch && !hadTargetPotential) continue;

    potCodes = potCodes.filter(function (c) { return !targetInvoices[c]; });
    const newPotStr = potCodes.join(',');
    const currentStatus = mutasiValues[r][statusCol];
    const currentAlasan = mutasiValues[r][alasanCol];
    const credit = Number(mutasiValues[r][creditCol]) || 0;

    let newStatus = currentStatus, newAlasan = currentAlasan;
    if (credit <= 0) {
      newStatus = 'DEBIT'; newAlasan = '';
    } else if (potCodes.length === 0) {
      newStatus = 'BELUM_DIPROSES'; newAlasan = '';
    }

    mutasiSh.getRange(r + 1, potCol + 1).setValue(newPotStr);
    if (newStatus !== currentStatus) mutasiSh.getRange(r + 1, statusCol + 1).setValue(newStatus);
    if (newAlasan !== currentAlasan) mutasiSh.getRange(r + 1, alasanCol + 1).setValue(newAlasan);
    mutasiRowsUpdated++;
  }

  logAudit('RESET_TEST_DATA', 'Rekonsiliasi', tanggal, 'Reset hasil match/potential untuk invoice tanggal ' + tanggal + ': ' + targetList.join(', '));
  return { tanggal: tanggal, targetInvoices: targetList, matchRowsDeleted: rowsToDelete.length, mutasiRowsUpdated: mutasiRowsUpdated };
}

function resetRekonsiliasiHasilUntukTanggal(tanggal) {
  requireRole('Supervisor');
  return rkResetHasilUntukTanggal_(tanggal);
}

// Reset PENUH data hasil testing modul Rekonsiliasi Pembayaran (tombol admin
// "Reset Data Rekonsiliasi"). Mengosongkan baris di bawah header pada
// Bank_Mutasi, Bank_Mutasi_Match, Upload_History, plus entri AuditTrail
// milik modul ini (Module diawali "Rekonsiliasi"). TIDAK PERNAH menyentuh:
// Master_Alias_Buyer, PenjualanTransaksi (invoice/Payment), atau header
// sheet manapun — hanya clearContent baris data.
function rkResetAllData_(actingUserEmail) {
  const counts = {};

  [RK_SHEETS.MUTASI, RK_SHEETS.MATCH, RK_SHEETS.UPLOAD_HISTORY].forEach(function (name) {
    const sh = getDb().getSheetByName(name);
    if (!sh) { counts[name] = 0; return; }
    const lastRow = sh.getLastRow();
    counts[name] = lastRow > 1 ? lastRow - 1 : 0;
    if (lastRow > 1) sh.deleteRows(2, lastRow - 1);
  });

  // Hanya entri AuditTrail milik modul Rekonsiliasi yang dibuang — sheet
  // AuditTrail dipakai bersama seluruh aplikasi, jangan disentuh selain itu.
  const auditSh = getSheet(SHEET_NAMES.AUDIT);
  const auditValues = auditSh.getDataRange().getValues();
  const auditHeaders = auditValues[0];
  const moduleCol = auditHeaders.indexOf('Module');
  let auditDeleted = 0;
  if (moduleCol > -1) {
    const rowsToDelete = [];
    for (let r = 1; r < auditValues.length; r++) {
      if (String(auditValues[r][moduleCol] || '').indexOf('Rekonsiliasi') === 0) rowsToDelete.push(r + 1);
    }
    rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (rowNum) { auditSh.deleteRow(rowNum); });
    auditDeleted = rowsToDelete.length;
  }

  logAudit('RESET_ALL_TEST_DATA', 'Rekonsiliasi', 'ALL',
    'Reset penuh Bank_Mutasi(' + counts[RK_SHEETS.MUTASI] + ')/Bank_Mutasi_Match(' + counts[RK_SHEETS.MATCH] +
    ')/Upload_History(' + counts[RK_SHEETS.UPLOAD_HISTORY] + ') + ' + auditDeleted + ' entri audit lama, oleh ' + (actingUserEmail || '-'));

  return {
    mutasiDeleted: counts[RK_SHEETS.MUTASI] || 0,
    matchDeleted: counts[RK_SHEETS.MATCH] || 0,
    uploadHistoryDeleted: counts[RK_SHEETS.UPLOAD_HISTORY] || 0,
    auditDeleted: auditDeleted
  };
}

function resetDataRekonsiliasiPembayaran() {
  const user = requireRole('Admin');
  return rkResetAllData_(user.Email);
}

/**
 * TEST MODE SEMENTARA — preview jumlah baris yang akan dihapus oleh
 * resetBankMutasiTanggal(tanggal), dipakai untuk menampilkan angka pasti
 * di dialog konfirmasi SEBELUM benar-benar menghapus apa pun. Read-only.
 */
function previewResetBankMutasiTanggal(tanggal) {
  requireRole('Supervisor');
  const target = extractDateOnly(tanggal);
  if (!target) throw new Error('Tanggal tidak valid.');

  const mutasiRows = sheetToObjects(RK_SHEETS.MUTASI);
  const targetMutasi = mutasiRows.filter(function (r) { return extractDateOnly(r.Tanggal) === target; });
  const fpSet = {};
  targetMutasi.forEach(function (r) { if (r.Fingerprint) fpSet[r.Fingerprint] = true; });

  const matchRows = sheetToObjects(RK_SHEETS.MATCH);
  const targetMatch = matchRows.filter(function (m) { return fpSet[m.Fingerprint]; });

  return { tanggal: target, mutasiCount: targetMutasi.length, matchCount: targetMatch.length };
}

/**
 * TEST MODE SEMENTARA — hapus SELURUH baris Bank_Mutasi bertanggal
 * `tanggal` (kasus: 13/08/2026), beserta baris Bank_Mutasi_Match yang
 * menunjuk ke fingerprint-nya, SUPAYA fingerprint-nya lenyap dari dedup
 * check rkUploadMutasi() dan file CSV yang sama bisa diupload ulang
 * sebagai transaksi BENAR-BENAR BARU — sehingga ikut diproses ulang oleh
 * rkExtractNamaPengirim_ versi TERBARU (parser BFST), bukan cuma di-skip
 * sebagai "sudah ada" seperti yang terjadi sebelumnya.
 *
 * SENGAJA di-scope ketat:
 *  - HANYA baris Bank_Mutasi/Bank_Mutasi_Match bertanggal PERSIS `target`
 *    yang dihapus — tanggal lain (369 baris lainnya) TIDAK disentuh.
 *  - Master_Alias_Buyer TIDAK disentuh sama sekali.
 *  - PenjualanTransaksi/kolom Payment (data invoice) TIDAK disentuh.
 *  - Logic dedup/fingerprint di rkFingerprint_()/rkUploadMutasi() TIDAK
 *    diubah — fungsi ini hanya menghapus DATA-nya, bukan mekanismenya.
 *  - Upload_History TIDAK dihapus (bukan bagian dari dedup check, murni
 *    log riwayat, tidak perlu direset untuk keperluan test ini).
 */
function resetBankMutasiTanggal(tanggal) {
  const user = requireRole('Supervisor');
  const target = extractDateOnly(tanggal);
  if (!target) throw new Error('Tanggal tidak valid.');

  const sh = rkMutasiSheet_();
  const lastRow = sh.getLastRow();
  const fpSet = {};
  const mutasiRowsToDelete = [];
  if (lastRow >= 2) {
    const tglCol = RK_MUTASI_HEADERS.indexOf('Tanggal');
    const fpCol = RK_MUTASI_HEADERS.indexOf('Fingerprint');
    const values = sh.getRange(2, 1, lastRow - 1, RK_MUTASI_HEADERS.length).getValues();
    values.forEach(function (row, idx) {
      if (extractDateOnly(row[tglCol]) !== target) return;
      mutasiRowsToDelete.push(idx + 2); // +2: offset header (row 1) + array 0-based
      if (row[fpCol]) fpSet[row[fpCol]] = true;
    });
  }

  let matchDeleted = 0;
  if (Object.keys(fpSet).length) {
    const matchSh = rkMatchSheet_();
    const matchLastRow = matchSh.getLastRow();
    if (matchLastRow >= 2) {
      const matchFpCol = RK_MATCH_HEADERS.indexOf('Fingerprint');
      const matchValues = matchSh.getRange(2, 1, matchLastRow - 1, RK_MATCH_HEADERS.length).getValues();
      const matchRowsToDelete = [];
      matchValues.forEach(function (row, idx) {
        if (fpSet[row[matchFpCol]]) matchRowsToDelete.push(idx + 2);
      });
      matchRowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (rowNum) { matchSh.deleteRow(rowNum); });
      matchDeleted = matchRowsToDelete.length;
    }
  }

  mutasiRowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (rowNum) { sh.deleteRow(rowNum); });

  logAudit('RESET_TEST_DATA_TANGGAL', 'Rekonsiliasi.Mutasi', target,
    'TEST MODE: hapus ' + mutasiRowsToDelete.length + ' baris Bank_Mutasi + ' + matchDeleted +
    ' baris Bank_Mutasi_Match untuk tanggal ' + target + ' (supaya bisa diupload ulang sebagai transaksi baru), oleh ' + user.Email);

  return { tanggal: target, mutasiDeleted: mutasiRowsToDelete.length, matchDeleted: matchDeleted };
}

function rkGetAllocationsByInvoice_() {
  rkMatchSheet_();
  const rows = sheetToObjects(RK_SHEETS.MATCH);
  const map = {};
  rows.forEach(function (m) {
    const inv = String(m.NomorInvoice || '').trim();
    if (!inv) return;
    (map[inv] = map[inv] || []).push(m);
  });
  return map;
}

function rkComputeInvoiceStatus_(nilai, allocatedTotal) {
  if (allocatedTotal <= RK_EPS) return 'BELUM BAYAR';
  if (allocatedTotal > nilai + RK_EPS) return 'LEBIH BAYAR';
  if (allocatedTotal >= nilai - RK_EPS) return 'LUNAS';
  return 'PARTIAL';
}

function rkPaymentTiming_(tglInvoice, tglLunas) {
  if (!tglInvoice || !tglLunas) return '-';
  const a = new Date(tglInvoice), b = new Date(tglLunas);
  const diffDays = Math.round((b.getTime() - a.getTime()) / 86400000);
  if (diffDays < 0) return 'H' + diffDays;
  return 'H+' + diffDays;
}

function rkAllocate_(newMatchesArr, fingerprint, nomorInvoice, amount, method, matchedBy, prevStatus, newStatus) {
  newMatchesArr.push([
    'MTC-' + Utilities.getUuid(), fingerprint, nomorInvoice, amount, method,
    new Date(), matchedBy, prevStatus || '', newStatus || '', ''
  ]);
}

// Cari kombinasi (subset) invoice dari `items` yang jumlah outstanding-nya
// == target (toleransi RK_EPS) — dipakai saat 1 transfer bank membayar
// beberapa invoice sekaligus. Berhenti begitu menemukan 2 kombinasi
// (cukup untuk tahu apakah hasilnya unik atau ambigu, sekaligus membatasi
// biaya pencarian). Kalau daftar invoice terlalu banyak (>20), jangan coba
// sama sekali — biar aman masuk PERLU VERIFIKASI daripada menebak.
function rkFindSubsetSumCombos_(items, target) {
  const results = [];
  const n = items.length;
  if (n === 0 || n > 20) return results;
  function backtrack(idx, remaining, chosen) {
    if (results.length >= 2) return;
    if (chosen.length > 0 && Math.abs(remaining) <= RK_EPS) {
      results.push(chosen.slice());
    }
    if (idx >= n || remaining < -RK_EPS || results.length >= 2) return;
    chosen.push(items[idx]);
    backtrack(idx + 1, remaining - items[idx].outstanding, chosen);
    chosen.pop();
    backtrack(idx + 1, remaining, chosen);
  }
  backtrack(0, target, []);
  return results;
}

function rkUpdateMutasiRows_(updatesArr) {
  if (!updatesArr.length) return;
  const sh = rkMutasiSheet_();
  const statusCol = RK_MUTASI_HEADERS.indexOf('Status') + 1;
  const potCol = RK_MUTASI_HEADERS.indexOf('PotentialInvoices') + 1;
  const alasanCol = RK_MUTASI_HEADERS.indexOf('Alasan') + 1;
  updatesArr.forEach(function (u) {
    sh.getRange(u.row, statusCol).setValue(u.status);
    sh.getRange(u.row, potCol).setValue(u.potential || '');
    sh.getRange(u.row, alasanCol).setValue(u.alasan || '');
  });
}

// ---------- Reconciliation engine ----------
// Dijalankan otomatis setelah tiap upload, dan bisa dipanggil ulang
// (mis. setelah alias baru ditambahkan) untuk memproses ulang transaksi
// yang statusnya masih belum final.

// ---------- Payment writeback (kolom "Payment" existing di PenjualanTransaksi) ----------
// Modul ini HANYA membaca kolom lain & menulis ke satu kolom existing yang
// header-nya persis "Payment" (dicari by nama header, bukan huruf kolom
// statis — supaya aman walau posisi kolom berubah di kemudian hari).
// Ditulis HANYA saat invoice berstatus LUNAS/LEBIH BAYAR, match valid, dan
// row invoice ditemukan tunggal. Tidak pernah overwrite nilai existing yang
// berbeda — kasus itu dilaporkan sebagai CONFLICT untuk direview manual.

function rkGetPenjualanHeaderIndex_() {
  const sh = getSheet(SHEET_NAMES.PENJUALAN_FULL);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  return { sh: sh, headers: headers, invoiceCol: headers.indexOf('NomorInvoice'), paymentCol: headers.indexOf('Payment'), buyerCol: headers.indexOf('Pembeli') };
}

function rkBuildInvoiceRowMap_() {
  const ctx = rkGetPenjualanHeaderIndex_();
  const map = {};
  if (ctx.invoiceCol === -1 || ctx.paymentCol === -1) return { ctx: ctx, map: map };
  const lastRow = ctx.sh.getLastRow();
  if (lastRow < 2) return { ctx: ctx, map: map };
  const values = ctx.sh.getRange(2, 1, lastRow - 1, ctx.sh.getLastColumn()).getValues();
  values.forEach(function (v, i) {
    const inv = String(v[ctx.invoiceCol] || '').trim();
    if (!inv) return;
    (map[inv] = map[inv] || []).push({ rowNum: i + 2, payment: v[ctx.paymentCol], buyer: ctx.buyerCol > -1 ? v[ctx.buyerCol] : '' });
  });
  return { ctx: ctx, map: map };
}

function rkBuildMutasiByFp_() {
  const map = {};
  sheetToObjects(RK_SHEETS.MUTASI).forEach(function (m) { map[m.Fingerprint] = m; });
  return map;
}

function rkPaymentTimestampStr_(dateObj) {
  return Utilities.formatDate(dateObj, 'Asia/Jakarta', "dd/MM/yy, HH:mm:ss");
}

// Cari transfer yang membuat kumulatif alokasi PERTAMA KALI mencapai nilai
// invoice (bukan asal transfer terakhir — kalau ada alokasi tambahan
// setelah lunas, waktu lunas tetap di titik saat nilai tercapai).
function rkFindSettlementTransfer_(allocsSortedAsc, nilai) {
  let cum = 0;
  for (let i = 0; i < allocsSortedAsc.length; i++) {
    cum += allocsSortedAsc[i].amount;
    if (cum >= nilai - RK_EPS) return allocsSortedAsc[i];
  }
  return allocsSortedAsc.length ? allocsSortedAsc[allocsSortedAsc.length - 1] : null;
}

// Evaluasi (DRY-RUN, tidak menulis apa pun) status writeback kolom Payment
// untuk satu invoice. invoiceRowMap didapat dari rkBuildInvoiceRowMap_().
function rkEvaluateInvoicePayment_(nomorInvoice, invoiceAgg, allocations, mutasiByFp, invoiceRowMap) {
  const totalMasuk = allocations.reduce(function (s, a) { return s + Number(a.AllocatedAmount || 0); }, 0);
  const status = rkComputeInvoiceStatus_(invoiceAgg.nilai, totalMasuk);
  const rows = invoiceRowMap.map[String(nomorInvoice).trim()] || [];

  const result = {
    nomorInvoice: nomorInvoice, invoiceStatus: status, totalMasuk: totalMasuk,
    existingPayment: '', computedPaymentStr: '', matchMethod: '', sourceFingerprint: '',
    writeback: 'NOT_UPDATED - PARTIAL', failReason: ''
  };

  if (invoiceRowMap.ctx.invoiceCol === -1 || invoiceRowMap.ctx.paymentCol === -1) {
    result.writeback = 'FAILED';
    result.failReason = 'Kolom NomorInvoice/Payment tidak ditemukan di sheet PenjualanTransaksi.';
    return result;
  }
  if (rows.length !== 1) {
    result.writeback = rows.length === 0 ? 'FAILED' : 'CONFLICT';
    result.failReason = rows.length === 0
      ? 'Invoice tidak ditemukan di PenjualanTransaksi.'
      : ('Ditemukan ' + rows.length + ' baris dengan Nomor Invoice yang sama — tidak unik.');
    return result;
  }
  result.existingPayment = String(rows[0].payment || '').trim();

  if (status !== 'LUNAS' && status !== 'LEBIH BAYAR') {
    return result; // tetap NOT_UPDATED - PARTIAL (mencakup PARTIAL & BELUM BAYAR)
  }

  const allocsWithTime = allocations.map(function (a) {
    const tx = mutasiByFp[a.Fingerprint];
    return { amount: Number(a.AllocatedAmount || 0), tanggal: tx ? tx.Tanggal : '', jam: tx ? tx.Jam : '', method: a.MatchMethod, fingerprint: a.Fingerprint };
  }).filter(function (a) { return a.tanggal; }).sort(function (a, b) { return (a.tanggal + a.jam).localeCompare(b.tanggal + b.jam); });

  const settlement = rkFindSettlementTransfer_(allocsWithTime, invoiceAgg.nilai);
  if (!settlement) {
    result.writeback = 'NOT_UPDATED - NEED_VERIFICATION';
    result.failReason = 'Tidak dapat menentukan waktu transfer pelunasan.';
    return result;
  }

  const settlementDate = new Date(settlement.tanggal + 'T' + settlement.jam);
  result.computedPaymentStr = rkPaymentTimestampStr_(settlementDate);
  result.matchMethod = settlement.method;
  result.sourceFingerprint = settlement.fingerprint;

  if (result.existingPayment) {
    result.writeback = (result.existingPayment === result.computedPaymentStr) ? 'ALREADY_MATCHED' : 'CONFLICT';
  } else {
    result.writeback = 'PENDING_WRITE';
  }
  return result;
}

// Menulis SATU cell (kolom Payment) untuk row invoice yang sudah divalidasi
// unik oleh rkEvaluateInvoicePayment_. Tidak pernah menyentuh cell lain.
function rkWritePayment_(nomorInvoice, evalResult, invoiceRowMap, actingUser) {
  const rows = invoiceRowMap.map[String(nomorInvoice).trim()] || [];
  if (rows.length !== 1) return Object.assign({}, evalResult, { writeback: rows.length === 0 ? 'FAILED' : 'CONFLICT' });
  const rowNum = rows[0].rowNum;
  invoiceRowMap.ctx.sh.getRange(rowNum, invoiceRowMap.ctx.paymentCol + 1).setValue(evalResult.computedPaymentStr);
  logAudit('UPDATE_PAYMENT', 'Rekonsiliasi.Writeback', nomorInvoice, JSON.stringify({
    buyer: rows[0].buyer, oldPayment: evalResult.existingPayment, newPayment: evalResult.computedPaymentStr,
    sourceTransactionId: evalResult.sourceFingerprint, matchMethod: evalResult.matchMethod, updatedBy: actingUser
  }));
  return Object.assign({}, evalResult, { writeback: 'UPDATED' });
}

// Evaluasi + (kalau autoWrite) langsung tulis. Dipakai baik oleh engine
// otomatis (setelah upload/konfirmasi manual) maupun oleh trigger manual.
function rkProcessPaymentWriteback_(nomorInvoice, invoiceAgg, allocations, mutasiByFp, invoiceRowMap, autoWrite, actingUser) {
  let evalResult = rkEvaluateInvoicePayment_(nomorInvoice, invoiceAgg, allocations, mutasiByFp, invoiceRowMap);
  if (evalResult.writeback === 'PENDING_WRITE') {
    evalResult = autoWrite
      ? rkWritePayment_(nomorInvoice, evalResult, invoiceRowMap, actingUser)
      : Object.assign({}, evalResult, { writeback: 'NOT_UPDATED - PENDING' });
  }
  return evalResult;
}

// Tombol "UPDATE PAYMENT" manual (dipakai saat invoice LUNAS tapi belum
// tertulis, mis. status berubah lewat penambahan alias tanpa upload baru).
// NONAKTIF: kolom "Payment" asli ada di spreadsheet Trading terpisah (bukan
// di PenjualanTransaksi/salinan yang dipakai app ini) — menulis ke sini
// tidak akan pernah terlihat oleh user dan berisiko tertimpa sync. Modul
// Rekonsiliasi HANYA menampilkan status+jam bayar (lihat paymentHasilRekon
// pada getRekonsiliasiPerInvoice/getRekonsiliasiHasilInvoiceDriven); user
// mengisi manual ke sheet Trading aslinya sendiri. Fungsi ini sengaja
// dibiarkan sebagai no-op dry-run, bukan dihapus, kalau nanti alamat sheet
// sumber yang benar sudah dikonfirmasi dan writeback mau diaktifkan lagi.
function triggerPaymentWriteback(nomorInvoice) {
  const user = requireRole('Supervisor');
  const invoices = rkGetInvoiceAggregates_();
  const inv = invoices[nomorInvoice];
  if (!inv) throw new Error('Invoice tidak ditemukan.');
  const allocByInvoice = rkGetAllocationsByInvoice_();
  const mutasiByFp = rkBuildMutasiByFp_();
  const invoiceRowMap = rkBuildInvoiceRowMap_();
  return rkProcessPaymentWriteback_(nomorInvoice, inv, allocByInvoice[nomorInvoice] || [], mutasiByFp, invoiceRowMap, false, user.Email);
}

// NONAKTIF (lihat catatan di atas triggerPaymentWriteback) — tidak pernah
// menulis ke kolom Payment mana pun.
function resolvePaymentConflict(nomorInvoice, action) {
  requireRole('Supervisor');
  const invoices = rkGetInvoiceAggregates_();
  const inv = invoices[nomorInvoice];
  if (!inv) throw new Error('Invoice tidak ditemukan.');
  const allocByInvoice = rkGetAllocationsByInvoice_();
  const mutasiByFp = rkBuildMutasiByFp_();
  const invoiceRowMap = rkBuildInvoiceRowMap_();
  return rkEvaluateInvoicePayment_(nomorInvoice, inv, allocByInvoice[nomorInvoice] || [], mutasiByFp, invoiceRowMap);
}

function rkRunReconciliationEngine_() {
  rkMutasiSheet_();
  const mutasiRows = sheetToObjects(RK_SHEETS.MUTASI);
  const invoices = rkGetInvoiceAggregates_();
  const allocByInvoice = rkGetAllocationsByInvoice_();
  const aliasMap = rkBuildAliasMap_();

  function outstandingOf(inv) {
    const allocs = allocByInvoice[inv.nomorInvoice] || [];
    const total = allocs.reduce(function (s, a) { return s + Number(a.AllocatedAmount || 0); }, 0);
    return inv.nilai - total;
  }

  const byBuyer = {};
  Object.keys(invoices).forEach(function (k) {
    const inv = invoices[k];
    const out = outstandingOf(inv);
    if (out > RK_EPS) {
      const key = rkNormalizeName_(inv.buyer);
      (byBuyer[key] = byBuyer[key] || []).push(Object.assign({}, inv, { outstanding: out }));
    }
  });
  Object.keys(byBuyer).forEach(function (k) { byBuyer[k].sort(function (a, b) { return a.tanggal < b.tanggal ? -1 : 1; }); });

  let autoMatched = 0, perluVerifikasi = 0;
  const newMatches = [];
  const mutasiUpdates = [];

  mutasiRows.forEach(function (tx) {
    const credit = Number(tx.Credit) || 0;
    if (credit <= 0) {
      if (tx.Status !== 'DEBIT') mutasiUpdates.push({ row: tx._row, status: 'DEBIT', potential: '', alasan: 'Transaksi debit, bukan dana masuk' });
      return;
    }
    if (tx.Status === 'MATCHED' || tx.Status === 'MANUAL_VERIFIED') return; // sudah final

    const extracted = tx.NamaPengirim;
    let buyerKey = extracted ? aliasMap[rkNormalizeName_(extracted)] : null;
    if (!buyerKey && extracted) {
      const norm = rkNormalizeName_(extracted);
      if (byBuyer[norm]) buyerKey = norm;
    }

    let status = 'PERLU_VERIFIKASI', reason = 'Nama buyer tidak ditemukan', potential = [];

    if (!extracted) {
      reason = 'Nama pengirim tidak tersedia';
      status = 'UNIDENTIFIED';
      const candidates = [];
      Object.keys(byBuyer).forEach(function (k) {
        byBuyer[k].forEach(function (inv) { if (Math.abs(inv.outstanding - credit) <= RK_EPS) candidates.push(inv); });
      });
      if (candidates.length === 1) {
        status = 'POTENTIAL_MATCH';
        potential = candidates;
        reason = 'Nama pengirim tidak tersedia, ditemukan potensi berdasarkan nominal';
      }
    } else if (buyerKey && byBuyer[buyerKey] && byBuyer[buyerKey].length) {
      const list = byBuyer[buyerKey];
      const exact = list.find(function (inv) { return Math.abs(inv.outstanding - credit) <= RK_EPS; });
      if (exact) {
        rkAllocate_(newMatches, tx.Fingerprint, exact.nomorInvoice, credit, 'AUTO - Nama + Nominal', 'system', tx.Status, 'MATCHED');
        exact.outstanding -= credit;
        status = 'MATCHED';
        autoMatched++;
      } else if (list.length === 1) {
        const inv = list[0];
        const alloc = Math.min(credit, inv.outstanding);
        rkAllocate_(newMatches, tx.Fingerprint, inv.nomorInvoice, alloc, 'AUTO - Nama + Akumulasi', 'system', tx.Status, 'MATCHED');
        inv.outstanding -= alloc;
        status = 'MATCHED';
        autoMatched++;
      } else {
        // 1 transfer bisa membayar beberapa invoice sekaligus — cari
        // kombinasi (subset) invoice yang totalnya persis sama dengan
        // nominal transfer ini.
        const combos = rkFindSubsetSumCombos_(list, credit);
        if (combos.length === 1) {
          const combo = combos[0];
          combo.forEach(function (inv) {
            rkAllocate_(newMatches, tx.Fingerprint, inv.nomorInvoice, inv.outstanding, 'AUTO - Nama + Kombinasi Invoice', 'system', tx.Status, 'MATCHED');
            inv.outstanding = 0;
          });
          status = 'MATCHED';
          autoMatched++;
        } else if (combos.length > 1) {
          status = 'PERLU_VERIFIKASI';
          reason = 'Ditemukan lebih dari satu kombinasi invoice yang cocok dengan nominal transfer';
          const seen = {};
          potential = [];
          combos.forEach(function (c) { c.forEach(function (inv) { if (!seen[inv.nomorInvoice]) { seen[inv.nomorInvoice] = true; potential.push(inv); } }); });
          perluVerifikasi++;
        } else {
          status = 'PERLU_VERIFIKASI';
          reason = 'Buyer memiliki beberapa invoice outstanding';
          potential = list.slice();
          perluVerifikasi++;
        }
      }
    } else if (extracted) {
      reason = 'Nama buyer tidak ditemukan';
      status = 'PERLU_VERIFIKASI';
      perluVerifikasi++;
    }

    mutasiUpdates.push({
      row: tx._row, status: status,
      potential: potential.map(function (i) { return i.nomorInvoice; }).join(','),
      alasan: reason
    });
  });

  if (newMatches.length) {
    const matchSh = rkMatchSheet_();
    matchSh.getRange(matchSh.getLastRow() + 1, 1, newMatches.length, RK_MATCH_HEADERS.length).setValues(newMatches);
  }
  rkUpdateMutasiRows_(mutasiUpdates);

  // Payment writeback: proses invoice yang baru saja mendapat alokasi baru.
  const touchedInvoices = {};
  newMatches.forEach(function (m) { touchedInvoices[m[2]] = true; }); // m[2] = NomorInvoice (lihat RK_MATCH_HEADERS)
  const writebackLog = [];
  if (Object.keys(touchedInvoices).length) {
    const freshAlloc = rkGetAllocationsByInvoice_();
    const mutasiByFpFresh = rkBuildMutasiByFp_();
    const invoiceRowMap = rkBuildInvoiceRowMap_();
    const actingUser = getCurrentUser().Email;
    Object.keys(touchedInvoices).forEach(function (inv) {
      const invAgg = invoices[inv];
      if (!invAgg) return;
      // autoWrite=false: kolom "Payment" asli ada di spreadsheet Trading
      // terpisah (bukan di PenjualanTransaksi/salinan yang dipakai app ini),
      // jadi modul ini hanya MENAMPILKAN status+jam bayar hasil rekonsiliasi
      // — user isi manual ke sheet aslinya sendiri, tidak ditulis otomatis.
      writebackLog.push(rkProcessPaymentWriteback_(inv, invAgg, freshAlloc[inv] || [], mutasiByFpFresh, invoiceRowMap, false, actingUser));
    });
  }

  return { autoMatched: autoMatched, perluVerifikasi: perluVerifikasi, writeback: writebackLog };
}

function rerunRekonsiliasiEngine() {
  requireRole('Supervisor');
  return rkRunReconciliationEngine_();
}

/**
 * TEST MODE SEMENTARA — reparse ulang kolom NamaPengirim di Bank_Mutasi
 * dari kolom Remark memakai rkExtractNamaPengirim_ TERBARU, TERBATAS pada
 * baris bertanggal `tanggal` (default kasus: 13/08/2026) SAJA.
 *
 * Kenapa perlu: rkUploadMutasi() men-skip baris yang fingerprint-nya
 * sudah ada (lihat komentar di dalamnya) SEBELUM NamaPengirim dihitung —
 * jadi baris lama yang sempat ke-parse dengan versi rkExtractNamaPengirim_
 * SEBELUM perbaikan format BFST tidak pernah ter-update otomatis oleh
 * upload ulang file yang sama (fingerprint identik -> selalu dianggap
 * "sudah ada"). Fungsi ini adalah perbaikan data SEKALI JALAN yang
 * di-scope ke tanggal tertentu supaya tidak menyentuh 369 baris lain di
 * luar tanggal itu.
 *
 * HANYA menulis kolom NamaPengirim. TIDAK menyentuh Fingerprint, Tanggal,
 * Jam, Remark, Debit, Credit, TellerID, Status, PotentialInvoices,
 * Alasan, SourceFile, UploadedAt, UploadedBy — dedup/fingerprint logic
 * production TIDAK diubah, TIDAK ADA baris yang dihapus atau ditambah.
 * Baris yang hasil parse-nya SAMA dengan nilai lama tidak ditulis ulang
 * (hemat write, idempotent kalau dipanggil berkali-kali).
 */
function reparseNamaPengirimUntukTanggal(tanggal) {
  requireRole('Supervisor');
  const target = extractDateOnly(tanggal);
  if (!target) throw new Error('Tanggal tidak valid.');

  const sh = rkMutasiSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { checked: 0, updated: 0, rows: [] };

  const tglCol = RK_MUTASI_HEADERS.indexOf('Tanggal') + 1;
  const remarkCol = RK_MUTASI_HEADERS.indexOf('Remark') + 1;
  const creditCol = RK_MUTASI_HEADERS.indexOf('Credit') + 1;
  const namaCol = RK_MUTASI_HEADERS.indexOf('NamaPengirim') + 1;

  const values = sh.getRange(2, 1, lastRow - 1, RK_MUTASI_HEADERS.length).getValues();
  let checked = 0, updated = 0;
  const changedRows = [];

  values.forEach(function (row, idx) {
    const tgl = extractDateOnly(row[tglCol - 1]);
    if (tgl !== target) return;
    checked++;
    const credit = Number(row[creditCol - 1]) || 0;
    if (credit <= 0) return; // baris debit tidak punya NamaPengirim

    const remark = row[remarkCol - 1];
    const namaLama = String(row[namaCol - 1] || '');
    const namaBaru = rkExtractNamaPengirim_(remark);
    if (namaBaru !== namaLama) {
      const sheetRow = idx + 2; // +2: offset header (row 1) + array 0-based
      sh.getRange(sheetRow, namaCol).setValue(namaBaru);
      changedRows.push({ row: sheetRow, remark: String(remark), namaLama: namaLama, namaBaru: namaBaru });
      updated++;
    }
  });

  if (updated) {
    logAudit('REPARSE_NAMA_PENGIRIM', 'Rekonsiliasi.Mutasi', target,
      updated + ' dari ' + checked + ' baris NamaPengirim diperbarui ulang dari Remark (parser BFST terbaru) untuk tanggal ' + target);
  }

  return { checked: checked, updated: updated, rows: changedRows };
}

/**
 * TEST MODE SEMENTARA — Rekonsiliasi Pembayaran invoice Trading tanggal
 * 13 Agustus 2026 (per instruksi user). Simulasi proses matching yang
 * SAMA PERSIS dengan rkRunReconciliationEngine_ (resolusi buyer via
 * Master Alias Buyer, nominal dicocokkan ke Total After Tax, kombinasi
 * subset-sum kalau 1 transfer membayar >1 invoice) untuk invoice pada
 * `tanggal`, TERLEPAS dari RK_CUTOFF_DATE dan status kolom Payment
 * (lihat rkGetInvoiceAggregates_(testModeDate)).
 *
 * MURNI DRY-RUN / READ-ONLY:
 *  - TIDAK PERNAH menulis ke Bank_Mutasi, Bank_Mutasi_Match, atau kolom
 *    Payment — hasil hanya dihitung di memori untuk ditampilkan.
 *  - Baris Bank_Mutasi yang sudah berstatus MATCHED/MANUAL_VERIFIED
 *    (sudah dipakai invoice production lain) TIDAK ikut disimulasikan
 *    ulang di sini, sama seperti aturan rkRunReconciliationEngine_ —
 *    supaya dana yang sudah benar-benar teralokasi ke invoice production
 *    tidak "direbut" simulasi test mode.
 *  - File CSV mutasi bank tetap diupload lewat flow Upload & Ringkasan
 *    existing (menulis ke Bank_Mutasi seperti biasa) — fungsi ini hanya
 *    MEMBACA baris yang sudah ada di sana.
 */
function previewRekonsiliasiTestMode(tanggal) {
  requireRole('Viewer');
  const target = extractDateOnly(tanggal);
  if (!target) throw new Error('Tanggal TEST MODE tidak valid.');

  rkMutasiSheet_();
  const invoices = rkGetInvoiceAggregates_(target);
  const testInvoices = {};
  Object.keys(invoices).forEach(function (k) {
    if (invoices[k].tanggal === target) testInvoices[k] = invoices[k];
  });

  const mutasiRows = sheetToObjects(RK_SHEETS.MUTASI);
  const aliasMap = rkBuildAliasMap_();
  const totalMutasiCount = mutasiRows.filter(function (m) { return Number(m.Credit) > 0; }).length;

  // Grouping per buyer, SAMA seperti rkRunReconciliationEngine_, tapi
  // salinan in-memory (outstanding/allocs) — TIDAK PERNAH ditulis balik.
  const byBuyer = {};
  Object.keys(testInvoices).forEach(function (k) {
    const inv = testInvoices[k];
    const key = rkNormalizeName_(inv.buyer);
    (byBuyer[key] = byBuyer[key] || []).push(Object.assign({}, inv, { outstanding: inv.nilai, allocs: [] }));
  });
  Object.keys(byBuyer).forEach(function (k) { byBuyer[k].sort(function (a, b) { return a.tanggal < b.tanggal ? -1 : 1; }); });

  mutasiRows.forEach(function (tx) {
    const credit = Number(tx.Credit) || 0;
    if (credit <= 0) return;
    if (tx.Status === 'MATCHED' || tx.Status === 'MANUAL_VERIFIED') return; // sudah dipakai invoice production lain

    const extracted = tx.NamaPengirim;
    let buyerKey = extracted ? aliasMap[rkNormalizeName_(extracted)] : null;
    if (!buyerKey && extracted) {
      const norm = rkNormalizeName_(extracted);
      if (byBuyer[norm]) buyerKey = norm;
    }
    if (!buyerKey || !byBuyer[buyerKey] || !byBuyer[buyerKey].length) return;

    const list = byBuyer[buyerKey];
    const exact = list.find(function (inv) { return Math.abs(inv.outstanding - credit) <= RK_EPS; });
    if (exact) {
      exact.allocs.push({ nominal: credit, namaPengirim: extracted, tanggal: tx.Tanggal, jam: tx.Jam });
      exact.outstanding -= credit;
      return;
    }
    if (list.length === 1) {
      const inv = list[0];
      const alloc = Math.min(credit, inv.outstanding);
      if (alloc > 0) {
        inv.allocs.push({ nominal: alloc, namaPengirim: extracted, tanggal: tx.Tanggal, jam: tx.Jam });
        inv.outstanding -= alloc;
      }
      return;
    }
    const combos = rkFindSubsetSumCombos_(list, credit);
    if (combos.length === 1) {
      combos[0].forEach(function (inv) {
        inv.allocs.push({ nominal: inv.outstanding, namaPengirim: extracted, tanggal: tx.Tanggal, jam: tx.Jam });
        inv.outstanding = 0;
      });
    }
  });

  const result = [];
  Object.keys(testInvoices).forEach(function (k) {
    const inv = testInvoices[k];
    const buyerKey = rkNormalizeName_(inv.buyer);
    const simInv = (byBuyer[buyerKey] || []).find(function (x) { return x.nomorInvoice === inv.nomorInvoice; });
    const allocs = simInv ? simInv.allocs.slice().sort(function (a, b) { return (a.tanggal + a.jam).localeCompare(b.tanggal + b.jam); }) : [];
    const totalMasuk = allocs.reduce(function (s, a) { return s + a.nominal; }, 0);

    let status, namaPengirim = '-', tanggalJamTransfer = '-';
    if (totalMasuk > RK_EPS) {
      status = rkComputeInvoiceStatus_(inv.nilai, totalMasuk);
      const last = allocs[allocs.length - 1];
      namaPengirim = last.namaPengirim || '-';
      tanggalJamTransfer = last.tanggal + ' ' + last.jam;
    } else if (totalMutasiCount === 0) {
      status = 'MENUNGGU_REKONSILIASI';
    } else {
      status = 'BELUM_DITEMUKAN';
    }

    result.push({
      nomorInvoice: inv.nomorInvoice, tanggalInvoice: inv.tanggal, buyer: inv.buyer, nilaiInvoice: inv.nilai,
      namaPengirim: namaPengirim, totalDanaMasuk: totalMasuk, jumlahTransfer: allocs.length,
      tanggalJamTransfer: tanggalJamTransfer, selisih: inv.nilai - totalMasuk, status: status,
      testMode: true
    });
  });

  result.sort(function (a, b) { return a.nomorInvoice.localeCompare(b.nomorInvoice); });
  return result;
}

// ---------- Upload mutasi ----------
// rows: array of { tanggal, jam, remark, debit, credit, tellerId } — sudah
// diparse di client (mendukung CSV & XLSX lewat SheetJS) supaya satu jalur
// pemrosesan untuk kedua format file.

function rkUploadMutasi(fileName, rows) {
  const user = requireRole('Supervisor');
  const sh = rkMutasiSheet_();
  const existingFp = rkGetExistingFingerprints_(sh);
  const now = new Date();
  let totalBaris = rows.length, baru = 0, sudahAda = 0, gagal = 0;
  const newRows = [];
  let minDate = '', maxDate = '';

  rows.forEach(function (r) {
    try {
      const tanggal = rkParseTanggal_(r.tanggal);
      if (!tanggal) { gagal++; return; }
      const jam = rkParseJam_(r.jam);
      const remark = String(r.remark || '').trim();
      const debit = rkParseNumber_(r.debit);
      const credit = rkParseNumber_(r.credit);
      const teller = String(r.tellerId || '').trim();
      const fp = rkFingerprint_(tanggal, jam, remark, debit, credit, teller);
      if (existingFp[fp]) { sudahAda++; return; }
      existingFp[fp] = true;
      const nama = credit > 0 ? rkExtractNamaPengirim_(remark) : '';
      newRows.push([
        'BM-' + Utilities.getUuid(), fp, tanggal, jam, remark, debit, credit, teller,
        nama, credit > 0 ? 'BELUM_DIPROSES' : 'DEBIT', '', '', fileName, now, user.Email
      ]);
      baru++;
      if (!minDate || tanggal < minDate) minDate = tanggal;
      if (!maxDate || tanggal > maxDate) maxDate = tanggal;
    } catch (e) {
      gagal++;
    }
  });

  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, RK_MUTASI_HEADERS.length).setValues(newRows);
  }

  const histSh = rkUploadHistorySheet_();
  histSh.appendRow([
    'UP-' + Utilities.getUuid(),
    Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy'),
    Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm:ss'),
    fileName, minDate || '-', maxDate || '-', totalBaris, baru, sudahAda, gagal, user.Email
  ]);

  const recon = rkRunReconciliationEngine_();
  logAudit('UPLOAD', 'Rekonsiliasi.Mutasi', fileName, totalBaris + ' baris, ' + baru + ' baru, ' + sudahAda + ' sudah ada');

  return { totalBaris: totalBaris, baru: baru, sudahAda: sudahAda, gagal: gagal, autoMatched: recon.autoMatched, perluVerifikasi: recon.perluVerifikasi };
}

// ---------- Manual verification ----------

function rkFindMutasiByFingerprint_(fingerprint) {
  const rows = sheetToObjects(RK_SHEETS.MUTASI);
  return rows.find(function (r) { return r.Fingerprint === fingerprint; });
}

function konfirmasiMatchRekonsiliasi(fingerprint, nomorInvoice, amount) {
  const user = requireRole('Supervisor');
  const tx = rkFindMutasiByFingerprint_(fingerprint);
  if (!tx) throw new Error('Transaksi bank tidak ditemukan.');
  const alloc = Number(amount) || Number(tx.Credit) || 0;
  const matchSh = rkMatchSheet_();
  matchSh.appendRow(['MTC-' + Utilities.getUuid(), fingerprint, nomorInvoice, alloc, 'MANUAL_VERIFIED', new Date(), user.Email, tx.Status, 'MANUAL_VERIFIED', '']);
  rkUpdateMutasiRows_([{ row: tx._row, status: 'MANUAL_VERIFIED', potential: '', alasan: 'Dikonfirmasi manual ke ' + nomorInvoice }]);
  logAudit('MANUAL_MATCH', 'Rekonsiliasi', fingerprint, 'Konfirmasi -> ' + nomorInvoice + ' (' + alloc + ')');

  const invoices = rkGetInvoiceAggregates_();
  const invAgg = invoices[nomorInvoice];
  let writeback = null;
  if (invAgg) {
    const freshAlloc = rkGetAllocationsByInvoice_();
    const mutasiByFp = rkBuildMutasiByFp_();
    const invoiceRowMap = rkBuildInvoiceRowMap_();
    writeback = rkProcessPaymentWriteback_(nomorInvoice, invAgg, freshAlloc[nomorInvoice] || [], mutasiByFp, invoiceRowMap, false, user.Email);
  }
  return { success: true, writeback: writeback };
}

function tandaiUnidentifiedRekonsiliasi(fingerprint) {
  const user = requireRole('Supervisor');
  const tx = rkFindMutasiByFingerprint_(fingerprint);
  if (!tx) throw new Error('Transaksi bank tidak ditemukan.');
  rkUpdateMutasiRows_([{ row: tx._row, status: 'UNIDENTIFIED', potential: '', alasan: 'Ditandai tidak teridentifikasi (manual)' }]);
  logAudit('MANUAL_MATCH', 'Rekonsiliasi', fingerprint, 'Ditandai UNIDENTIFIED oleh ' + user.Email);
  return { success: true };
}

// ---------- Read functions untuk dashboard ----------

function rkMatchFilter_(tx, filter) {
  if (!filter) return true;
  if (filter.startDate && tx.Tanggal < filter.startDate) return false;
  if (filter.endDate && tx.Tanggal > filter.endDate) return false;
  if (filter.status && tx.Status !== filter.status) return false;
  if (filter.namaPengirim && String(tx.NamaPengirim || '').toLowerCase().indexOf(String(filter.namaPengirim).toLowerCase()) === -1) return false;
  if (filter.nominalMin && Number(tx.Credit) < Number(filter.nominalMin)) return false;
  if (filter.nominalMax && Number(tx.Credit) > Number(filter.nominalMax)) return false;
  return true;
}

function getRekonsiliasiDashboard(filter) {
  requireRole('Viewer');
  rkMutasiSheet_();
  const mutasi = sheetToObjects(RK_SHEETS.MUTASI).filter(function (r) { return Number(r.Credit) > 0; });
  const matches = sheetToObjects(RK_SHEETS.MATCH);
  const matchByFp = {};
  matches.forEach(function (m) { (matchByFp[m.Fingerprint] = matchByFp[m.Fingerprint] || []).push(m); });
  const invoices = rkGetInvoiceAggregates_();

  const summary = { lunas: 0, partial: 0, perluVerifikasi: 0, tidakDikenali: 0 };
  const rows = [];

  const allocByInvoiceForSummary = rkGetAllocationsByInvoice_();
  Object.keys(invoices).forEach(function (k) {
    const inv = invoices[k];
    const allocs = allocByInvoiceForSummary[inv.nomorInvoice] || [];
    const total = allocs.reduce(function (s, a) { return s + Number(a.AllocatedAmount || 0); }, 0);
    const st = rkComputeInvoiceStatus_(inv.nilai, total);
    if (st === 'LUNAS') summary.lunas++;
    else if (st === 'PARTIAL' || st === 'LEBIH BAYAR') summary.partial++;
  });

  mutasi.filter(function (tx) { return rkMatchFilter_(tx, filter); }).forEach(function (tx) {
    const allocs = matchByFp[tx.Fingerprint] || [];
    let matchedInvoice = '-', buyer = '-', method = '-';
    if (allocs.length) {
      matchedInvoice = allocs.map(function (a) { return a.NomorInvoice; }).join(', ');
      method = allocs[allocs.length - 1].MatchMethod;
      const firstInv = invoices[allocs[0].NomorInvoice];
      buyer = firstInv ? firstInv.buyer : '-';
    }
    if (filter && filter.buyer && String(buyer).toLowerCase().indexOf(String(filter.buyer).toLowerCase()) === -1) return;
    if (filter && filter.noInvoice && matchedInvoice.toLowerCase().indexOf(String(filter.noInvoice).toLowerCase()) === -1) return;

    rows.push({
      fingerprint: tx.Fingerprint, tanggal: tx.Tanggal, jam: tx.Jam, remark: tx.Remark,
      namaPengirim: tx.NamaPengirim, nominal: Number(tx.Credit), matchedInvoice: matchedInvoice,
      buyer: buyer, status: tx.Status, method: method, alasan: tx.Alasan,
      potentialInvoices: tx.PotentialInvoices ? String(tx.PotentialInvoices).split(',').filter(Boolean) : []
    });

    if (tx.Status === 'PERLU_VERIFIKASI' || tx.Status === 'POTENTIAL_MATCH') summary.perluVerifikasi++;
    else if (tx.Status === 'UNIDENTIFIED') summary.tidakDikenali++;
  });

  rows.sort(function (a, b) { return (b.tanggal + b.jam).localeCompare(a.tanggal + a.jam); });
  return { summary: summary, rows: rows };
}

// ---------- Tab "Hasil Rekonsiliasi" — INVOICE-DRIVEN ----------
// Baris utama adalah invoice Trading yang Payment-nya kosong (lihat
// rkGetInvoiceAggregates_ — sudah memfilter invoice Payment terisi keluar
// total). Mutasi bank di sini hanya jadi "bukti" yang mengisi kolom kanan
// (Nama Pengirim, Total Dana Masuk, Jumlah Transfer, dst), bukan baris
// utama. Status per invoice:
//   LUNAS/PARTIAL/LEBIH BAYAR : sudah ada alokasi dana masuk
//   PERLU_VERIFIKASI          : ada transaksi bank yang menyebut invoice ini
//                                sebagai salah satu kandidat, tapi belum
//                                dikonfirmasi user
//   BELUM_DITEMUKAN           : sudah ada data mutasi ter-upload tapi tidak
//                                ada satu pun yang cocok untuk invoice ini
//   MENUNGGU_REKONSILIASI     : belum ada mutasi yang pernah diupload sama
//                                sekali (Bank_Mutasi masih kosong)
function getRekonsiliasiHasilInvoiceDriven(filter) {
  requireRole('Viewer');
  rkMutasiSheet_();
  const invoices = rkGetInvoiceAggregates_(); // sudah hanya invoice Payment kosong
  const allocByInvoice = rkGetAllocationsByInvoice_();
  const mutasiRows = sheetToObjects(RK_SHEETS.MUTASI);
  const mutasiByFp = {};
  mutasiRows.forEach(function (m) { mutasiByFp[m.Fingerprint] = m; });
  const totalMutasiCount = mutasiRows.filter(function (m) { return Number(m.Credit) > 0; }).length;

  const potentialByInvoice = {};
  mutasiRows.forEach(function (m) {
    if ((m.Status === 'PERLU_VERIFIKASI' || m.Status === 'POTENTIAL_MATCH') && m.PotentialInvoices) {
      String(m.PotentialInvoices).split(',').filter(Boolean).forEach(function (code) {
        (potentialByInvoice[code] = potentialByInvoice[code] || []).push({
          fingerprint: m.Fingerprint, namaPengirim: m.NamaPengirim, nominal: Number(m.Credit),
          tanggal: m.Tanggal, jam: m.Jam, remark: m.Remark, alasan: m.Alasan
        });
      });
    }
  });

  const result = [];
  Object.keys(invoices).forEach(function (k) {
    const inv = invoices[k];
    if (filter && filter.buyer && String(inv.buyer).toLowerCase().indexOf(String(filter.buyer).toLowerCase()) === -1) return;
    if (filter && filter.noInvoice && inv.nomorInvoice.toLowerCase().indexOf(String(filter.noInvoice).toLowerCase()) === -1) return;
    if (filter && filter.startDate && inv.tanggal < filter.startDate) return;
    if (filter && filter.endDate && inv.tanggal > filter.endDate) return;

    const allocs = (allocByInvoice[inv.nomorInvoice] || []).slice().sort(function (a, b) {
      const ta = mutasiByFp[a.Fingerprint], tb = mutasiByFp[b.Fingerprint];
      const da = ta ? ta.Tanggal + ta.Jam : '', db = tb ? tb.Tanggal + tb.Jam : '';
      return da.localeCompare(db);
    });
    const totalMasuk = allocs.reduce(function (s, a) { return s + Number(a.AllocatedAmount || 0); }, 0);
    const potentials = potentialByInvoice[inv.nomorInvoice] || [];

    let status, namaPengirim = '-', tanggalJamTransfer = '-';
    if (totalMasuk > RK_EPS) {
      status = rkComputeInvoiceStatus_(inv.nilai, totalMasuk);
      const last = mutasiByFp[allocs[allocs.length - 1].Fingerprint];
      namaPengirim = last ? last.NamaPengirim : '-';
      tanggalJamTransfer = last ? (last.Tanggal + ' ' + last.Jam) : '-';
    } else if (potentials.length) {
      status = 'PERLU_VERIFIKASI';
      namaPengirim = potentials[0].namaPengirim || '-';
    } else if (totalMutasiCount === 0) {
      status = 'MENUNGGU_REKONSILIASI';
    } else {
      status = 'BELUM_DITEMUKAN';
    }

    if (filter && filter.status && status !== filter.status) return;
    if (filter && filter.namaPengirim && String(namaPengirim).toLowerCase().indexOf(String(filter.namaPengirim).toLowerCase()) === -1) return;

    result.push({
      nomorInvoice: inv.nomorInvoice, tanggalInvoice: inv.tanggal, buyer: inv.buyer, nilaiInvoice: inv.nilai,
      namaPengirim: namaPengirim, totalDanaMasuk: totalMasuk, jumlahTransfer: allocs.length,
      tanggalJamTransfer: tanggalJamTransfer, selisih: inv.nilai - totalMasuk, status: status,
      // Kalau invoice ini cuma punya SATU alokasi, simpan fingerprint transfer
      // sumbernya — dipakai frontend untuk mengelompokkan invoice-invoice
      // yang dibayar oleh satu transaksi bank yang sama ("1 transfer -> N invoice").
      singleTransferFingerprint: allocs.length === 1 ? allocs[0].Fingerprint : null,
      potentialCandidates: potentials
    });
  });

  result.sort(function (a, b) { return b.tanggalInvoice.localeCompare(a.tanggalInvoice); });
  return result;
}

function getRekonsiliasiPerInvoice(filter) {
  requireRole('Viewer');
  rkMutasiSheet_();
  const invoices = rkGetInvoiceAggregates_();
  const allocByInvoice = rkGetAllocationsByInvoice_();
  const mutasiByFp = rkBuildMutasiByFp_();
  const invoiceRowMap = rkBuildInvoiceRowMap_(); // dry-run saja di sini — tidak menulis

  const result = [];
  Object.keys(invoices).forEach(function (k) {
    const inv = invoices[k];
    if (filter && filter.buyer && String(inv.buyer).toLowerCase().indexOf(String(filter.buyer).toLowerCase()) === -1) return;
    if (filter && filter.noInvoice && inv.nomorInvoice.toLowerCase().indexOf(String(filter.noInvoice).toLowerCase()) === -1) return;
    if (filter && filter.startDate && inv.tanggal < filter.startDate) return;
    if (filter && filter.endDate && inv.tanggal > filter.endDate) return;

    const allocs = (allocByInvoice[inv.nomorInvoice] || []).slice().sort(function (a, b) {
      const ta = mutasiByFp[a.Fingerprint], tb = mutasiByFp[b.Fingerprint];
      const da = ta ? ta.Tanggal + ta.Jam : '', db = tb ? tb.Tanggal + tb.Jam : '';
      return da.localeCompare(db);
    });
    const totalMasuk = allocs.reduce(function (s, a) { return s + Number(a.AllocatedAmount || 0); }, 0);
    const status = rkComputeInvoiceStatus_(inv.nilai, totalMasuk);
    if (filter && filter.status && status !== filter.status) return;

    let transferPertama = '-', transferTerakhir = '-', paymentTiming = '-';
    if (allocs.length) {
      const first = mutasiByFp[allocs[0].Fingerprint];
      const last = mutasiByFp[allocs[allocs.length - 1].Fingerprint];
      transferPertama = first ? (first.Tanggal + ' ' + first.Jam) : '-';
      if (status === 'LUNAS' && last) {
        transferTerakhir = last.Tanggal + ' ' + last.Jam;
        paymentTiming = rkPaymentTiming_(inv.tanggal, last.Tanggal);
      } else if (last) {
        transferTerakhir = last.Tanggal + ' ' + last.Jam;
      }
    }

    const paymentEval = rkEvaluateInvoicePayment_(inv.nomorInvoice, inv, allocByInvoice[inv.nomorInvoice] || [], mutasiByFp, invoiceRowMap);

    result.push({
      nomorInvoice: inv.nomorInvoice, tanggalInvoice: inv.tanggal, buyer: inv.buyer, nilaiInvoice: inv.nilai,
      totalDanaMasuk: totalMasuk, outstanding: inv.nilai - totalMasuk, jumlahTransfer: allocs.length,
      transferPertama: transferPertama, transferTerakhir: transferTerakhir, paymentTiming: paymentTiming, status: status,
      paymentExisting: paymentEval.existingPayment || '-',
      paymentHasilRekon: paymentEval.computedPaymentStr || '-',
      paymentWriteback: paymentEval.writeback === 'PENDING_WRITE' ? 'NOT_UPDATED - PENDING' : paymentEval.writeback
    });
  });

  result.sort(function (a, b) { return b.tanggalInvoice.localeCompare(a.tanggalInvoice); });
  return result;
}

function getRekonsiliasiInvoiceDetail(nomorInvoice) {
  requireRole('Viewer');
  const invoices = rkGetInvoiceAggregates_();
  const inv = invoices[nomorInvoice];
  if (!inv) throw new Error('Invoice tidak ditemukan.');
  const allocByInvoice = rkGetAllocationsByInvoice_();
  const mutasiByFp = {};
  sheetToObjects(RK_SHEETS.MUTASI).forEach(function (m) { mutasiByFp[m.Fingerprint] = m; });

  const allocs = (allocByInvoice[nomorInvoice] || []).map(function (a) {
    const tx = mutasiByFp[a.Fingerprint] || {};
    return {
      tanggal: tx.Tanggal || '-', jam: tx.Jam || '-', remark: tx.Remark || '-',
      amount: Number(a.AllocatedAmount || 0), method: a.MatchMethod, matchedBy: a.MatchedBy, matchedAt: a.MatchedAt
    };
  }).sort(function (a, b) { return (a.tanggal + a.jam).localeCompare(b.tanggal + b.jam); });

  const totalMasuk = allocs.reduce(function (s, a) { return s + a.amount; }, 0);
  const status = rkComputeInvoiceStatus_(inv.nilai, totalMasuk);
  return {
    nomorInvoice: inv.nomorInvoice, tanggalInvoice: inv.tanggal, buyer: inv.buyer, nilaiInvoice: inv.nilai,
    totalDanaMasuk: totalMasuk, outstanding: inv.nilai - totalMasuk, jumlahTransfer: allocs.length,
    status: status, transfers: allocs,
    pelunasan: (status === 'LUNAS' && allocs.length) ? (allocs[allocs.length - 1].tanggal + ' ' + allocs[allocs.length - 1].jam) : '-'
  };
}

function getRekonsiliasiPerluVerifikasi(filter) {
  requireRole('Viewer');
  rkMutasiSheet_();
  const invoices = rkGetInvoiceAggregates_();
  const rows = sheetToObjects(RK_SHEETS.MUTASI).filter(function (r) {
    return ['PERLU_VERIFIKASI', 'POTENTIAL_MATCH', 'UNIDENTIFIED'].indexOf(r.Status) > -1 && rkMatchFilter_(r, filter);
  });

  return rows.map(function (tx) {
    const potentialInvoiceCodes = tx.PotentialInvoices ? String(tx.PotentialInvoices).split(',').filter(Boolean) : [];
    const potentialCandidates = potentialInvoiceCodes.map(function (code) {
      const inv = invoices[code];
      return inv ? { nomorInvoice: inv.nomorInvoice, buyer: inv.buyer, nilai: inv.nilai, tanggal: inv.tanggal } : null;
    }).filter(Boolean);

    let potentialBuyer = '-';
    if (potentialCandidates.length) potentialBuyer = potentialCandidates[0].buyer;
    else if (tx.NamaPengirim) potentialBuyer = tx.NamaPengirim;

    return {
      fingerprint: tx.Fingerprint, tanggal: tx.Tanggal, jam: tx.Jam, remark: tx.Remark,
      nominal: Number(tx.Credit), namaPengirim: tx.NamaPengirim || '-', potentialBuyer: potentialBuyer,
      potentialInvoices: potentialCandidates, alasan: tx.Alasan, status: tx.Status
    };
  }).sort(function (a, b) { return (b.tanggal + b.jam).localeCompare(a.tanggal + a.jam); });
}

function getRiwayatUploadMutasi() {
  requireRole('Viewer');
  rkUploadHistorySheet_();
  return sheetToObjects(RK_SHEETS.UPLOAD_HISTORY).sort(function (a, b) {
    return (b.TanggalUpload + b.JamUpload).localeCompare(a.TanggalUpload + a.JamUpload);
  });
}

function getRekonsiliasiInvoiceOptions() {
  requireRole('Viewer');
  const invoices = rkGetInvoiceAggregates_();
  return Object.keys(invoices).map(function (k) {
    const inv = invoices[k];
    return { nomorInvoice: inv.nomorInvoice, buyer: inv.buyer, nilai: inv.nilai, tanggal: inv.tanggal };
  }).sort(function (a, b) { return b.tanggal.localeCompare(a.tanggal); });
}


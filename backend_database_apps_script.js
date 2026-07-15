// PERHATIAN: Pastikan file HTML Anda bernama 'index.html'
// =========================================================================
// JIKA ANDA MEMBUAT SCRIPT INI SECARA TERPISAH (STANDALONE SCRIPT),
// ANDA WAJIB MENGISI ID SPREADSHEET DI BAWAH INI:
// Contoh: const SPREADSHEET_ID = "1A2B3c4d5E6F7g8H9i0J_kLmNoP";
const SPREADSHEET_ID = ""; // Kosongkan HANYA JIKA ini Bound Script (Menu Ekstensi -> Apps Script di Sheets)
// =========================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('LPG Control System - Gacoan')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDB() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } catch (e) {
      throw new Error("ID SPREADSHEET SALAH atau Anda tidak memiliki akses ke Spreadsheet tersebut.");
    }
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("SCRIPT TIDAK MENGENALI DATABASE! Alasan: Anda mendeploy Script ini sebagai 'Standalone' tanpa ID Spreadsheet. Solusi: Isi variabel SPREADSHEET_ID di baris paling atas Code.gs.");
  }
  return ss;
}

function checkAndInitSheets() {
  var ss = getDB();
  var sheets = {
    "MasterTabung": ["ID", "Label", "EmptyWeight", "GasCapacity", "CurrentGrossWeight", "LastUpdate"],
    "Logs": ["ID", "Timestamp", "TabungID", "Label", "Shift", "GrossWeight", "NetGas", "Stocker", "Status"],
    "StockOpname": ["ID", "Date", "Type", "Supplier", "Qty12KG", "TotalBiaya", "ReturQty", "SelisihRetur", "Notes", "PIC"],
    "DailyRekap": ["Tanggal", "Beli12KG", "TotalBiaya", "ReturQty", "SelisihRetur", "GasTerpakai", "SisaPenuh", "SisaTerpakai", "SisaKosong", "PIC", "WaktuSimpan", "Notes"]
  };

  for (var name in sheets) {
    if (!ss.getSheetByName(name)) {
      var sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
      if (name === "MasterTabung") {
        sheet.appendRow(["LPG-001", "Dapur Utama Wok 1", 15.0, 12.0, 27.0, getFormattedDateTime()]);
        sheet.appendRow(["LPG-002", "Stasiun Penggorengan 1", 15.0, 12.0, 20.0, getFormattedDateTime()]);
      }
    }
  }
  return true;
}

// 🚀 FUNGSI BARU (Mengurangi Loading 4x lipat!)
// Ini mencegah error/timeout karena menarik semua data di 1 panggilan jaringan
function getInitialData() {
  try {
    checkAndInitSheets();
    return {
      masterTabung: getMasterTabung(),
      logs: getLogs(),
      opname: getOpnameLogs(),
      rekap: getDailyRekap()
    };
  } catch (e) {
    throw new Error("Gagal mengambil data dari Spreadsheet: " + e.message);
  }
}

function getFormattedDateTime() {
  return Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm");
}

function getMasterTabung() {
  var data = getDB().getSheetByName("MasterTabung").getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0].toString(), label: data[i][1].toString(),
      emptyWeight: parseFloat(data[i][2]), gasCapacity: parseFloat(data[i][3]),
      currentGrossWeight: parseFloat(data[i][4]), lastUpdate: data[i][5].toString()
    });
  }
  return result;
}

function getLogs() {
  var data = getDB().getSheetByName("Logs").getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0].toString(), timestamp: data[i][1].toString(), tabungId: data[i][2].toString(),
      label: data[i][3].toString(), shift: data[i][4].toString(), grossWeight: parseFloat(data[i][5]),
      netGas: parseFloat(data[i][6]), stocker: data[i][7].toString(), status: data[i][8].toString()
    });
  }
  return result.reverse(); 
}

function getOpnameLogs() {
  var data = getDB().getSheetByName("StockOpname").getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    result.push({
      id: data[i][0].toString(), date: data[i][1].toString(), type: data[i][2].toString(),
      supplier: data[i][3].toString(), qty12: parseInt(data[i][4]) || 0,
      totalBiaya: parseFloat(data[i][5]) || 0, returQty: parseInt(data[i][6]) || 0,
      selisihRetur: parseInt(data[i][7]) || 0, notes: data[i][8].toString(), pic: data[i][9].toString()
    });
  }
  return result.reverse();
}

function getDailyRekap() {
  var data = getDB().getSheetByName("DailyRekap").getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    result.push({
      tanggal: data[i][0].toString(), beli12: parseInt(data[i][1]) || 0, totalBiaya: parseFloat(data[i][2]) || 0,
      returQty: parseInt(data[i][3]) || 0, selisihRetur: parseInt(data[i][4]) || 0, gasTerpakai: parseFloat(data[i][5]) || 0,
      sisaPenuh: parseInt(data[i][6]) || 0, sisaTerpakai: parseInt(data[i][7]) || 0, sisaKosong: parseInt(data[i][8]) || 0,
      pic: data[i][9].toString(), waktuSimpan: data[i][10].toString(), notes: data[i][11] ? data[i][11].toString() : ""
    });
  }
  return result.reverse();
}

// ------------------------------------------------------------------
// FUNGSI SUBMIT (Ditambahkan penanganan Error DB)
// ------------------------------------------------------------------
function registerNewTabungBackend(id, label, emptyWeight, gasCapacity) {
  var sheet = getDB().getSheetByName("MasterTabung");
  sheet.appendRow([id, label, emptyWeight, gasCapacity, emptyWeight + gasCapacity, getFormattedDateTime()]);
  return true;
}

function submitLpgLogBackend(tabungId, currentGross, shift, stocker) {
  var ss = getDB();
  var masterSheet = ss.getSheetByName("MasterTabung");
  var logSheet = ss.getSheetByName("Logs");
  var masterData = masterSheet.getDataRange().getValues();
  
  var label = "", emptyWt = 0, capacity = 0, netGas = 0, status = "Terpakai";
  
  for (var i = 1; i < masterData.length; i++) {
    if (masterData[i][0].toString() === tabungId) {
      label = masterData[i][1].toString(); emptyWt = parseFloat(masterData[i][2]); capacity = parseFloat(masterData[i][3]);
      netGas = currentGross - emptyWt; var pct = (netGas / capacity) * 100;
      if (pct >= 80) status = "Penuh"; else if (pct < 15) status = "Kosong";
      masterSheet.getRange(i + 1, 5).setValue(currentGross); masterSheet.getRange(i + 1, 6).setValue(getFormattedDateTime());
      break;
    }
  }
  var logId = "L-" + new Date().getTime();
  logSheet.appendRow([logId, getFormattedDateTime(), tabungId, label, shift, currentGross, netGas, stocker, status]);
  return true;
}

function deleteLogBackend(logId) {
  var sheet = getDB().getSheetByName("Logs"); var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][0].toString() === logId) { sheet.deleteRow(i + 1); return true; } }
  throw new Error("Log tidak ditemukan");
}

function submitOpnameLogBackend(date, supplier, qty12, totalBiaya, returQty, selisih, notes, pic) {
  getDB().getSheetByName("StockOpname").appendRow(["SO-" + new Date().getTime(), date, "KEDATANGAN", supplier, qty12, totalBiaya, returQty, selisih, notes, pic]);
  return true;
}

function deleteOpnameLogBackend(id) {
  var sheet = getDB().getSheetByName("StockOpname"); var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][0].toString() === id) { sheet.deleteRow(i + 1); return true; } }
}

function submitDailyClosingBackend(tanggal, beli12, totalBiaya, returQty, selisihRetur, gasTerpakai, sisaPenuh, sisaTerpakai, sisaKosong, pic, notes) {
  var sheet = getDB().getSheetByName("DailyRekap"); var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === tanggal) {
      sheet.getRange(i + 1, 2, 1, 11).setValues([[beli12, totalBiaya, returQty, selisihRetur, gasTerpakai, sisaPenuh, sisaTerpakai, sisaKosong, pic, getFormattedDateTime(), notes]]);
      return true;
    }
  }
  sheet.appendRow([tanggal, beli12, totalBiaya, returQty, selisihRetur, gasTerpakai, sisaPenuh, sisaTerpakai, sisaKosong, pic, getFormattedDateTime(), notes]);
  return true;
}

function deleteDailyRekapBackend(tanggal) {
  var sheet = getDB().getSheetByName("DailyRekap"); var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][0].toString() === tanggal) { sheet.deleteRow(i + 1); return true; } }
}
function doGet(e) {
  // 1. Jika Plesk meminta halaman tampilan utama (index.html GAS)
  if (e && e.parameter && e.parameter.action === 'getPage') {
    var htmlContent = HtmlService.createTemplateFromFile('Index').evaluate().getContent();
    return ContentService.createTextOutput(htmlContent).setMimeType(ContentService.MimeType.TEXT);
  }
  
  // 2. Jika form meminta daftar dosen peserta
  if (e && e.parameter && e.parameter.action === 'getDosen') {
    var list = getDaftarDosen();
    return ContentService.createTextOutput(JSON.stringify(list)).setMimeType(ContentService.MimeType.JSON);
  }
  
  // Fallback standar jika diakses langsung
  return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('RSVP Capacity Building Dosen')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  // Menangani pengiriman data form dari web
  try {
    var formObj = JSON.parse(e.postData.contents);
    var result = simpanRSVP(formObj);
    return ContentService.createTextOutput(JSON.stringify({ status: result })).setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "Error: " + error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Mengambil daftar nama dari sheet "Peserta"
function getDaftarDosen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPeserta = ss.getSheetByName('Peserta');
  var sheetRekap = ss.getSheetByName('Rekap_RSVP');
  
  if (!sheetPeserta) return ["Error: Sheet 'Peserta' tidak ditemukan"];
  
  var dataPeserta = sheetPeserta.getRange(1, 1, sheetPeserta.getLastRow(), 1).getValues();
  var listPeserta = dataPeserta.map(function(row) { return row[0]; });
  
  var listRekap = [];
  if (sheetRekap && sheetRekap.getLastRow() > 1) {
    var dataRekap = sheetRekap.getRange(2, 2, sheetRekap.getLastRow() - 1, 1).getValues();
    listRekap = dataRekap.map(function(row) { return row[0]; });
  }
  
  var listTersedia = listPeserta.filter(function(dosen) {
    return listRekap.indexOf(dosen) === -1 && dosen.trim() !== "";
  });
  
  return listTersedia;
}

// Menyimpan data hasil input dari form web
function simpanRSVP(formObj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetRekap = ss.getSheetByName('Rekap_RSVP');
    
    // PERBAIKAN: Jika sheet Rekap_RSVP tidak ada, buatkan otomatis
    if (!sheetRekap) {
      sheetRekap = ss.insertSheet('Rekap_RSVP');
      // Buat header otomatis
      sheetRekap.appendRow(['Timestamp', 'Nama Dosen', 'Kehadiran H1 (Sabtu)', 'Alasan H1', 'Kehadiran H2 (Minggu)', 'Alasan H2', 'Catatan']);
    }
    
    // Validasi alasan
    var alasanH1 = (formObj.h1 === 'Tidak Hadir' || formObj.h1 === 'Tentatif') ? (formObj.alasanH1 || '-') : '-';
    var alasanH2 = (formObj.h2 === 'Tidak Hadir' || formObj.h2 === 'Tentatif') ? (formObj.alasanH2 || '-') : '-';
    
    sheetRekap.appendRow([
      new Date(), // Timestamp
      formObj.namaDosen,
      formObj.h1,
      alasanH1,
      formObj.h2,
      alasanH2,
      formObj.catatan || '-'
    ]);
    
    return "Sukses";
  } catch(e) {
    return "Error: " + e.toString();
  }
}

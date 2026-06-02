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

// Mengambil daftar nama dari sheet "Peserta" beserta status RSVP
function getDaftarDosen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPeserta = ss.getSheetByName('Peserta');
  var sheetRekap = ss.getSheetByName('Rekap_RSVP');
  
  if (!sheetPeserta) return [{nama: "Error: Sheet 'Peserta' tidak ditemukan", sudahIsi: false}];
  
  var dataPeserta = sheetPeserta.getRange(1, 1, sheetPeserta.getLastRow(), 1).getValues();
  
  var rekapMap = {};
  // Mapping data yang sudah masuk di Rekap_RSVP
  if (sheetRekap && sheetRekap.getLastRow() > 1) {
    var dataRekap = sheetRekap.getRange(2, 1, sheetRekap.getLastRow() - 1, 7).getValues();
    for (var i = 0; i < dataRekap.length; i++) {
      var row = dataRekap[i];
      var nama = row[1]; // Kolom B (Nama Dosen)
      if (nama) {
        rekapMap[nama] = {
          h1: row[2],
          alasanH1: row[3],
          h2: row[4],
          alasanH2: row[5],
          catatan: row[6]
        };
      }
    }
  }
  
  var listSemua = [];
  // Gabungkan daftar peserta dengan data rekapan
  for (var j = 0; j < dataPeserta.length; j++) {
    var namaDosen = dataPeserta[j][0];
    if (namaDosen && namaDosen.toString().trim() !== "") {
      if (rekapMap[namaDosen]) {
        listSemua.push({
          nama: namaDosen,
          sudahIsi: true,
          dataLama: rekapMap[namaDosen]
        });
      } else {
        listSemua.push({
          nama: namaDosen,
          sudahIsi: false,
          dataLama: null
        });
      }
    }
  }
  
  return listSemua;
}

// Menyimpan data hasil input dari form web
function simpanRSVP(formObj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetRekap = ss.getSheetByName('Rekap_RSVP');
    
    // Jika sheet Rekap_RSVP tidak ada, buatkan otomatis
    if (!sheetRekap) {
      sheetRekap = ss.insertSheet('Rekap_RSVP');
      sheetRekap.appendRow(['Timestamp', 'Nama Dosen', 'Kehadiran H1 (Sabtu)', 'Alasan H1', 'Kehadiran H2 (Minggu)', 'Alasan H2', 'Catatan']);
    }
    
    // Validasi alasan
    var alasanH1 = (formObj.h1 === 'Tidak Hadir' || formObj.h1 === 'Tentatif') ? (formObj.alasanH1 || '-') : '-';
    var alasanH2 = (formObj.h2 === 'Tidak Hadir' || formObj.h2 === 'Tentatif') ? (formObj.alasanH2 || '-') : '-';
    
    var dataBaru = [
      new Date(), // Timestamp
      formObj.namaDosen,
      formObj.h1,
      alasanH1,
      formObj.h2,
      alasanH2,
      formObj.catatan || '-'
    ];

    // Cek apakah data dosen ini sudah ada di sheet Rekap (Logika UPDATE)
    var lastRow = sheetRekap.getLastRow();
    var rowToUpdate = -1;
    
    if (lastRow > 1) {
      var namaData = sheetRekap.getRange(2, 2, lastRow - 1, 1).getValues();
      for (var i = 0; i < namaData.length; i++) {
        if (namaData[i][0] === formObj.namaDosen) {
          rowToUpdate = i + 2; // +2 karena indeks array mulai 0 dan baris 1 adalah header
          break;
        }
      }
    }

    if (rowToUpdate !== -1) {
      // Dosen sudah ada -> UPDATE / TIMPA data di baris tersebut
      sheetRekap.getRange(rowToUpdate, 1, 1, 7).setValues([dataBaru]);
    } else {
      // Dosen belum ada -> TAMBAH baris baru
      sheetRekap.appendRow(dataBaru);
    }
    
    return "Sukses";
  } catch(e) {
    return "Error: " + e.toString();
  }
}

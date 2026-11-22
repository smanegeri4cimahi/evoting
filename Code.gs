// Menangani permintaan baca data Live Score dari GitHub Pages (GET)
function doGet(e) {
  var sheet = getOrCreateSheetSuara();
  var data = sheet.getDataRange().getValues();
  
  var rekapOsis = {};
  var rekapMpk = {};
  var totalSuara = 0;
  
  for (var i = 1; i < data.length; i++) {
    var pilihanOsis = String(data[i][1]).trim();
    var pilihanMpk  = String(data[i][2]).trim();
    
    if (pilihanOsis !== "") rekapOsis[pilihanOsis] = (rekapOsis[pilihanOsis] || 0) + 1;
    if (pilihanMpk !== "") rekapMpk[pilihanMpk] = (rekapMpk[pilihanMpk] || 0) + 1;
    if (pilihanOsis !== "" || pilihanMpk !== "") totalSuara++;
  }
  
  var hasil = {
    status: "sukses",
    osis: rekapOsis,
    mpk: rekapMpk,
    total: totalSuara
  };
  
  return ContentService.createTextOutput(JSON.stringify(hasil))
    .setMimeType(ContentService.MimeType.JSON);
}

// Menangani penyimpanan suara dari GitHub Pages (POST) dengan LockService
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheetSuara();
    var waktu = new Date();
    
    sheet.appendRow([waktu, data.osis, data.mpk]);
    SpreadsheetApp.flush();
    
    return ContentService.createTextOutput(JSON.stringify({ status: "sukses" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "gagal", pesan: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheetSuara() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Suara');
  if (!sheet) {
    sheet = ss.insertSheet('Suara');
    sheet.appendRow(['Waktu', 'Pilihan OSIS', 'Pilihan MPK']);
  }
  return sheet;
}

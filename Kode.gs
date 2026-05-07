function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Registrasi HUT SKST Ke-27')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var formObject = JSON.parse(e.postData.contents);
    // Membuka Spreadsheet yang terikat dengan script ini
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Jika header kosong, buat header secara otomatis
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["No", "Timestamp", "Nama", "SAP", "Komisariat/Unit Kerja", "Pesan & Kesan", "URL Bukti Kehadiran"]);
      sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#d0e0e3"); // Styling Header
    }
    
    var fileUrl = "-";
    // Proses File Upload (Drive) jika dilampirkan
    if (formObject.fileData && formObject.fileData.data) {
       // Menggunakan ID folder KMST 2026 yang diberikan
       var targetFolderId = "1GqJNSwWX8hnGhv0SJstDvGojGWxDrtL5";
       var folder;
       try {
           folder = DriveApp.getFolderById(targetFolderId);
       } catch (errorDrive) {
           // Fallback menggunakan nama folder jika metode ById ditolak (sering terjadi karena butuh Re-Otorisasi)
           var folderIterator = DriveApp.getFoldersByName("Dokumentasi HUT SKST");
           if (folderIterator.hasNext()) {
               folder = folderIterator.next();
           } else {
               throw new Error("Folder gagal diakses! Solusi: Buka editor script, klik tombol 'Run' (Jalankan) dan klik 'Review Permissions' untuk memberi izin akses folder.");
           }
       }
       
       var blob = Utilities.newBlob(Utilities.base64Decode(formObject.fileData.data), formObject.fileData.mimeType, formObject.fileData.filename);
       var file = folder.createFile(blob);
       fileUrl = file.getUrl();
    }
    
    // Cek duplikasi SAP/NIK (Jika diisi)
    var inputSAP = formObject.nomorSAP ? formObject.nomorSAP.trim() : "";
    if (inputSAP !== "") {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) { // Mulai dari 1 untuk melewati baris header
        if (data[i][3] == inputSAP) { // Indeks 3 merujuk ke kolom "Nomor SAP"
          return { success: false, message: "Nomor SAP/NIK ini (" + inputSAP + ") sudah terdaftar." };
        }
      }
    }
    
    // Auto increment "No"
    var lastRow = sheet.getLastRow();
    var no = (lastRow === 0) ? 1 : lastRow; // karena isi form di appendRow ada di baris lastRow+1
    
    var timestamp = new Date();
    sheet.appendRow([
      no,
      timestamp,
      formObject.namaKaryawan,
      formObject.nomorSAP,
      formObject.unitKerja,
      formObject.pesanKesan || "-",
      fileUrl
    ]);
    
    var result = { success: true, message: "Pendaftaran berhasil disimpan!" };
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    var result = { success: false, message: "Error: " + error.message };
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }
}

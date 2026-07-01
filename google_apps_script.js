/**
 * ============================================================================
 * SCRIPT GOOGLE SPREADSHEET UNTUK UNDANGAN PERNIKAHAN DANIEL & NADINE
 * ============================================================================
 * 
 * Struktur Kolom di Spreadsheet Anda (Baris 1 sebagai Header):
 * Kolom A (1) : TAMU     (Nama Tamu Undangan)
 * Kolom B (2) : LINK     (Link Undangan khusus untuk tamu tersebut)
 * Kolom C (3) : SHARE WA (Link WhatsApp otomatis siap kirim)
 * Kolom D (4) : HADIR    (Status kehadiran: "Hadir" atau "Tidak Hadir")
 * Kolom E (5) : UCAPAN   (Doa restu / ucapan dari tamu)
 * 
 * CARA PASANG:
 * 1. Buka Google Spreadsheet Anda
 * 2. Klik menu: Ekstensi > Apps Script
 * 3. Hapus semua kode bawaan, lalu paste seluruh kode di file ini.
 * 4. Ganti BASE_URL di bawah sesuai URL web undangan Anda jika sudah di-publish.
 * 5. Simpan proyek (Ctrl + S / Cmd + S).
 * 6. Klik tombol "Terapkan" (Deploy) > "Deployment Baru" (New Deployment).
 * 7. Pilih jenis: "Aplikasi Web" (Web App).
 * 8. Pada bagian "Akses siapa saja" (Who has access), pilih: "Siapa saja" (Anyone).
 * 9. Salin URL Web App yang muncul, dan tempelkan ke variabel `GOOGLE_SHEET_URL` di index.html baris 2446.
 */

var BASE_URL = "https://mikaeladityan.github.io/nadine-daniel/";

/**
 * Fitur 1: OTOMATIS MEMBUAT LINK & SHARE WA SAAT NAMA TAMU DIKETIK
 * Trigger ini berjalan otomatis setiap kali Anda mengetik nama di Kolom A (TAMU).
 */
function onEdit(e) {
  // Jika tombol "Jalankan" diklik secara manual di editor, e akan undefined.
  if (!e || !e.source) {
    Logger.log("Fungsi onEdit berjalan otomatis saat Anda mengetik di Spreadsheet. Untuk membuat link massal, jalankan fungsi generateAllLinks.");
    return;
  }
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // Jika yang diedit adalah Kolom 1 (TAMU) dan bukan baris header (Baris > 1)
  if (range.getColumn() === 1 && range.getRow() > 1) {
    var row = range.getRow();
    var nama = range.getValue().toString().trim();
    
    if (nama) {
      var linkUrl = BASE_URL + "?to=" + encodeURIComponent(nama);
      var waMsg = "Kepada Yth. Bapak/Ibu/Saudara/i " + nama + "\n\nTanpa mengurangi rasa hormat, kami mengundang Bapak/Ibu/Saudara/i untuk hadir pada acara pernikahan kami.\n\nDetail undangan dapat dibuka melalui tautan berikut:\n" + linkUrl + "\n\nMerupakan suatu kehormatan dan kebahagiaan bagi kami apabila Bapak/Ibu/Saudara/i berkenan hadir dan memberikan doa restu.\n\nTerima kasih.";
      var shareWaUrl = "https://api.whatsapp.com/send?text=" + encodeURIComponent(waMsg);
      
      // Isi otomatis Kolom B (LINK) dan Kolom C (SHARE WA)
      sheet.getRange(row, 2).setValue(linkUrl);
      sheet.getRange(row, 3).setValue(shareWaUrl);
    } else {
      // Kosongkan jika nama di-hapus
      sheet.getRange(row, 2).clearContent();
      sheet.getRange(row, 3).clearContent();
    }
  }
}

/**
 * Fitur 2: MENERIMA DATA RSVP & UCAPAN DARI WEBSITE (POST REQUEST)
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Tunggu maksimal 10 detik agar tidak bentrok antar tamu
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataObj = {};
    
    if (e.postData && e.postData.contents) {
      try {
        dataObj = JSON.parse(e.postData.contents);
      } catch(err) {
        dataObj = e.parameter || {};
      }
    } else {
      dataObj = e.parameter || {};
    }
    
    var nama = (dataObj.nama || "").trim();
    var attend = (dataObj.attend || "").trim();
    var ucapan = (dataObj.ucapan || "").trim();
    var hadirStr = attend === "hadir" ? "Hadir" : (attend === "tidak" ? "Tidak Hadir" : attend);
    
    if (!nama) return ContentService.createTextOutput("Error: Nama kosong");
    
    var linkUrl = BASE_URL + "?to=" + encodeURIComponent(nama);
    var waMsg = "Kepada Yth. Bapak/Ibu/Saudara/i " + nama + "\n\nTanpa mengurangi rasa hormat, kami mengundang Bapak/Ibu/Saudara/i untuk hadir pada acara pernikahan kami.\n\nDetail undangan dapat dibuka melalui tautan berikut:\n" + linkUrl + "\n\nMerupakan suatu kehormatan dan kebahagiaan bagi kami apabila Bapak/Ibu/Saudara/i berkenan hadir dan memberikan doa restu.\n\nTerima kasih.";
    var shareWaUrl = "https://api.whatsapp.com/send?text=" + encodeURIComponent(waMsg);
    
    // Cari apakah nama tamu sudah ada di Kolom A (TAMU)
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === nama.toLowerCase()) {
        rowIndex = i + 1; // Index baris Spreadsheet mulai dari 1
        break;
      }
    }
    
    if (rowIndex > 0) {
      // Jika tamu sudah terdaftar, update kolom HADIR & UCAPAN
      sheet.getRange(rowIndex, 4).setValue(hadirStr);
      sheet.getRange(rowIndex, 5).setValue(ucapan);
      if (!sheet.getRange(rowIndex, 2).getValue()) sheet.getRange(rowIndex, 2).setValue(linkUrl);
      if (!sheet.getRange(rowIndex, 3).getValue()) sheet.getRange(rowIndex, 3).setValue(shareWaUrl);
    } else {
      // Jika tamu belum ada, tambahkan baris baru di paling bawah
      sheet.appendRow([nama, linkUrl, shareWaUrl, hadirStr, ucapan]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({result: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({result: "error", error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fitur 3: MENGIRIM DATA UCAPAN KE WEBSITE UNTUK DITAMPILKAN (GET REQUEST)
 */
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var wishes = [];
  
  for (var i = 1; i < data.length; i++) {
    var nama = data[i][0];
    var hadir = data[i][3];
    var ucapan = data[i][4];
    
    if (nama && (ucapan || hadir)) {
      var attendVal = "hadir";
      if (hadir && hadir.toString().toLowerCase().indexOf("tidak") !== -1) {
        attendVal = "tidak";
      }
      wishes.push({
        nama: nama,
        attend: attendVal,
        ucapan: ucapan || "",
        time: "Tersimpan"
      });
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(wishes))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * FUNGSI BANTUAN (OPSIONAL):
 * Jalankan fungsi ini 1x secara manual di menu Apps Script jika Anda sudah punya daftar nama tamu
 * sebelumnya dan ingin langsung membuatkan LINK & SHARE WA untuk seluruh baris sekaligus.
 */
function generateAllLinks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var nama = data[i][0] ? data[i][0].toString().trim() : "";
    if (nama) {
      var linkUrl = BASE_URL + "?to=" + encodeURIComponent(nama);
      var waMsg = "Kepada Yth. Bapak/Ibu/Saudara/i " + nama + "\n\nTanpa mengurangi rasa hormat, kami mengundang Bapak/Ibu/Saudara/i untuk hadir pada acara pernikahan kami.\n\nDetail undangan dapat dibuka melalui tautan berikut:\n" + linkUrl + "\n\nMerupakan suatu kehormatan dan kebahagiaan bagi kami apabila Bapak/Ibu/Saudara/i berkenan hadir dan memberikan doa restu.\n\nTerima kasih.";
      var shareWaUrl = "https://api.whatsapp.com/send?text=" + encodeURIComponent(waMsg);
      
      sheet.getRange(i + 1, 2).setValue(linkUrl);
      sheet.getRange(i + 1, 3).setValue(shareWaUrl);
    }
  }
}

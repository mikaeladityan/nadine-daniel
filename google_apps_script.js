/**
 * ============================================================================
 * SCRIPT GOOGLE SPREADSHEET UNTUK UNDANGAN PERNIKAHAN DANIEL & NADINE
 * ============================================================================
 * 
 * Struktur Kolom di Spreadsheet Anda (Baris 1 sebagai Header):
 * Kolom A (1) : TAMU     (Nama Tamu Undangan)
 * Kolom B (2) : LANG     (Bahasa: "ID" atau "EN" - default: "ID")
 * Kolom C (3) : LINK     (Link Undangan khusus untuk tamu: ID -> /?to=... | EN -> /en/?to=...)
 * Kolom D (4) : SHARE WA (Link WhatsApp otomatis siap kirim dengan teks ID/EN)
 * Kolom E (5) : HADIR    (Status kehadiran: "Hadir" atau "Tidak Hadir")
 * Kolom F (6) : UCAPAN   (Doa restu / ucapan dari tamu)
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
 * Fitur 1: OTOMATIS MEMBUAT LINK & SHARE WA SAAT NAMA ATAU BAHASA DIKETIK/DIGANTI
 * Trigger ini berjalan otomatis setiap kali Anda mengetik nama di Kolom A (TAMU)
 * atau mengubah bahasa di Kolom B (LANG) menjadi "ID" atau "EN".
 */
function onEdit(e) {
  // Jika tombol "Jalankan" diklik secara manual di editor, e akan undefined.
  if (!e || !e.source) {
    Logger.log("Fungsi onEdit berjalan otomatis saat Anda mengetik di Spreadsheet. Untuk membuat link massal, jalankan fungsi generateAllLinks.");
    return;
  }
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var col = range.getColumn();
  
  // Jika yang diedit adalah Kolom 1 (TAMU) atau Kolom 2 (LANG) dan bukan baris header (Baris > 1)
  if ((col === 1 || col === 2) && startRow > 1) {
    for (var r = startRow; r < startRow + numRows; r++) {
      var nama = sheet.getRange(r, 1).getValue().toString().trim();
      var langVal = sheet.getRange(r, 2).getValue().toString().trim().toUpperCase();
      
      if (nama) {
        // Default ke ID jika kolom LANG kosong atau bukan EN/ID
        if (langVal !== "EN" && langVal !== "ID") {
          langVal = "ID";
          sheet.getRange(r, 2).setValue("ID");
        }
        
        var linkUrl = langVal === "EN" 
          ? (BASE_URL + "en/?to=" + encodeURIComponent(nama)) 
          : (BASE_URL + "?to=" + encodeURIComponent(nama));
          
        var waMsg = langVal === "EN"
          ? ("Dear Mr./Mrs./Ms./Family " + nama + "\n\nWe cordially invite you to celebrate our wedding reception.\n\nYou can access the invitation details through the link below:\n" + linkUrl + "\n\nIt is a great honor and happiness for our families if you could attend and give your blessings.\n\nThank you.")
          : ("Kepada Yth. Bapak/Ibu/Saudara/i " + nama + "\n\nTanpa mengurangi rasa hormat, kami mengundang Bapak/Ibu/Saudara/i untuk hadir pada acara pernikahan kami.\n\nDetail undangan dapat dibuka melalui tautan berikut:\n" + linkUrl + "\n\nMerupakan suatu kehormatan dan kebahagiaan bagi kami apabila Bapak/Ibu/Saudara/i berkenan hadir dan memberikan doa restu.\n\nTerima kasih.");
          
        var shareWaUrl = "https://api.whatsapp.com/send?text=" + encodeURIComponent(waMsg);
        
        // Isi otomatis Kolom C (LINK) dan Kolom D (SHARE WA)
        sheet.getRange(r, 3).setValue(linkUrl);
        sheet.getRange(r, 4).setValue(shareWaUrl);
      } else {
        // Kosongkan LANG, LINK, dan SHARE WA jika nama di-hapus
        sheet.getRange(r, 2, 1, 3).clearContent();
      }
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
    var lang = (dataObj.lang || "ID").toString().trim().toUpperCase();
    if (lang !== "EN") lang = "ID";
    
    var hadirStr = attend === "hadir" ? "Hadir" : (attend === "tidak" ? "Tidak Hadir" : attend);
    
    if (!nama) return ContentService.createTextOutput("Error: Nama kosong");
    
    var linkUrl = lang === "EN" 
      ? (BASE_URL + "en/?to=" + encodeURIComponent(nama)) 
      : (BASE_URL + "?to=" + encodeURIComponent(nama));
      
    var waMsg = lang === "EN"
      ? ("Dear Mr./Mrs./Ms./Family " + nama + "\n\nWe cordially invite you to celebrate our wedding reception.\n\nYou can access the invitation details through the link below:\n" + linkUrl + "\n\nIt is a great honor and happiness for our families if you could attend and give your blessings.\n\nThank you.")
      : ("Kepada Yth. Bapak/Ibu/Saudara/i " + nama + "\n\nTanpa mengurangi rasa hormat, kami mengundang Bapak/Ibu/Saudara/i untuk hadir pada acara pernikahan kami.\n\nDetail undangan dapat dibuka melalui tautan berikut:\n" + linkUrl + "\n\nMerupakan suatu kehormatan dan kebahagiaan bagi kami apabila Bapak/Ibu/Saudara/i berkenan hadir dan memberikan doa restu.\n\nTerima kasih.");
      
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
      // Jika tamu sudah terdaftar, update kolom HADIR (Kolom E/5) & UCAPAN (Kolom F/6)
      sheet.getRange(rowIndex, 5).setValue(hadirStr);
      sheet.getRange(rowIndex, 6).setValue(ucapan);
      if (!sheet.getRange(rowIndex, 2).getValue()) sheet.getRange(rowIndex, 2).setValue(lang);
      if (!sheet.getRange(rowIndex, 3).getValue()) sheet.getRange(rowIndex, 3).setValue(linkUrl);
      if (!sheet.getRange(rowIndex, 4).getValue()) sheet.getRange(rowIndex, 4).setValue(shareWaUrl);
    } else {
      // Jika tamu belum ada, tambahkan baris baru di paling bawah
      sheet.appendRow([nama, lang, linkUrl, shareWaUrl, hadirStr, ucapan]);
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
    var hadir = data[i][4];  // Kolom E (5) - HADIR
    var ucapan = data[i][5]; // Kolom F (6) - UCAPAN
    
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
 * FUNGSI BANTUAN:
 * Jalankan fungsi ini 1x secara manual di menu Apps Script jika Anda sudah punya daftar nama tamu
 * sebelumnya dan ingin langsung membuatkan LANG, LINK & SHARE WA untuk seluruh baris sekaligus.
 */
function generateAllLinks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var nama = data[i][0] ? data[i][0].toString().trim() : "";
    var langVal = data[i][1] ? data[i][1].toString().trim().toUpperCase() : "";
    
    if (nama) {
      if (langVal !== "EN" && langVal !== "ID") {
        langVal = "ID";
        sheet.getRange(i + 1, 2).setValue("ID");
      }
      var linkUrl = langVal === "EN" 
        ? (BASE_URL + "en/?to=" + encodeURIComponent(nama)) 
        : (BASE_URL + "?to=" + encodeURIComponent(nama));
        
      var waMsg = langVal === "EN" 
        ? ("Dear Mr./Mrs./Ms./Family " + nama + "\n\nWe cordially invite you to celebrate our wedding reception.\n\nYou can access the invitation details through the link below:\n" + linkUrl + "\n\nIt is a great honor and happiness for our families if you could attend and give your blessings.\n\nThank you.")
        : ("Kepada Yth. Bapak/Ibu/Saudara/i " + nama + "\n\nTanpa mengurangi rasa hormat, kami mengundang Bapak/Ibu/Saudara/i untuk hadir pada acara pernikahan kami.\n\nDetail undangan dapat dibuka melalui tautan berikut:\n" + linkUrl + "\n\nMerupakan suatu kehormatan dan kebahagiaan bagi kami apabila Bapak/Ibu/Saudara/i berkenan hadir dan memberikan doa restu.\n\nTerima kasih.");
        
      var shareWaUrl = "https://api.whatsapp.com/send?text=" + encodeURIComponent(waMsg);
      
      sheet.getRange(i + 1, 3).setValue(linkUrl);
      sheet.getRange(i + 1, 4).setValue(shareWaUrl);
    }
  }
}

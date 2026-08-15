/**
 * ============================================================================
 * MODUL GENERATE NOMOR INVESTIGASI & PENYIMPANAN DATA INSIDEN + LAMPIRAN
 * File: generatenomerinvestigasi.gs
 * ============================================================================
 */

/**
 * Auto-Generate Nomor Insiden dengan Reset Urutan Harian ke 0001 per Tanggal & Site
 * Format: DDMMYYYY-INV-SITE-0001
 */
function generateIncidentID(siteName, tglKejadian) {
  const ss = getTransaksiSS();
  const sheet = ss.getSheetByName(CONFIG.TAB_INSIDEN);
  
  // Format Tanggal Kejadian (DDMMYYYY)
  let dateObj = tglKejadian ? new Date(tglKejadian) : new Date();
  let dd = String(dateObj.getDate()).padStart(2, '0');
  let mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  let yyyy = dateObj.getFullYear();
  let dateStr = `${dd}${mm}${yyyy}`;

  // Pembersihan Nama Site
  let cleanSite = String(siteName).replace(/[^a-zA-Z0-9]/g, '').toUpperCase();

  // Hitung jumlah laporan pada tanggal & site yang sama
  let lastRow = sheet.getLastRow();
  let dailyCounter = 1;

  if (lastRow > 1) {
    const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Kolom B (No Insiden)
    const prefixTarget = `${dateStr}-INV-${cleanSite}`;
    
    data.forEach(row => {
      let noInc = String(row[0]);
      if (noInc.indexOf(prefixTarget) === 0) {
        dailyCounter++;
      }
    });
  }

  let counterStr = String(dailyCounter).padStart(4, '0');
  return `${dateStr}-INV-${cleanSite}-${counterStr}`;
}

/**
 * Simpan Data Laporan Insiden Lengkap (Mendukung Multi-Karyawan Terlibat)
 */
function saveIncidentReport(token, formData) {
  try {
    const session = requireSession(token);
    const ss = getTransaksiSS();
    const sheet = ss.getSheetByName(CONFIG.TAB_INSIDEN);

    // Auto-Generate 1 ID Insiden yang sama untuk seluruh karyawan di laporan ini
    const noInsiden = generateIncidentID(formData.site, formData.tglKejadian);
    const timestamp = new Date();

    const empList = (formData.karyawanList && formData.karyawanList.length > 0) 
      ? formData.karyawanList 
      : [{}];

    // Simpan baris data di Sheet Transaksi sejumlah karyawan yang terlibat
    empList.forEach(emp => {
      const rowData = [
        timestamp,                          // 1. Timestamp
        noInsiden,                          // 2. No Insiden
        session.Nama || session.Username,   // 3. Nama Pelapor
        formData.hari || '',                // 4. Hari
        formData.tglKejadian || '',         // 5. Tanggal
        formData.bulan || '',               // 6. Bulan
        formData.jamKejadian || '',         // 7. Jam
        formData.shift || '',               // 8. Shift
        formData.lokasi || '',              // 9. Lokasi Spesifik
        formData.kronologis || '',          // 10. Kronologis
        formData.site || '',                // 11. Site/Business Unit (BU)
        emp.perusahaan || '',               // 12. Perusahan
        emp.departemen || '',               // 13. Departement
        emp.klasifikasi || '',              // 14. Klasifikasi Kecelakaan
        emp.karyawan || '',                 // 15. Karyawan Terlibat / Nama
        emp.jabatan || '',                  // 16. Jabatan
        emp.umur || '',                     // 17. Umur (Tahun)
        emp.masaKerja || '',                // 18. Masa Kerja
        formData.alatTerlibat || '',        // 19. Alat Terlibat
        formData.jenisAlat || '',           // 20. Jenis Alat
        formData.lossCost || '',            // 21. Loss Cost
        formData.jenisKontak || '',         // 22. Jenis Kontak
        formData.sumberKecelakaan || '',    // 23. Sumber Kecelakaan
        formData.tta || '',                 // 24. Tindakan Tidak Aman (TTA)
        formData.ketTta || '',              // 25. Keterangan TTA
        formData.kta || '',                 // 26. Kondisi Tidak Aman (KTA)
        formData.ketKta || '',              // 27. Keterangan KTA
        formData.faktorManusia || '',       // 28. Faktor Manusia
        formData.ketManusia || '',          // 29. Keterangan Faktor Manusia
        formData.faktorPekerjaan || '',     // 30. Faktor Pekerjaan
        formData.ketPekerjaan || '',        // 31. Keterangan Faktor Pekerjaan
        formData.kurangKendali || '',       // 32. Kurang Kendali Manajemen
        formData.ketKendali || '',          // 33. Keterangan Kurang Kendali
        formData.tindakanPerbaikan || '',   // 34. Tindakan Perbaikan
        formData.tindakanPencegahan || '',  // 35. Tindakan Pencegahan
        formData.dueDate || '',             // 36. Due Date
        '',                                 // 37. Completion Date (Kosong saat input awal)
        'Open'                              // 38. Status Laporan (Default: Open)
      ];

      sheet.appendRow(rowData);
    });

    return { 
      success: true, 
      message: 'Data insiden berhasil disimpan!', 
      noInsiden: noInsiden 
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Upload Lampiran Media ke Google Drive dengan Sub-Folder Otomatis per Nomor Insiden
 */
function uploadAttachment(token, noInsiden, fileObj, tipeLampiran) {
  try {
    const session = requireSession(token);
    const ss = getTransaksiSS();
    
    // Pastikan Tab Sheet 'Lampiran' Tersedia
    let sheetLampiran = ss.getSheetByName('Lampiran');
    if (!sheetLampiran) {
      sheetLampiran = ss.insertSheet('Lampiran');
      sheetLampiran.appendRow(['ID_Lampiran', 'No Insiden', 'Tipe', 'Nama File', 'URL Drive', 'MimeType', 'Diupload Oleh', 'Diupload Pada']);
    }

    // ID Folder Utama Parent
    const mainFolderId = (typeof CONFIG !== 'undefined' && CONFIG.FOLDER_DRIVE_ID) 
      ? CONFIG.FOLDER_DRIVE_ID 
      : '1VF_GLSbe8h_bvJ5LtR7kRylsO79HOTx_';

    // 1. Ambil Folder Utama
    let mainFolder;
    try {
      mainFolder = DriveApp.getFolderById(mainFolderId.trim());
    } catch (eFolder) {
      mainFolder = DriveApp.createFolder("HSE_Lampiran_Incidents");
    }

    // 2. Buat Sub-Folder Khusus Berdasarkan Nomor Insiden
    let incidentFolder;
    const subFolders = mainFolder.getFoldersByName(noInsiden);
    if (subFolders.hasNext()) {
      incidentFolder = subFolders.next();
    } else {
      incidentFolder = mainFolder.createFolder(noInsiden);
    }

    // 3. Dekode Base64 & Simpan File Ke Dalam Sub-Folder Tersebut
    let cleanBase64 = fileObj.bytes || '';
    if (cleanBase64.indexOf(',') !== -1) {
      cleanBase64 = cleanBase64.split(',')[1];
    }

    const decodedBytes = Utilities.base64Decode(cleanBase64);
    const blob = Utilities.newBlob(decodedBytes, fileObj.mimeType || 'image/png', fileObj.fileName);
    const file = incidentFolder.createFile(blob);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(eShare) {}

    // 4. Catat Metadata ke Tab Sheet "Lampiran"
    const idLampiran = 'LMP-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    const rowLampiran = [
      idLampiran,                       // 1. ID_Lampiran
      noInsiden,                        // 2. No Insiden
      tipeLampiran || 'Foto Kejadian',  // 3. Tipe
      fileObj.fileName,                 // 4. Nama File
      file.getUrl(),                    // 5. URL Drive
      fileObj.mimeType,                 // 6. MimeType
      session.Nama || session.Username, // 7. Diupload Oleh
      new Date()                        // 8. Diupload Pada
    ];

    sheetLampiran.appendRow(rowLampiran);

    return { success: true, url: file.getUrl() };
  } catch (err) {
    Logger.log("Error Upload Attachment: " + err.message);
    return { success: false, message: err.message };
  }
}
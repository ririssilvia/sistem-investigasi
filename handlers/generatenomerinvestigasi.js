/**
 * ============================================================================
 * HANDLER GENERATE NOMOR INVESTIGASI & TRANSAKSI LAPORAN INSIDEN
 * File: handlers/generatenomerinvestigasi.js
 * ============================================================================
 */

/**
 * Auto-Generate Nomor Insiden Reset Urutan Harian per Tanggal & Site
 * Format: DDMMYYYY-INV-SITE-0001
 */
function generateIncidentID(siteName, tglKejadian) {
  const ss = getTransaksiSS();
  const sheet = ss.getSheetByName(CONFIG.TAB_INSIDEN);
  
  let dateObj = tglKejadian ? new Date(tglKejadian) : new Date();
  let dd = String(dateObj.getDate()).padStart(2, '0');
  let mm = String(dateObj.getMonth() + 1).padStart(2, '0');
  let yyyy = dateObj.getFullYear();
  let dateStr = `${dd}${mm}${yyyy}`;

  let cleanSite = String(siteName).replace(/[^a-zA-Z0-9]/g, '').toUpperCase();

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
 * Simpan Data Laporan Insiden (Mendukung Multi-Karyawan per 1 Insiden ID)
 */
function saveIncidentReport(token, formData) {
  try {
    const session = requireSession(token);
    const ss = getTransaksiSS();
    const sheet = ss.getSheetByName(CONFIG.TAB_INSIDEN);

    const noInsiden = generateIncidentID(formData.site, formData.tglKejadian);
    const timestamp = new Date();

    const empList = (formData.karyawanList && formData.karyawanList.length > 0) 
      ? formData.karyawanList 
      : [{}];

    // Simpan baris transaksi sebanyak jumlah karyawan yang didaftarkan
    empList.forEach(emp => {
      const rowData = [
        timestamp,                          // 1. Timestamp
        noInsiden,                          // 2. No Insiden
        session.Nama || session.Username,   // 3. Pelapor
        formData.hari || '',                // 4. Hari
        formData.tglKejadian || '',         // 5. Tanggal
        formData.bulan || '',               // 6. Bulan
        formData.jamKejadian || '',         // 7. Jam
        formData.shift || '',               // 8. Shift
        formData.lokasi || '',              // 9. Lokasi Spesifik
        formData.kronologis || '',          // 10. Kronologis
        formData.site || '',                // 11. Site/Business Unit
        emp.perusahaan || '',               // 12. Perusahaan
        emp.departemen || '',               // 13. Departemen
        emp.klasifikasi || '',              // 14. Klasifikasi
        emp.karyawan || '',                 // 15. Karyawan Terlibat / Nama
        emp.jabatan || '',                  // 16. Jabatan
        emp.umur || '',                     // 17. Umur
        emp.masaKerja || '',                // 18. Masa Kerja
        formData.alatTerlibat || '',        // 19. Alat Terlibat
        formData.jenisAlat || '',           // 20. Jenis Alat
        formData.lossCost || '',            // 21. Loss Cost
        formData.jenisKontak || '',         // 22. Jenis Kontak
        formData.sumberKecelakaan || '',    // 23. Sumber Kecelakaan
        formData.tta || '',                 // 24. TTA
        formData.ketTta || '',              // 25. Ket TTA
        formData.kta || '',                 // 26. KTA
        formData.ketKta || '',              // 27. Ket KTA
        formData.faktorManusia || '',       // 28. Faktor Manusia
        formData.ketManusia || '',          // 29. Ket Faktor Manusia
        formData.faktorPekerjaan || '',     // 30. Faktor Pekerjaan
        formData.ketPekerjaan || '',        // 31. Ket Faktor Pekerjaan
        formData.kurangKendali || '',       // 32. Kurang Kendali
        formData.ketKendali || '',          // 33. Ket Kurang Kendali
        formData.tindakanPerbaikan || '',   // 34. Perbaikan
        formData.tindakanPencegahan || '',  // 35. Pencegahan
        formData.dueDate || '',             // 36. Due Date
        '',                                 // 37. Completion Date
        'Open'                              // 38. Status Laporan
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

/**
 * Mengambil Seluruh Dropdown Master Data untuk Form Input Insiden
 */
/**
 * Helper Aman: Mengambil data 1 kolom dari Tab Master tanpa pernah mengembalikan null/error
 */
function getMasterColumnSafe(tabName, colHeader) {
  try {
    const ss = getMasterSS(); // Mengambil Spreadsheet Master
    if (!ss) return [];

    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; // Cuma header atau kosong

    const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    
    let targetIdx = 0; // Default kolom 0 (Kolom A)
    if (colHeader) {
      const idxFound = headers.indexOf(String(colHeader).trim().toLowerCase());
      if (idxFound !== -1) targetIdx = idxFound;
    }

    let result = [];
    for (let i = 1; i < data.length; i++) {
      let val = String(data[i][targetIdx]).trim();
      if (val !== "" && val !== "null" && val !== "undefined") {
        result.push(val);
      }
    }
    return result;
  } catch (e) {
    Logger.log(`Error reading tab ${tabName}: ${e.message}`);
    return [];
  }
}

/**
 * Helper Aman: Mengambil data 1 kolom dari Tab Master tanpa pernah mengembalikan null/error
 */
function getMasterColumnSafe(tabName, colHeader) {
  try {
    const ss = getMasterSS(); // Mengambil Spreadsheet Master
    if (!ss) return [];

    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; // Cuma header atau kosong

    const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    
    let targetIdx = 0; // Default kolom 0 (Kolom A)
    if (colHeader) {
      const idxFound = headers.indexOf(String(colHeader).trim().toLowerCase());
      if (idxFound !== -1) targetIdx = idxFound;
    }

    let result = [];
    for (let i = 1; i < data.length; i++) {
      let val = String(data[i][targetIdx]).trim();
      if (val !== "" && val !== "null" && val !== "undefined") {
        result.push(val);
      }
    }
    return result;
  } catch (e) {
    Logger.log(`Error reading tab ${tabName}: ${e.message}`);
    return [];
  }
}

/**
 * Fungsi Utama Mengambil Master Data (Tahan Banting & Anti-Null)
 */
function getFormMasterData(token) {
  try {
    // Ambil Data Master secara langsung
    const masterData = {
      jam: getMasterColumnSafe(CONFIG.TAB_JAM, 'Jam'),
      site: getMasterColumnSafe(CONFIG.TAB_SITE, 'Site/Business Unit (BU)'),
      perusahaan: getMasterColumnSafe(CONFIG.TAB_PERUSAHAAN, 'Perusahan'),
      departemen: getMasterColumnSafe(CONFIG.TAB_DEPARTEMEN, 'Departement'),
      lokasi: getMasterColumnSafe(CONFIG.TAB_LOKASI, 'Lokasi Spesifik'),
      jabatan: getMasterColumnSafe(CONFIG.TAB_JABATAN, 'Jabatan'),
      umur: getMasterColumnSafe(CONFIG.TAB_UMUR, 'Umur (Tahun)'),
      masaKerja: getMasterColumnSafe(CONFIG.TAB_MASA_KERJA, 'Masa Kerja'),
      kategoriAlat: getMasterColumnSafe(CONFIG.TAB_KATEGORI_ALAT, 'Alat Terlibat'),
      jenisAlat: getMasterColumnSafe(CONFIG.TAB_JENIS_ALAT, 'Jenis Alat'),
      jenisKontak: getMasterColumnSafe(CONFIG.TAB_JENIS_KONTAK, 'Jenis Kontak'),
      sumberKecelakaan: getMasterColumnSafe(CONFIG.TAB_SUMBER_KECELAKAAN, 'Sumber Kecelakaan'),
      klasifikasi: getMasterColumnSafe(CONFIG.TAB_KLASIFIKASI, 'Klasifikasi Kecelakaan'),
      tta: getMasterColumnSafe(CONFIG.TAB_TTA, 'Penyebab Langsung Tindakan Tidak Aman (TTA)'),
      ketTta: getMasterColumnSafe(CONFIG.TAB_KET_TTA, 'Keterangan Penyebab Langsung Tindakan Tidak Aman (TTA)'),
      kta: getMasterColumnSafe(CONFIG.TAB_KTA, 'Penyebab Langsung Kondisi Tidak Aman (KTA)'),
      ketKta: getMasterColumnSafe(CONFIG.TAB_KET_KTA, 'Keterangan Penyebab Langsung Kondisi Tidak Aman (KTA)'),
      faktorManusia: getMasterColumnSafe(CONFIG.TAB_FAKTOR_MANUSIA, 'Penyebab Dasar Faktor Manusia'),
      ketFaktorManusia: getMasterColumnSafe(CONFIG.TAB_KET_FAKTOR_MANUSIA, 'Keterangan Penyebab Dasar Faktor Manusia'),
      faktorPekerjaan: getMasterColumnSafe(CONFIG.TAB_FAKTOR_PEKERJAAN, 'Penyebab Dasar Faktor Pekerjaan'),
      ketFaktorPekerjaan: getMasterColumnSafe(CONFIG.TAB_KET_FAKTOR_PEKERJAAN, 'Keterangan Penyebab Dasar Faktor Pekerjaan'),
      kurangKendali: getMasterColumnSafe(CONFIG.TAB_KURANG_KENDALI, 'Kurang Kendali Manajemen'),
      ketKurangKendali: getMasterColumnSafe(CONFIG.TAB_KET_KURANG_KENDALI, 'Keterangan Kurang Kendali Manajemen'),
    };

    // Sertakan default jam jika tab jam di master belum diisi (00:00 - 23:00)
    if (masterData.jam.length === 0) {
      for (let i = 0; i < 24; i++) {
        masterData.jam.push(String(i).padStart(2, '0') + ':00');
      }
    }

    return {
      success: true,
      master: masterData
    };
  } catch (err) {
    return {
      success: false,
      message: err.message,
      master: {}
    };
  }
}

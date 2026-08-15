/**
 * HANDLER MASTER DATA DROPDOWN FORM
 * File: handlers/Masterdata.js
 */
function getFormMasterData(token) {
  try {
    const session = requireSession(token);
    const ss = getMasterSS();
    
    function getColumnData(tabKey, colHeaderName) {
      try {
        const sheetName = CONFIG[tabKey];
        if (!sheetName) return [];
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return [];
        const values = sheet.getDataRange().getValues();
        if (values.length <= 1) return [];
        
        let colIndex = 0;
        if (colHeaderName) {
          const idx = values[0].indexOf(colHeaderName);
          if (idx !== -1) colIndex = idx;
        }

        const list = values.slice(1).map(r => r[colIndex]).filter(v => v !== "" && v !== null && v !== undefined);
        return Array.from(new Set(list));
      } catch(e) {
        return [];
      }
    }

    const masterData = {
      site: getColumnData("TAB_SITE"),
      jam: getColumnData("TAB_JAM"),
      lokasi: getColumnData("TAB_LOKASI"),
      perusahaan: getColumnData("TAB_PERUSAHAAN"),
      departemen: getColumnData("TAB_DEPARTEMEN"),
      klasifikasi: getColumnData("TAB_KLASIFIKASI"),
      jabatan: getColumnData("TAB_JABATAN"),
      umur: getColumnData("TAB_UMUR"),
      masaKerja: getColumnData("TAB_MASA_KERJA"),
      kategoriAlat: getColumnData("TAB_KATEGORI_ALAT"),
      jenisAlat: getColumnData("TAB_JENIS_ALAT"),
      jenisKontak: getColumnData("TAB_JENIS_KONTAK"),
      sumberKecelakaan: getColumnData("TAB_SUMBER_KECELAKAAN"),
      tta: getColumnData("TAB_TTA"),
      ketTta: getColumnData("TAB_KET_TTA"),
      kta: getColumnData("TAB_KTA"),
      ketKta: getColumnData("TAB_KET_KTA"),
      faktorManusia: getColumnData("TAB_FAKTOR_MANUSIA"),
      ketFaktorManusia: getColumnData("TAB_KET_FAKTOR_MANUSIA"),
      faktorPekerjaan: getColumnData("TAB_FAKTOR_PEKERJAAN"),
      ketFaktorPekerjaan: getColumnData("TAB_KET_FAKTOR_PEKERJAAN"),
      kurangKendali: getColumnData("TAB_KURANG_KENDALI"),
      ketKurangKendali: getColumnData("TAB_KET_KURANG_KENDALI")
    };

    return { success: true, master: masterData };
  } catch (err) {
    Logger.log("Error getFormMasterData: " + err.message);
    return { success: false, message: err.message };
  }
}
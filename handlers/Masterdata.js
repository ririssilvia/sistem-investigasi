/**
 * HANDLER MASTER DATA DROPDOWN FORM
 * File: handlers/Masterdata.js
 */
function getFormMasterData(token) {
  try {
    const session = requireSession(token);
    const ss = getMasterSS(); // Mengambil Spreadsheet Master
    
    // Helper Ambil Nilai Kolom Pertama dari Tab Sheet Master
    function getColumnData(sheetName, colIndex) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];
      const values = sheet.getRange(2, colIndex || 1, lastRow - 1, 1).getValues();
      const list = values.map(r => r[0]).filter(v => v !== "" && v !== null && v !== undefined);
      return Array.from(new Set(list)); // Hapus Duplikat
    }

    // Ambil data murni dari Tab Sheet Master
    const masterData = {
      site: getColumnData("Site"),
      jam: getColumnData("Jam"),
      lokasi: getColumnData("Lokasi"),
      perusahaan: getColumnData("Perusahaan"),
      departemen: getColumnData("Departemen"),
      klasifikasi: getColumnData("Klasifikasi"),
      jabatan: getColumnData("Jabatan"),
      umur: getColumnData("Umur"),
      masaKerja: getColumnData("MasaKerja"),
      kategoriAlat: getColumnData("KategoriAlat"),
      jenisAlat: getColumnData("JenisAlat"),
      jenisKontak: getColumnData("JenisKontak"),
      sumberKecelakaan: getColumnData("SumberKecelakaan"),
      tta: getColumnData("TTA"),
      ketTta: getColumnData("KetTTA"),
      kta: getColumnData("KTA"),
      ketKta: getColumnData("KetKTA"),
      faktorManusia: getColumnData("FaktorManusia"),
      ketFaktorManusia: getColumnData("KetFaktorManusia"),
      faktorPekerjaan: getColumnData("FaktorPekerjaan"),
      ketFaktorPekerjaan: getColumnData("KetFaktorPekerjaan"),
      kurangKendali: getColumnData("KurangKendali"),
      ketKurangKendali: getColumnData("KetKurangKendali")
    };

    return { success: true, master: masterData };
  } catch (err) {
    Logger.log("Error getFormMasterData: " + err.message);
    return { success: false, message: err.message };
  }
}
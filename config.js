/**
 * Config.gs
 * Menyimpan semua konfigurasi ID Spreadsheet, Nama Tab, dan Folder Drive
 */

const CONFIG = {
  SS_ID_MASTER: '1n960O0SrhGZESgLRGOZlkUfwDumOIDPCL5jrM2Dwm0U',
  SS_ID_TRANSAKSI: '1hb2yyHvtivYdsceI5h70goR0g8JuRFyNdEuAPExUrOs',

  // Nama tab di DB_Master_System
  TAB_USERS: 'Master_User',
  TAB_JAM: 'Master_Jam_Kejadian',
  TAB_SITE: 'Master_Site',
  TAB_PERUSAHAAN: 'Master_Perusahaan',
  TAB_DEPARTEMEN: 'Master_Departemen',
  TAB_LOKASI: 'Master_Lokasi',
  TAB_JABATAN: 'Master_Jabatan',
  TAB_UMUR: 'Master_Kelompok_Umur',
  TAB_MASA_KERJA: 'Master_Masa_Kerja',
  TAB_KATEGORI_ALAT: 'Master_Kategori_Alat',
  TAB_JENIS_ALAT: 'Master_Jenis_Alat',
  TAB_JENIS_KONTAK: 'Master_Jenis_Kontak',
  TAB_SUMBER_KECELAKAAN: 'Master_Sumber_Kecelakaan',
  TAB_KLASIFIKASI: 'Master_Klasifikasi',
  TAB_TTA: 'Master_TTA',
  TAB_KET_TTA: 'Master_Keterangan_TTA',
  TAB_KTA: 'Master_KTA',
  TAB_KET_KTA: 'Master_Keterangan_KTA',
  TAB_FAKTOR_MANUSIA: 'Master_Faktor_Manusia',
  TAB_KET_FAKTOR_MANUSIA: 'Master_Keterangan_Faktor_Manusia',
  TAB_FAKTOR_PEKERJAAN: 'Master_Faktor_Pekerjaan',
  TAB_KET_FAKTOR_PEKERJAAN: 'Master_Keterangan_Faktor_Pekerjaan',
  TAB_KURANG_KENDALI: 'Master_Kurang_Kendali',
  TAB_KET_KURANG_KENDALI: 'Master_Keterangan_Kurang_Kendali',

  // Nama tab di DB_Transaksi_Investigasi
  TAB_INSIDEN: 'Data_Input',
  TAB_LAMPIRAN: 'Lampiran',
  TAB_RIWAYAT: 'Riwayat_Status',

  // Folder ID Google Drive untuk simpan upload lampiran
  LAMPIRAN_FOLDER_ID: '1VF_GLSbe8h_bvJ5LtR7kRylsO79HOTx_'
};

/** Helper: Buka spreadsheet Master */
function getMasterSS() {
  return SpreadsheetApp.openById(CONFIG.SS_ID_MASTER);
}

/** Helper: Buka spreadsheet Transaksi */
function getTransaksiSS() {
  return SpreadsheetApp.openById(CONFIG.SS_ID_TRANSAKSI);
}

/** Helper: Ambil semua data satu tab sebagai array of object (row 1 = header) */
function sheetToObjects(spreadsheet, tabName) {
  const sheet = spreadsheet.getSheetByName(tabName);
  if (!sheet) throw new Error('Tab tidak ditemukan: ' + tabName);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

/** Helper: Ambil satu kolom saja dari tab master sebagai array string */
function getMasterColumn(tabName, columnName) {
  const objs = sheetToObjects(getMasterSS(), tabName);
  return objs.map(o => o[columnName]).filter(v => v !== '' && v != null);
}
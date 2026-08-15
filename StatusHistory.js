/**
 * StatusHistory.gs — Update status insiden & catat log riwayat
 */

function updateIncidentStatus(token, idInsiden, statusBaru, catatan) {
  const session = requireSession(token);
  const ss = getTransaksiSS();
  
  const sheetInput = ss.getSheetByName(CONFIG.TAB_INSIDEN);
  const dataInput = sheetInput.getDataRange().getValues();
  const headers = dataInput[0];
  
  let idColIdx = headers.indexOf("No Insiden"); 
  if (idColIdx === -1) idColIdx = headers.indexOf("ID Insiden");

  let statusColIdx = headers.indexOf("Status Laporan"); 
  if (statusColIdx === -1) statusColIdx = headers.indexOf("Status");

  let statusLama = "-";
  let rowIndexToUpdate = -1;

  for (let i = 1; i < dataInput.length; i++) {
    if (String(dataInput[i][idColIdx]).trim() === String(idInsiden).trim()) {
      statusLama = dataInput[i][statusColIdx] || "Draft";
      rowIndexToUpdate = i + 1; 
      break;
    }
  }

  if (rowIndexToUpdate === -1) {
    return { success: false, message: "ID Insiden tidak ditemukan!" };
  }

  // Update nilai status di sheet utama
  sheetInput.getRange(rowIndexToUpdate, statusColIdx + 1).setValue(statusBaru);

  // Catat audit log ke sheet Riwayat_Status
  const sheetRiwayat = ss.getSheetByName(CONFIG.TAB_RIWAYAT);
  const logId = "HIS-" + Utilities.getUuid().substring(0, 8).toUpperCase();
  
  sheetRiwayat.appendRow([
    logId,
    idInsiden,
    statusLama,
    statusBaru,
    catatan || "-",
    new Date(),
    session.Nama || session.Username
  ]);

  return { 
    success: true, 
    message: `Status insiden ${idInsiden} berhasil diperbarui menjadi '${statusBaru}'` 
  };
}

/**
 * Menyimpan Laporan Baru dari FormInput.html
 */
function saveIncidentReport(token, formData) {
  const session = requireSession(token);

  try {
    const ss = getTransaksiSS();
    const sheet = ss.getSheetByName(CONFIG.TAB_INSIDEN);
    
    const year = new Date().getFullYear();
    const siteCode = (session.Site || "PUSAT").toUpperCase();
    const randomNum = Math.floor(1000 + Math.random() * 9000);
    const idInsiden = `INC-${siteCode}-${year}-${randomNum}`;

    // Tentukan Site (Kunci site user, terkecuali role Admin yang memilih)
    const finalSite = String(session.Role).toLowerCase() === 'admin' ? (formData.site || session.Site) : session.Site;

    sheet.appendRow([
      new Date(),                         // Timestamp
      idInsiden,                          // No Insiden
      session.Nama || session.Username,   // Nama Pelapor
      formData.tglKejadian,               // Tanggal
      "", "",                             // Hari & Bulan (Formulir/Auto)
      formData.jamKejadian,               // Jam
      "",                                 // Shift
      formData.lokasi,                    // Lokasi Spesifik
      formData.kronologis,                // Kronologis
      finalSite,                          // Site / Business Unit (BU)
      formData.perusahaan,                // Perusahaan
      formData.departemen,                // Departemen
      formData.klasifikasi                // Klasifikasi Kecelakaan
    ]);

    // Berikan log awal 'Open'
    updateIncidentStatus(token, idInsiden, "Open", "Laporan baru dibuat via Web App");

    return { success: true, idInsiden: idInsiden };

  } catch (error) {
    return { success: false, message: "Gagal menyimpan laporan: " + error.toString() };
  }
}
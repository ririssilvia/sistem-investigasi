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

  sheetInput.getRange(rowIndexToUpdate, statusColIdx + 1).setValue(statusBaru);

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
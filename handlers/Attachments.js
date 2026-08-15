/**
 * Attachments.gs — Mengunggah file ke Google Drive
 */
function uploadAttachment(token, idInsiden, fileObject, jenisLampiran) {
  const session = requireSession(token);
  try {
    const ss = getTransaksiSS();
    const inputSheet = ss.getSheetByName(CONFIG.TAB_INSIDEN);
    if (!inputSheet) throw new Error('Tab Data_Input tidak ditemukan.');
    const values = inputSheet.getDataRange().getDisplayValues();
    if (values.length <= 1) throw new Error('Data insiden kosong.');
    const headers = values[0].map(v => String(v || '').trim());
    const idIdx = incidentHeaderIndex_(headers, ['No Insiden','Nomor Insiden','No. Insiden','ID Insiden']);
    const statusIdx = headers.length >= 38 ? 37 : incidentHeaderIndex_(headers, ['Status','Status Laporan']);
    const siteIdx = incidentHeaderIndex_(headers, ['Site/Business Unit (BU)','Site']);
    if (idIdx < 0 || statusIdx < 0) throw new Error('Kolom No Insiden/Status tidak ditemukan.');
    const matches = values.slice(1).filter(r => String(r[idIdx] || '').trim() === String(idInsiden).trim());
    if (!matches.length) throw new Error('Data insiden tidak ditemukan.');
    if (!incidentIsAdmin_(session) && siteIdx >= 0 && String(matches[0][siteIdx] || '').trim().toLowerCase() !== String(session.Site || '').trim().toLowerCase()) throw new Error('Anda tidak memiliki akses ke laporan ini.');
    if (String(matches[0][statusIdx] || '').trim().toLowerCase() === 'close') throw new Error('Laporan yang sudah Close tidak dapat menerima lampiran baru.');

    const parentFolder = DriveApp.getFolderById(CONFIG.LAMPIRAN_FOLDER_ID);
    let targetFolder;
    const subFolders = parentFolder.getFoldersByName(idInsiden);
    if (subFolders.hasNext()) targetFolder = subFolders.next();
    else targetFolder = parentFolder.createFolder(idInsiden);
    const blob = Utilities.newBlob(Utilities.base64Decode(fileObject.bytes), fileObject.mimeType, fileObject.fileName);
    const file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const sheet = ss.getSheetByName(CONFIG.TAB_LAMPIRAN);
    const newId = 'ATT-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    sheet.appendRow([newId,idInsiden,jenisLampiran,file.getName(),file.getUrl(),file.getId(),new Date(),session.Nama||session.Username]);
    return {success:true,message:'File berhasil diunggah!',fileUrl:file.getUrl(),fileId:newId};
  } catch (error) {
    return {success:false,message:'Gagal mengunggah file: ' + error.message};
  }
}
/**
 * Attachments.gs — Mengunggah file ke Google Drive
 */

function uploadAttachment(token, idInsiden, fileObject, jenisLampiran) {
  const session = requireSession(token);

  try {
    const parentFolder = DriveApp.getFolderById(CONFIG.LAMPIRAN_FOLDER_ID);
    
    // Sub-folder dinamis berdasarkan ID Insiden
    let targetFolder;
    const subFolders = parentFolder.getFoldersByName(idInsiden);
    if (subFolders.hasNext()) {
      targetFolder = subFolders.next();
    } else {
      targetFolder = parentFolder.createFolder(idInsiden);
    }

    // Decode & Save File
    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileObject.bytes), 
      fileObject.mimeType, 
      fileObject.fileName
    );
    const file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const ss = getTransaksiSS();
    const sheet = ss.getSheetByName(CONFIG.TAB_LAMPIRAN);
    
    const newId = "ATT-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    
    sheet.appendRow([
      newId,
      idInsiden,
      jenisLampiran,
      file.getName(),
      file.getUrl(),
      file.getId(),
      new Date(),
      session.Nama || session.Username
    ]);

    return { 
      success: true, 
      message: "File berhasil diunggah!", 
      fileUrl: file.getUrl(),
      fileId: newId
    };

  } catch (error) {
    return { success: false, message: "Gagal mengunggah file: " + error.toString() };
  }
}
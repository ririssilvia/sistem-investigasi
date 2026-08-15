/** StatusHistory.gs — Update status insiden & catat log riwayat */
function updateIncidentStatus(token, idInsiden, statusBaru, catatan) {
  try {
    const session = requireSession(token);
    const ss = getTransaksiSS();
    const sheetInput = ss.getSheetByName(CONFIG.TAB_INSIDEN);
    if (!sheetInput) throw new Error('Tab Data_Input tidak ditemukan.');
    const dataInput = sheetInput.getDataRange().getValues();
    if (dataInput.length <= 1) throw new Error('Data insiden kosong.');
    const headers = dataInput[0].map(v => String(v || '').trim());
    const idColIdx = incidentHeaderIndex_(headers, ['No Insiden', 'Nomor Insiden', 'No. Insiden', 'ID Insiden']);
    const statusColIdx = incidentHeaderIndex_(headers, ['Status Laporan', 'Status']);
    const siteColIdx = incidentHeaderIndex_(headers, ['Site/Business Unit (BU)', 'Site']);
    if (idColIdx === -1 || statusColIdx === -1) throw new Error('Kolom identitas/status tidak ditemukan.');

    const matches = [];
    for (let i = 1; i < dataInput.length; i++) {
      if (String(dataInput[i][idColIdx] || '').trim() === String(idInsiden).trim()) matches.push(i + 1);
    }
    if (!matches.length) return { success: false, message: 'ID Insiden tidak ditemukan!' };

    if (String(session.Role || '').toLowerCase() !== 'admin' && siteColIdx !== -1 &&
        String(dataInput[matches[0] - 1][siteColIdx] || '').trim().toLowerCase() !== String(session.Site || '').trim().toLowerCase()) {
      throw new Error('Anda tidak memiliki akses untuk mengubah status laporan ini.');
    }

    const statusLama = String(dataInput[matches[0] - 1][statusColIdx] || 'Draft');
    matches.forEach(rowNumber => sheetInput.getRange(rowNumber, statusColIdx + 1).setValue(statusBaru));

    const sheetRiwayat = ss.getSheetByName(CONFIG.TAB_RIWAYAT);
    if (sheetRiwayat) {
      const logId = 'HIS-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      sheetRiwayat.appendRow([logId, idInsiden, statusLama, statusBaru, catatan || '-', new Date(), session.Nama || session.Username]);
    }
    return { success: true, message: `Status insiden ${idInsiden} berhasil diperbarui menjadi '${statusBaru}'`, updatedRows: matches.length };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

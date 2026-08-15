/** StatusHistory.gs — Update status insiden & catat log riwayat */
function updateIncidentStatus(token, idInsiden, statusBaru, catatan) {
  try {
    const session = requireSession(token);
    const isAdmin = String(session.Role || '').trim().toLowerCase() === 'admin';
    const ss = getTransaksiSS();
    const sheetInput = ss.getSheetByName(CONFIG.TAB_INSIDEN);
    if (!sheetInput) throw new Error('Tab Data_Input tidak ditemukan.');
    const dataInput = sheetInput.getDataRange().getValues();
    if (dataInput.length <= 1) throw new Error('Data insiden kosong.');
    const headers = dataInput[0].map(v => String(v || '').trim());
    const idColIdx = incidentHeaderIndex_(headers, ['No Insiden', 'Nomor Insiden', 'No. Insiden', 'ID Insiden']);
    const statusColIdx = headers.length >= 38 ? 37 : incidentHeaderIndex_(headers, ['Status', 'Status Laporan']);
    const completionColIdx = headers.length >= 37 ? 36 : incidentHeaderIndex_(headers, ['Completion Date', 'Tanggal Completion']);
    const siteColIdx = incidentHeaderIndex_(headers, ['Site/Business Unit (BU)', 'Site']);
    if (idColIdx === -1 || statusColIdx === -1) throw new Error('Kolom identitas/status tidak ditemukan.');
    if (headers.length >= 38 && statusColIdx !== 37) throw new Error('Struktur status Data_Input tidak sesuai: AL harus Status.');
    if (headers.length >= 37 && completionColIdx !== 36) throw new Error('Struktur completion date Data_Input tidak sesuai: AK harus Completion Date.');

    const matches = [];
    for (let i = 1; i < dataInput.length; i++) if (String(dataInput[i][idColIdx] || '').trim() === String(idInsiden).trim()) matches.push(i + 1);
    if (!matches.length) return { success: false, message: 'ID Insiden tidak ditemukan!' };
    if (!isAdmin && siteColIdx !== -1 && String(dataInput[matches[0] - 1][siteColIdx] || '').trim().toLowerCase() !== String(session.Site || '').trim().toLowerCase()) throw new Error('Anda tidak memiliki akses untuk mengubah status laporan ini.');

    const statusLama = String(dataInput[matches[0] - 1][statusColIdx] || '').trim();
    if (String(statusBaru || '').trim().toLowerCase() !== 'close') throw new Error('Update Laporan hanya dapat menggunakan status Close.');
    if (statusLama.toLowerCase() === 'close' && !isAdmin) throw new Error('Laporan ini sudah berstatus Close dan tidak dapat di-update kembali.');

    const completionDate = new Date();
    matches.forEach(rowNumber => {
      sheetInput.getRange(rowNumber, statusColIdx + 1).setValue('Close');
      sheetInput.getRange(rowNumber, completionColIdx + 1).setValue(completionDate);
    });
    const sheetRiwayat = ss.getSheetByName(CONFIG.TAB_RIWAYAT);
    if (sheetRiwayat) {
      const logId = 'HIS-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      const activity = statusLama.toLowerCase() === 'close' ? 'Laporan Close diperbarui oleh Admin' : 'Laporan ditutup';
      sheetRiwayat.appendRow([logId, idInsiden, statusLama || 'Open', 'Close', catatan || activity, completionDate, session.Nama || session.Username]);
    }
    return { success: true, message: `Laporan ${idInsiden} berhasil ditutup.`, updatedRows: matches.length, completionDate: completionDate.toISOString() };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

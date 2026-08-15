/** Ambil data detail insiden berdasarkan Nomor Insiden */
function getIncidentDetailForEdit(token, incidentId) {
  try {
    const session = requireSession(token);
    const sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
    if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');

    const values = sheet.getDataRange().getDisplayValues();
    if (values.length <= 1) throw new Error('Data insiden kosong.');

    const headers = values[0].map(v => String(v || '').trim());
    const idIdx = incidentHeaderIndex_(headers, ['No Insiden', 'Nomor Insiden', 'No. Insiden', 'ID Insiden']);

    if (idIdx === -1) throw new Error('Kolom No Insiden tidak ditemukan.');

    // Ambil semua baris transaksi yang memiliki No Insiden yang sama
    const rows = values.slice(1).filter(r => String(r[idIdx] || '').trim() === String(incidentId).trim());

    if (!rows.length) throw new Error('Data insiden tidak ditemukan: ' + incidentId);

    return JSON.stringify({
      success: true,
      headers: headers,
      rows: rows,
      session: session
    });
  } catch (err) {
    return JSON.stringify({ success: false, message: err.message });
  }
}

/** Simpan Update Data Insiden ke Spreadsheet */
function updateIncidentData(token, incidentId, updatedRowsData) {
  try {
    const session = requireSession(token);
    const sheet = getTransaksiSS().getSheetByName(CONFIG.TAB_INSIDEN);
    if (!sheet) throw new Error('Tab Data_Input tidak ditemukan.');

    const values = sheet.getDataRange().getDisplayValues();
    const headers = values[0].map(v => String(v || '').trim());
    const idIdx = incidentHeaderIndex_(headers, ['No Insiden', 'Nomor Insiden', 'No. Insiden', 'ID Insiden']);

    if (idIdx === -1) throw new Error('Kolom No Insiden tidak ditemukan.');

    // 1. Hapus baris-baris lama dengan No Insiden tersebut (mulai dari bawah ke atas agar indeks tidak bergeser)
    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][idIdx] || '').trim() === String(incidentId).trim()) {
        sheet.deleteRow(i + 1);
      }
    }

    // 2. Masukkan baris data yang baru di-update
    if (updatedRowsData && updatedRowsData.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, updatedRowsData.length, updatedRowsData[0].length).setValues(updatedRowsData);
    }

    // 3. Catat ke tab Riwayat_Status jika ada
    const riwayatSheet = getTransaksiSS().getSheetByName(CONFIG.TAB_RIWAYAT);
    if (riwayatSheet) {
      riwayatSheet.appendRow([
        new Date(),
        incidentId,
        'Updated',
        session.Nama + ' (' + session.Site + ')',
        'Pembaruan data laporan insiden'
      ]);
    }

    return JSON.stringify({ success: true, message: 'Data insiden berhasil diperbarui!' });
  } catch (err) {
    return JSON.stringify({ success: false, message: err.message });
  }
}
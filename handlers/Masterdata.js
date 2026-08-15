/**
 * HANDLER MASTER DATA DROPDOWN FORM
 * Dropdown menampilkan ISI/NAMA master, bukan ID.
 * File: handlers/Masterdata.js
 */
function getFormMasterData(token) {
  try {
    requireSession(token);
    const ss = getMasterSS();

    function findDisplayColumn(headers) {
      const normalized = headers.map(h => String(h || '').trim().toLowerCase());

      // Prioritas nama/isi yang umum dipakai pada tabel master.
      const preferred = [
        'nama', 'name', 'nama site', 'site', 'nama perusahaan', 'perusahaan',
        'nama departemen', 'departemen', 'nama lokasi', 'lokasi',
        'nama jabatan', 'jabatan', 'keterangan', 'deskripsi', 'description',
        'uraian', 'value', 'nilai', 'label', 'jam', 'shift'
      ];

      for (const p of preferred) {
        const idx = normalized.indexOf(p);
        if (idx !== -1) return idx;
      }

      // Jika kolom pertama adalah ID/Kode, ambil kolom kedua sebagai isi.
      const first = normalized[0] || '';
      const firstIsId = /^(id|kode|code|no|nomor|key|uuid)$/i.test(first) ||
                        first.includes('id_') || first.endsWith('_id') ||
                        first.includes('kode');
      if (firstIsId && headers.length > 1) return 1;

      // Fallback: kolom pertama tetap dipakai jika memang tidak ada ID.
      return 0;
    }

    function getColumnData(tabKey) {
      const sheetName = CONFIG[tabKey];
      if (!sheetName) throw new Error('CONFIG master tidak ditemukan untuk key: ' + tabKey);

      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error('Tab master tidak ditemukan: ' + sheetName);

      const values = sheet.getDataRange().getValues();
      if (!values || values.length <= 1) return [];

      const headers = values[0];
      const colIndex = findDisplayColumn(headers);

      return Array.from(new Set(
        values.slice(1)
          .map(row => row[colIndex])
          .filter(v => v !== '' && v !== null && v !== undefined)
          .map(v => String(v).trim())
          .filter(v => v !== '')
      ));
    }

    const masterData = {
      site: getColumnData('TAB_SITE'),
      jam: getColumnData('TAB_JAM'),
      lokasi: getColumnData('TAB_LOKASI'),
      perusahaan: getColumnData('TAB_PERUSAHAAN'),
      departemen: getColumnData('TAB_DEPARTEMEN'),
      klasifikasi: getColumnData('TAB_KLASIFIKASI'),
      jabatan: getColumnData('TAB_JABATAN'),
      umur: getColumnData('TAB_UMUR'),
      masaKerja: getColumnData('TAB_MASA_KERJA'),
      kategoriAlat: getColumnData('TAB_KATEGORI_ALAT'),
      jenisAlat: getColumnData('TAB_JENIS_ALAT'),
      jenisKontak: getColumnData('TAB_JENIS_KONTAK'),
      sumberKecelakaan: getColumnData('TAB_SUMBER_KECELAKAAN'),
      tta: getColumnData('TAB_TTA'),
      ketTta: getColumnData('TAB_KET_TTA'),
      kta: getColumnData('TAB_KTA'),
      ketKta: getColumnData('TAB_KET_KTA'),
      faktorManusia: getColumnData('TAB_FAKTOR_MANUSIA'),
      ketFaktorManusia: getColumnData('TAB_KET_FAKTOR_MANUSIA'),
      faktorPekerjaan: getColumnData('TAB_FAKTOR_PEKERJAAN'),
      ketFaktorPekerjaan: getColumnData('TAB_KET_FAKTOR_PEKERJAAN'),
      kurangKendali: getColumnData('TAB_KURANG_KENDALI'),
      ketKurangKendali: getColumnData('TAB_KET_KURANG_KENDALI')
    };

    return { success: true, master: masterData };
  } catch (err) {
    Logger.log('Error getFormMasterData: ' + err.stack);
    return { success: false, message: 'Gagal memuat Master Data: ' + err.message };
  }
}

/**
 * HANDLER MASTER DATA DROPDOWN FORM
 *
 * ATURAN:
 * 1. Dropdown menampilkan ISI/NAMA master, BUKAN ID/KODE.
 * 2. Nilai yang dikirim dari form juga ISI/NAMA tersebut.
 * 3. ID/KODE hanya diabaikan dan tidak pernah dikirim sebagai option.
 */
function getFormMasterData(token) {
  try {
    requireSession(token);
    const ss = getMasterSS();

    // Kata kunci kolom tampilan untuk masing-masing master.
    // Ini dibuat lebih spesifik supaya kolom ID seperti "ID_JAM" tidak terpilih.
    const displayKeywords = {
      TAB_SITE: ['nama site', 'site name', 'site', 'nama'],
      TAB_JAM: ['jam kejadian', 'jam', 'waktu kejadian', 'waktu', 'time'],
      TAB_LOKASI: ['nama lokasi', 'lokasi spesifik', 'lokasi', 'location', 'nama'],
      TAB_PERUSAHAAN: ['nama perusahaan', 'perusahaan', 'company', 'nama'],
      TAB_DEPARTEMEN: ['nama departemen', 'departemen', 'department', 'nama'],
      TAB_KLASIFIKASI: ['klasifikasi kecelakaan', 'klasifikasi', 'classification', 'nama'],
      TAB_JABATAN: ['nama jabatan', 'jabatan', 'position', 'job title', 'nama'],
      TAB_UMUR: ['kelompok umur', 'umur', 'usia', 'age', 'nama'],
      TAB_MASA_KERJA: ['masa kerja', 'tenure', 'lama kerja', 'nama'],
      TAB_KATEGORI_ALAT: ['kategori alat', 'nama kategori', 'kategori', 'nama'],
      TAB_JENIS_ALAT: ['jenis alat', 'nama jenis alat', 'nama alat', 'jenis', 'nama'],
      TAB_JENIS_KONTAK: ['jenis kontak', 'kontak', 'contact', 'nama'],
      TAB_SUMBER_KECELAKAAN: ['sumber kecelakaan', 'sumber', 'source', 'nama'],
      TAB_TTA: ['tindakan tidak aman', 'tta', 'nama tindakan', 'tindakan', 'nama'],
      TAB_KET_TTA: ['keterangan tta', 'keterangan', 'deskripsi', 'uraian', 'nama'],
      TAB_KTA: ['kondisi tidak aman', 'kta', 'nama kondisi', 'kondisi', 'nama'],
      TAB_KET_KTA: ['keterangan kta', 'keterangan', 'deskripsi', 'uraian', 'nama'],
      TAB_FAKTOR_MANUSIA: ['faktor manusia', 'human factor', 'faktor', 'nama'],
      TAB_KET_FAKTOR_MANUSIA: ['keterangan faktor manusia', 'keterangan', 'deskripsi', 'uraian', 'nama'],
      TAB_FAKTOR_PEKERJAAN: ['faktor pekerjaan', 'job factor', 'faktor', 'nama'],
      TAB_KET_FAKTOR_PEKERJAAN: ['keterangan faktor pekerjaan', 'keterangan', 'deskripsi', 'uraian', 'nama'],
      TAB_KURANG_KENDALI: ['kurang kendali', 'management control', 'kendali', 'nama'],
      TAB_KET_KURANG_KENDALI: ['keterangan kurang kendali', 'keterangan', 'deskripsi', 'uraian', 'nama']
    };

    function normalizeHeader(value) {
      return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/[\s_-]+/g, ' ');
    }

    function isIdOrCodeHeader(header) {
      const h = normalizeHeader(header);
      return /^(id|kode|code|no|nomor|key|uuid|id master|kode master)$/.test(h) ||
             /^(id|kode|code)\s+/.test(h) ||
             /\s+(id|kode|code)$/.test(h) ||
             h.includes('created by') ||
             h.includes('created at') ||
             h.includes('updated by') ||
             h.includes('updated at');
    }

    function findDisplayColumn(headers, tabKey) {
      const normalized = headers.map(normalizeHeader);
      const keywords = (displayKeywords[tabKey] || []).map(normalizeHeader);

      // Skor kandidat. Kolom yang persis sama dengan keyword mendapat skor tertinggi.
      let bestIndex = -1;
      let bestScore = -1;

      normalized.forEach((header, index) => {
        if (!header || isIdOrCodeHeader(header)) return;

        let score = 0;
        keywords.forEach(keyword => {
          if (!keyword) return;
          if (header === keyword) score = Math.max(score, 100);
          else if (header.includes(keyword)) score = Math.max(score, 80);
          else {
            const words = keyword.split(' ');
            if (words.some(word => word.length >= 3 && header.includes(word))) {
              score = Math.max(score, 50);
            }
          }
        });

        // Kolom kedua biasanya merupakan kolom isi jika kolom pertama ID.
        if (score === 0 && index === 1 && isIdOrCodeHeader(normalized[0])) score = 20;

        if (score > bestScore) {
          bestScore = score;
          bestIndex = index;
        }
      });

      if (bestIndex !== -1) return bestIndex;

      // Fallback terakhir: ambil kolom pertama yang bukan ID/Kode.
      for (let i = 0; i < normalized.length; i++) {
        if (normalized[i] && !isIdOrCodeHeader(normalized[i])) return i;
      }

      return 0;
    }

    function getDisplayData(tabKey) {
      const sheetName = CONFIG[tabKey];
      if (!sheetName) throw new Error('CONFIG master tidak ditemukan untuk key: ' + tabKey);

      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error('Tab master tidak ditemukan: ' + sheetName);

      const values = sheet.getDataRange().getValues();
      if (!values || values.length <= 1) return [];

      const headers = values[0];
      const displayColumn = findDisplayColumn(headers, tabKey);

      Logger.log('Master ' + sheetName + ' menggunakan kolom tampilan: ' + headers[displayColumn]);

      return Array.from(new Set(
        values.slice(1)
          .map(row => row[displayColumn])
          .filter(value => value !== '' && value !== null && value !== undefined)
          .map(value => String(value).trim())
          .filter(value => value !== '')
      ));
    }

    const masterData = {
      site: getDisplayData('TAB_SITE'),
      jam: getDisplayData('TAB_JAM'),
      lokasi: getDisplayData('TAB_LOKASI'),
      perusahaan: getDisplayData('TAB_PERUSAHAAN'),
      departemen: getDisplayData('TAB_DEPARTEMEN'),
      klasifikasi: getDisplayData('TAB_KLASIFIKASI'),
      jabatan: getDisplayData('TAB_JABATAN'),
      umur: getDisplayData('TAB_UMUR'),
      masaKerja: getDisplayData('TAB_MASA_KERJA'),
      kategoriAlat: getDisplayData('TAB_KATEGORI_ALAT'),
      jenisAlat: getDisplayData('TAB_JENIS_ALAT'),
      jenisKontak: getDisplayData('TAB_JENIS_KONTAK'),
      sumberKecelakaan: getDisplayData('TAB_SUMBER_KECELAKAAN'),
      tta: getDisplayData('TAB_TTA'),
      ketTta: getDisplayData('TAB_KET_TTA'),
      kta: getDisplayData('TAB_KTA'),
      ketKta: getDisplayData('TAB_KET_KTA'),
      faktorManusia: getDisplayData('TAB_FAKTOR_MANUSIA'),
      ketFaktorManusia: getDisplayData('TAB_KET_FAKTOR_MANUSIA'),
      faktorPekerjaan: getDisplayData('TAB_FAKTOR_PEKERJAAN'),
      ketFaktorPekerjaan: getDisplayData('TAB_KET_FAKTOR_PEKERJAAN'),
      kurangKendali: getDisplayData('TAB_KURANG_KENDALI'),
      ketKurangKendali: getDisplayData('TAB_KET_KURANG_KENDALI')
    };

    return { success: true, master: masterData };
  } catch (err) {
    Logger.log('Error getFormMasterData: ' + err.stack);
    return {
      success: false,
      message: 'Gagal memuat Master Data: ' + err.message
    };
  }
}

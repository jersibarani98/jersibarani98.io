/**
 * Modul Perhitungan Status Otomatis (SIMONI LKD-LAD)
 * Tanggung Jawab: Menghitung skor kinerja berdasarkan data fakta terbaru.
 * Versi: Dynamic Header Mapping (Aman terhadap pergeseran kolom).
 * Catatan: Identitas Desa tidak masuk dalam bobot penilaian (0 poin).
 */

/**
 * 1. UPDATE STATUS LKD
 * Bobot Skor Total: 100
 */
function updateLkdStatus(idEntry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('DATA_LKD_MASTER');
  if (!masterSheet) return;

  const fullData = masterSheet.getDataRange().getValues();
  const headers = fullData[0];
  const rowData = fullData.find(r => r[0] === idEntry);
  if (!rowData) return;

  // Helper untuk mengambil nilai berdasarkan nama header
  const getV = (hName) => {
    const idx = headers.indexOf(hName);
    return idx > -1 ? rowData[idx] : null;
  };

  let skor = 0;

  // A. LEGALITAS (10 Poin)
  if (getV("SK_NOMOR") && getV("SK_TAHUN")) skor += 5;
  if (getV("PERDES_ADA") === "Ya") skor += 5;

  // B. SARANA PRASARANA (10 Poin)
  // Identitas Desa (Alamat Kantor) = 0 Poin (Tidak dimasukkan bobot)
  if (getV("ADA_SEKRETARIAT") === "Ada") {
    skor += 10; // Bobot sarana ditingkatkan karena identitas dihapus
  }

  // C. PENINGKATAN KAPASITAS / BIMTEK (20 Poin)
  if (getV("KAPASITAS_ADA") === "Ya") {
    skor += 20; 
  }

  // D. DOKUMEN KERJA / AD-ART / SOP (20 Poin)
  if (getV("DOK_KERJA_ADA") === "Ya") {
    skor += 20;
  }

  // E. PARTISIPASI PERENCANAAN (10 Poin)
  if (getV("IKUT_MUSDES") === "Ya" || getV("IKUT_MUSRENBANG") === "Ya") {
    skor += 10;
  }

  // F. DUKUNGAN ANGGARAN (15 Poin)
  if (getV("ANGGARAN_ADA") === "Ya" || (Number(getV("ANGGARAN_JUMLAH")) > 0)) {
    skor += 15;
  }

  // G. AKUNTABILITAS / LAPORAN AKHIR (15 Poin)
  if (getV("ADA_LAPORAN_AKHIR") === "Ya") {
    skor += 15;
  }

  saveToStatusLog(idEntry, skor);
}

/**
 * 2. UPDATE STATUS LAD
 * Bobot Skor Total: 100
 */
function updateLadStatus(idEntry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('DATA_LAD_MASTER');
  if (!masterSheet) return;

  const fullData = masterSheet.getDataRange().getValues();
  const headers = fullData[0];
  const rowData = fullData.find(r => r[0] === idEntry);
  if (!rowData) return;

  const getV = (hName) => {
    const idx = headers.indexOf(hName);
    return idx > -1 ? rowData[idx] : null;
  };

  let skor = 0;

  // A. LEGALITAS (15 Poin)
  // Identitas Desa (Alamat Kantor) = 0 Poin
  if (getV("SK_NOMOR") && getV("SK_TAHUN")) {
    skor += 15; // Bobot SK ditingkatkan
  }

  // B. SARANA FISIK ADAT (15 Poin)
  if (getV("ADA_BALAI") === "Ada" || getV("ADA_RUMAH_ADAT") === "Ada") {
    skor += 15;
  }

  // C. KAPASITAS & BIMTEK (20 Poin)
  if (getV("KAPASITAS_ADA") === "Ya") {
    skor += 20;
  }

  // D. DOKUMEN KERJA ADAT (20 Poin)
  if (getV("DOK_KERJA_ADA") === "Ya") {
    skor += 20;
  }

  // E. KEAKTIFAN FUNGSIONAL (15 Poin)
  if (getV("SENGKETA_JENIS") && getV("SENGKETA_JENIS") !== "") {
    skor += 15;
  }

  // F. ANGGARAN & LAPORAN (15 Poin)
  if (Number(getV("ANGGARAN_JUMLAH")) > 0) skor += 7;
  if (getV("ADA_LAPORAN_AKHIR") === "Ya") skor += 8;

  saveToStatusLog(idEntry, skor);
}

/**
 * 3. HELPER: SIMPAN KE LOG STATUS
 */
function saveToStatusLog(idEntry, skor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('STATUS_LOG');
  if (!logSheet) return;

  let status = "TIDAK AKTIF";
  if (skor >= 80) status = "AKTIF";
  else if (skor >= 50) status = "KURANG AKTIF";

  const logData = logSheet.getDataRange().getValues();
  const index = logData.findIndex(l => l[0] === idEntry);
  
  if (index > -1) {
    logSheet.getRange(index + 1, 2, 1, 3).setValues([[skor, status, new Date()]]);
  } else {
    logSheet.appendRow([idEntry, skor, status, new Date()]);
  }
}
/**
 * Modul Manajemen User Otomatis
 * Proyek: SIMONI LKD-LAD
 * Tanggung Jawab: Menghasilkan akun berdasarkan data hirarki di sheet WILAYAH dengan kata sandi acak (Alfanumerik).
 */

function generateRegionalUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wilayahSheet = ss.getSheetByName('WILAYAH');
  const userSheet = ss.getSheetByName('USERS');
  
  if (!wilayahSheet || !userSheet) {
    return "Error: Sheet WILAYAH atau USERS tidak ditemukan. Jalankan setupDatabase() terlebih dahulu.";
  }

  const wilayahData = wilayahSheet.getDataRange().getValues();
  if (wilayahData.length < 2) return "Error: Sheet WILAYAH masih kosong.";
  
  // 1. Ambil data user yang sudah ada untuk menghindari duplikasi username
  const existingUsers = userSheet.getDataRange().getValues();
  const existingUsernames = new Set(existingUsers.slice(1).map(row => row[1]));
  
  // Data wilayah tanpa header
  const rows = wilayahData.slice(1);

  // Variabel penampung proses batch
  let internalProcessed = new Set();
  let newUsersBatch = [];
  
  /**
   * SISTEM KATA SANDI ACAK (ALFANUMERIK):
   * Setiap akun akan mendapatkan 6 karakter kombinasi huruf dan angka yang unik.
   */

  // 1. PROSES AKUN PROVINSI
  const idProvSultra = "74"; 
  const userProv = "sultra_prov";
  if (!existingUsernames.has(userProv)) {
    newUsersBatch.push([
      Utilities.getUuid(), 
      userProv, 
      generateRandomPassword(6), // Password acak alfanumerik untuk Provinsi
      "PROVINSI", 
      idProvSultra, "", "", "", "AKTIF"
    ]);
    internalProcessed.add(userProv);
  }

  // 2. ITERASI SEMUA BARIS DI SHEET WILAYAH
  rows.forEach(row => {
    const idDesa = row[0];       // Kolom A: ID_DESA
    const kabName = row[2];      // Kolom C: KABUPATEN
    const kecName = row[3];      // Kolom D: KECAMATAN
    const desaName = row[4];     // Kolom E: DESA_KELURAHAN
    
    if (!idDesa || !kabName || !kecName || !desaName) return;

    // Parsing ID standar: 74.XX.XX.XXXX
    const idParts = idDesa.toString().split('.');
    const idKab = idParts.length > 1 ? idParts.slice(0, 2).join('.') : idDesa;
    const idKec = idParts.length > 2 ? idParts.slice(0, 3).join('.') : idDesa;

    // A. AKUN KABUPATEN
    const usernameKab = "kab_" + kabName.toString().toLowerCase().replace(/\s+/g, '_');
    if (!existingUsernames.has(usernameKab) && !internalProcessed.has(usernameKab)) {
      newUsersBatch.push([
        Utilities.getUuid(), 
        usernameKab, 
        generateRandomPassword(6), // Password acak alfanumerik untuk Kabupaten
        "KABUPATEN", 
        idProvSultra, idKab, "", "", "AKTIF"
      ]);
      internalProcessed.add(usernameKab);
    }

    // B. AKUN KECAMATAN
    const usernameKec = "kec_" + kecName.toString().toLowerCase().replace(/\s+/g, '_');
    if (!existingUsernames.has(usernameKec) && !internalProcessed.has(usernameKec)) {
      newUsersBatch.push([
        Utilities.getUuid(), 
        usernameKec, 
        generateRandomPassword(6), // Password acak alfanumerik untuk Kecamatan
        "KECAMATAN", 
        idProvSultra, idKab, idKec, "", "AKTIF"
      ]);
      internalProcessed.add(usernameKec);
    }

    // C. AKUN DESA
    const usernameDesa = "desa_" + desaName.toString().toLowerCase().replace(/\s+/g, '_');
    let finalUserDesa = usernameDesa;
    if (existingUsernames.has(finalUserDesa) || internalProcessed.has(finalUserDesa)) {
      finalUserDesa = usernameDesa + "_" + idDesa.toString().replace(/\./g, '');
    }

    if (!existingUsernames.has(finalUserDesa) && !internalProcessed.has(finalUserDesa)) {
      newUsersBatch.push([
        Utilities.getUuid(), 
        finalUserDesa, 
        generateRandomPassword(6), // Password acak alfanumerik untuk Desa
        "DESA", 
        idProvSultra, idKab, idKec, idDesa, "AKTIF"
      ]);
      internalProcessed.add(finalUserDesa);
    }
  });

  // 3. SIMPAN SEMUA DATA SEKALIGUS
  if (newUsersBatch.length > 0) {
    const lastRow = userSheet.getLastRow();
    userSheet.getRange(lastRow + 1, 1, newUsersBatch.length, newUsersBatch[0].length)
             .setValues(newUsersBatch);
    return "Selesai! Berhasil membuat " + newUsersBatch.length + " akun pengguna wilayah dengan password kombinasi huruf dan angka.";
  } else {
    return "Tidak ada akun baru. Semua wilayah sudah terdaftar.";
  }
}

/**
 * Fungsi pembantu untuk menghasilkan kata sandi acak (Alfanumerik)
 * @param {number} length Panjang karakter yang diinginkan
 * @return {string} Kata sandi dalam bentuk string
 */
function generateRandomPassword(length) {
  let result = '';
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  const charactersLength = characters.length;
  for (let i = 0; i < length; i++) {
    result += characters.charAt(Math.floor(Math.random() * charactersLength));
  }
  return result;
}

/**
 * Helper: Cek apakah username sudah terdaftar
 */
function checkUserExists(sheet, username) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  return data.some(row => row[0] === username);
}
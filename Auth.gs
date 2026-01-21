/**
 * Modul Autentikasi Sistem SIMONI LKD-LAD (Versi Regional)
 * Optimasi: Mendeteksi cakupan wilayah berdasarkan Role Pengguna.
 */

function authenticateUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('USERS');
  const wilayahSheet = ss.getSheetByName('WILAYAH');
  
  if (!userSheet) return { success: false, message: 'Tabel pengguna tidak ditemukan.' };

  const userData = userSheet.getDataRange().getValues();
  const inputUser = username.toString().trim().toLowerCase();
  const inputPass = password.toString().trim();
  const scriptUrl = ScriptApp.getService().getUrl();

  for (let i = 1; i < userData.length; i++) {
    const storedUser = userData[i][1] ? userData[i][1].toString().trim().toLowerCase() : ""; 
    const storedPass = userData[i][2] ? userData[i][2].toString().trim() : ""; 
    const userRole = userData[i][3] ? userData[i][3].toString().toUpperCase().trim() : "";
    
    // Ambil parameter wilayah dari sheet USERS
    const idProv = userData[i][4];
    const idKab  = userData[i][5];
    const idKec  = userData[i][6];
    const idDesa = userData[i][7];
    
    if (storedUser === inputUser) {
      if (inputPass === storedPass) {
        
        let displayRegion = "Prov. Sulawesi Tenggara";
        
        // --- LOGIKA IDENTIFIKASI NAMA WILAYAH BERDASARKAN HIERARKI ---
        if (wilayahSheet) {
          const wData = wilayahSheet.getDataRange().getValues();
          
          if (userRole === "KABUPATEN" && idKab) {
            // Mencari baris wilayah yang sesuai dengan ID Kabupaten
            const row = wData.find(r => r[0].toString().startsWith(idKab.toString()));
            if (row) displayRegion = "Kab/Kota: " + row[2];
          } 
          else if (userRole === "KECAMATAN" && idKec) {
            // Mencari baris wilayah yang sesuai dengan ID Kecamatan
            const row = wData.find(r => r[0].toString().startsWith(idKec.toString()));
            if (row) displayRegion = "Kecamatan: " + row[3];
          } 
          else if (userRole === "DESA" && idDesa) {
            // Mencari baris spesifik Desa
            const row = wData.find(r => r[0].toString() === idDesa.toString());
            if (row) displayRegion = "Desa: " + row[4];
          }
        }

        // Tentukan halaman tujuan (Desa langsung ke Form, Admin ke Dashboard)
        const targetPage = (userRole === "DESA") ? "form" : "dashboard";
        
        return { 
          success: true, 
          url: scriptUrl,
          targetPage: targetPage,
          user: {
            id: userData[i][0],
            username: userData[i][1],
            role: userRole,
            idProv: idProv,
            idKab: idKab,
            idKec: idKec,
            idDesa: idDesa,
            namaWilayah: displayRegion 
          }
        };
      } else {
        return { success: false, message: 'Password salah.' };
      }
    }
  }
  return { success: false, message: 'Username tidak terdaftar.' };
}
/**
 * SISTEM MONITORING LEMBAGA KEMASYARAKATAN DESA & LEMBAGA ADAT DESA (SIMONI LKD-LAD)
 * PEMERINTAH PROVINSI SULAWESI TENGGARA
 * * Modul: Code.gs (Versi Komulatif Maksimal - 2026)
 * Tanggung Jawab: 
 * 1. Navigasi Halaman (doGet)
 * 2. Autentikasi & Manajemen Sesi
 * 3. Deteksi Otomatis Identitas Desa & Kepala Desa
 * 4. API Dashboard (Statistik, Anggaran, Ranking)
 * 5. Logika Penyimpanan Data Fakta (LKD & LAD)
 */

/**
 * 1. FUNGSI NAVIGASI UTAMA (doGet)
 * Mengatur alur masuk pengguna dan proteksi halaman berdasarkan Role.
 */
function doGet(e) {
  const user = getUserSession();
  let page = e.parameter.page || 'index';

  // Proteksi Keamanan: Jika sesi kosong, arahkan ke login (index)
  if (!user) {
    page = 'index';
  } else {
    // Jika sudah login tapi mengakses index, arahkan ke beranda sesuai Role
    if (page.toLowerCase() === 'index') {
      if (user.role === 'DESA') {
        page = 'form';
      } else {
        page = 'dashboard';
      }
    }
    // Proteksi Role: Akun Desa dilarang mengakses Dashboard Admin
    if (user.role === 'DESA' && page.toLowerCase() === 'dashboard') {
      page = 'form';
    }
  }

  // Pemetaan Nama File (Case Sensitive sesuai instruksi penamaan)
  let fileName = 'Index'; 
  switch(page.toLowerCase()) {
    case 'dashboard': 
      fileName = 'Dashboard'; 
      break;
    case 'form':      
      fileName = 'FormInput'; 
      break;
    default:          
      fileName = 'Index';
  }

  try {
    const template = HtmlService.createTemplateFromFile(fileName);
    template.user = user; // Mengirim data user ke sisi klien

    return template.evaluate()
        .setTitle('SIMONI LKD-LAD SULTRA')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("<h3>Terjadi Kesalahan Sistem</h3><p>Gagal memuat halaman: " + err.message + "</p>");
  }
}

/**
 * 2. AUTO-DETECT IDENTITAS DESA (FITUR SEKALI ISI)
 * Mencari data Kepala Desa dan Alamat Kantor Desa yang sudah pernah diinput sebelumnya.
 */
function getExistingVillageIdentity() {
  const user = getUserSession();
  if (!user || !user.idDesa) {
    return null;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToCheck = ['DATA_LKD_MASTER', 'DATA_LAD_MASTER'];
  
  for (let sName of sheetsToCheck) {
    const sheet = ss.getSheetByName(sName);
    if (!sheet) continue;
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    // Iterasi dari baris terbawah ke atas (untuk mendapatkan data paling mutakhir)
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      // Kolom index 1 adalah ID_DESA
      if (row[1].toString() === user.idDesa.toString()) {
        
        let identityFound = {
          nama: "",
          gender: "",
          hp: "",
          alamat: ""
        };

        if (sName === 'DATA_LKD_MASTER') {
          // Berdasarkan Setup.gs: Nama(56), Gen(57), HP(58), Alamat Kantor(59)
          identityFound.nama = row[56];
          identityFound.gender = row[57];
          identityFound.hp = row[58];
          identityFound.alamat = row[59];
        } else {
          // Berdasarkan Setup.gs (LAD): Nama(51), Gen(52), HP(53), Alamat Kantor(54)
          identityFound.nama = row[51];
          identityFound.gender = row[52];
          identityFound.hp = row[53];
          identityFound.alamat = row[54];
        }
        
        // Kembalikan hanya jika nama Kades sudah terisi
        if (identityFound.nama && identityFound.nama !== "") {
          return identityFound;
        }
      }
    }
  }
  return null;
}

/**
 * 3. API DASHBOARD DATA FLOW
 * Menyuplai statistik, data anggaran per wilayah, dan daftar laporan lembaga.
 */
function getDashboardData() {
  const user = getUserSession();
  if (!user) return null;

  try {
    // 3.1 Ambil statistik dasar dari modul Visual.gs
    const stats = typeof getDashboardStats === 'function' 
                  ? getDashboardStats(user) 
                  : { total:0, aktif:0, kurang:0, tidak:0, lkdDist: {}, bimtekPct: 0 };
    
    // 3.2 Ambil data agregat wilayah (ranking & progress)
    const regionalData = typeof getRegionalAggregateData === 'function' 
                         ? getRegionalAggregateData(user) 
                         : { progress: [], lkdAgregat: [], ladAgregat: [], ranking: { top: [], bottom: [] } };
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lkdSheet = ss.getSheetByName('DATA_LKD_MASTER');
    const ladSheet = ss.getSheetByName('DATA_LAD_MASTER');
    const logSheet = ss.getSheetByName('STATUS_LOG');
    
    // 3.3 Normalisasi Log Status (ID -> Skor/Status)
    const logsRaw = logSheet ? logSheet.getDataRange().getValues() : [];
    const logMap = new Map();
    logsRaw.forEach(l => {
      if(l[0]) logMap.set(l[0].toString().trim().toUpperCase(), { skor: l[1], status: l[2] });
    });

    const list = [];
    const budgetMap = {}; // Map untuk akumulasi anggaran per Kabupaten
    
    // Helper: Filter Wilayah untuk Akun Kabupaten
    const isVisible = (rowKabName) => {
      if (user.role === "PROVINSI") return true;
      if (user.role === "KABUPATEN") {
        const userKabClean = user.namaWilayah.replace("Kab/Kota: ", "").trim().toUpperCase();
        return rowKabName.toString().trim().toUpperCase() === userKabClean;
      }
      return false;
    };

    // Helper: Pembersih Format Angka Mata Uang
    const cleanNumber = (val) => {
      if (!val) return 0;
      const cleaned = val.toString().replace(/[^0-9]/g, "");
      return parseInt(cleaned) || 0;
    };

    // 3.4 Proses Data LKD (Analisis Indikator & Anggaran)
    if (lkdSheet && lkdSheet.getLastRow() > 1) {
      const dataLkd = lkdSheet.getDataRange().getValues();
      for (let i = 1; i < dataLkd.length; i++) {
        const r = dataLkd[i];
        if (!r[0] || !isVisible(r[2])) continue; 
        
        const kabKey = r[2].toString().trim().toUpperCase();
        if (!budgetMap[kabKey]) budgetMap[kabKey] = { total: 0, count: 0 };

        const id = r[0].toString().trim().toUpperCase();
        const log = logMap.get(id) || { skor: 0, status: "BELUM INPUT" };
        
        let terpenuhi = [];
        let tdkTerpenuhi = [];
        
        // Analisis Indikator Sederhana untuk Tooltip Dashboard
        (r[10] && r[11]) ? terpenuhi.push("Legalitas") : tdkTerpenuhi.push("Legalitas");
        (r[13] === "Ya") ? terpenuhi.push("Regulasi") : tdkTerpenuhi.push("Regulasi");
        
        const nominal = cleanNumber(r[33]);
        if (r[30] === "Ya" || nominal > 0) {
          terpenuhi.push("Anggaran");
          budgetMap[kabKey].total += nominal;
          budgetMap[kabKey].count += 1;
        } else {
          tdkTerpenuhi.push("Anggaran");
        }

        list.push({
          id: r[0], nama: r[8], kab: r[2], kec: r[3], des: r[4],
          kat: "LKD (" + r[5] + ")", skor: log.skor, status: log.status,
          terpenuhi: terpenuhi.join(", "), tidak_terpenuhi: tdkTerpenuhi.join(", "),
          anggaran_val: nominal
        });
      }
    }
    
    // 3.5 Proses Data LAD
    if (ladSheet && ladSheet.getLastRow() > 1) {
      const dataLad = ladSheet.getDataRange().getValues();
      for (let i = 1; i < dataLad.length; i++) {
        const r = dataLad[i];
        if (!r[0] || !isVisible(r[2])) continue;
        
        const kabKey = r[2].toString().trim().toUpperCase();
        if (!budgetMap[kabKey]) budgetMap[kabKey] = { total: 0, count: 0 };

        const id = r[0].toString().trim().toUpperCase();
        const log = logMap.get(id) || { skor: 0, status: "BELUM INPUT" };
        const nominalLad = cleanNumber(r[29]);

        if (nominalLad > 0) {
          budgetMap[kabKey].total += nominalLad;
          budgetMap[kabKey].count += 1;
        }

        list.push({
          id: r[0], nama: r[5], kab: r[2], kec: r[3], des: r[4],
          kat: "LAD", skor: log.skor, status: log.status,
          anggaran_val: nominalLad
        });
      }
    }

    // Sinkronisasi Anggaran ke Data Progres Wilayah
    if (regionalData.progress) {
      regionalData.progress.forEach(p => {
        const bData = budgetMap[p.name.toUpperCase()] || { total: 0, count: 0 };
        p.budgetTotal = bData.total;
        p.budgetCount = bData.count;
      });
    }

    return {
      total: stats.total || 0, 
      aktif: stats.aktif || 0, 
      kurang: stats.kurang || 0, 
      tidak: stats.tidak || 0,
      perdesLkd: stats.perdesLkd || 0, 
      perdesLad: stats.perdesLad || 0,
      lkdDist: stats.lkdDist || {}, 
      bimtekPct: stats.bimtekPct || 0, 
      ranking: regionalData.ranking || { top: [], bottom: [] }, 
      list: list, 
      progress: regionalData.progress, 
      lkdAgregat: regionalData.lkdAgregat, 
      ladAgregat: regionalData.ladAgregat
    };
  } catch (e) { 
    return { total: 0, list: [], error: e.toString() }; 
  }
}

/**
 * 4. SIMPAN DATA LKD (Lembaga Kemasyarakatan Desa)
 * Menyimpan data lengkap termasuk Identitas Kades, Kapasitas, Dokumen, dan Laporan.
 */
function saveLkdData(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DATA_LKD_MASTER');
    const wilayahSheet = ss.getSheetByName('WILAYAH');
    const pengurusSheet = ss.getSheetByName('DETAIL_PENGURUS_LKD');

    // Identifikasi Wilayah secara presisi
    const wData = wilayahSheet.getDataRange().getValues();
    const wRow = wData.find(r => r[0].toString() === payload.idDesa.toString());
    const kab = wRow ? wRow[2] : "";
    const kec = wRow ? wRow[3] : "";
    const des = wRow ? wRow[4] : "";
    
    // Generate UUID untuk ID Entry
    const idEntry = "LKD-" + Utilities.getUuid().substring(0,8).toUpperCase();
    
    // Proses Upload Berkas ke Google Drive
    const skUrl = uploadToDrive(payload.skFile, "SK_" + idEntry, "SK_LKD");
    const perdesUrl = uploadToDrive(payload.perdesFile, "PERDES_" + idEntry, "PERDES_LKD");
    const sekreUrl = uploadToDrive(payload.sekreFile, "FOTO_" + idEntry, "FOTO_SEKRE");
    const laporanUrl = uploadToDrive(payload.laporanFile, "LAPORAN_" + idEntry, "LAPORAN_LKD");
    const dokumenUrl = uploadToDrive(payload.dokumenFile, "DOKUMEN_" + idEntry, "DOK_KERJA_LKD");

    /**
     * PEMETAAN KOLOM (Sesuai Setup.gs):
     * 0-41: Metadata & Input Dasar LKD
     * 43-48: Peningkatan Kapasitas
     * 49-51: Dokumen Kerja
     * 52-55: Laporan Akhir Kegiatan
     * 56-59: Identitas Desa & Kepala Desa
     */
    const rowDataMaster = [
      idEntry, payload.idDesa, kab, kec, des,                    // 0-4
      payload.jenis, payload.subUnitPkk, payload.unitNo,         // 5-7
      payload.namaLembaga, payload.alamat,                       // 8-9
      payload.skNomor, payload.skTahun, skUrl,                   // 10-12
      payload.perdesAda, payload.perdesNomor, payload.perdesTahun, perdesUrl, // 13-16
      payload.adaSekre, sekreUrl, payload.saranaKerja, "-",      // 17-20
      "-", "-", "-", "-",                                         // 21-24 (Placeholder File Dokumen Lama)
      payload.ikutMusdes ? "Ya" : "Tidak", 
      payload.ikutMusren ? "Ya" : "Tidak", 
      payload.ikutRpjm ? "Ya" : "Tidak", 
      payload.ikutRkp ? "Ya" : "Tidak", "-",                     // 25-29
      payload.anggaranAda, payload.anggaranTahun, payload.anggaranSumber, 
      payload.anggaranJumlah, "-", payload.anggaranUntuk,         // 30-35
      "-", payload.laporanTahun, payload.laporanCatatan,         // 36-38
      payload.penginputNama, payload.penginputJabatan, payload.penginputHp, // 39-41
      new Date(),                                                // 42 (TGL_INPUT)
      // DATA BARU (HASIL DISKUSI TERAKHIR)
      payload.kapasitasAda, payload.kapasitasNama, payload.kapasitasJenis, payload.kapasitasMateri, payload.kapasitasPenyelenggara, payload.kapasitasTahun, // 43-48
      payload.dokumenAda, payload.dokumenJenis, dokumenUrl,      // 49-51
      payload.adaLaporan, payload.laporanTahun, laporanUrl, payload.laporanCatatan, // 52-55
      // IDENTITAS KADES (KOLOM TERAKHIR)
      payload.kadesNama, payload.kadesGender, payload.kadesHp, payload.kadesAlamat // 56-59
    ];
    
    sheet.appendRow(rowDataMaster);

    // Simpan Detail Pengurus (dengan Jenis Kelamin & Nomor HP)
    if (payload.pengurus && payload.pengurus.length > 0) {
      payload.pengurus.forEach(p => {
        pengurusSheet.appendRow([
          Utilities.getUuid().substring(0,5), idEntry, kab, kec, des, 
          p.nama, p.gender, p.jabatan, "-", "-", p.hp
        ]);
      });
    }

    // Trigger Perhitungan Skor Otomatis
    if (typeof updateLkdStatus === 'function') {
      updateLkdStatus(idEntry);
    }
    
    return { success: true, message: "Data LKD berhasil disimpan. ID: " + idEntry };
  } catch (e) { 
    return { success: false, message: "Kesalahan Simpan LKD: " + e.toString() }; 
  }
}

/**
 * 5. SIMPAN DATA LAD (Lembaga Adat Desa)
 */
function saveLadData(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DATA_LAD_MASTER');
    const wilayahSheet = ss.getSheetByName('WILAYAH');
    const pengurusSheet = ss.getSheetByName('DETAIL_PENGURUS_LAD');

    const wData = wilayahSheet.getDataRange().getValues();
    const wRow = wData.find(r => r[0].toString() === payload.idDesa.toString());
    const kab = wRow ? wRow[2] : "";
    const kec = wRow ? wRow[3] : "";
    const des = wRow ? wRow[4] : "";
    
    const idEntry = "LAD-" + Utilities.getUuid().substring(0,8).toUpperCase();
    
    const skUrl = uploadToDrive(payload.skFile, "SK_LAD_" + idEntry, "SK_LAD");
    const laporanUrl = uploadToDrive(payload.laporanFile, "LAPORAN_LAD_" + idEntry, "LAPORAN_LAD");
    const dokumenUrl = uploadToDrive(payload.dokumenFile, "DOK_LAD_" + idEntry, "DOK_KERJA_LAD");

    /**
     * PEMETAAN KOLOM LAD:
     * 0-35: Metadata & Input Dasar
     * 37: TGL_INPUT
     * 38-43: Kapasitas
     * 44-46: Dokumen Kerja
     * 47-50: Laporan Akhir
     * 51-54: Identitas Desa/Kades
     */
    const rowDataLad = [
      idEntry, payload.idDesa, kab, kec, des,                    // 0-4
      payload.namaLad, payload.wilayahAdat, payload.tahunBerdiri, 
      payload.alamatBalai, payload.sejarahAdat,                   // 5-9
      payload.skNomor, payload.skTahun, skUrl, "-", "-", "-",     // 10-15
      payload.adaBalai, payload.adaRumahAdat, payload.saranaLain, "-", // 16-19
      "-", "-", "-", "-",                                         // 20-23
      payload.sengketaJenis, payload.sengketaTahun, payload.sengketaHasil, "-", // 24-27
      payload.anggaranSumber, payload.anggaranJumlah, "-",        // 28-30
      "-", "-",                                                   // 31-32
      payload.penginputNama, payload.penginputJabatan, payload.penginputHp, // 33-35
      new Date(),                                                // 36 (TGL_INPUT)
      // DATA BARU
      payload.kapasitasAda, payload.kapasitasNama, payload.kapasitasJenis, payload.kapasitasMateri, payload.kapasitasPenyelenggara, payload.kapasitasTahun, // 37-42
      payload.dokumenAda, payload.dokumenJenis, dokumenUrl,      // 43-45
      payload.adaLaporan, payload.laporanTahun, laporanUrl, payload.laporanCatatan, // 46-49
      // IDENTITAS DESA (KOLOM TERAKHIR)
      payload.kadesNama, payload.kadesGender, payload.kadesHp, payload.kadesAlamat // 50-53
    ];
    
    sheet.appendRow(rowDataLad);

    if (payload.pengurus) {
      payload.pengurus.forEach(p => {
        pengurusSheet.appendRow([
          Utilities.getUuid().substring(0,5), idEntry, kab, kec, des, 
          p.nama, p.gender, p.jabatan, p.hp
        ]);
      });
    }

    if (typeof updateLadStatus === 'function') {
      updateLadStatus(idEntry);
    }
    
    return { success: true, message: "Data LAD berhasil disimpan. ID: " + idEntry };
  } catch (e) { 
    return { success: false, message: "Kesalahan Simpan LAD: " + e.toString() }; 
  }
}

/**
 * 6. UTILITY: UPLOAD TO DRIVE
 * Mengkonversi base64 dari form menjadi file rill di Google Drive.
 */
function uploadToDrive(base64Data, fileName, folderName) {
  if (!base64Data || base64Data === "-" || base64Data.length < 50) {
    return "-";
  }
  try {
    const root = DriveApp.getRootFolder();
    let parentFolder = root.getFoldersByName("SIMONI_SULTRA_UPLOADS").hasNext() 
                       ? root.getFoldersByName("SIMONI_SULTRA_UPLOADS").next() 
                       : root.createFolder("SIMONI_SULTRA_UPLOADS");
    
    let targetFolder = parentFolder.getFoldersByName(folderName).hasNext() 
                       ? parentFolder.getFoldersByName(folderName).next() 
                       : parentFolder.createFolder(folderName);
    
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const file = targetFolder.createFile(Utilities.newBlob(bytes, contentType, fileName));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (e) { 
    return "Error Upload: " + e.toString(); 
  }
}

/**
 * 7. AUTHENTICATION & SESSION MANAGEMENT
 */
function loginUser(username, password) {
  const result = authenticateUser(username, password); 
  if (result.success) {
    PropertiesService.getUserProperties().setProperty('current_user', JSON.stringify(result.user));
  }
  return result;
}

function getUserSession() {
  const userJson = PropertiesService.getUserProperties().getProperty('current_user');
  if (!userJson) return null;
  try {
    return JSON.parse(userJson);
  } catch (e) {
    return null;
  }
}

/**
 * LOGOUT: Menghapus sesi dan mengembalikan URL aplikasi dalam satu kali panggilan (Fast Transition).
 */
function logoutUser() {
  PropertiesService.getUserProperties().deleteProperty('current_user');
  return { 
    success: true, 
    url: ScriptApp.getService().getUrl() 
  };
}

/**
 * 8. API BRIDGE & HELPER
 */
function getUrl() { 
  return ScriptApp.getService().getUrl(); 
}

function apiGeneratePdf(id) { 
  return generatePdfLaporan(id); 
}

function apiGetExcelData(type) { 
  return getDownloadData(type); 
}
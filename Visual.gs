/**
 * Modul Visualisasi & Statistik (Versi Strategis & Analitis - Robust)
 * Tanggung Jawab: Menyediakan data rill, sebaran jenis, ranking wilayah, dan statistik Bimtek.
 * Perbaikan: Penanganan nilai null dan sinkronisasi string untuk menghindari data kosong.
 */

function getDashboardStats(user) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lkdSheet = ss.getSheetByName('DATA_LKD_MASTER');
    const ladSheet = ss.getSheetByName('DATA_LAD_MASTER');
    const logSheet = ss.getSheetByName('STATUS_LOG');
    const kapasitasSheet = ss.getSheetByName('DETAIL_KAPASITAS_LKD');
    
    // Normalisasi Filter Kabupaten
    const kabFilter = (user && user.role === "KABUPATEN" && user.namaWilayah) 
                      ? user.namaWilayah.replace("Kab/Kota: ", "").trim().toUpperCase() 
                      : null;

    // --- 1. AMBIL DATA MASTER DENGAN PENANGANAN ERROR ---
    const lkdRows = lkdSheet && lkdSheet.getLastRow() > 1 ? lkdSheet.getDataRange().getValues().slice(1) : [];
    const ladRows = ladSheet && ladSheet.getLastRow() > 1 ? ladSheet.getDataRange().getValues().slice(1) : [];

    // Filter baris LKD berdasarkan wilayah (Case Insensitive & Null Safe)
    const filteredLkd = lkdRows.filter(r => {
      if (!r[0]) return false; // ID_ENTRY harus ada
      if (!kabFilter) return true; // Admin Provinsi lihat semua
      const rowKab = r[2] ? r[2].toString().trim().toUpperCase() : "";
      return rowKab === kabFilter;
    });

    // Filter baris LAD berdasarkan wilayah
    const filteredLad = ladRows.filter(r => {
      if (!r[0]) return false;
      if (!kabFilter) return true;
      const rowKab = r[2] ? r[2].toString().trim().toUpperCase() : "";
      return rowKab === kabFilter;
    });

    // --- 2. ANALISIS SEBARAN JENIS LKD ---
    const lkdDistribution = {};
    filteredLkd.forEach(r => {
      const jenisRaw = r[5] ? r[5].toString().trim() : "Lainnya";
      // Gunakan format Proper Case untuk tampilan grafik agar rapi
      const jenis = jenisRaw.charAt(0).toUpperCase() + jenisRaw.slice(1).toLowerCase();
      lkdDistribution[jenis] = (lkdDistribution[jenis] || 0) + 1;
    });

    // --- 3. ANALISIS CAKUPAN PERDES (SINKRONISASI ID_DESA) ---
    const desaPunyaPerdesLkd = new Set();
    filteredLkd.forEach(r => { 
      const perdesAda = r[13] ? r[13].toString().trim().toUpperCase() : "TIDAK";
      if (perdesAda === "YA" || perdesAda === "ADA") {
        if (r[1]) desaPunyaPerdesLkd.add(r[1].toString().trim()); 
      }
    });

    const desaPunyaPerdesLad = new Set();
    filteredLad.forEach(r => { 
      const fileUrl = r[13] ? r[13].toString().trim() : "";
      if (fileUrl.startsWith("http")) {
        if (r[1]) desaPunyaPerdesLad.add(r[1].toString().trim());
      }
    });

    // --- 4. ANALISIS KAPASITAS (BIMTEK) ---
    let bimtekCount = 0;
    if (kapasitasSheet && kapasitasSheet.getLastRow() > 1) {
      const kapData = kapasitasSheet.getDataRange().getValues().slice(1);
      const pesertaBimtek = new Set();
      
      // Ambil ID LKD yang masuk dalam filter wilayah saat ini
      const validLkdIdsInWilayah = new Set(filteredLkd.map(lkd => lkd[0].toString().trim().toUpperCase()));

      kapData.forEach(r => {
        if (r[1]) {
          const id = r[1].toString().trim().toUpperCase();
          if (!kabFilter || validLkdIdsInWilayah.has(id)) {
            pesertaBimtek.add(id);
          }
        }
      });
      bimtekCount = pesertaBimtek.size;
    }

    const allValidIds = [...filteredLkd.map(r => r[0]), ...filteredLad.map(r => r[0])];
    
    const stats = { 
      total: allValidIds.length, 
      aktif: 0, kurang: 0, tidak: 0,
      perdesLkd: desaPunyaPerdesLkd.size,
      perdesLad: desaPunyaPerdesLad.size,
      lkdDist: lkdDistribution,
      bimtekPct: allValidIds.length > 0 ? Math.round((bimtekCount / allValidIds.length) * 100) : 0
    };

    if (allValidIds.length === 0) return stats;

    // --- 5. SINKRONISASI STATUS LOG (PENCARIAN AMAN) ---
    const logsRaw = logSheet ? logSheet.getDataRange().getValues() : [];
    const logMap = new Map();
    logsRaw.forEach(l => {
      if(l[0]) logMap.set(l[0].toString().trim().toUpperCase(), l[2] ? l[2].toString().trim().toUpperCase() : "");
    });

    allValidIds.forEach(id => {
      const idKey = id ? id.toString().trim().toUpperCase() : "";
      const status = logMap.get(idKey) || "BELUM INPUT";
      if (status === 'AKTIF') stats.aktif++;
      else if (status === 'KURANG AKTIF') stats.kurang++;
      else if (status === 'TIDAK AKTIF') stats.tidak++;
    });
    
    return stats;
  } catch (e) { 
    console.error("Error getDashboardStats: " + e.toString());
    return { total: 0, aktif: 0, kurang: 0, tidak: 0, perdesLkd: 0, perdesLad: 0, lkdDist: {}, bimtekPct: 0 }; 
  }
}

function getRegionalAggregateData(user) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const wilayahSheet = ss.getSheetByName('WILAYAH');
    const lkdSheet = ss.getSheetByName('DATA_LKD_MASTER');
    const ladSheet = ss.getSheetByName('DATA_LAD_MASTER');
    const logSheet = ss.getSheetByName('STATUS_LOG');

    const kabFilter = (user && user.role === "KABUPATEN" && user.namaWilayah) 
                      ? user.namaWilayah.replace("Kab/Kota: ", "").trim().toUpperCase() 
                      : null;

    const result = { progress: [], lkdAgregat: [], ladAgregat: [], ranking: { top: [], bottom: [] } };
    if (!wilayahSheet) return result;

    const lkdData = lkdSheet && lkdSheet.getLastRow() > 1 ? lkdSheet.getDataRange().getValues().slice(1).filter(r => r[0]) : [];
    const ladData = ladSheet && ladSheet.getLastRow() > 1 ? ladSheet.getDataRange().getValues().slice(1).filter(r => r[0]) : [];
    const logsRaw = logSheet ? logSheet.getDataRange().getValues() : [];
    
    const logMap = new Map();
    logsRaw.forEach(l => { if(l[0]) logMap.set(l[0].toString().trim().toUpperCase(), l[2] ? l[2].toString().trim().toUpperCase() : ""); });

    // 1. PROSES REKAPITULASI STATUS (FILTER KABUPATEN)
    const processRekap = (dataArr) => {
      const map = {};
      dataArr.forEach(r => {
        const rowKab = r[2] ? r[2].toString().trim().toUpperCase() : "";
        if (kabFilter && rowKab !== kabFilter) return;

        const displayKab = r[2] ? r[2].toString().trim() : "Tidak Diketahui";
        if (!map[displayKab]) map[displayKab] = { name: displayKab, total: 0, aktif: 0, kurang: 0, tidak: 0 };
        
        const id = r[0].toString().trim().toUpperCase();
        const status = logMap.get(id) || "BELUM INPUT";
        
        map[displayKab].total++;
        if (status === 'AKTIF') map[displayKab].aktif++;
        else if (status === 'KURANG AKTIF') map[displayKab].kurang++;
        else if (status === 'TIDAK AKTIF') map[displayKab].tidak++;
      });
      return Object.values(map);
    };

    result.lkdAgregat = processRekap(lkdData);
    result.ladAgregat = processRekap(ladData);

    // 2. PROSES PROGRES DAN RANKING
    const wData = wilayahSheet.getDataRange().getValues().slice(1).filter(r => {
      if (!kabFilter) return true;
      const rowKab = r[2] ? r[2].toString().trim().toUpperCase() : "";
      return rowKab === kabFilter;
    });

    const targetPerKab = {};
    wData.forEach(r => { 
      const kab = r[2] ? r[2].toString().trim() : "";
      if(kab) targetPerKab[kab] = (targetPerKab[kab] || 0) + 1; 
    });

    const perdesLkdPerKab = {};
    const perdesLadPerKab = {};
    const uniqueDesaPerdesLkd = new Set();
    const uniqueDesaPerdesLad = new Set();

    lkdData.forEach(r => {
      const kab = r[2] ? r[2].toString().trim() : "";
      const perdesAda = r[13] ? r[13].toString().trim().toUpperCase() : "";
      if ((perdesAda === "YA" || perdesAda === "ADA") && !uniqueDesaPerdesLkd.has(r[1])) {
        uniqueDesaPerdesLkd.add(r[1]);
        perdesLkdPerKab[kab] = (perdesLkdPerKab[kab] || 0) + 1;
      }
    });

    ladData.forEach(r => {
      const kab = r[2] ? r[2].toString().trim() : "";
      const fileUrl = r[13] ? r[13].toString().trim() : "";
      if (fileUrl.startsWith("http") && !uniqueDesaPerdesLad.has(r[1])) {
        uniqueDesaPerdesLad.add(r[1]);
        perdesLadPerKab[kab] = (perdesLadPerKab[kab] || 0) + 1;
      }
    });

    const realisasiPerKab = {};
    const uniqueDesaInput = new Set();
    [...lkdData, ...ladData].forEach(r => {
      if(r[1] && !uniqueDesaInput.has(r[1].toString().trim())) {
        uniqueDesaInput.add(r[1].toString().trim());
        const kab = r[2] ? r[2].toString().trim() : "";
        if(kab) realisasiPerKab[kab] = (realisasiPerKab[kab] || 0) + 1;
      }
    });

    const progressList = Object.keys(targetPerKab).map(kab => {
      const jml = targetPerKab[kab] || 0;
      const input = realisasiPerKab[kab] || 0;
      const pct = Math.round((input / jml) * 100) || 0;
      return { 
        name: kab, jmlDes: jml, inputDes: input, pct: pct,
        perdesLkd: perdesLkdPerKab[kab] || 0, perdesLad: perdesLadPerKab[kab] || 0
      };
    });

    result.progress = progressList.sort((a,b) => a.name.localeCompare(b.name));

    if (!kabFilter) { 
      const sortedByProgress = [...progressList].sort((a, b) => b.pct - a.pct);
      result.ranking.top = sortedByProgress.slice(0, 5);
      result.ranking.bottom = sortedByProgress.slice(-5).reverse();
    }

    return result;
  } catch (e) {
    console.error("Error getRegionalAggregateData: " + e.toString());
    return { progress: [], lkdAgregat: [], ladAgregat: [], ranking: { top: [], bottom: [] } };
  }
}

function recalculateAllStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lkdSheet = ss.getSheetByName('DATA_LKD_MASTER');
  const ladSheet = ss.getSheetByName('DATA_LAD_MASTER');
  let count = 0;
  if (lkdSheet) lkdSheet.getDataRange().getValues().slice(1).forEach(row => { if (row[0]) { updateLkdStatus(row[0]); count++; } });
  if (ladSheet) ladSheet.getDataRange().getValues().slice(1).forEach(row => { if (row[0]) { updateLadStatus(row[0]); count++; } });
  return "Berhasil memproses: " + count + " data.";
}
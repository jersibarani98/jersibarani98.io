/**
 * Modul Export & Pelaporan SIMONI LKD-LAD (Versi Lengkap: Sarana & Partisipasi)
 * Tanggung Jawab: Mengambil data dari database dan memetakan ke Template PDF.
 */

function generatePdfLaporan(idEntry) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lkdMaster = ss.getSheetByName('DATA_LKD_MASTER');
    const ladMaster = ss.getSheetByName('DATA_LAD_MASTER');
    const logSheet = ss.getSheetByName('STATUS_LOG');
    const pengurusLkdSheet = ss.getSheetByName('DETAIL_PENGURUS_LKD');
    const pengurusLadSheet = ss.getSheetByName('DETAIL_PENGURUS_LAD');

    let d = {}; 
    let listPengurus = [];
    let idStr = idEntry ? idEntry.toString().trim().toUpperCase() : "";
    let jenisLembaga = idStr.startsWith("LKD") ? "LKD" : "LAD";
    let masterSheet = (jenisLembaga === "LKD") ? lkdMaster : ladMaster;

    if (!masterSheet) throw new Error("Tabel master database tidak ditemukan.");

    const fullData = masterSheet.getDataRange().getValues();
    const headers = fullData[0];
    const r = fullData.find(row => row[0] && row[0].toString().trim().toUpperCase() === idStr);
    
    if (!r) throw new Error("Data " + jenisLembaga + " tidak ditemukan.");

    const getVal = (headerName) => {
      const idx = headers.indexOf(headerName);
      return (idx === -1) ? "" : (r[idx] === undefined || r[idx] === null ? "" : r[idx]);
    };

    if (jenisLembaga === "LKD") {
      d = {
        kab: getVal("KABUPATEN"), kec: getVal("KECAMATAN"), des: getVal("DESA"), 
        jenis: getVal("JENIS_LKD"), nama: getVal("NAMA_LEMBAGA"), 
        sk_no: getVal("SK_NOMOR"), sk_thn: getVal("SK_TAHUN"), 
        perdes: getVal("PERDES_ADA"), perdes_no: getVal("PERDES_NOMOR"),
        // Sarana & Partisipasi
        sekre: getVal("ADA_SEKRETARIAT"), 
        sarana_kerja: getVal("SARANA_KERJA"),
        musdes: getVal("IKUT_MUSDES"), 
        musren: getVal("IKUT_MUSRENBANG"),
        rpjm: getVal("IKUT_RPJM"),
        rkp: getVal("IKUT_RKP"),
        // Anggaran & Laporan
        angg_ada: getVal("ANGGARAN_ADA"), angg_sbr: getVal("ANGGARAN_SUMBER"), 
        angg_jml: Number(getVal("ANGGARAN_JUMLAH")) || 0, angg_utk: getVal("ANGGARAN_UNTUK"),
        kap_ada: getVal("KAPASITAS_ADA"), kap_nama: getVal("KAPASITAS_NAMA"), kap_jns: getVal("KAPASITAS_JENIS"), 
        kap_mtr: getVal("KAPASITAS_MATERI"), kap_pny: getVal("KAPASITAS_PENYELENGGARA"), kap_thn: getVal("KAPASITAS_TAHUN"),
        dok_ada: getVal("DOK_KERJA_ADA"), dok_jns: getVal("DOK_KERJA_JENIS"),
        lap_akhir_ada: getVal("ADA_LAPORAN_AKHIR"), lap_akhir_thn: getVal("LAPORAN_AKHIR_TAHUN"),
        // Identitas Desa
        kades_nama: getVal("KADES_NAMA"), kades_gen: getVal("KADES_GENDER"), 
        kades_hp: getVal("KADES_HP"), kades_almt: getVal("KANTOR_DESA_ALAMAT")
      };
      if (pengurusLkdSheet) {
        listPengurus = pengurusLkdSheet.getDataRange().getValues()
          .filter(row => row[1] && row[1].toString().trim().toUpperCase() === idStr)
          .map(row => ({ nama: row[5], gen: row[6], jab: row[7], hp: row[10] }));
      }
    } else {
      // Mapping LAD (Sesuai kebutuhan)
      d = {
        kab: getVal("KABUPATEN"), kec: getVal("KECAMATAN"), des: getVal("DESA"), 
        nama: getVal("NAMA_LAD"), sk_no: getVal("SK_NOMOR"), sk_thn: getVal("SK_TAHUN"),
        balai: getVal("ADA_BALAI"), rumah_adat: getVal("ADA_RUMAH_ADAT"), sarana_lain: getVal("SARANA_LAIN"),
        angg_jml: Number(getVal("ANGGARAN_JUMLAH")) || 0,
        kap_ada: getVal("KAPASITAS_ADA"), kap_nama: getVal("KAPASITAS_NAMA"), 
        kap_jns: getVal("KAPASITAS_JENIS"), kap_mtr: getVal("KAPASITAS_MATERI"),
        kap_pny: getVal("KAPASITAS_PENYELENGGARA"), kap_thn: getVal("KAPASITAS_TAHUN"),
        lap_akhir_ada: getVal("ADA_LAPORAN_AKHIR"), kades_nama: getVal("KADES_NAMA"), 
        kades_almt: getVal("KANTOR_DESA_ALAMAT"), kades_gen: getVal("KADES_GENDER"), kades_hp: getVal("KADES_HP")
      };
      if (pengurusLadSheet) {
        listPengurus = pengurusLadSheet.getDataRange().getValues()
          .filter(row => row[1] && row[1].toString().trim().toUpperCase() === idStr)
          .map(row => ({ nama: row[5], gen: row[6], jab: row[7], hp: row[8] }));
      }
    }

    const logFound = logSheet ? logSheet.getDataRange().getValues().slice().reverse().find(l => l[0] && l[0].toString().trim().toUpperCase() === idStr) : null;
    const logInfo = logFound || [idStr, 0, "BELUM DINILAI"];

    const template = HtmlService.createTemplateFromFile(jenisLembaga === "LKD" ? 'ReportTemplate' : 'ReportTemplateLAD');
    template.d = d;
    template.pengurus = listPengurus;
    template.status_monev = logInfo[2];
    template.skor_total = logInfo[1];
    template.tgl_cetak = Utilities.formatDate(new Date(), "GMT+7", "dd MMMM yyyy");

    const htmlOutput = template.evaluate().getContent();
    const pdf = Utilities.newBlob(htmlOutput, 'text/html', 'temp.html').getAs('application/pdf').setName("LAPORAN_" + d.nama + ".pdf");
    return { success: true, base64: Utilities.base64Encode(pdf.getBytes()), filename: pdf.getName() };
  } catch (err) { return { success: false, message: "Gagal Cetak: " + err.toString() }; }
}
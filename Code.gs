// ================================================================
// GOOGLE APPS SCRIPT — Backend untuk index.html
// Bidang Transmigrasi Kabupaten Sijunjung
//
// CARA PENGGUNAAN:
//   1. Buka script.google.com → buat proyek baru
//   2. Tempel seluruh kode ini → ganti SPREADSHEET_ID dengan ID spreadsheet Anda
//   3. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone
//   4. Salin URL deployment → tempel ke APPS_SCRIPT_URL di index.html
// ================================================================

// ================================================================
// KONFIGURASI — Ganti dengan ID Google Spreadsheet Anda
// Format: https://docs.google.com/spreadsheets/d/<ID_INI>/edit
// ================================================================
const SPREADSHEET_ID = "GANTI_DENGAN_ID_SPREADSHEET_ANDA";

// Nama sheet (tab) di dalam spreadsheet
const SHEET_KK      = "DataKK";       // Sheet untuk data Kepala Keluarga
const SHEET_ANGGOTA = "DataAnggota";  // Sheet untuk data Anggota

// ================================================================
// CORS Helper — wajib agar index.html (file lokal) bisa akses API
// ================================================================
function setCorsHeaders(output) {
  return output
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}

// Tangani preflight OPTIONS request
function doOptions(e) {
  return setCorsHeaders(
    ContentService.createTextOutput("")
      .setMimeType(ContentService.MimeType.TEXT)
  );
}

// ================================================================
// ENTRY POINT — menerima semua request POST dari index.html
// ================================================================
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    if      (action === "get_all")       result = actionGetAll();
    else if (action === "simpan_kk")     result = actionSimpanKK(body.data);
    else if (action === "simpan_anggota")result = actionSimpanAnggota(body.data);
    else if (action === "update_kk")     result = actionUpdateKK(body.data);
    else if (action === "hapus_kk")      result = actionHapusKK(body.nokk);
    else if (action === "cari")          result = actionCari(body.query);
    else result = { success: false, message: "Action tidak dikenal: " + action };

    return setCorsHeaders(
      ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON)
    );

  } catch (err) {
    return setCorsHeaders(
      ContentService.createTextOutput(
        JSON.stringify({ success: false, message: "Error: " + err.message })
      ).setMimeType(ContentService.MimeType.JSON)
    );
  }
}

// Juga tangani GET (untuk test di browser)
function doGet(e) {
  const result = actionGetAll();
  return setCorsHeaders(
    ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON)
  );
}

// ================================================================
// HELPER — Pastikan sheet ada, buat jika belum ada
// ================================================================
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    // Style header
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#c9974a");
    headerRange.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Header kolom untuk masing-masing sheet
const HEADERS_KK = [
  "ID", "No. KK", "Nama KK", "NIK KK", "Jenis Kelamin",
  "Pendidikan", "Blok/Wilayah", "Link Berkas KK (GDrive)",
  "Link KTP (GDrive)", "Timestamp"
];

const HEADERS_ANGGOTA = [
  "ID", "No. KK", "Nama Anggota", "NIK Anggota",
  "Hubungan", "Jenis Kelamin", "Pendidikan", "Timestamp"
];

// ================================================================
// ACTION: get_all — ambil semua data KK, Anggota, dan statistik
// ================================================================
function actionGetAll() {
  const ss          = getSpreadsheet();
  const sheetKK     = getOrCreateSheet(ss, SHEET_KK, HEADERS_KK);
  const sheetAng    = getOrCreateSheet(ss, SHEET_ANGGOTA, HEADERS_ANGGOTA);

  const rawKK  = sheetKK.getDataRange().getValues();
  const rawAng = sheetAng.getDataRange().getValues();

  // Baris pertama adalah header, skip
  const kkData = rawKK.slice(1).map(r => ({
    id:              String(r[0]),
    nokk:            String(r[1]),
    nama_kk:         String(r[2]),
    nik_kk:          String(r[3]),
    jenis_kelamin:   String(r[4]),
    pendidikan:      String(r[5]),
    wilayah:         String(r[6]),
    gdrive_link:     String(r[7]),
    gdrive_ktp_link: String(r[8]),
    has_berkas:      r[7] ? "Ya" : "Tidak"
  })).filter(k => k.nokk && k.nokk !== "undefined" && k.nokk !== "");

  const angData = rawAng.slice(1).map(r => ({
    id:               String(r[0]),
    nokk:             String(r[1]),
    nama_anggota:     String(r[2]),
    nik_anggota:      String(r[3]),
    hubungan:         String(r[4]),
    jenis_kelamin:    String(r[5]),
    jenis_kelamin_label: r[5] === "L" ? "Laki-laki" : (r[5] === "P" ? "Perempuan" : "-"),
    pendidikan:       String(r[6])
  })).filter(a => a.nokk && a.nokk !== "undefined" && a.nokk !== "");

  // Hitung statistik
  const totalKK       = kkData.length;
  const totalAnggota  = angData.length;
  const totalWarga    = totalKK + totalAnggota;
  const kkDenganBerkas = kkData.filter(k => k.has_berkas === "Ya").length;

  return {
    success:  true,
    kk:       kkData,
    anggota:  angData,
    stats: {
      total_kk:          totalKK,
      total_anggota:     totalAnggota,
      total_warga:       totalWarga,
      kk_dengan_berkas:  kkDenganBerkas
    }
  };
}

// ================================================================
// ACTION: simpan_kk — tambah baris baru ke sheet DataKK
// ================================================================
function actionSimpanKK(data) {
  if (!data || !data.nokk || !data.nama_kk) {
    return { success: false, message: "Data tidak lengkap (nokk & nama_kk wajib)" };
  }

  const ss      = getSpreadsheet();
  const sheet   = getOrCreateSheet(ss, SHEET_KK, HEADERS_KK);
  const allData = sheet.getDataRange().getValues();

  // Cek duplikat No. KK (kolom ke-2, index 1)
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][1]) === String(data.nokk)) {
      return { success: false, message: "No. KK " + data.nokk + " sudah terdaftar!" };
    }
  }

  const id        = "KK" + Date.now();
  const timestamp = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });

  sheet.appendRow([
    id,
    data.nokk,
    data.nama_kk,
    data.nik_kk          || "",
    data.jenis_kelamin   || "",
    data.pendidikan      || "",
    data.wilayah         || "",
    data.gdrive_link     || "",
    data.gdrive_ktp_link || "",
    timestamp
  ]);

  return { success: true, message: "Data KK berhasil disimpan", id: id };
}

// ================================================================
// ACTION: simpan_anggota — tambah baris baru ke sheet DataAnggota
// ================================================================
function actionSimpanAnggota(data) {
  if (!data || !data.nokk || !data.nama_anggota) {
    return { success: false, message: "Data tidak lengkap (nokk & nama_anggota wajib)" };
  }

  // Validasi: pastikan KK induk ada
  const ss      = getSpreadsheet();
  const sheetKK = getOrCreateSheet(ss, SHEET_KK, HEADERS_KK);
  const kkRows  = sheetKK.getDataRange().getValues().slice(1);
  const kkExists = kkRows.some(r => String(r[1]) === String(data.nokk));
  if (!kkExists) {
    return { success: false, message: "No. KK " + data.nokk + " tidak ditemukan di data KK" };
  }

  const sheetAng = getOrCreateSheet(ss, SHEET_ANGGOTA, HEADERS_ANGGOTA);

  // Cek duplikat NIK anggota (kolom ke-4, index 3)
  const existingAng = sheetAng.getDataRange().getValues().slice(1);
  if (data.nik_anggota) {
    for (let i = 0; i < existingAng.length; i++) {
      if (String(existingAng[i][3]) === String(data.nik_anggota)) {
        return { success: false, message: "NIK " + data.nik_anggota + " sudah terdaftar sebagai anggota!" };
      }
    }
  }

  const id        = "ANG" + Date.now();
  const timestamp = new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" });

  sheetAng.appendRow([
    id,
    data.nokk,
    data.nama_anggota,
    data.nik_anggota    || "",
    data.hubungan       || "",
    data.jenis_kelamin  || "",
    data.pendidikan     || "",
    timestamp
  ]);

  return { success: true, message: "Anggota berhasil ditambahkan", id: id };
}

// ================================================================
// ACTION: update_kk — perbarui data KK berdasarkan ID
// ================================================================
function actionUpdateKK(data) {
  if (!data || !data.id) {
    return { success: false, message: "ID tidak ditemukan" };
  }

  const ss    = getSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_KK, HEADERS_KK);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      const rowNum = i + 1; // +1 karena getValues() 0-indexed, sheet 1-indexed
      sheet.getRange(rowNum, 3).setValue(data.nama_kk         || rows[i][2]);
      sheet.getRange(rowNum, 4).setValue(data.nik_kk          || rows[i][3]);
      sheet.getRange(rowNum, 5).setValue(data.jenis_kelamin   || rows[i][4]);
      sheet.getRange(rowNum, 6).setValue(data.pendidikan      || rows[i][5]);
      sheet.getRange(rowNum, 7).setValue(data.wilayah         !== undefined ? data.wilayah : rows[i][6]);
      sheet.getRange(rowNum, 8).setValue(data.gdrive_link     !== undefined ? data.gdrive_link : rows[i][7]);
      sheet.getRange(rowNum, 9).setValue(data.gdrive_ktp_link !== undefined ? data.gdrive_ktp_link : rows[i][8]);
      return { success: true, message: "Data KK berhasil diperbarui" };
    }
  }

  return { success: false, message: "Data dengan ID " + data.id + " tidak ditemukan" };
}

// ================================================================
// ACTION: hapus_kk — hapus baris KK + semua anggotanya
// ================================================================
function actionHapusKK(nokk) {
  if (!nokk) return { success: false, message: "No. KK tidak diberikan" };

  const ss      = getSpreadsheet();
  const sheetKK = getOrCreateSheet(ss, SHEET_KK, HEADERS_KK);
  const sheetAng = getOrCreateSheet(ss, SHEET_ANGGOTA, HEADERS_ANGGOTA);

  // Hapus dari DataKK
  const kkRows = sheetKK.getDataRange().getValues();
  let kkDeleted = false;
  // Loop dari bawah agar row index tidak bergeser
  for (let i = kkRows.length - 1; i >= 1; i--) {
    if (String(kkRows[i][1]) === String(nokk)) {
      sheetKK.deleteRow(i + 1);
      kkDeleted = true;
    }
  }

  if (!kkDeleted) {
    return { success: false, message: "No. KK " + nokk + " tidak ditemukan" };
  }

  // Hapus semua anggota dari DataAnggota
  const angRows = sheetAng.getDataRange().getValues();
  let deletedCount = 0;
  for (let i = angRows.length - 1; i >= 1; i--) {
    if (String(angRows[i][1]) === String(nokk)) {
      sheetAng.deleteRow(i + 1);
      deletedCount++;
    }
  }

  return {
    success:  true,
    message:  "KK dan " + deletedCount + " anggota berhasil dihapus"
  };
}

// ================================================================
// ACTION: cari — cari KK berdasarkan nama / No. KK / NIK
// ================================================================
function actionCari(query) {
  if (!query || query.length < 2) {
    return { success: false, message: "Query pencarian terlalu pendek" };
  }

  const allData = actionGetAll();
  if (!allData.success) return allData;

  const q = query.toLowerCase();

  const matched = allData.kk.filter(k =>
    k.nama_kk.toLowerCase().includes(q) ||
    k.nokk.includes(q)                  ||
    k.nik_kk.includes(q)                ||
    (k.wilayah && k.wilayah.toLowerCase().includes(q))
  );

  // Sertakan data anggota masing-masing hasil pencarian
  const results = matched.map(kk => ({
    ...kk,
    anggota: allData.anggota.filter(a => a.nokk === kk.nokk)
  }));

  return { success: true, results: results };
}

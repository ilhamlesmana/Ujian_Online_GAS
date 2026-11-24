/**
 * ============================================================================
 * SISTEM UJIAN ONLINE SEKOLAH DIGITAL V2.5 (DENGAN ROLE ADMIN)
 * Fitur Lengkap:
 * 1. Role Administrator: Kelola Data Guru (DATA_GURU) & Siswa (DATA_SISWA)
 * 2. Relasi Mapel Per Tingkat Kelas (Kelas X, XI, XII) & Isolasi Guru
 * 3. Siswa Otomatis Hanya Melihat Ujian Sesuai Tingkat Kelasnya
 * 4. Pengaturan Profil & Ubah Password Mandiri untuk Siswa (DATA_SISWA)
 * 5. Batch Input Banyak Soal Sekaligus, Anti-Cheat Fullscreen, & Tampilkan Nilai
 * ============================================================================
 */

const SPREADSHEET_ID = "GANTI_DENGAN_ID_SPREADSHEET_ANDA";
const EMAIL_GURU_NOTIFIKASI = "GANTI_DENGAN_EMAIL_NOTIFIKASI@gmail.com";

const SHEET_NAMES = {
  ADMIN: "DATA_ADMIN",
  MASTER_MAPEL: "MASTER_MAPEL",
  MASTER_JENIS: "MASTER_JENIS_PENILAIAN",
  KELAS: "DATA_KELAS",
  GURU: "DATA_GURU",
  SISWA: "DATA_SISWA",
  SOAL: "BANK_SOAL",
  PENGATURAN: "PENGATURAN_UJIAN",
  HASIL: "HASIL_UJIAN",
  WEB_SETTINGS: "PENGATURAN_WEBSITE",
  PEMBELI: "DATA_PEMBELI"
};

let _cachedSs = null;
let _cachedMasterSs = null;
let _activeTargetSpreadsheetId = "";

/**
 * Mengekstrak ID Spreadsheet murni apabila pembeli memasukkan URL Google Sheet lengkap
 * maupun ID Spreadsheet langsung.
 */
function extractCleanSpreadsheetId_(rawInput) {
  const str = String(rawInput || "").trim();
  if (!str || str === "GANTI_DENGAN_ID_SPREADSHEET_ANDA") return "";
  const urlMatch = str.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/i);
  if (urlMatch && urlMatch[1]) return urlMatch[1];
  return str;
}

/**
 * Membuka Master Spreadsheet milik Penjual (Owner) tempat sheet DATA_PEMBELI berada.
 */
function getMasterOwnerSpreadsheet_() {
  if (_cachedMasterSs) return _cachedMasterSs;
  let ss = null;
  try {
    const cleanMasterId = extractCleanSpreadsheetId_(SPREADSHEET_ID);
    if (cleanMasterId) {
      ss = SpreadsheetApp.openById(cleanMasterId);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
  } catch (err) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) {
    throw new Error("Master Spreadsheet Penjual tidak ditemukan.");
  }
  _cachedMasterSs = ss;
  return ss;
}

/**
 * Memastikan sheet DATA_PEMBELI tersedia di Master Spreadsheet milik Anda (Penjual)
 * untuk mengontrol status AKTIF / NONAKTIF setiap Spreadsheet ID pembeli.
 */
function ensureDataPembeliSheet_(masterSs) {
  const ss = masterSs || getMasterOwnerSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAMES.PEMBELI);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.PEMBELI);
    const headers = ["spreadsheet_id", "nama_pembeli_sekolah", "status", "tanggal_daftar", "catatan"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#fef08a")
      .setFontColor("#854d0e");
    sheet.appendRow([
      ss.getId(),
      "Master Server Utama (Milik Penjual)",
      "AKTIF",
      Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd"),
      "Jangan dihapus - ID Spreadsheet Master"
    ]);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

/**
 * Memeriksa apakah Spreadsheet ID pembeli terdaftar dan berstatus AKTIF di sheet DATA_PEMBELI.
 */
function verifyBuyerSpreadsheetAccess_(rawBuyerSpreadsheetId) {
  const masterSs = getMasterOwnerSpreadsheet_();
  const masterId = masterSs.getId();
  const cleanTargetId = extractCleanSpreadsheetId_(rawBuyerSpreadsheetId);

  // Jika tidak mengirim spreadsheetId khusus, atau sama dengan Master Spreadsheet Anda sendiri -> Izinkan
  if (!cleanTargetId || cleanTargetId === masterId) {
    return { allowed: true, targetId: masterId, isMaster: true };
  }

  const sheetPembeli = ensureDataPembeliSheet_(masterSs);
  const lastRow = sheetPembeli.getLastRow();
  if (lastRow <= 1) {
    return {
      allowed: false,
      message: "ID Spreadsheet (" + cleanTargetId + ") belum terdaftar di sistem. Silakan hubungi Developer untuk aktivasi."
    };
  }

  const rows = sheetPembeli.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let i = 0; i < rows.length; i++) {
    const rowSsId = extractCleanSpreadsheetId_(rows[i][0]);
    const namaSekolah = String(rows[i][1] || "Pembeli").trim();
    const status = String(rows[i][2] || "AKTIF").trim().toUpperCase();
    const catatan = String(rows[i][4] || "").trim();

    if (rowSsId && rowSsId === cleanTargetId) {
      if (status === "AKTIF" || status === "ACTIVE" || status === "YA" || status === "OPEN" || status === "1") {
        return { allowed: true, targetId: cleanTargetId, isMaster: false, namaSekolah: namaSekolah };
      } else {
        return {
          allowed: false,
          message:
            "Akses Aplikasi untuk " +
            namaSekolah +
            " saat ini sedang DINONAKTIFKAN (Status: " +
            status +
            ")." +
            (catatan ? " Catatan: " + catatan : " Silakan hubungi Administrator/Developer.")
        };
      }
    }
  }

  return {
    allowed: false,
    message:
      "ID Spreadsheet (" +
      cleanTargetId +
      ") belum diaktifkan di server pusat. Pastikan Anda telah memberikan ID Spreadsheet ini kepada Developer dan membagikan akses Editor pada Google Sheet Anda."
  };
}

/**
 * ============================================================================
 * GENERATOR 9 SHEET BERELASI (JALANKAN SEKALI DARI EDITOR APPS SCRIPT)
 * ============================================================================
 */
function setupDatabaseDanDataAwal() {
  const ss = getSpreadsheet_(true);

  const seedData = {
    "DATA_ADMIN": [
      ["id_admin", "nama_lengkap", "username", "password"],
      ["ADM-01", "Administrator", "admin", "admin123"]
    ],
    "MASTER_MAPEL": [
      ["kode_mapel", "nama_mapel", "kelas", "label_mapel_kelas", "kelompok"],
      ["MP-1-MTK", "Matematika", "1", "Matematika (Kelas 1)", "Umum / Wajib"],
      ["MP-1-BIN", "Bahasa Indonesia", "1", "Bahasa Indonesia (Kelas 1)", "Umum / Wajib"],
      ["MP-4-IPA", "IPAS / Sains", "4", "IPAS / Sains (Kelas 4)", "Umum / Wajib"],
      ["MP-6-MTK", "Matematika", "6", "Matematika (Kelas 6)", "Umum / Wajib"],
      ["MP-7-MTK", "Matematika", "7", "Matematika (Kelas 7)", "Umum / Wajib"],
      ["MP-7-IPA", "Ilmu Pengetahuan Alam (IPA)", "7", "Ilmu Pengetahuan Alam (IPA) (Kelas 7)", "Umum / Wajib"],
      ["MP-8-BIG", "Bahasa Inggris", "8", "Bahasa Inggris (Kelas 8)", "Umum / Wajib"],
      ["MP-9-BIN", "Bahasa Indonesia", "9", "Bahasa Indonesia (Kelas 9)", "Umum / Wajib"],
      ["MP-10-MTK", "Matematika", "10", "Matematika (Kelas 10)", "Umum / Wajib"],
      ["MP-10-FIS", "Fisika", "10", "Fisika (Kelas 10)", "Sains / Pilihan"],
      ["MP-10-BIN", "Bahasa Indonesia", "10", "Bahasa Indonesia (Kelas 10)", "Umum / Wajib"],
      ["MP-10-BIG", "Bahasa Inggris", "10", "Bahasa Inggris (Kelas 10)", "Umum / Wajib"],
      ["MP-11-MTK", "Matematika", "11", "Matematika (Kelas 11)", "Umum / Wajib"],
      ["MP-11-FIS", "Fisika", "11", "Fisika (Kelas 11)", "Sains / Pilihan"],
      ["MP-11-BIN", "Bahasa Indonesia", "11", "Bahasa Indonesia (Kelas 11)", "Umum / Wajib"],
      ["MP-12-MTK", "Matematika", "12", "Matematika (Kelas 12)", "Umum / Wajib"],
      ["MP-12-INF", "Informatika", "12", "Informatika (Kelas 12)", "Teknologi / Pilihan"]
    ],
    "MASTER_JENIS_PENILAIAN": [
      ["kode_jenis", "nama_jenis", "keterangan", "pemilik_guru", "mapel_relasi"],
      ["JP-01", "Ulangan Harian 1", "Penilaian Harian Matematika Bab 1", "guru_mtk_x", "Matematika (Kelas 10)"],
      ["JP-02", "PTS Ganjil", "Penilaian Tengah Semester Ganjil", "guru_mtk_x", "Matematika (Kelas 10)"],
      ["JP-03", "Ulangan Harian 1", "Penilaian Harian Fisika Bab 1", "guru_mtk_x", "Fisika (Kelas 10)"],
      ["JP-04", "Ulangan Harian 1", "Penilaian Harian Matematika Kelas 11", "guru_mtk_xi", "Matematika (Kelas 11)"],
      ["JP-05", "Ujian Sekolah", "Ujian Kelulusan Matematika Kelas 12", "guru_mtk_xi", "Matematika (Kelas 12)"],
      ["JP-06", "Ulangan Harian 1", "Penilaian Harian Bahasa Indonesia", "guru_bahasa", "Bahasa Indonesia (Kelas 10)"],
      ["JP-07", "Kuis / Latihan", "Evaluasi Mingguan Bahasa Inggris", "guru_bahasa", "Bahasa Inggris (Kelas 10)"],
      ["JP-08", "PAS Ganjil", "Penilaian Akhir Semester Informatika", "guru_tik", "Informatika (Kelas 12)"]
    ],
    "DATA_KELAS": [
      ["kode_kelas", "nama_kelas", "tingkat_kelas", "wali_kelas"],
      ["KLS-01", "1-A", "1", "Siti Aminah, M.Pd."],
      ["KLS-02", "7-A", "7", "Budi Santoso, S.Pd."],
      ["KLS-03", "10-A", "10", "Budi Santoso, S.Pd."],
      ["KLS-04", "10-B", "10", "Siti Aminah, M.Pd."],
      ["KLS-05", "11-A", "11", "Ratna Sari, M.Pd."],
      ["KLS-06", "12-A", "12", "Andi Pratama, S.Kom."]
    ],
    "DATA_GURU": [
      ["nip_atau_id", "nama_lengkap", "mata_pelajaran", "username", "password"],
      ["198501012010011001", "Budi Santoso, S.Pd.", "Matematika (Kelas X), Fisika (Kelas X)", "guru_mtk_x", "guru123"],
      ["198703152012012002", "Ratna Sari, M.Pd.", "Matematika (Kelas XI), Matematika (Kelas XII)", "guru_mtk_xi", "guru123"],
      ["198805122014022003", "Siti Aminah, M.Pd.", "Bahasa Indonesia (Kelas X), Bahasa Inggris (Kelas X)", "guru_bahasa", "guru123"],
      ["199011202018031004", "Andi Pratama, S.Kom.", "Informatika (Kelas XII)", "guru_tik", "guru123"]
    ],
    "DATA_SISWA": [
      ["nis", "nama_lengkap", "kelas", "username", "password"],
      ["2026001", "Ahmad Fauzi", "X-IPA-1", "siswa_x1", "siswa123"],
      ["2026002", "Dewi Lestari", "X-IPA-2", "siswa_x2", "siswa123"],
      ["2026003", "Rizky Ramadhan", "XI-IPA-1", "siswa_xi1", "siswa123"],
      ["2026004", "Nabila Putri", "XII-IPA-1", "siswa_xii1", "siswa123"]
    ],
    "BANK_SOAL": [
      ["id_soal", "mata_pelajaran", "pertanyaan", "id_gambar", "id_suara", "opsi_a", "opsi_b", "opsi_c", "opsi_d", "opsi_e", "kunci_jawaban"],
      ["SOAL-1001", "Matematika (Kelas X) - Ulangan Harian 1", "Hasil dari 15 + 25 x 2 adalah ...", "", "", "55", "65", "80", "70", "60", "B"],
      ["SOAL-1002", "Matematika (Kelas X) - Ulangan Harian 1", "Jika 3x - 5 = 16, maka nilai x yang memenuhi adalah ...", "", "", "5", "6", "7", "8", "9", "C"],
      ["SOAL-1003", "Matematika (Kelas X) - Ulangan Harian 1", "Nilai dari 2 pangkat 5 (2^5) adalah ...", "", "", "10", "16", "25", "32", "64", "D"],
      ["SOAL-1004", "Matematika (Kelas X) - PTS Ganjil", "Himpunan penyelesaian dari |2x - 1| = 5 adalah ...", "", "", "{-2, 3}", "{2, -3}", "{-2, -3}", "{1, 3}", "{0, 3}", "A"],
      ["SOAL-1005", "Fisika (Kelas X) - Ulangan Harian 1", "Satuan Internasional (SI) untuk besaran kuat arus listrik adalah ...", "", "", "Volt", "Watt", "Ohm", "Ampere", "Joule", "D"],
      ["SOAL-1006", "Fisika (Kelas X) - Ulangan Harian 1", "Mobil bergerak lurus dengan kecepatan tetap 20 m/s selama 10 sekon. Jarak tempuhnya adalah ...", "", "", "100 m", "150 m", "200 m", "250 m", "300 m", "C"],
      ["SOAL-2001", "Matematika (Kelas XI) - Ulangan Harian 1", "Turunan pertama dari fungsi f(x) = 3x² + 4x - 5 adalah ...", "", "", "3x + 4", "6x + 4", "6x - 5", "3x² + 4", "6x", "B"],
      ["SOAL-2002", "Matematika (Kelas XI) - Ulangan Harian 1", "Nilai determinan matriks A = [[2, 1], [4, 3]] adalah ...", "", "", "1", "2", "3", "4", "5", "B"],
      ["SOAL-2003", "Matematika (Kelas XII) - Ujian Sekolah", "Hasil dari integral tak tentu ∫ (4x³ - 6x²) dx adalah ...", "", "", "x⁴ - 2x³ + C", "x⁴ - 3x³ + C", "12x² - 12x + C", "4x⁴ - 6x³ + C", "2x⁴ - 2x³ + C", "A"],
      ["SOAL-3001", "Bahasa Indonesia (Kelas X) - Ulangan Harian 1", "Ide pokok yang menjadi dasar pengembangan sebuah paragraf disebut ...", "", "", "Kalimat penjelas", "Gagasan utama", "Kesimpulan", "Rangkuman", "Alur cerita", "B"],
      ["SOAL-4001", "Informatika (Kelas XII) - PAS Ganjil", "Manakah di bawah ini yang termasuk bahasa pemrograman web sisi klien (Frontend)?", "", "", "JavaScript", "MySQL", "PostgreSQL", "Apache", "Linux", "A"]
    ],
    "PENGATURAN_UJIAN": [
      ["mata_pelajaran", "status_ujian", "durasi_menit", "acak_soal", "acak_jawaban", "tampilkan_nilai"],
      ["Matematika (Kelas X) - Ulangan Harian 1", "OPEN", 30, "YA", "YA", "YA"],
      ["Matematika (Kelas X) - PTS Ganjil", "OPEN", 60, "YA", "YA", "TIDAK"],
      ["Fisika (Kelas X) - Ulangan Harian 1", "OPEN", 40, "YA", "YA", "YA"],
      ["Matematika (Kelas XI) - Ulangan Harian 1", "OPEN", 45, "YA", "YA", "YA"],
      ["Matematika (Kelas XII) - Ujian Sekolah", "OPEN", 90, "YA", "YA", "TIDAK"],
      ["Bahasa Indonesia (Kelas X) - Ulangan Harian 1", "OPEN", 45, "YA", "YA", "YA"],
      ["Informatika (Kelas XII) - PAS Ganjil", "CLOSE", 60, "TIDAK", "YA", "YA"]
    ],
    "HASIL_UJIAN": [
      ["timestamp", "nis", "nama_lengkap", "kelas", "mata_pelajaran", "jawaban_siswa", "nilai_akhir", "catatan_kecurangan"],
      ["2026-11-08 08:30:00", "2026001", "Ahmad Fauzi", "X-IPA-1", "Matematika (Kelas X) - Ulangan Harian 1", "{\"SOAL-1001\":\"B\",\"SOAL-1002\":\"C\",\"SOAL-1003\":\"A\"}", 67, "[]"]
    ],
    "PENGATURAN_WEBSITE": [
      ["kunci", "nilai"],
      ["nama_website", "Sistem Ujian Online Sekolah Digital"],
      ["subtitle_website", "Portal Akademik & Evaluasi Digital"],
      ["logo_url", ""],
      ["login_judul", "CBT Sekolah Digital"],
      ["login_subtitle", "Terintegrasi Relasi Tingkat Kelas (Kelas 1 s/d Kelas 12) & Multi-Evaluasi"],
      ["login_footer", "Isolasi Mapel Per Rombel Kelas & Anti-Cheat Fullscreen."],
      ["login_bg_gradient", "linear-gradient(135deg, #1e3a8a 0%, #2563eb 50%, #0ea5e9 100%)"]
    ]
  };

  for (const [sheetName, rows] of Object.entries(seedData)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }

    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange(1, 1, 1, rows[0].length)
      .setFontWeight("bold")
      .setBackground("#dbeafe")
      .setFontColor("#1e3a8a");
    sheet.autoResizeColumns(1, rows[0].length);
  }

  // Buat juga tabel kontrol DATA_PEMBELI di Master Spreadsheet Anda (tidak menghapus data pembeli yang sudah ada)
  try {
    ensureDataPembeliSheet_(getMasterOwnerSpreadsheet_());
  } catch (e) {}

  Logger.log("SUKSES! Database V3.0 (beserta DATA_ADMIN & DATA_PEMBELI) berhasil dibuat!");
}

/**
 * Render halaman utama SPA atau melayani request API JSON (Mode Hybrid Vercel / Hosting Eksternal)
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    const reqSsId = (e.parameter.spreadsheetId || e.parameter.ssId || "").trim();
    return handleExternalApiRequest_(
      e.parameter.action,
      e.parameter.args ? JSON.parse(e.parameter.args) : [],
      reqSsId
    );
  }
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Sistem Ujian Online Sekolah Digital")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, shrink-to-fit=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Endpoint REST API (POST) agar frontend index.html yang di-hosting di Vercel / Netlify / Cloudflare
 * dapat memanggil seluruh fungsi backend di Google Apps Script & Google Spreadsheet tanpa terkena blokir CORS.
 */
function doPost(e) {
  try {
    const rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    const body = JSON.parse(rawBody);
    const action = String(body.action || (e && e.parameter && e.parameter.action) || "").trim();
    const args = Array.isArray(body.args) ? body.args : [];
    const reqSsId = String(body.spreadsheetId || (e && e.parameter && e.parameter.spreadsheetId) || "").trim();
    return handleExternalApiRequest_(action, args, reqSsId);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: "Format request API tidak valid: " + err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Dispatcher fungsi backend yang diizinkan untuk dipanggil dari aplikasi Vercel / Web Eksternal
 * Sekaligus memvalidasi apakah ID Spreadsheet Pembeli berstatus AKTIF di sheet DATA_PEMBELI.
 */
function handleExternalApiRequest_(action, args, buyerSpreadsheetId) {
  const cleanAction = String(action || "").trim();

  // Cegah pemanggilan fungsi internal/private (yang diakhiri underscore _) atau fungsi bawaan
  if (
    !cleanAction ||
    cleanAction.endsWith("_") ||
    cleanAction === "doGet" ||
    cleanAction === "doPost" ||
    cleanAction === "setupDatabaseDanDataAwal"
  ) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: "Fungsi API '" + cleanAction + "' tidak diizinkan." })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  // 1. Cek aktivasi Spreadsheet ID Pembeli di sheet DATA_PEMBELI milik Anda
  const accessCheck = verifyBuyerSpreadsheetAccess_(buyerSpreadsheetId);
  if (!accessCheck.allowed) {
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "error",
        code: "SPREADSHEET_NOT_ACTIVE",
        message: accessCheck.message
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  // 2. Arahkan koneksi database ke Spreadsheet milik Pembeli selama request ini berlangsung
  _activeTargetSpreadsheetId = accessCheck.targetId || "";
  _cachedSs = null;

  const targetFn = typeof globalThis !== "undefined" ? globalThis[cleanAction] : this[cleanAction];
  if (typeof targetFn !== "function") {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: "Fungsi API '" + cleanAction + "' tidak ditemukan di code.gs." })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const result = targetFn.apply(null, Array.isArray(args) ? args : []);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err && err.message ? err.message : String(err) })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Membuka koneksi Spreadsheet (Otomatis mengarah ke Spreadsheet Pembeli jika dipanggil dari Frontend Pembeli,
 * dan otomatis men-setup 10 sheet database awal jika Spreadsheet Pembeli masih baru/kosong!)
 */
function getSpreadsheet_(forceCheckSchema) {
  if (_cachedSs && !forceCheckSchema) {
    return _cachedSs;
  }

  let ss = null;
  const targetId = extractCleanSpreadsheetId_(_activeTargetSpreadsheetId) || extractCleanSpreadsheetId_(SPREADSHEET_ID);

  try {
    if (targetId) {
      ss = SpreadsheetApp.openById(targetId);
    } else {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
  } catch (err) {
    if (_activeTargetSpreadsheetId) {
      throw new Error(
        "Gagal mengakses Google Spreadsheet Pembeli (ID: " +
        _activeTargetSpreadsheetId +
        "). Pastikan Spreadsheet tersebut sudah di-Share (Bagikan) dengan akses 'Editor' ke email server atau 'Siapa saja yang memiliki link dapat mengedit'."
      );
    }
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  if (!ss) {
    throw new Error("Spreadsheet tidak ditemukan. Pastikan SPREADSHEET_ID sudah diisi dan diberi akses Editor.");
  }

  // Jika Spreadsheet milik pembeli masih kosong (baru memiliki Sheet1 bawaan dan belum punya DATA_ADMIN),
  // otomatis buatkan ke-10 sheet database beserta akun admin default agar pembeli langsung bisa login!
  if (forceCheckSchema || !ss.getSheetByName(SHEET_NAMES.ADMIN)) {
    ensureSheetsExist_(ss);
  }
  _cachedSs = ss;
  return ss;
}

/**
 * Inisialisasi ke-9 sheet jika belum ada (termasuk DATA_ADMIN & DATA_KELAS default)
 */
function ensureSheetsExist_(ss) {
  const schemas = {
    [SHEET_NAMES.ADMIN]: ["id_admin", "nama_lengkap", "username", "password"],
    [SHEET_NAMES.MASTER_MAPEL]: ["kode_mapel", "nama_mapel", "kelas", "label_mapel_kelas", "kelompok"],
    [SHEET_NAMES.MASTER_JENIS]: ["kode_jenis", "nama_jenis", "keterangan", "pemilik_guru", "mapel_relasi"],
    [SHEET_NAMES.KELAS]: ["kode_kelas", "nama_kelas", "tingkat_kelas", "wali_kelas"],
    [SHEET_NAMES.GURU]: ["nip_atau_id", "nama_lengkap", "mata_pelajaran", "username", "password"],
    [SHEET_NAMES.SISWA]: ["nis", "nama_lengkap", "kelas", "username", "password"],
    [SHEET_NAMES.SOAL]: ["id_soal", "mata_pelajaran", "pertanyaan", "id_gambar", "id_suara", "opsi_a", "opsi_b", "opsi_c", "opsi_d", "opsi_e", "kunci_jawaban"],
    [SHEET_NAMES.PENGATURAN]: ["mata_pelajaran", "status_ujian", "durasi_menit", "acak_soal", "acak_jawaban", "tampilkan_nilai"],
    [SHEET_NAMES.HASIL]: ["timestamp", "nis", "nama_lengkap", "kelas", "mata_pelajaran", "jawaban_siswa", "nilai_akhir", "catatan_kecurangan"],
    [SHEET_NAMES.WEB_SETTINGS]: ["kunci", "nilai"]
  };

  for (const [name, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#dbeafe");
      if (name === SHEET_NAMES.ADMIN) {
        sheet.appendRow(["ADM-01", "Administrator", "admin", "admin123"]);
      }
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#dbeafe");
      if (name === SHEET_NAMES.ADMIN) {
        sheet.appendRow(["ADM-01", "Administrator", "admin", "admin123"]);
      }
    }
  }

  ensureMasterJenisSchemaAndIsolation_(ss);
  ensureKelasSheetSeeded_(ss);
  ensureHasilUjianSchema_(ss);
}

/**
 * Memastikan sheet HASIL_UJIAN memiliki kolom ke-8: catatan_kecurangan
 */
function ensureHasilUjianSchema_(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.HASIL);
  if (!sheet || sheet.getLastRow() === 0) return;
  const lastCol = Math.max(sheet.getLastColumn(), 8);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim().toLowerCase());
  if (headers[7] !== "catatan_kecurangan") {
    sheet.getRange(1, 8).setValue("catatan_kecurangan").setFontWeight("bold").setBackground("#dbeafe");
  }
}

/**
 * Konversi Tingkat Kelas Universal (Kelas 1 s/d Kelas 12 & Romawi I s/d XII)
 * Menghasilkan angka kanonik "1" s/d "12" agar semua jenjang kelas (SD/SMP/SMA) terhubung otomatis.
 */
function normalizeGradeLevel_(rawVal) {
  let str = String(rawVal || "").trim().toUpperCase();
  if (!str) return "";
  str = str.replace(/^KELAS\s+/i, "").trim();
  const token = str.split(/[-_\s/]+/)[0];

  const romanMap = {
    "I": "1", "II": "2", "III": "3", "IV": "4", "V": "5", "VI": "6",
    "VII": "7", "VIII": "8", "IX": "9", "X": "10", "XI": "11", "XII": "12"
  };
  if (romanMap[token]) return romanMap[token];

  const num = parseInt(token, 10);
  if (!isNaN(num) && num >= 1 && num <= 12) {
    return String(num);
  }
  return token;
}

/**
 * Memastikan setiap mata pelajaran & tingkat kelas (Kelas 1 - 12) yang diampu guru terdaftar di MASTER_MAPEL
 */
function ensureGuruMapelInMasterMapel_(ss, mapelArray) {
  if (!mapelArray || !Array.isArray(mapelArray) || mapelArray.length === 0) return;
  const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_MAPEL);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const existingRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];

  mapelArray.forEach(mp => {
    const cleanMp = String(mp || "").trim();
    if (!cleanMp) return;
    const parsed = splitMapelDanJenis_(cleanMp);
    const normGrade = normalizeGradeLevel_(parsed.tingkat_kelas) || "10";
    const namaMurni = parsed.nama_mapel_murni || cleanMp;

    const alreadyExists = existingRows.some(r => {
      const rLabel = String(r[3] || "").trim();
      const rParsed = splitMapelDanJenis_(rLabel || String(r[1] || ""));
      const rGrade = normalizeGradeLevel_(r[2] || rParsed.tingkat_kelas);
      return (
        rLabel.toUpperCase() === cleanMp.toUpperCase() ||
        (rParsed.nama_mapel_murni.toUpperCase() === namaMurni.toUpperCase() && rGrade === normGrade)
      );
    });

    if (!alreadyExists) {
      const abbr = namaMurni.replace(/[^A-Za-z]/g, "").slice(0, 3).toUpperCase() || "MPL";
      const kode = "MP-" + normGrade + "-" + abbr + "-" + (existingRows.length + 1);
      const label = `${namaMurni} (Kelas ${normGrade})`;
      const newRow = [kode, namaMurni, normGrade, label, "Mapel Sekolah"];
      sheet.appendRow(newRow);
      existingRows.push(newRow);
    }
  });
}

/**
 * Memastikan sheet MASTER_JENIS_PENILAIAN memiliki kolom pemilik_guru & mapel_relasi
 * serta memigrasikan baris lama yang belum memiliki pemilik_guru agar terisolasi per guru.
 */
function ensureMasterJenisSchemaAndIsolation_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.MASTER_JENIS);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 3);
  if (lastRow === 0) {
    const headers = ["kode_jenis", "nama_jenis", "keterangan", "pemilik_guru", "mapel_relasi"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#dbeafe");
    return;
  }

  const headerRow = sheet.getRange(1, 1, 1, Math.max(lastCol, 5)).getValues()[0].map(h => String(h || "").trim().toLowerCase());
  if (headerRow[3] !== "pemilik_guru") {
    sheet.getRange(1, 4).setValue("pemilik_guru").setFontWeight("bold").setBackground("#dbeafe");
  }
  if (headerRow[4] !== "mapel_relasi") {
    sheet.getRange(1, 5).setValue("mapel_relasi").setFontWeight("bold").setBackground("#dbeafe");
  }

  if (lastRow <= 1) return;

  const allValues = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const hasAnyOwnedRow = allValues.some(r => String(r[3] || "").trim() !== "");
  const hasUnownedRow = allValues.some(r => String(r[1] || "").trim() !== "" && String(r[3] || "").trim() === "");

  if (!hasUnownedRow) return;

  // Jika terdapat baris lama tanpa pemilik_guru, migrasikan menjadi data terisolasi per guru
  const ownedRows = allValues.filter(r => String(r[1] || "").trim() !== "" && String(r[3] || "").trim() !== "");

  if (!hasAnyOwnedRow) {
    const sheetGuru = ss.getSheetByName(SHEET_NAMES.GURU);
    const sheetPengaturan = ss.getSheetByName(SHEET_NAMES.PENGATURAN);
    const sheetSoal = ss.getSheetByName(SHEET_NAMES.SOAL);

    const guruRows = (sheetGuru && sheetGuru.getLastRow() > 1)
      ? sheetGuru.getRange(2, 1, sheetGuru.getLastRow() - 1, 5).getValues()
      : [];
    const pengaturanRows = (sheetPengaturan && sheetPengaturan.getLastRow() > 1)
      ? sheetPengaturan.getRange(2, 1, sheetPengaturan.getLastRow() - 1, 1).getValues()
      : [];
    const soalRows = (sheetSoal && sheetSoal.getLastRow() > 1)
      ? sheetSoal.getRange(2, 1, sheetSoal.getLastRow() - 1, 2).getValues()
      : [];

    const existingPakets = [];
    pengaturanRows.forEach(r => { if (r[0]) existingPakets.push(String(r[0]).trim()); });
    soalRows.forEach(r => { if (r[1]) existingPakets.push(String(r[1]).trim()); });

    let counter = 1;
    guruRows.forEach(g => {
      const username = String(g[3] || g[0] || "").trim();
      const mapelList = parseGuruMapelList_(g[2]);
      if (!username || mapelList.length === 0) return;

      const addedPairs = new Set();
      mapelList.forEach(mp => {
        existingPakets.forEach(pkt => {
          const parsed = splitMapelDanJenis_(pkt);
          if (isAllowedForGuru_(pkt, [mp]) && parsed.jenis_ujian) {
            const pairKey = mp.toUpperCase() + "::" + parsed.jenis_ujian.toUpperCase();
            if (!addedPairs.has(pairKey)) {
              addedPairs.add(pairKey);
              ownedRows.push([
                "JP-" + String(counter++).padStart(2, "0"),
                parsed.jenis_ujian,
                "Evaluasi " + mp,
                username,
                mp
              ]);
            }
          }
        });

        const defaultJenis = "Ulangan Harian 1";
        const defKey = mp.toUpperCase() + "::" + defaultJenis.toUpperCase();
        if (!addedPairs.has(defKey)) {
          addedPairs.add(defKey);
          ownedRows.push([
            "JP-" + String(counter++).padStart(2, "0"),
            defaultJenis,
            "Penilaian Harian " + mp,
            username,
            mp
          ]);
        }
      });
    });
  }

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }
  if (ownedRows.length > 0) {
    sheet.getRange(2, 1, ownedRows.length, 5).setValues(ownedRows);
  }
}

/**
 * Memastikan sheet DATA_KELAS menggunakan 4 kolom bersih (kode_kelas, nama_kelas, tingkat_kelas, wali_kelas)
 * tanpa kolom jurusan, serta mendukung Kelas 1 s/d Kelas 12.
 */
function ensureKelasSheetSeeded_(ss) {
  let sheetKelas = ss.getSheetByName(SHEET_NAMES.KELAS);
  if (!sheetKelas) return;

  // Jika sheet DATA_KELAS lama masih memiliki kolom ke-4 'jurusan', hapus kolom jurusan otomatis
  if (sheetKelas.getLastColumn() >= 5) {
    const hdr = sheetKelas.getRange(1, 1, 1, 5).getValues()[0].map(h => String(h || "").trim().toLowerCase());
    if (hdr[3] === "jurusan") {
      sheetKelas.deleteColumn(4);
    }
  }

  if (sheetKelas.getLastRow() > 1) return;

  const defaultKelasMap = {
    "1-A": { tingkat: "1", wali: "Siti Aminah, M.Pd." },
    "7-A": { tingkat: "7", wali: "Budi Santoso, S.Pd." },
    "X-IPA-1": { tingkat: "10", wali: "Budi Santoso, S.Pd." },
    "X-IPA-2": { tingkat: "10", wali: "Siti Aminah, M.Pd." },
    "XI-IPA-1": { tingkat: "11", wali: "Ratna Sari, M.Pd." },
    "XII-IPA-1": { tingkat: "12", wali: "Andi Pratama, S.Kom." }
  };

  const sheetSiswa = ss.getSheetByName(SHEET_NAMES.SISWA);
  if (sheetSiswa && sheetSiswa.getLastRow() > 1) {
    const siswaRows = sheetSiswa.getRange(2, 1, sheetSiswa.getLastRow() - 1, 3).getValues();
    siswaRows.forEach(r => {
      const kls = String(r[2] || "").trim().toUpperCase();
      if (kls && !defaultKelasMap[kls]) {
        const tingkat = normalizeGradeLevel_(kls) || "10";
        defaultKelasMap[kls] = { tingkat: tingkat, wali: "-" };
      }
    });
  }

  const rowsToInsert = [];
  let idx = 1;
  Object.keys(defaultKelasMap).forEach(namaKls => {
    const info = defaultKelasMap[namaKls];
    rowsToInsert.push([
      "KLS-" + String(idx++).padStart(2, "0"),
      namaKls,
      normalizeGradeLevel_(info.tingkat) || "10",
      info.wali
    ]);
  });

  if (rowsToInsert.length > 0) {
    sheetKelas.getRange(2, 1, rowsToInsert.length, 4).setValues(rowsToInsert);
  }
}

/**
 * Helper Cepat: Membaca isi sheet menjadi Array of Objects
 */
function getSheetData_(sheetName) {
  const ss = getSpreadsheet_(false);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ensureSheetsExist_(ss);
    sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
  } else {
    if (sheetName === SHEET_NAMES.MASTER_JENIS) {
      ensureMasterJenisSchemaAndIsolation_(ss);
    } else if (sheetName === SHEET_NAMES.KELAS) {
      ensureKelasSheetSeeded_(ss);
    } else if (sheetName === SHEET_NAMES.HASIL) {
      ensureHasilUjianSchema_(ss);
    }
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];

  const rows = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = rows[0].map(h => String(h).trim().toLowerCase());
  const tz = Session.getScriptTimeZone() || "Asia/Jakarta";
  const result = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[0] === "" && row[1] === "" && row[2] === "") continue;

    const obj = { _rowIndex: i + 1 };
    for (let c = 0; c < headers.length; c++) {
      let val = row[c];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, tz, "yyyy-MM-dd HH:mm:ss");
      }
      obj[headers[c]] = val !== undefined && val !== null ? String(val).trim() : "";
    }
    result.push(obj);
  }
  return result;
}

function parseGuruMapelList_(rawMapelString) {
  if (!rawMapelString) return [];
  return String(rawMapelString)
    .split(/[,;|]+/)
    .map(m => m.trim())
    .filter(Boolean);
}

function isPureGradeToken_(token) {
  const t = String(token || "").trim().toUpperCase();
  if (!t) return false;
  return /^(?:[1-9]|1[0-2]|I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)$/.test(t);
}

function splitMapelDanJenis_(labelLengkap) {
  const raw = String(labelLengkap || "").trim();
  if (!raw) {
    return {
      mapel_induk: "",
      nama_mapel_murni: "",
      tingkat_kelas: "SEMUA",
      target_rombel: [],
      jenis_ujian: "Ulangan Harian 1",
      paket_lengkap: ""
    };
  }

  const parts = raw.split(/\s+-\s+/);
  const mapelInduk = parts[0].trim();
  const jenisUjian = parts.length >= 2 ? parts.slice(1).join(" - ").trim() : "Ujian Umum";

  let namaMurni = mapelInduk;
  let tingkatKelas = "SEMUA";
  let targetRombel = [];
  const matchKelas = mapelInduk.match(/\(Kelas\s+([^)]+)\)/i);
  if (matchKelas) {
    const rawKelasContent = String(matchKelas[1] || "").trim();
    const tokens = rawKelasContent
      .split(/\s*\+\s*/)
      .map(t => t.trim())
      .filter(Boolean);

    if (tokens.length === 1 && isPureGradeToken_(tokens[0])) {
      tingkatKelas = normalizeGradeLevel_(tokens[0]);
      targetRombel = [];
    } else if (tokens.length > 0) {
      targetRombel = tokens.map(t => t.toUpperCase());
      tingkatKelas = normalizeGradeLevel_(tokens[0]) || "SEMUA";
    }
    namaMurni = mapelInduk.replace(/\s*\(Kelas\s+[^)]+\)/i, "").trim();
  }

  return {
    mapel_induk: mapelInduk,
    nama_mapel_murni: namaMurni,
    tingkat_kelas: tingkatKelas,
    target_rombel: targetRombel,
    jenis_ujian: jenisUjian,
    paket_lengkap: parts.length >= 2 ? (mapelInduk + " - " + jenisUjian) : mapelInduk
  };
}

function extractStudentGradeLevel_(kelasSiswa) {
  const str = String(kelasSiswa || "").trim().toUpperCase();
  if (!str) return "";

  // 1. Cek relasi langsung dari tabel DATA_KELAS terlebih dahulu
  try {
    const dataKelas = getSheetData_(SHEET_NAMES.KELAS);
    const foundKelas = dataKelas.find(k => String(k.nama_kelas || "").trim().toUpperCase() === str);
    if (foundKelas && foundKelas.tingkat_kelas) {
      return normalizeGradeLevel_(foundKelas.tingkat_kelas);
    }
  } catch (e) {}

  // 2. Fallback ekstraksi dari awalan nama kelas (mendukung Kelas 1 s/d 12 maupun I s/d XII)
  return normalizeGradeLevel_(str);
}

function isExamForStudentClass_(labelPaketUjian, kelasSiswa) {
  const parsed = splitMapelDanJenis_(labelPaketUjian);
  const kelasSiswaClean = String(kelasSiswa || "").trim().toUpperCase();

  // Jika paket ujian dikhususkan untuk rombel tertentu (misal: 10-A + 10-B + 10-C atau 10-PPLG-1 + 10-PPLG-2)
  if (parsed.target_rombel && parsed.target_rombel.length > 0) {
    if (!kelasSiswaClean) return false;
    return parsed.target_rombel.some(r => r.toUpperCase() === kelasSiswaClean);
  }

  // Fallback mode umum per tingkat kelas (misal: Kelas 10)
  const targetKelas = normalizeGradeLevel_(parsed.tingkat_kelas);
  if (!targetKelas || targetKelas === "SEMUA" || targetKelas === "ALL") return true;

  const tingkatSiswa = normalizeGradeLevel_(extractStudentGradeLevel_(kelasSiswaClean));
  return targetKelas === tingkatSiswa || parsed.tingkat_kelas === kelasSiswaClean;
}

function isAllowedForGuru_(labelDiSheet, allowedMapelArray) {
  if (!allowedMapelArray || allowedMapelArray.length === 0) return false;
  const parsedSheet = splitMapelDanJenis_(labelDiSheet);
  const indukUpper = parsedSheet.mapel_induk.toUpperCase();
  const fullUpper = String(labelDiSheet || "").trim().toUpperCase();
  const sheetNamaMurni = parsedSheet.nama_mapel_murni.toUpperCase();
  const sheetGradeNorm = normalizeGradeLevel_(parsedSheet.tingkat_kelas);
  const sheetRombels = (parsedSheet.target_rombel || []).slice().sort();

  return allowedMapelArray.some(m => {
    const mUpper = String(m).trim().toUpperCase();
    if (indukUpper === mUpper || fullUpper === mUpper || fullUpper.startsWith(mUpper + " - ")) {
      return true;
    }
    const parsedAllowed = splitMapelDanJenis_(m);
    const allowedNamaMurni = parsedAllowed.nama_mapel_murni.toUpperCase();
    const allowedGradeNorm = normalizeGradeLevel_(parsedAllowed.tingkat_kelas);
    const allowedRombels = (parsedAllowed.target_rombel || []).slice().sort();

    if (!sheetNamaMurni || allowedNamaMurni !== sheetNamaMurni) return false;

    // Jika salah satu menggunakan rombel spesifik (misal 10-A + 10-B atau 10-PPLG-1),
    // wajib cocok secara KETAT agar guru dengan mapel & tingkat sama tetapi rombel berbeda tetap terisolasi!
    if (sheetRombels.length > 0 || allowedRombels.length > 0) {
      if (sheetRombels.length !== allowedRombels.length) return false;
      return sheetRombels.every((r, idx) => r === allowedRombels[idx]);
    }

    return Boolean(sheetGradeNorm && allowedGradeNorm === sheetGradeNorm);
  });
}

/**
 * Helper: Memeriksa apakah suatu label mapel (baik tingkat umum maupun rombel spesifik)
 * merujuk pada Master Mapel yang sama (nama_mapel_murni + tingkat_kelas).
 */
function isSameMasterMapel_(labelDiSheet, masterNamaMurni, masterGradeNorm) {
  const parsed = splitMapelDanJenis_(labelDiSheet);
  const sNama = parsed.nama_mapel_murni.toUpperCase();
  const sGrade = normalizeGradeLevel_(parsed.tingkat_kelas);
  return Boolean(
    sNama &&
    sNama === String(masterNamaMurni || "").trim().toUpperCase() &&
    sGrade &&
    sGrade === normalizeGradeLevel_(masterGradeNorm)
  );
}

function shuffleArray_(arr) {
  const copy = arr.slice();
  for (let i = copy.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const temp = copy[i];
    copy[i] = copy[j];
    copy[j] = temp;
  }
  return copy;
}

function formatMediaUrl_(input, type) {
  if (!input) return "";
  const str = String(input).trim();
  if (!str) return "";

  if (str.startsWith("http://") || str.startsWith("https://")) {
    const driveMatch = str.match(/[-\w]{25,}/);
    if (str.includes("drive.google.com") && driveMatch) {
      const fileId = driveMatch[0];
      return type === "audio"
        ? "https://docs.google.com/uc?export=open&id=" + fileId
        : "https://lh3.googleusercontent.com/d/" + fileId;
    }
    return str;
  }

  return type === "audio"
    ? "https://docs.google.com/uc?export=open&id=" + str
    : "https://lh3.googleusercontent.com/d/" + str;
}

/**
 * ============================================================================
 * 1. MODUL AUTENTIKASI & PENGATURAN PROFIL SISWA
 * ============================================================================
 */
function loginUser(username, password) {
  try {
    const userClean = String(username || "").trim();
    const passClean = String(password || "").trim();

    if (!userClean || !passClean) {
      return { status: "error", message: "Username dan Password wajib diisi!" };
    }

    // 1. Cek ke sheet DATA_ADMIN
    const dataAdmin = getSheetData_(SHEET_NAMES.ADMIN);
    const foundAdmin = dataAdmin.find(
      a => String(a.username).trim() === userClean && String(a.password).trim() === passClean
    );

    if (foundAdmin) {
      return {
        status: "success",
        role: "admin",
        data: {
          id: foundAdmin.id_admin,
          nama: foundAdmin.nama_lengkap,
          username: foundAdmin.username
        }
      };
    }

    // 2. Cek ke sheet DATA_GURU
    const dataGuru = getSheetData_(SHEET_NAMES.GURU);
    const foundGuru = dataGuru.find(
      g => String(g.username).trim() === userClean && String(g.password).trim() === passClean
    );

    if (foundGuru) {
      const mapelList = parseGuruMapelList_(foundGuru.mata_pelajaran);
      return {
        status: "success",
        role: "guru",
        data: {
          id: foundGuru.nip_atau_id,
          nama: foundGuru.nama_lengkap,
          mapel: foundGuru.mata_pelajaran,
          mapel_list: mapelList,
          username: foundGuru.username
        }
      };
    }

    // 3. Cek ke sheet DATA_SISWA
    const dataSiswa = getSheetData_(SHEET_NAMES.SISWA);
    const foundSiswa = dataSiswa.find(
      s => String(s.username).trim() === userClean && String(s.password).trim() === passClean
    );

    if (foundSiswa) {
      return {
        status: "success",
        role: "siswa",
        data: {
          nis: foundSiswa.nis,
          nama: foundSiswa.nama_lengkap,
          kelas: foundSiswa.kelas,
          tingkat: extractStudentGradeLevel_(foundSiswa.kelas),
          username: foundSiswa.username
        }
      };
    }

    return { status: "error", message: "Username atau Password salah!" };
  } catch (err) {
    return { status: "error", message: "Gagal memproses login: " + err.message };
  }
}

/**
 * Akses Cepat Login (1-Klik) berdasarkan Role ('admin', 'guru', 'siswa')
 * Mengambil akun pertama yang tersedia pada tabel DATA_ADMIN, DATA_GURU, atau DATA_SISWA.
 */
function quickLoginByRole(roleName) {
  try {
    const role = String(roleName || "").trim().toLowerCase();
    if (role === "admin") {
      const dataAdmin = getSheetData_(SHEET_NAMES.ADMIN);
      if (dataAdmin.length === 0) {
        return { status: "error", message: "Data Administrator belum ada di sheet DATA_ADMIN." };
      }
      const foundAdmin = dataAdmin[0];
      return {
        status: "success",
        role: "admin",
        data: {
          id: foundAdmin.id_admin,
          nama: foundAdmin.nama_lengkap,
          username: foundAdmin.username
        }
      };
    }

    if (role === "guru") {
      const dataGuru = getSheetData_(SHEET_NAMES.GURU);
      if (dataGuru.length === 0) {
        return { status: "error", message: "Data Guru belum ada di sheet DATA_GURU." };
      }
      const foundGuru = dataGuru[0];
      const mapelList = parseGuruMapelList_(foundGuru.mata_pelajaran);
      return {
        status: "success",
        role: "guru",
        data: {
          id: foundGuru.nip_atau_id,
          nama: foundGuru.nama_lengkap,
          mapel: foundGuru.mata_pelajaran,
          mapel_list: mapelList,
          username: foundGuru.username
        }
      };
    }

    if (role === "siswa") {
      const dataSiswa = getSheetData_(SHEET_NAMES.SISWA);
      if (dataSiswa.length === 0) {
        return { status: "error", message: "Data Siswa belum ada di sheet DATA_SISWA." };
      }
      const foundSiswa = dataSiswa[0];
      return {
        status: "success",
        role: "siswa",
        data: {
          nis: foundSiswa.nis,
          nama: foundSiswa.nama_lengkap,
          kelas: foundSiswa.kelas,
          tingkat: extractStudentGradeLevel_(foundSiswa.kelas),
          username: foundSiswa.username
        }
      };
    }

    return { status: "error", message: "Role tidak dikenali." };
  } catch (err) {
    return { status: "error", message: "Gagal memproses akses cepat: " + err.message };
  }
}

/**
 * FITUR BARU V2.5: Mengubah Password Siswa dari Halaman Pengaturan Profil Siswa
 */
function ubahPasswordSiswa(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const nis = String(payload.nis || "").trim();
    const passwordLama = String(payload.password_lama || "").trim();
    const passwordBaru = String(payload.password_baru || "").trim();

    if (!nis || !passwordLama || !passwordBaru) {
      return { status: "error", message: "Password lama dan password baru wajib diisi!" };
    }

    if (passwordBaru.length < 5) {
      return { status: "error", message: "Password baru minimal harus terdiri dari 5 karakter!" };
    }

    if (passwordLama === passwordBaru) {
      return { status: "error", message: "Password baru tidak boleh sama dengan password lama!" };
    }

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.SISWA);
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return { status: "error", message: "Data siswa tidak ditemukan di Spreadsheet." };
    }

    // Kolom DATA_SISWA: A=nis(1), B=nama_lengkap(2), C=kelas(3), D=username(4), E=password(5)
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

    for (let i = 0; i < data.length; i++) {
      const rowNis = String(data[i][0]).trim();
      const rowPass = String(data[i][4]).trim();

      if (rowNis === nis) {
        if (rowPass !== passwordLama) {
          return { status: "error", message: "Password saat ini (lama) yang Anda masukkan salah!" };
        }

        const targetRow = i + 2;
        sheet.getRange(targetRow, 5).setValue(passwordBaru);

        return {
          status: "success",
          message: "Password akun Anda berhasil diperbarui! Gunakan password baru ini saat login berikutnya."
        };
      }
    }

    return { status: "error", message: "Akun siswa dengan NIS tersebut tidak ditemukan." };
  } catch (err) {
    return { status: "error", message: "Gagal mengubah password: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mengubah Password Guru dari Halaman Profil Guru
 */
function ubahPasswordGuru(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const nip = String(payload.nip_atau_id || "").trim();
    const username = String(payload.username || "").trim();
    const passwordLama = String(payload.password_lama || "").trim();
    const passwordBaru = String(payload.password_baru || "").trim();

    if ((!nip && !username) || !passwordLama || !passwordBaru) {
      return { status: "error", message: "Password lama dan password baru wajib diisi!" };
    }

    if (passwordBaru.length < 5) {
      return { status: "error", message: "Password baru minimal harus terdiri dari 5 karakter!" };
    }

    if (passwordLama === passwordBaru) {
      return { status: "error", message: "Password baru tidak boleh sama dengan password lama!" };
    }

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.GURU);
    const lastRow = sheet ? sheet.getLastRow() : 0;

    if (lastRow <= 1) {
      return { status: "error", message: "Data guru tidak ditemukan di Spreadsheet." };
    }

    // Kolom DATA_GURU: A=nip_atau_id(1), B=nama_lengkap(2), C=mata_pelajaran(3), D=username(4), E=password(5)
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

    for (let i = 0; i < data.length; i++) {
      const rowNip = String(data[i][0]).trim();
      const rowUser = String(data[i][3]).trim();
      const rowPass = String(data[i][4]).trim();

      const matched =
        (nip && rowNip.toLowerCase() === nip.toLowerCase()) ||
        (username && rowUser.toLowerCase() === username.toLowerCase());

      if (matched) {
        if (rowPass !== passwordLama) {
          return { status: "error", message: "Password saat ini (lama) yang Anda masukkan salah!" };
        }

        const targetRow = i + 2;
        sheet.getRange(targetRow, 5).setValue(passwordBaru);

        return {
          status: "success",
          message: "Password akun guru Anda berhasil diperbarui! Gunakan password baru ini saat login berikutnya."
        };
      }
    }

    return { status: "error", message: "Akun guru tidak ditemukan di Spreadsheet." };
  } catch (err) {
    return { status: "error", message: "Gagal mengubah password guru: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * PENGATURAN WEBSITE & HALAMAN LOGIN (TAMPIL DI SEMUA HALAMAN)
 */
function getDefaultWebsiteSettings_() {
  return {
    nama_website: "Sistem Ujian Online Sekolah Digital",
    subtitle_website: "Portal Akademik & Evaluasi Digital",
    logo_url: "",
    login_judul: "CBT Sekolah Digital",
    login_subtitle: "Terintegrasi Relasi Tingkat Kelas (Kelas 1 s/d Kelas 12) & Multi-Evaluasi",
    login_footer: "Isolasi Mapel Per Rombel Kelas & Anti-Cheat Fullscreen.",
    login_bg_gradient: "linear-gradient(135deg, #1e3a8a 0%, #2563eb 50%, #0ea5e9 100%)"
  };
}

function normalizeLogoUrl_(rawLogo) {
  const str = String(rawLogo || "").trim();
  if (!str) return "";
  if (str.startsWith("data:image/")) return str;
  return formatMediaUrl_(str, "image");
}

function getPengaturanWebsite() {
  try {
    const defaults = getDefaultWebsiteSettings_();
    const rows = getSheetData_(SHEET_NAMES.WEB_SETTINGS);
    const settings = Object.assign({}, defaults);

    rows.forEach(r => {
      const k = String(r.kunci || "").trim();
      if (k && Object.prototype.hasOwnProperty.call(settings, k)) {
        const val = String(r.nilai !== undefined ? r.nilai : "").trim();
        if (k === "logo_url") {
          settings[k] = val;
        } else if (val !== "") {
          settings[k] = val;
        }
      }
    });

    settings.logo_url_resolved = normalizeLogoUrl_(settings.logo_url);
    return { status: "success", settings: settings };
  } catch (err) {
    const fallback = getDefaultWebsiteSettings_();
    fallback.logo_url_resolved = "";
    return { status: "success", settings: fallback };
  }
}

function simpanPengaturanWebsite(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const defaults = getDefaultWebsiteSettings_();
    const nextSettings = {
      nama_website: String(payload.nama_website || "").trim() || defaults.nama_website,
      subtitle_website: String(payload.subtitle_website || "").trim() || defaults.subtitle_website,
      logo_url: String(payload.logo_url || "").trim(),
      login_judul: String(payload.login_judul || "").trim() || defaults.login_judul,
      login_subtitle: String(payload.login_subtitle || "").trim() || defaults.login_subtitle,
      login_footer: String(payload.login_footer || "").trim() || defaults.login_footer,
      login_bg_gradient: String(payload.login_bg_gradient || "").trim() || defaults.login_bg_gradient
    };

    if (nextSettings.logo_url.length > 48000) {
      return {
        status: "error",
        message: "Ukuran data logo terlalu besar untuk disimpan di sel Spreadsheet. Gunakan gambar yang lebih kecil atau tautan URL / ID Google Drive."
      };
    }

    const ss = getSpreadsheet_(false);
    let sheet = ss.getSheetByName(SHEET_NAMES.WEB_SETTINGS);
    if (!sheet) {
      ensureSheetsExist_(ss);
      sheet = ss.getSheetByName(SHEET_NAMES.WEB_SETTINGS);
    }

    sheet.clear();
    const rows = [["kunci", "nilai"]];
    Object.keys(nextSettings).forEach(k => {
      rows.push([k, nextSettings[k]]);
    });
    sheet.getRange(1, 1, rows.length, 2).setValues(rows);
    sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#dbeafe").setFontColor("#1e3a8a");

    nextSettings.logo_url_resolved = normalizeLogoUrl_(nextSettings.logo_url);
    return {
      status: "success",
      message: "Pengaturan Website & Halaman Login berhasil disimpan dan langsung diterapkan!",
      settings: nextSettings
    };
  } catch (err) {
    return { status: "error", message: "Gagal menyimpan pengaturan website: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * 2. MODUL SISWA (BERELASI DENGAN TINGKAT KELAS SISWA)
 * ============================================================================
 */
function getDataDashboardSiswa(nis) {
  try {
    const nisClean = String(nis || "").trim();
    const dataSiswa = getSheetData_(SHEET_NAMES.SISWA);
    const siswaObj = dataSiswa.find(s => String(s.nis).trim() === nisClean);
    const kelasSiswa = siswaObj ? siswaObj.kelas : "";

    const pengaturanList = getSheetData_(SHEET_NAMES.PENGATURAN);
    const bankSoal = getSheetData_(SHEET_NAMES.SOAL);
    const hasilList = getSheetData_(SHEET_NAMES.HASIL);

    const soalCountMap = {};
    bankSoal.forEach(s => {
      const m = String(s.mata_pelajaran || "").trim().toUpperCase();
      if (m) soalCountMap[m] = (soalCountMap[m] || 0) + 1;
    });

    const showScoreMap = {};
    pengaturanList.forEach(p => {
      const key = String(p.mata_pelajaran || "").trim().toUpperCase();
      const tampil = String(p.tampilkan_nilai || "YA").trim().toUpperCase();
      showScoreMap[key] = (tampil !== "TIDAK");
    });

    const hasilSiswa = hasilList.filter(h => String(h.nis).trim() === nisClean);

    const pengaturanSesuaiKelas = pengaturanList.filter(p =>
      isExamForStudentClass_(p.mata_pelajaran, kelasSiswa)
    );

    const daftarUjian = pengaturanSesuaiKelas.map(p => {
      const paketName = String(p.mata_pelajaran || "").trim();
      const paketKey = paketName.toUpperCase();
      const parsedInfo = splitMapelDanJenis_(paketName);
      const isShowScore = showScoreMap[paketKey] !== false;

      const riwayatMapel = hasilSiswa.filter(
        h => String(h.mata_pelajaran || "").trim().toUpperCase() === paketKey
      );
      const lastSubmission = riwayatMapel.length > 0 ? riwayatMapel[riwayatMapel.length - 1] : null;

      return {
        mata_pelajaran: paketName,
        mapel_induk: parsedInfo.mapel_induk,
        nama_mapel_murni: parsedInfo.nama_mapel_murni,
        tingkat_kelas: parsedInfo.tingkat_kelas,
        jenis_ujian: parsedInfo.jenis_ujian,
        status_ujian: String(p.status_ujian || "CLOSE").toUpperCase(),
        durasi_menit: parseInt(p.durasi_menit, 10) || 60,
        acak_soal: String(p.acak_soal || "TIDAK").toUpperCase(),
        acak_jawaban: String(p.acak_jawaban || "TIDAK").toUpperCase(),
        tampilkan_nilai: isShowScore ? "YA" : "TIDAK",
        jumlah_soal: soalCountMap[paketKey] || 0,
        status_pengerjaan: lastSubmission ? "SUDAH" : "BELUM",
        nilai_akhir: (lastSubmission && isShowScore) ? Number(lastSubmission.nilai_akhir) : null,
        waktu_submit: lastSubmission ? lastSubmission.timestamp : null
      };
    });

    const riwayatLengkap = hasilSiswa.map(h => {
      const paketName = String(h.mata_pelajaran || "").trim();
      const paketKey = paketName.toUpperCase();
      const parsedInfo = splitMapelDanJenis_(paketName);
      const isShowScore = showScoreMap[paketKey] !== false;

      return {
        timestamp: h.timestamp,
        nis: h.nis,
        nama_lengkap: h.nama_lengkap,
        kelas: h.kelas,
        mata_pelajaran: paketName,
        mapel_induk: parsedInfo.mapel_induk,
        tingkat_kelas: parsedInfo.tingkat_kelas,
        jenis_ujian: parsedInfo.jenis_ujian,
        tampilkan_nilai: isShowScore ? "YA" : "TIDAK",
        nilai_akhir: isShowScore ? Number(h.nilai_akhir) : null
      };
    }).reverse();

    return {
      status: "success",
      profil_siswa: siswaObj ? {
        nis: siswaObj.nis,
        nama_lengkap: siswaObj.nama_lengkap,
        kelas: siswaObj.kelas,
        tingkat: extractStudentGradeLevel_(siswaObj.kelas),
        username: siswaObj.username
      } : null,
      kelas_siswa: kelasSiswa,
      tingkat_siswa: extractStudentGradeLevel_(kelasSiswa),
      ujian: daftarUjian,
      riwayat: riwayatLengkap,
      website_settings: getPengaturanWebsite().settings
    };
  } catch (err) {
    return { status: "error", message: "Gagal memuat dashboard siswa: " + err.message };
  }
}

function getSoalUjian(mapel) {
  try {
    const mapelClean = String(mapel || "").trim();
    const mapelKey = mapelClean.toUpperCase();

    const pengaturanList = getSheetData_(SHEET_NAMES.PENGATURAN);
    const config = pengaturanList.find(
      p => String(p.mata_pelajaran || "").trim().toUpperCase() === mapelKey
    );

    if (!config) {
      return { status: "error", message: "Konfigurasi untuk paket penilaian ini belum diatur." };
    }

    if (String(config.status_ujian).toUpperCase() !== "OPEN") {
      return { status: "error", message: "Penilaian ini sedang DITUTUP (CLOSE)." };
    }

    const durasiMenit = parseInt(config.durasi_menit, 10) || 60;
    const isAcakSoal = String(config.acak_soal).toUpperCase() === "YA";
    const isAcakJawaban = String(config.acak_jawaban).toUpperCase() === "YA";
    const tampilkanNilai = String(config.tampilkan_nilai || "YA").toUpperCase();

    const allSoal = getSheetData_(SHEET_NAMES.SOAL);
    let soalMapel = allSoal.filter(
      s => String(s.mata_pelajaran || "").trim().toUpperCase() === mapelKey
    );

    if (soalMapel.length === 0) {
      return { status: "error", message: "Belum ada butir soal untuk " + mapelClean };
    }

    if (isAcakSoal) {
      soalMapel = shuffleArray_(soalMapel);
    }

    const sanitizedSoal = soalMapel.map((item, idx) => {
      const optionKeys = ["A", "B", "C", "D", "E"];
      let options = [];

      optionKeys.forEach(k => {
        const colName = "opsi_" + k.toLowerCase();
        const teksOpsi = item[colName] !== undefined ? String(item[colName]).trim() : "";
        if (teksOpsi !== "") {
          options.push({
            key: k,
            text: teksOpsi
          });
        }
      });

      if (isAcakJawaban) {
        options = shuffleArray_(options);
      }

      return {
        no: idx + 1,
        id_soal: item.id_soal,
        mata_pelajaran: item.mata_pelajaran,
        pertanyaan: item.pertanyaan,
        url_gambar: formatMediaUrl_(item.id_gambar, "image"),
        url_suara: formatMediaUrl_(item.id_suara, "audio"),
        opsi: options
      };
    });

    const parsedInfo = splitMapelDanJenis_(config.mata_pelajaran);

    return {
      status: "success",
      mata_pelajaran: config.mata_pelajaran,
      mapel_induk: parsedInfo.mapel_induk,
      tingkat_kelas: parsedInfo.tingkat_kelas,
      jenis_ujian: parsedInfo.jenis_ujian,
      durasi_menit: durasiMenit,
      tampilkan_nilai: tampilkanNilai,
      total_soal: sanitizedSoal.length,
      soal: sanitizedSoal
    };
  } catch (err) {
    return { status: "error", message: "Gagal mengambil soal ujian: " + err.message };
  }
}

function submitUjian(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const nis = String(payload.nis || "").trim();
    const nama = String(payload.nama_lengkap || "").trim();
    const kelas = String(payload.kelas || "").trim();
    const mapel = String(payload.mata_pelajaran || "").trim();
    const jawabanSiswa = payload.jawaban_siswa || {};

    if (!nis || !mapel) {
      return { status: "error", message: "Data identitas siswa atau paket penilaian tidak lengkap." };
    }

    const allSoal = getSheetData_(SHEET_NAMES.SOAL);
    const soalMapel = allSoal.filter(
      s => String(s.mata_pelajaran || "").trim().toUpperCase() === mapel.toUpperCase()
    );

    if (soalMapel.length === 0) {
      return { status: "error", message: "Bank soal tidak ditemukan untuk penilaian." };
    }

    const pengaturanList = getSheetData_(SHEET_NAMES.PENGATURAN);
    const config = pengaturanList.find(
      p => String(p.mata_pelajaran || "").trim().toUpperCase() === mapel.toUpperCase()
    );
    const tampilkanNilai = config
      ? (String(config.tampilkan_nilai || "YA").trim().toUpperCase() !== "TIDAK")
      : true;

    let jumlahBenar = 0;
    const totalSoal = soalMapel.length;

    soalMapel.forEach(s => {
      const idSoal = String(s.id_soal).trim();
      const kunci = String(s.kunci_jawaban || "").trim().toUpperCase();
      const jawab = String(jawabanSiswa[idSoal] || "").trim().toUpperCase();

      if (kunci && jawab && jawab === kunci) {
        jumlahBenar++;
      }
    });

    const nilaiAkhir = Math.round((jumlahBenar / totalSoal) * 100);
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || "Asia/Jakarta",
      "yyyy-MM-dd HH:mm:ss"
    );

    const ss = getSpreadsheet_(false);
    ensureHasilUjianSchema_(ss);
    const sheetHasil = ss.getSheetByName(SHEET_NAMES.HASIL);

    let rawCatatan = payload.catatan_kecurangan || [];
    if (typeof rawCatatan === "string") {
      try {
        rawCatatan = JSON.parse(rawCatatan);
      } catch (e) {
        rawCatatan = [];
      }
    }
    if (!Array.isArray(rawCatatan)) rawCatatan = [];
    const catatanJson = JSON.stringify(rawCatatan);

    sheetHasil.appendRow([
      timestamp,
      nis,
      nama,
      kelas,
      mapel,
      JSON.stringify(jawabanSiswa),
      nilaiAkhir,
      catatanJson
    ]);

    try {
      if (
        EMAIL_GURU_NOTIFIKASI &&
        EMAIL_GURU_NOTIFIKASI !== "GANTI_DENGAN_EMAIL_NOTIFIKASI@gmail.com" &&
        EMAIL_GURU_NOTIFIKASI.includes("@")
      ) {
        const subject = "[Hasil Penilaian] " + nama + " (" + kelas + ") - " + mapel;
        const body =
          "Halo Bapak/Ibu Guru,\n\n" +
          "Berikut adalah ringkasan hasil pengerjaan siswa:\n\n" +
          "Waktu Submit     : " + timestamp + "\n" +
          "NIS              : " + nis + "\n" +
          "Nama Lengkap     : " + nama + "\n" +
          "Kelas            : " + kelas + "\n" +
          "Paket Penilaian  : " + mapel + "\n" +
          "Jawaban Benar    : " + jumlahBenar + " dari " + totalSoal + " soal\n" +
          "Nilai Akhir      : " + nilaiAkhir + " / 100\n\n" +
          "Sistem Ujian Online Sekolah Digital V2.5";
        MailApp.sendEmail(EMAIL_GURU_NOTIFIKASI, subject, body);
      }
    } catch (emailErr) {
      console.warn("Gagal mengirim email notifikasi: " + emailErr.message);
    }

    return {
      status: "success",
      message: "Jawaban berhasil disimpan!",
      hasil: {
        timestamp: timestamp,
        nis: nis,
        nama_lengkap: nama,
        kelas: kelas,
        mata_pelajaran: mapel,
        tampilkan_nilai: tampilkanNilai ? "YA" : "TIDAK",
        benar: tampilkanNilai ? jumlahBenar : null,
        salah: tampilkanNilai ? (totalSoal - jumlahBenar) : null,
        total_soal: totalSoal,
        nilai_akhir: tampilkanNilai ? nilaiAkhir : null
      }
    };
  } catch (err) {
    return { status: "error", message: "Gagal menyimpan ujian: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function getDetailRiwayatUjian(nis, mapel) {
  try {
    const nisClean = String(nis || "").trim();
    const mapelClean = String(mapel || "").trim();
    const mapelKey = mapelClean.toUpperCase();

    const pengaturanList = getSheetData_(SHEET_NAMES.PENGATURAN);
    const config = pengaturanList.find(
      p => String(p.mata_pelajaran || "").trim().toUpperCase() === mapelKey
    );
    if (config && String(config.tampilkan_nilai || "YA").trim().toUpperCase() === "TIDAK") {
      return {
        status: "error",
        message: "Nilai dan pembahasan untuk penilaian ini belum ditampilkan oleh Guru."
      };
    }

    const hasilList = getSheetData_(SHEET_NAMES.HASIL);
    const rekamanSiswa = hasilList.filter(
      h =>
        String(h.nis).trim() === nisClean &&
        String(h.mata_pelajaran || "").trim().toUpperCase() === mapelKey
    );

    if (rekamanSiswa.length === 0) {
      return { status: "error", message: "Data riwayat ujian tidak ditemukan." };
    }

    const latestRecord = rekamanSiswa[rekamanSiswa.length - 1];
    let jawabanMap = {};
    try {
      jawabanMap = JSON.parse(latestRecord.jawaban_siswa || "{}");
    } catch (e) {
      jawabanMap = {};
    }

    const allSoal = getSheetData_(SHEET_NAMES.SOAL);
    const soalMapel = allSoal.filter(
      s => String(s.mata_pelajaran || "").trim().toUpperCase() === mapelKey
    );

    let benarCount = 0;
    const detailPembahasan = soalMapel.map((s, idx) => {
      const idSoal = String(s.id_soal).trim();
      const kunci = String(s.kunci_jawaban || "").trim().toUpperCase();
      const jawabanSiswaKey = String(jawabanMap[idSoal] || "").trim().toUpperCase();
      const isBenar = Boolean(kunci && jawabanSiswaKey === kunci);
      if (isBenar) benarCount++;

      const opsiMap = {
        A: s.opsi_a || "",
        B: s.opsi_b || "",
        C: s.opsi_c || "",
        D: s.opsi_d || "",
        E: s.opsi_e || ""
      };

      return {
        no: idx + 1,
        id_soal: idSoal,
        pertanyaan: s.pertanyaan,
        opsi: opsiMap,
        jawaban_siswa: jawabanSiswaKey || "-",
        teks_jawaban_siswa: jawabanSiswaKey ? (opsiMap[jawabanSiswaKey] || "") : "Tidak Dijawab",
        kunci_jawaban: kunci,
        teks_kunci_jawaban: kunci ? (opsiMap[kunci] || "") : "",
        status_benar: isBenar
      };
    });

    return {
      status: "success",
      ringkasan: {
        timestamp: latestRecord.timestamp,
        nis: latestRecord.nis,
        nama_lengkap: latestRecord.nama_lengkap,
        kelas: latestRecord.kelas,
        mata_pelajaran: latestRecord.mata_pelajaran,
        nilai_akhir: Number(latestRecord.nilai_akhir),
        total_soal: soalMapel.length,
        benar: benarCount,
        salah: soalMapel.length - benarCount
      },
      detail: detailPembahasan
    };
  } catch (err) {
    return { status: "error", message: "Gagal memuat detail pembahasan: " + err.message };
  }
}

/**
 * ============================================================================
 * 3. MODUL GURU (ISOLASI KETAT BERDASARKAN MAPEL + TINGKAT KELAS AMPUAN)
 * ============================================================================
 */
function getGuruProfile_(guruIdentifier) {
  const dataGuru = getSheetData_(SHEET_NAMES.GURU);
  if (!guruIdentifier) return null;

  const cleanId = String(guruIdentifier).trim().toLowerCase();
  return dataGuru.find(
    g =>
      String(g.username || "").trim().toLowerCase() === cleanId ||
      String(g.nip_atau_id || "").trim().toLowerCase() === cleanId
  ) || null;
}

function getAllowedMapelForGuru_(guruIdentifier) {
  const found = getGuruProfile_(guruIdentifier);
  if (!found) return [];
  return parseGuruMapelList_(found.mata_pelajaran);
}

function getDashboardGuru(guruIdentifier) {
  try {
    const guruObj = getGuruProfile_(guruIdentifier);
    if (!guruObj) {
      return {
        status: "error",
        message: "Akun guru tidak ditemukan di sheet DATA_GURU."
      };
    }

    const allowedMapel = parseGuruMapelList_(guruObj.mata_pelajaran);
    if (allowedMapel.length === 0) {
      return {
        status: "error",
        message: "Akun guru belum memiliki mata pelajaran ampuan di sheet DATA_GURU."
      };
    }

    const guruUsername = String(guruObj.username || guruObj.nip_atau_id || "").trim();
    const guruUserLower = guruUsername.toLowerCase();
    const guruNipLower = String(guruObj.nip_atau_id || "").trim().toLowerCase();

    const masterMapelRows = getSheetData_(SHEET_NAMES.MASTER_MAPEL);
    let masterJenisRows = getSheetData_(SHEET_NAMES.MASTER_JENIS);
    const allSoal = getSheetData_(SHEET_NAMES.SOAL);
    const allPengaturan = getSheetData_(SHEET_NAMES.PENGATURAN);
    const allHasil = getSheetData_(SHEET_NAMES.HASIL);
    const allSiswa = getSheetData_(SHEET_NAMES.SISWA);
    const allKelas = getSheetData_(SHEET_NAMES.KELAS);

    // Ambil daftar siswa & rombel kelas yang berelasi dengan mapel ampuan guru ini
    const siswaAmpuan = allSiswa
      .filter(s => allowedMapel.some(mp => isExamForStudentClass_(mp, s.kelas)))
      .map(s => ({
        nis: String(s.nis || "").trim(),
        nama_lengkap: String(s.nama_lengkap || "").trim(),
        kelas: String(s.kelas || "").trim().toUpperCase()
      }));

    const kelasAmpuanSet = new Set();
    allKelas.forEach(k => {
      const namaKls = String(k.nama_kelas || "").trim().toUpperCase();
      if (namaKls && allowedMapel.some(mp => isExamForStudentClass_(mp, namaKls))) {
        kelasAmpuanSet.add(namaKls);
      }
    });
    siswaAmpuan.forEach(s => { if (s.kelas) kelasAmpuanSet.add(s.kelas); });

    const bankSoalGuru = allSoal
      .filter(s => isAllowedForGuru_(s.mata_pelajaran, allowedMapel))
      .map(s => {
        const info = splitMapelDanJenis_(s.mata_pelajaran);
        return Object.assign({}, s, {
          mapel_induk: info.mapel_induk,
          nama_mapel_murni: info.nama_mapel_murni,
          tingkat_kelas: info.tingkat_kelas,
          jenis_ujian: info.jenis_ujian
        });
      });

    const pengaturanGuru = allPengaturan
      .filter(p => isAllowedForGuru_(p.mata_pelajaran, allowedMapel))
      .map(p => {
        const info = splitMapelDanJenis_(p.mata_pelajaran);
        return Object.assign({}, p, {
          mapel_induk: info.mapel_induk,
          nama_mapel_murni: info.nama_mapel_murni,
          tingkat_kelas: info.tingkat_kelas,
          jenis_ujian: info.jenis_ujian,
          tampilkan_nilai: String(p.tampilkan_nilai || "YA").toUpperCase() === "TIDAK" ? "TIDAK" : "YA"
        });
      });

    const hasilGuru = allHasil
      .filter(h => isAllowedForGuru_(h.mata_pelajaran, allowedMapel))
      .map(h => {
        const info = splitMapelDanJenis_(h.mata_pelajaran);
        let listKecurangan = [];
        try {
          const parsedLog = JSON.parse(h.catatan_kecurangan || "[]");
          if (Array.isArray(parsedLog)) listKecurangan = parsedLog;
        } catch (e) {
          listKecurangan = [];
        }
        return Object.assign({}, h, {
          mapel_induk: info.mapel_induk,
          nama_mapel_murni: info.nama_mapel_murni,
          tingkat_kelas: info.tingkat_kelas,
          jenis_ujian: info.jenis_ujian,
          pelanggaran_list: listKecurangan
        });
      })
      .reverse();

    // Filter MASTER_JENIS_PENILAIAN secara KETAT hanya milik guru yang sedang login
    let masterJenisGuru = masterJenisRows.filter(j => {
      const owner = String(j.pemilik_guru || "").trim().toLowerCase();
      if (!owner) return false;
      return owner === guruUserLower || (guruNipLower && owner === guruNipLower);
    });

    // Jika guru baru belum punya jenis penilaian sama sekali, buatkan jenis awal khusus untuk guru ini
    if (masterJenisGuru.length === 0) {
      const ss = getSpreadsheet_(false);
      const sheetJenis = ss.getSheetByName(SHEET_NAMES.MASTER_JENIS);
      const nextNumBase = sheetJenis ? sheetJenis.getLastRow() : 1;
      let addedCount = 0;

      allowedMapel.forEach(mp => {
        addedCount++;
        const kode = "JP-" + String(nextNumBase + addedCount).padStart(2, "0") + "-" + Math.floor(10 + Math.random() * 89);
        const newRow = [kode, "Ulangan Harian 1", "Penilaian Harian " + mp, guruUsername, mp];
        if (sheetJenis) sheetJenis.appendRow(newRow);
        masterJenisGuru.push({
          kode_jenis: kode,
          nama_jenis: "Ulangan Harian 1",
          keterangan: "Penilaian Harian " + mp,
          pemilik_guru: guruUsername,
          mapel_relasi: mp
        });
      });
    }

    const paketSet = new Set();
    bankSoalGuru.forEach(s => { if (s.mata_pelajaran) paketSet.add(String(s.mata_pelajaran).trim()); });
    pengaturanGuru.forEach(p => { if (p.mata_pelajaran) paketSet.add(String(p.mata_pelajaran).trim()); });
    hasilGuru.forEach(h => { if (h.mata_pelajaran) paketSet.add(String(h.mata_pelajaran).trim()); });

    const jenisSet = new Set();
    masterJenisGuru.forEach(j => { if (j.nama_jenis) jenisSet.add(String(j.nama_jenis).trim()); });
    bankSoalGuru.forEach(s => { if (s.jenis_ujian) jenisSet.add(String(s.jenis_ujian).trim()); });
    pengaturanGuru.forEach(p => { if (p.jenis_ujian) jenisSet.add(String(p.jenis_ujian).trim()); });

    const ss = getSpreadsheet_(false);
    ensureGuruMapelInMasterMapel_(ss, allowedMapel);
    const updatedMasterMapelRows = getSheetData_(SHEET_NAMES.MASTER_MAPEL);

    const masterMapelGuru = [];
    allowedMapel.forEach((am, idx) => {
      const parsedAm = splitMapelDanJenis_(am);
      const amGradeNorm = normalizeGradeLevel_(parsedAm.tingkat_kelas);
      const foundMm = updatedMasterMapelRows.find(m => {
        const label = String(m.label_mapel_kelas || m.nama_mapel || "").trim().toUpperCase();
        if (label === am.toUpperCase()) return true;
        const mParsed = splitMapelDanJenis_(label);
        const mGrade = normalizeGradeLevel_(m.kelas || mParsed.tingkat_kelas);
        return (
          mParsed.nama_mapel_murni.toUpperCase() === parsedAm.nama_mapel_murni.toUpperCase() &&
          mGrade === amGradeNorm
        );
      });
      if (foundMm) {
        masterMapelGuru.push(Object.assign({}, foundMm, {
          kelas: normalizeGradeLevel_(foundMm.kelas || parsedAm.tingkat_kelas),
          label_mapel_kelas: am
        }));
      } else {
        masterMapelGuru.push({
          kode_mapel: "MP-" + amGradeNorm + "-" + (idx + 1),
          nama_mapel: parsedAm.nama_mapel_murni,
          kelas: amGradeNorm,
          label_mapel_kelas: am,
          kelompok: "Mapel Ampuan Guru"
        });
      }
    });

    hasilGuru.forEach(h => {
      if (h.kelas) kelasAmpuanSet.add(String(h.kelas).trim().toUpperCase());
    });

    return {
      status: "success",
      allowed_mapel: allowedMapel,
      master_mapel_guru: masterMapelGuru,
      master_jenis_penilaian: masterJenisGuru,
      jenis_list: Array.from(jenisSet),
      paket_list: Array.from(paketSet).sort(),
      bank_soal: bankSoalGuru,
      pengaturan_ujian: pengaturanGuru,
      hasil_ujian: hasilGuru,
      siswa_ampuan: siswaAmpuan,
      kelas_ampuan: Array.from(kelasAmpuanSet).sort(),
      website_settings: getPengaturanWebsite().settings
    };
  } catch (err) {
    return { status: "error", message: "Gagal memuat data dashboard guru: " + err.message };
  }
}

function simpanSoalBatch(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const mapelInduk = String(payload.mapel_induk || "").trim();
    const jenisUjian = String(payload.jenis_ujian || "").trim();
    const paketLengkap = mapelInduk && jenisUjian
      ? (mapelInduk + " - " + jenisUjian)
      : String(payload.mata_pelajaran || "").trim();

    const daftarSoal = payload.daftar_soal || [];
    if (!paketLengkap || !Array.isArray(daftarSoal) || daftarSoal.length === 0) {
      return { status: "error", message: "Pilih Mata Pelajaran (Kelas), Jenis Penilaian, dan isi minimal 1 soal!" };
    }

    if (payload.guru_username) {
      const allowedMapel = getAllowedMapelForGuru_(payload.guru_username);
      if (!isAllowedForGuru_(paketLengkap, allowedMapel)) {
        return {
          status: "error",
          message: "Akses Ditolak! Anda hanya dapat mengelola mata pelajaran & tingkat kelas yang Anda ampu: " + allowedMapel.join(", ")
        };
      }
    }

    const parsedInfo = splitMapelDanJenis_(paketLengkap);
    const baseTime = new Date().getTime();
    const rowsToInsert = [];
    const createdObjects = [];

    for (let i = 0; i < daftarSoal.length; i++) {
      const item = daftarSoal[i];
      const pertanyaan = String(item.pertanyaan || "").trim();
      const idGambar = String(item.id_gambar || "").trim();
      const idSuara = String(item.id_suara || "").trim();
      const opsiA = String(item.opsi_a || "").trim();
      const opsiB = String(item.opsi_b || "").trim();
      const opsiC = String(item.opsi_c || "").trim();
      const opsiD = String(item.opsi_d || "").trim();
      const opsiE = String(item.opsi_e || "").trim();
      const kunci = String(item.kunci_jawaban || "A").trim().toUpperCase();

      if (!pertanyaan || !opsiA || !opsiB) {
        return {
          status: "error",
          message: "Baris Soal ke-" + (i + 1) + " belum lengkap (Pertanyaan, Opsi A, dan Opsi B wajib diisi)!"
        };
      }

      const newId = "SOAL-" + (baseTime + i) + "-" + Math.floor(100 + Math.random() * 900);
      rowsToInsert.push([
        newId,
        paketLengkap,
        pertanyaan,
        idGambar,
        idSuara,
        opsiA,
        opsiB,
        opsiC,
        opsiD,
        opsiE,
        kunci
      ]);

      createdObjects.push({
        id_soal: newId,
        mata_pelajaran: paketLengkap,
        mapel_induk: parsedInfo.mapel_induk,
        nama_mapel_murni: parsedInfo.nama_mapel_murni,
        tingkat_kelas: parsedInfo.tingkat_kelas,
        jenis_ujian: parsedInfo.jenis_ujian,
        pertanyaan: pertanyaan,
        id_gambar: idGambar,
        id_suara: idSuara,
        opsi_a: opsiA,
        opsi_b: opsiB,
        opsi_c: opsiC,
        opsi_d: opsiD,
        opsi_e: opsiE,
        kunci_jawaban: kunci
      });
    }

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.SOAL);
    const startRow = sheet.getLastRow() + 1;

    sheet.getRange(startRow, 1, rowsToInsert.length, 11).setValues(rowsToInsert);
    const isNewPengaturan = ensureMapelInPengaturan_(ss, paketLengkap);

    return {
      status: "success",
      message: rowsToInsert.length + " butir soal berhasil disimpan ke [" + paketLengkap + "]!",
      inserted_soal: createdObjects,
      new_pengaturan: isNewPengaturan ? {
        mata_pelajaran: paketLengkap,
        mapel_induk: parsedInfo.mapel_induk,
        nama_mapel_murni: parsedInfo.nama_mapel_murni,
        tingkat_kelas: parsedInfo.tingkat_kelas,
        jenis_ujian: parsedInfo.jenis_ujian,
        status_ujian: "CLOSE",
        durasi_menit: 45,
        acak_soal: "YA",
        acak_jawaban: "YA",
        tampilkan_nilai: "YA"
      } : null
    };
  } catch (err) {
    return { status: "error", message: "Gagal menyimpan soal batch: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function simpanSoal(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const idSoalInput = String(payload.id_soal || "").trim();
    if (!idSoalInput) {
      return simpanSoalBatch({
        guru_username: payload.guru_username,
        mapel_induk: payload.mapel_induk,
        jenis_ujian: payload.jenis_ujian,
        mata_pelajaran: payload.mata_pelajaran,
        daftar_soal: [payload]
      });
    }

    const mapelInduk = String(payload.mapel_induk || "").trim();
    const jenisUjian = String(payload.jenis_ujian || "").trim();
    const paketLengkap = mapelInduk && jenisUjian
      ? (mapelInduk + " - " + jenisUjian)
      : String(payload.mata_pelajaran || "").trim();

    if (payload.guru_username) {
      const allowedMapel = getAllowedMapelForGuru_(payload.guru_username);
      if (!isAllowedForGuru_(paketLengkap, allowedMapel)) {
        return { status: "error", message: "Akses Ditolak! Anda tidak mengampu kelas/mapel ini." };
      }
    }

    const parsedInfo = splitMapelDanJenis_(paketLengkap);
    const pertanyaan = String(payload.pertanyaan || "").trim();
    const idGambar = String(payload.id_gambar || "").trim();
    const idSuara = String(payload.id_suara || "").trim();
    const opsiA = String(payload.opsi_a || "").trim();
    const opsiB = String(payload.opsi_b || "").trim();
    const opsiC = String(payload.opsi_c || "").trim();
    const opsiD = String(payload.opsi_d || "").trim();
    const opsiE = String(payload.opsi_e || "").trim();
    const kunci = String(payload.kunci_jawaban || "A").trim().toUpperCase();

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.SOAL);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "error", message: "Soal tidak ditemukan." };

    const ids = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === idSoalInput) {
        if (payload.guru_username) {
          const allowedMapel = getAllowedMapelForGuru_(payload.guru_username);
          if (!isAllowedForGuru_(ids[i][1], allowedMapel)) {
            return { status: "error", message: "Akses Ditolak! Soal ini milik kelas/guru lain." };
          }
        }

        const rowNum = i + 2;
        sheet.getRange(rowNum, 1, 1, 11).setValues([[
          idSoalInput,
          paketLengkap,
          pertanyaan,
          idGambar,
          idSuara,
          opsiA,
          opsiB,
          opsiC,
          opsiD,
          opsiE,
          kunci
        ]]);
        ensureMapelInPengaturan_(ss, paketLengkap);
        return {
          status: "success",
          message: "Soal (" + idSoalInput + ") berhasil diperbarui!",
          updated_soal: {
            id_soal: idSoalInput,
            mata_pelajaran: paketLengkap,
            mapel_induk: parsedInfo.mapel_induk,
            nama_mapel_murni: parsedInfo.nama_mapel_murni,
            tingkat_kelas: parsedInfo.tingkat_kelas,
            jenis_ujian: parsedInfo.jenis_ujian,
            pertanyaan: pertanyaan,
            id_gambar: idGambar,
            id_suara: idSuara,
            opsi_a: opsiA,
            opsi_b: opsiB,
            opsi_c: opsiC,
            opsi_d: opsiD,
            opsi_e: opsiE,
            kunci_jawaban: kunci
          }
        };
      }
    }

    return { status: "error", message: "ID Soal tidak ditemukan." };
  } catch (err) {
    return { status: "error", message: "Gagal memperbarui soal: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function ensureMapelInPengaturan_(ss, paketLengkap) {
  const sheetPengaturan = ss.getSheetByName(SHEET_NAMES.PENGATURAN);
  const lastRow = sheetPengaturan.getLastRow();
  if (lastRow > 1) {
    const colA = sheetPengaturan.getRange(2, 1, lastRow - 1, 1).getValues();
    const exists = colA.some(r => String(r[0] || "").trim().toUpperCase() === paketLengkap.toUpperCase());
    if (exists) return false;
  }
  sheetPengaturan.appendRow([paketLengkap, "CLOSE", 45, "YA", "YA", "YA"]);
  return true;
}

function hapusSoal(id_soal, guruUsername) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const targetId = String(id_soal || "").trim();
    if (!targetId) return { status: "error", message: "ID Soal tidak valid." };

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.SOAL);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "error", message: "Bank soal kosong." };

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).trim() === targetId) {
        if (guruUsername) {
          const allowedMapel = getAllowedMapelForGuru_(guruUsername);
          if (!isAllowedForGuru_(data[i][1], allowedMapel)) {
            return { status: "error", message: "Akses Ditolak! Soal ini milik kelas/guru lain." };
          }
        }
        sheet.deleteRow(i + 2);
        return { status: "success", message: "Soal (" + targetId + ") berhasil dihapus." };
      }
    }

    return { status: "error", message: "Soal tidak ditemukan." };
  } catch (err) {
    return { status: "error", message: "Gagal menghapus soal: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function updatePengaturanUjian(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    let paketLengkap = String(payload.mata_pelajaran || "").trim();
    if (payload.mapel_induk && payload.jenis_ujian) {
      paketLengkap = String(payload.mapel_induk).trim() + " - " + String(payload.jenis_ujian).trim();
    }

    if (!paketLengkap) {
      return { status: "error", message: "Mata pelajaran (kelas) dan jenis penilaian wajib dipilih!" };
    }

    if (payload.guru_username) {
      const allowedMapel = getAllowedMapelForGuru_(payload.guru_username);
      if (!isAllowedForGuru_(paketLengkap, allowedMapel)) {
        return { status: "error", message: "Akses Ditolak! Anda tidak mengampu kelas/mapel ini." };
      }
    }

    const statusUjian = String(payload.status_ujian || "CLOSE").trim().toUpperCase();
    const durasiMenit = Math.max(1, parseInt(payload.durasi_menit, 10) || 45);
    const acakSoal = String(payload.acak_soal || "TIDAK").trim().toUpperCase();
    const acakJawaban = String(payload.acak_jawaban || "TIDAK").trim().toUpperCase();
    const tampilkanNilai = String(payload.tampilkan_nilai || "YA").trim().toUpperCase();

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.PENGATURAN);
    const lastRow = sheet.getLastRow();

    const parsed = splitMapelDanJenis_(paketLengkap);
    const updatedObj = {
      mata_pelajaran: paketLengkap,
      mapel_induk: parsed.mapel_induk,
      nama_mapel_murni: parsed.nama_mapel_murni,
      tingkat_kelas: parsed.tingkat_kelas,
      jenis_ujian: parsed.jenis_ujian,
      status_ujian: statusUjian,
      durasi_menit: String(durasiMenit),
      acak_soal: acakSoal,
      acak_jawaban: acakJawaban,
      tampilkan_nilai: tampilkanNilai
    };

    if (lastRow > 1) {
      const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < colA.length; i++) {
        if (String(colA[i][0]).trim().toUpperCase() === paketLengkap.toUpperCase()) {
          const rowNum = i + 2;
          sheet.getRange(rowNum, 1, 1, 6).setValues([[
            paketLengkap,
            statusUjian,
            durasiMenit,
            acakSoal,
            acakJawaban,
            tampilkanNilai
          ]]);
          return {
            status: "success",
            message: "Pengaturan [" + paketLengkap + "] berhasil disimpan!",
            pengaturan: updatedObj
          };
        }
      }
    }

    sheet.appendRow([paketLengkap, statusUjian, durasiMenit, acakSoal, acakJawaban, tampilkanNilai]);
    return {
      status: "success",
      message: "Paket penilaian baru [" + paketLengkap + "] berhasil ditambahkan!",
      pengaturan: updatedObj
    };
  } catch (err) {
    return { status: "error", message: "Gagal menyimpan pengaturan ujian: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function tambahMasterJenisPenilaian(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const namaJenis = String(payload.nama_jenis || "").trim();
    const keterangan = String(payload.keterangan || "Penilaian Akademik").trim();
    const guruInput = String(payload.guru_username || payload.pemilik_guru || "").trim();
    let mapelRelasi = String(payload.mapel_relasi || "").trim();

    if (!namaJenis) return { status: "error", message: "Nama Jenis Penilaian wajib diisi!" };
    if (!guruInput) return { status: "error", message: "Identitas guru tidak terdeteksi. Silakan login ulang." };

    const guruObj = getGuruProfile_(guruInput);
    if (!guruObj) {
      return { status: "error", message: "Akun guru tidak ditemukan di DATA_GURU." };
    }

    const allowedMapel = parseGuruMapelList_(guruObj.mata_pelajaran);
    const guruUsername = String(guruObj.username || guruObj.nip_atau_id || guruInput).trim();
    const guruUserLower = guruUsername.toLowerCase();
    const guruNipLower = String(guruObj.nip_atau_id || "").trim().toLowerCase();

    if (!mapelRelasi || mapelRelasi.toUpperCase() === "SEMUA") {
      mapelRelasi = allowedMapel.length === 1 ? allowedMapel[0] : "SEMUA";
    } else {
      const isAllowed = allowedMapel.some(am => am.toUpperCase() === mapelRelasi.toUpperCase());
      if (!isAllowed) {
        return {
          status: "error",
          message: "Akses Ditolak! Mata pelajaran relasi tersebut bukan ampuan Anda."
        };
      }
    }

    const ss = getSpreadsheet_(false);
    ensureMasterJenisSchemaAndIsolation_(ss);
    const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_JENIS);
    const lastRow = sheet.getLastRow();
    const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];

    // Cek duplikasi HANYA pada jenis penilaian milik guru ini sendiri
    const existsForThisGuru = rows.some(r => {
      const rowNama = String(r[1] || "").trim().toUpperCase();
      const rowOwner = String(r[3] || "").trim().toLowerCase();
      const rowMapel = String(r[4] || "SEMUA").trim().toUpperCase();
      const isSameOwner = (rowOwner === guruUserLower || (guruNipLower && rowOwner === guruNipLower));
      const isSameMapel = (rowMapel === mapelRelasi.toUpperCase() || rowMapel === "SEMUA" || mapelRelasi.toUpperCase() === "SEMUA");
      return isSameOwner && rowNama === namaJenis.toUpperCase() && isSameMapel;
    });

    if (existsForThisGuru) {
      return {
        status: "error",
        message: "Jenis Penilaian \"" + namaJenis + "\" sudah terdaftar pada daftar evaluasi Anda!"
      };
    }

    // Generate kode unik yang aman
    let maxNum = rows.length;
    rows.forEach(r => {
      const m = String(r[0] || "").match(/^JP-(\d+)/i);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > maxNum) maxNum = n;
      }
    });
    const kodeBaru = "JP-" + String(maxNum + 1).padStart(2, "0");

    sheet.appendRow([kodeBaru, namaJenis, keterangan, guruUsername, mapelRelasi]);

    return {
      status: "success",
      message: "Jenis Penilaian \"" + namaJenis + "\" berhasil ditambahkan khusus untuk akun Anda!",
      item: {
        kode_jenis: kodeBaru,
        nama_jenis: namaJenis,
        keterangan: keterangan,
        pemilik_guru: guruUsername,
        mapel_relasi: mapelRelasi
      }
    };
  } catch (err) {
    return { status: "error", message: "Gagal menambah jenis penilaian: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function hapusMasterJenisPenilaian(kodeJenis, guruIdentifier) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const target = String(kodeJenis || "").trim();
    const ss = getSpreadsheet_(false);
    ensureMasterJenisSchemaAndIsolation_(ss);
    const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_JENIS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: "error", message: "Data kosong." };

    const guruObj = guruIdentifier ? getGuruProfile_(guruIdentifier) : null;
    const guruUserLower = guruObj ? String(guruObj.username || "").trim().toLowerCase() : String(guruIdentifier || "").trim().toLowerCase();
    const guruNipLower = guruObj ? String(guruObj.nip_atau_id || "").trim().toLowerCase() : "";

    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).trim() === target) {
        const rowOwner = String(data[i][3] || "").trim().toLowerCase();
        if (guruUserLower && rowOwner && rowOwner !== guruUserLower && rowOwner !== guruNipLower) {
          return {
            status: "error",
            message: "Akses Ditolak! Jenis penilaian ini milik guru lain dan tidak dapat Anda hapus."
          };
        }
        sheet.deleteRow(i + 2);
        return { status: "success", message: "Jenis penilaian (" + target + ") berhasil dihapus dari daftar Anda." };
      }
    }
    return { status: "error", message: "Data jenis penilaian tidak ditemukan." };
  } catch (err) {
    return { status: "error", message: "Gagal menghapus jenis penilaian: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function hapusHasilUjianSiswa(timestamp, nis, mapel, guruUsername) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const tsClean = String(timestamp || "").trim();
    const nisClean = String(nis || "").trim();
    const mapelClean = String(mapel || "").trim();

    if (guruUsername) {
      const allowedMapel = getAllowedMapelForGuru_(guruUsername);
      if (!isAllowedForGuru_(mapelClean, allowedMapel)) {
        return { status: "error", message: "Akses Ditolak! Nilai ini milik kelas/guru lain." };
      }
    }

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.HASIL);
    const data = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone() || "Asia/Jakarta";

    for (let i = data.length - 1; i >= 1; i--) {
      let rowTs = data[i][0];
      if (rowTs instanceof Date) {
        rowTs = Utilities.formatDate(rowTs, tz, "yyyy-MM-dd HH:mm:ss");
      } else {
        rowTs = String(rowTs).trim();
      }

      const rowNis = String(data[i][1]).trim();
      const rowMapel = String(data[i][4]).trim().toUpperCase();

      if (rowNis === nisClean && rowMapel === mapelClean.toUpperCase() && (!tsClean || rowTs === tsClean)) {
        sheet.deleteRow(i + 1);
        return {
          status: "success",
          message: "Hasil ujian siswa (NIS: " + nisClean + ") berhasil direset."
        };
      }
    }

    return { status: "error", message: "Data hasil ujian siswa tidak ditemukan." };
  } catch (err) {
    return { status: "error", message: "Gagal mereset hasil ujian: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * 4. MODUL ADMINISTRATOR (MANAJEMEN DATA GURU, TABEL KELAS, & DATA SISWA)
 * ============================================================================
 */
function getDataAdminDashboard() {
  try {
    const ss = getSpreadsheet_(false);
    const rawGuru = getSheetData_(SHEET_NAMES.GURU);
    const rawSiswa = getSheetData_(SHEET_NAMES.SISWA);
    const rawKelas = getSheetData_(SHEET_NAMES.KELAS);

    // Pastikan seluruh mapel ampuan dari semua guru terdaftar di MASTER_MAPEL
    const allGuruMapels = [];
    rawGuru.forEach(g => {
      parseGuruMapelList_(g.mata_pelajaran).forEach(mp => allGuruMapels.push(mp));
    });
    ensureGuruMapelInMasterMapel_(ss, allGuruMapels);
    const masterMapelRaw = getSheetData_(SHEET_NAMES.MASTER_MAPEL);
    const masterMapel = masterMapelRaw.map(m => {
      const namaMapel = String(m.nama_mapel || "").trim();
      const kelasNorm = normalizeGradeLevel_(m.kelas || m.label_mapel_kelas) || "10";
      const labelMapel = String(m.label_mapel_kelas || "").trim() || (namaMapel ? namaMapel + " (Kelas " + kelasNorm + ")" : "");
      return {
        kode_mapel: String(m.kode_mapel || "").trim(),
        nama_mapel: namaMapel,
        kelas: kelasNorm,
        label_mapel_kelas: labelMapel,
        kelompok: String(m.kelompok || "Umum / Wajib").trim() || "Umum / Wajib"
      };
    });

    const dataGuru = rawGuru.map(g => ({
      nip_atau_id: String(g.nip_atau_id || "").trim(),
      nama_lengkap: String(g.nama_lengkap || "").trim(),
      mata_pelajaran: String(g.mata_pelajaran || "").trim(),
      username: String(g.username || "").trim(),
      password: String(g.password || "").trim()
    }));

    const dataSiswa = rawSiswa.map(s => ({
      nis: String(s.nis || "").trim(),
      nama_lengkap: String(s.nama_lengkap || "").trim(),
      kelas: String(s.kelas || "").trim().toUpperCase(),
      username: String(s.username || "").trim(),
      password: String(s.password || "").trim()
    }));

    const dataKelas = rawKelas.map(k => ({
      kode_kelas: String(k.kode_kelas || "").trim(),
      nama_kelas: String(k.nama_kelas || "").trim().toUpperCase(),
      tingkat_kelas: normalizeGradeLevel_(k.tingkat_kelas || k.nama_kelas) || "10",
      wali_kelas: String(k.wali_kelas || "-").trim() || "-"
    }));

    // Pastikan jika ada kelas di DATA_SISWA yang belum terdaftar di DATA_KELAS ikut tersinkronisasi
    const existingNamaKelas = new Set(dataKelas.map(k => k.nama_kelas.toUpperCase()));
    const sheetKelas = ss.getSheetByName(SHEET_NAMES.KELAS);
    dataSiswa.forEach(s => {
      const kls = s.kelas;
      if (kls && !existingNamaKelas.has(kls)) {
        existingNamaKelas.add(kls);
        const tingkat = normalizeGradeLevel_(kls) || "10";
        const newKode = "KLS-" + String(dataKelas.length + 1).padStart(2, "0");
        const newObj = {
          kode_kelas: newKode,
          nama_kelas: kls,
          tingkat_kelas: tingkat,
          wali_kelas: "-"
        };
        dataKelas.push(newObj);
        if (sheetKelas) {
          sheetKelas.appendRow([newObj.kode_kelas, newObj.nama_kelas, newObj.tingkat_kelas, newObj.wali_kelas]);
        }
      }
    });

    const webSettingsRes = getPengaturanWebsite();

    return {
      status: "success",
      guru: dataGuru,
      siswa: dataSiswa,
      kelas: dataKelas,
      master_mapel: masterMapel,
      website_settings: webSettingsRes.settings
    };
  } catch (err) {
    return { status: "error", message: "Gagal memuat data dashboard admin: " + err.message };
  }
}

/**
 * CRUD TABEL MASTER_MAPEL (KELOLA MATA PELAJARAN DARI WEB ADMIN — MENDUKUNG KELAS 1 S/D 12)
 */
function helperGetNormalizedMasterMapel_() {
  return getSheetData_(SHEET_NAMES.MASTER_MAPEL).map(m => {
    const namaMapel = String(m.nama_mapel || "").trim();
    const kelasNorm = normalizeGradeLevel_(m.kelas || m.label_mapel_kelas) || "10";
    const labelMapel = String(m.label_mapel_kelas || "").trim() || (namaMapel ? namaMapel + " (Kelas " + kelasNorm + ")" : "");
    return {
      kode_mapel: String(m.kode_mapel || "").trim(),
      nama_mapel: namaMapel,
      kelas: kelasNorm,
      label_mapel_kelas: labelMapel,
      kelompok: String(m.kelompok || "Umum / Wajib").trim() || "Umum / Wajib"
    };
  });
}

function simpanDataMapel(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const isEdit = Boolean(payload.is_edit);
    const oldKode = String(payload.old_kode_mapel || payload.kode_mapel || "").trim();
    let kodeMapel = String(payload.kode_mapel || "").trim().toUpperCase();
    const namaMapel = String(payload.nama_mapel || "").trim();
    const kelasNorm = normalizeGradeLevel_(payload.kelas);
    const kelompok = String(payload.kelompok || "Umum / Wajib").trim() || "Umum / Wajib";

    if (!namaMapel || !kelasNorm) {
      return { status: "error", message: "Nama Mata Pelajaran dan Tingkat Kelas (1 - 12) wajib diisi!" };
    }

    const labelMapelKelas = namaMapel + " (Kelas " + kelasNorm + ")";

    const ss = getSpreadsheet_(false);
    let sheet = ss.getSheetByName(SHEET_NAMES.MASTER_MAPEL);
    if (!sheet) {
      ensureSheetsExist_(ss);
      sheet = ss.getSheetByName(SHEET_NAMES.MASTER_MAPEL);
    }

    const lastRow = sheet.getLastRow();
    const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];

    if (!kodeMapel) {
      const abbr = namaMapel.replace(/[^A-Za-z]/g, "").slice(0, 3).toUpperCase() || "MPL";
      let baseKode = "MP-" + kelasNorm + "-" + abbr;
      kodeMapel = baseKode;
      let suffix = 2;
      while (rows.some(r => String(r[0]).trim().toLowerCase() === kodeMapel.toLowerCase())) {
        kodeMapel = baseKode + "-" + suffix;
        suffix++;
      }
    }

    if (isEdit) {
      let targetRowIndex = -1;
      let oldLabel = "";
      let oldNamaMurni = "";
      let oldKelasNorm = "";

      for (let i = 0; i < rows.length; i++) {
        const rKode = String(rows[i][0]).trim();
        const rNama = String(rows[i][1]).trim();
        const rKelas = normalizeGradeLevel_(rows[i][2] || rows[i][3]);
        const rLabel = String(rows[i][3] || "").trim() || (rNama ? rNama + " (Kelas " + rKelas + ")" : "");

        if (rKode.toLowerCase() === oldKode.toLowerCase()) {
          targetRowIndex = i + 2;
          oldLabel = rLabel;
          oldNamaMurni = rNama;
          oldKelasNorm = rKelas;
        } else {
          if (rKode.toLowerCase() === kodeMapel.toLowerCase()) {
            return { status: "error", message: "Kode Mapel \"" + kodeMapel + "\" sudah digunakan!" };
          }
          if (rNama.toLowerCase() === namaMapel.toLowerCase() && rKelas === kelasNorm) {
            return { status: "error", message: "Mata Pelajaran \"" + labelMapelKelas + "\" sudah terdaftar!" };
          }
        }
      }

      if (targetRowIndex === -1) {
        return { status: "error", message: "Data mata pelajaran yang akan diedit tidak ditemukan." };
      }

      sheet.getRange(targetRowIndex, 1, 1, 5).setValues([[kodeMapel, namaMapel, kelasNorm, labelMapelKelas, kelompok]]);

      // Cascade Update: Jika nama_mapel atau kelas berubah, perbarui relasi sambil mempertahankan rombel spesifik guru!
      let updatedGuruCount = 0;
      let updatedGuruList = null;

      const buildCascadedLabel = (existingRaw) => {
        const p = splitMapelDanJenis_(existingRaw);
        if (p.target_rombel && p.target_rombel.length > 0 && oldKelasNorm === kelasNorm) {
          return namaMapel + " (Kelas " + p.target_rombel.join(" + ") + ")";
        }
        return labelMapelKelas;
      };

      if (oldLabel && oldLabel.toUpperCase() !== labelMapelKelas.toUpperCase()) {
        // 1. Update DATA_GURU
        const sheetGuru = ss.getSheetByName(SHEET_NAMES.GURU);
        if (sheetGuru && sheetGuru.getLastRow() > 1) {
          const gLastRow = sheetGuru.getLastRow();
          const gMapelRange = sheetGuru.getRange(2, 3, gLastRow - 1, 1);
          const gMapelValues = gMapelRange.getValues();
          let guruChanged = false;

          for (let j = 0; j < gMapelValues.length; j++) {
            const list = parseGuruMapelList_(gMapelValues[j][0]);
            let rowModified = false;
            const newList = list.map(item => {
              if (isSameMasterMapel_(item, oldNamaMurni, oldKelasNorm)) {
                rowModified = true;
                return buildCascadedLabel(item);
              }
              return item;
            });
            if (rowModified) {
              const deduped = [];
              newList.forEach(it => {
                if (!deduped.some(d => d.toUpperCase() === it.toUpperCase())) deduped.push(it);
              });
              gMapelValues[j][0] = deduped.join(", ");
              updatedGuruCount++;
              guruChanged = true;
            }
          }
          if (guruChanged) {
            gMapelRange.setValues(gMapelValues);
            updatedGuruList = getSheetData_(SHEET_NAMES.GURU).map(g => ({
              nip_atau_id: String(g.nip_atau_id || "").trim(),
              nama_lengkap: String(g.nama_lengkap || "").trim(),
              mata_pelajaran: String(g.mata_pelajaran || "").trim(),
              username: String(g.username || "").trim(),
              password: String(g.password || "").trim()
            }));
          }
        }

        // 2. Update MASTER_JENIS_PENILAIAN (kolom 5: mapel_relasi)
        const sheetJenis = ss.getSheetByName(SHEET_NAMES.MASTER_JENIS);
        if (sheetJenis && sheetJenis.getLastRow() > 1 && sheetJenis.getLastColumn() >= 5) {
          const jLastRow = sheetJenis.getLastRow();
          const jRange = sheetJenis.getRange(2, 5, jLastRow - 1, 1);
          const jValues = jRange.getValues();
          let jChanged = false;
          for (let j = 0; j < jValues.length; j++) {
            const rel = String(jValues[j][0] || "").trim();
            if (rel && rel.toUpperCase() !== "SEMUA" && isSameMasterMapel_(rel, oldNamaMurni, oldKelasNorm)) {
              jValues[j][0] = buildCascadedLabel(rel);
              jChanged = true;
            }
          }
          if (jChanged) jRange.setValues(jValues);
        }

        // 3. Update BANK_SOAL (kolom 2), PENGATURAN_UJIAN (kolom 1), HASIL_UJIAN (kolom 5)
        const cascadePaketSheet = (sheetName, colIdx) => {
          const sh = ss.getSheetByName(sheetName);
          if (!sh || sh.getLastRow() <= 1) return;
          const lRow = sh.getLastRow();
          const rng = sh.getRange(2, colIdx, lRow - 1, 1);
          const vals = rng.getValues();
          let chg = false;
          for (let k = 0; k < vals.length; k++) {
            const rawVal = String(vals[k][0] || "").trim();
            if (rawVal && isSameMasterMapel_(rawVal, oldNamaMurni, oldKelasNorm)) {
              const parsed = splitMapelDanJenis_(rawVal);
              const newInduk = buildCascadedLabel(parsed.mapel_induk);
              vals[k][0] = rawVal.includes(" - ")
                ? (newInduk + " - " + parsed.jenis_ujian)
                : newInduk;
              chg = true;
            }
          }
          if (chg) rng.setValues(vals);
        };

        cascadePaketSheet(SHEET_NAMES.SOAL, 2);
        cascadePaketSheet(SHEET_NAMES.PENGATURAN, 1);
        cascadePaketSheet(SHEET_NAMES.HASIL, 5);
      }

      const resp = {
        status: "success",
        message: "Mata Pelajaran (" + labelMapelKelas + ") berhasil diperbarui!" +
          (updatedGuruCount > 0 ? " (" + updatedGuruCount + " data guru terkait ikut diperbarui otomatis)" : ""),
        mapel: {
          kode_mapel: kodeMapel,
          nama_mapel: namaMapel,
          kelas: kelasNorm,
          label_mapel_kelas: labelMapelKelas,
          kelompok: kelompok
        }
      };
      if (updatedGuruList) resp.guru = updatedGuruList;
      return resp;
    } else {
      const dupKode = rows.some(r => String(r[0]).trim().toLowerCase() === kodeMapel.toLowerCase());
      if (dupKode) {
        return { status: "error", message: "Kode Mapel \"" + kodeMapel + "\" sudah terdaftar!" };
      }
      const dupMapel = rows.some(r => {
        const rNama = String(r[1]).trim();
        const rKelas = normalizeGradeLevel_(r[2] || r[3]);
        return rNama.toLowerCase() === namaMapel.toLowerCase() && rKelas === kelasNorm;
      });
      if (dupMapel) {
        return { status: "error", message: "Mata Pelajaran \"" + labelMapelKelas + "\" sudah terdaftar!" };
      }

      sheet.appendRow([kodeMapel, namaMapel, kelasNorm, labelMapelKelas, kelompok]);
      return {
        status: "success",
        message: "Mata Pelajaran baru (" + labelMapelKelas + ") berhasil ditambahkan!",
        mapel: {
          kode_mapel: kodeMapel,
          nama_mapel: namaMapel,
          kelas: kelasNorm,
          label_mapel_kelas: labelMapelKelas,
          kelompok: kelompok
        }
      };
    }
  } catch (err) {
    return { status: "error", message: "Gagal menyimpan data mata pelajaran: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function hapusDataMapel(kode_mapel) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const targetKode = String(kode_mapel || "").trim();
    if (!targetKode) {
      return { status: "error", message: "Kode Mata Pelajaran tidak valid." };
    }

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.MASTER_MAPEL);
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (lastRow <= 1) {
      return { status: "error", message: "Data master mata pelajaran kosong." };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).trim().toLowerCase() === targetKode.toLowerCase()) {
        const namaMapel = String(data[i][1] || "").trim();
        const kelasNorm = normalizeGradeLevel_(data[i][2] || data[i][3]) || "";
        const labelMapel = String(data[i][3] || "").trim() || (namaMapel + " (Kelas " + kelasNorm + ")");

        // Cek relasi dengan DATA_GURU (mencakup ampuan tingkat umum maupun rombel spesifik)
        const dataGuru = getSheetData_(SHEET_NAMES.GURU);
        const guruTerkait = dataGuru.filter(g =>
          parseGuruMapelList_(g.mata_pelajaran).some(mp => isSameMasterMapel_(mp, namaMapel, kelasNorm))
        );
        if (guruTerkait.length > 0) {
          return {
            status: "error",
            message: "Mata Pelajaran \"" + labelMapel + "\" tidak dapat dihapus karena masih diampu oleh " +
              guruTerkait.length + " guru aktif. Hapus mapel ini dari akun guru terkait terlebih dahulu."
          };
        }

        // Cek relasi dengan BANK_SOAL dan PENGATURAN_UJIAN
        const dataSoal = getSheetData_(SHEET_NAMES.SOAL);
        const soalTerkait = dataSoal.filter(s => isSameMasterMapel_(s.mata_pelajaran, namaMapel, kelasNorm));
        const dataPengaturan = getSheetData_(SHEET_NAMES.PENGATURAN);
        const paketTerkait = dataPengaturan.filter(p => isSameMasterMapel_(p.mata_pelajaran, namaMapel, kelasNorm));
        if (soalTerkait.length > 0 || paketTerkait.length > 0) {
          return {
            status: "error",
            message: "Mata Pelajaran \"" + labelMapel + "\" tidak dapat dihapus karena masih digunakan oleh " +
              soalTerkait.length + " butir soal dan " + paketTerkait.length + " paket ujian."
          };
        }

        sheet.deleteRow(i + 2);
        return {
          status: "success",
          message: "Mata Pelajaran \"" + labelMapel + "\" (" + targetKode + ") berhasil dihapus."
        };
      }
    }

    return { status: "error", message: "Data mata pelajaran dengan kode tersebut tidak ditemukan." };
  } catch (err) {
    return { status: "error", message: "Gagal menghapus data mata pelajaran: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * CRUD TABEL DATA_KELAS (BERELASI DENGAN DATA_SISWA — MENDUKUNG KELAS 1 S/D KELAS 12 TANPA JURUSAN)
 */
function simpanDataKelas(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const isEdit = Boolean(payload.is_edit);
    const oldKode = String(payload.old_kode_kelas || payload.kode_kelas || "").trim();
    let kodeKelas = String(payload.kode_kelas || "").trim().toUpperCase();
    const namaKelas = String(payload.nama_kelas || "").trim().toUpperCase();
    const tingkatKelas = normalizeGradeLevel_(payload.tingkat_kelas || namaKelas) || "1";
    const waliKelas = String(payload.wali_kelas || "-").trim() || "-";

    if (!namaKelas || !tingkatKelas) {
      return { status: "error", message: "Nama Rombel Kelas dan Tingkat Kelas (1 - 12) wajib diisi!" };
    }

    const ss = getSpreadsheet_(false);
    let sheet = ss.getSheetByName(SHEET_NAMES.KELAS);
    if (!sheet) {
      ensureSheetsExist_(ss);
      sheet = ss.getSheetByName(SHEET_NAMES.KELAS);
    } else {
      ensureKelasSheetSeeded_(ss);
    }

    const lastRow = sheet.getLastRow();
    const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 4).getValues() : [];

    if (!kodeKelas) {
      let maxNum = rows.length;
      rows.forEach(r => {
        const m = String(r[0] || "").match(/^KLS-(\d+)/i);
        if (m) {
          const n = parseInt(m[1], 10);
          if (n > maxNum) maxNum = n;
        }
      });
      kodeKelas = "KLS-" + String(maxNum + 1).padStart(2, "0");
    }

    if (isEdit) {
      let targetRowIndex = -1;
      let oldNamaKelas = "";

      for (let i = 0; i < rows.length; i++) {
        const rKode = String(rows[i][0]).trim();
        const rNama = String(rows[i][1]).trim().toUpperCase();
        if (rKode.toLowerCase() === oldKode.toLowerCase()) {
          targetRowIndex = i + 2;
          oldNamaKelas = rNama;
        } else {
          if (rKode.toLowerCase() === kodeKelas.toLowerCase()) {
            return { status: "error", message: "Kode Kelas \"" + kodeKelas + "\" sudah digunakan!" };
          }
          if (rNama === namaKelas) {
            return { status: "error", message: "Nama Rombel Kelas \"" + namaKelas + "\" sudah terdaftar!" };
          }
        }
      }

      if (targetRowIndex === -1) {
        return { status: "error", message: "Data kelas yang akan diedit tidak ditemukan." };
      }

      sheet.getRange(targetRowIndex, 1, 1, 4).setValues([[kodeKelas, namaKelas, tingkatKelas, waliKelas]]);

      // Relasi Cascade Update: Jika nama_kelas berubah, perbarui otomatis kelas pada DATA_SISWA & target rombel di DATA_GURU
      let updatedSiswaCount = 0;
      let updatedGuruList = null;
      if (oldNamaKelas && oldNamaKelas !== namaKelas) {
        const sheetSiswa = ss.getSheetByName(SHEET_NAMES.SISWA);
        if (sheetSiswa && sheetSiswa.getLastRow() > 1) {
          const sLastRow = sheetSiswa.getLastRow();
          const sKelasRange = sheetSiswa.getRange(2, 3, sLastRow - 1, 1);
          const sKelasValues = sKelasRange.getValues();
          let changed = false;
          for (let j = 0; j < sKelasValues.length; j++) {
            if (String(sKelasValues[j][0]).trim().toUpperCase() === oldNamaKelas) {
              sKelasValues[j][0] = namaKelas;
              updatedSiswaCount++;
              changed = true;
            }
          }
          if (changed) {
            sKelasRange.setValues(sKelasValues);
          }
        }

        const replaceRombelInLabel = (rawVal) => {
          const p = splitMapelDanJenis_(rawVal);
          if (!p.target_rombel || p.target_rombel.length === 0) return rawVal;
          let hit = false;
          const nextRombels = p.target_rombel.map(r => {
            if (r.toUpperCase() === oldNamaKelas) {
              hit = true;
              return namaKelas;
            }
            return r;
          });
          if (!hit) return rawVal;
          const newInduk = p.nama_mapel_murni + " (Kelas " + nextRombels.join(" + ") + ")";
          return rawVal.includes(" - ") ? (newInduk + " - " + p.jenis_ujian) : newInduk;
        };

        const sheetGuru = ss.getSheetByName(SHEET_NAMES.GURU);
        if (sheetGuru && sheetGuru.getLastRow() > 1) {
          const gLastRow = sheetGuru.getLastRow();
          const gRange = sheetGuru.getRange(2, 3, gLastRow - 1, 1);
          const gVals = gRange.getValues();
          let gChg = false;
          for (let j = 0; j < gVals.length; j++) {
            const list = parseGuruMapelList_(gVals[j][0]);
            let rowHit = false;
            const nextList = list.map(item => {
              const rep = replaceRombelInLabel(item);
              if (rep !== item) rowHit = true;
              return rep;
            });
            if (rowHit) {
              gVals[j][0] = nextList.join(", ");
              gChg = true;
            }
          }
          if (gChg) {
            gRange.setValues(gVals);
            updatedGuruList = getSheetData_(SHEET_NAMES.GURU).map(g => ({
              nip_atau_id: String(g.nip_atau_id || "").trim(),
              nama_lengkap: String(g.nama_lengkap || "").trim(),
              mata_pelajaran: String(g.mata_pelajaran || "").trim(),
              username: String(g.username || "").trim(),
              password: String(g.password || "").trim()
            }));
          }
        }

        const cascadeRombelSheet = (shName, colIdx) => {
          const sh = ss.getSheetByName(shName);
          if (!sh || sh.getLastRow() <= 1 || sh.getLastColumn() < colIdx) return;
          const lRow = sh.getLastRow();
          const rng = sh.getRange(2, colIdx, lRow - 1, 1);
          const vals = rng.getValues();
          let chg = false;
          for (let k = 0; k < vals.length; k++) {
            const orig = String(vals[k][0] || "").trim();
            if (!orig) continue;
            const rep = replaceRombelInLabel(orig);
            if (rep !== orig) {
              vals[k][0] = rep;
              chg = true;
            }
          }
          if (chg) rng.setValues(vals);
        };
        cascadeRombelSheet(SHEET_NAMES.MASTER_JENIS, 5);
        cascadeRombelSheet(SHEET_NAMES.SOAL, 2);
        cascadeRombelSheet(SHEET_NAMES.PENGATURAN, 1);
        cascadeRombelSheet(SHEET_NAMES.HASIL, 5);
      }

      const respKelas = {
        status: "success",
        message: "Data Kelas (" + namaKelas + " - Kelas " + tingkatKelas + ") berhasil diperbarui!" +
          (updatedSiswaCount > 0 ? " (" + updatedSiswaCount + " data siswa terkait ikut diperbarui otomatis)" : ""),
        old_nama_kelas: oldNamaKelas,
        kelas: {
          kode_kelas: kodeKelas,
          nama_kelas: namaKelas,
          tingkat_kelas: tingkatKelas,
          wali_kelas: waliKelas
        }
      };
      if (updatedGuruList) respKelas.guru = updatedGuruList;
      return respKelas;
    } else {
      const duplicateKode = rows.some(r => String(r[0]).trim().toLowerCase() === kodeKelas.toLowerCase());
      if (duplicateKode) {
        return { status: "error", message: "Kode Kelas \"" + kodeKelas + "\" sudah terdaftar!" };
      }
      const duplicateNama = rows.some(r => String(r[1]).trim().toUpperCase() === namaKelas);
      if (duplicateNama) {
        return { status: "error", message: "Nama Rombel Kelas \"" + namaKelas + "\" sudah terdaftar!" };
      }

      sheet.appendRow([kodeKelas, namaKelas, tingkatKelas, waliKelas]);
      return {
        status: "success",
        message: "Rombel Kelas baru (" + namaKelas + " - Kelas " + tingkatKelas + ") berhasil ditambahkan!",
        kelas: {
          kode_kelas: kodeKelas,
          nama_kelas: namaKelas,
          tingkat_kelas: tingkatKelas,
          wali_kelas: waliKelas
        }
      };
    }
  } catch (err) {
    return { status: "error", message: "Gagal menyimpan data kelas: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function hapusDataKelas(kode_kelas) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const targetKode = String(kode_kelas || "").trim();
    if (!targetKode) {
      return { status: "error", message: "Kode Kelas tidak valid." };
    }

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.KELAS);
    const lastRow = sheet ? sheet.getLastRow() : 0;
    if (lastRow <= 1) {
      return { status: "error", message: "Data kelas kosong." };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).trim().toLowerCase() === targetKode.toLowerCase()) {
        const namaKelas = String(data[i][1]).trim().toUpperCase();

        // Cek apakah masih ada siswa yang berelasi dengan kelas ini
        const dataSiswa = getSheetData_(SHEET_NAMES.SISWA);
        const siswaTerkait = dataSiswa.filter(s => String(s.kelas || "").trim().toUpperCase() === namaKelas);
        if (siswaTerkait.length > 0) {
          return {
            status: "error",
            message: "Kelas \"" + namaKelas + "\" tidak dapat dihapus karena masih berelasi dengan " +
              siswaTerkait.length + " siswa aktif. Ubah atau pindahkan kelas siswa tersebut terlebih dahulu."
          };
        }

        sheet.deleteRow(i + 2);
        return {
          status: "success",
          message: "Data Kelas \"" + namaKelas + "\" (" + targetKode + ") berhasil dihapus."
        };
      }
    }

    return { status: "error", message: "Data kelas dengan kode tersebut tidak ditemukan." };
  } catch (err) {
    return { status: "error", message: "Gagal menghapus data kelas: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function isUsernameTaken_(username, ignoreRole, ignorePrimaryKey) {
  const targetUser = String(username || "").trim().toLowerCase();
  const ignoreKey = String(ignorePrimaryKey || "").trim().toLowerCase();

  const admins = getSheetData_(SHEET_NAMES.ADMIN);
  for (let i = 0; i < admins.length; i++) {
    const u = String(admins[i].username || "").trim().toLowerCase();
    const id = String(admins[i].id_admin || "").trim().toLowerCase();
    if (u === targetUser && !(ignoreRole === "admin" && id === ignoreKey)) {
      return true;
    }
  }

  const gurus = getSheetData_(SHEET_NAMES.GURU);
  for (let i = 0; i < gurus.length; i++) {
    const u = String(gurus[i].username || "").trim().toLowerCase();
    const id = String(gurus[i].nip_atau_id || "").trim().toLowerCase();
    if (u === targetUser && !(ignoreRole === "guru" && id === ignoreKey)) {
      return true;
    }
  }

  const siswas = getSheetData_(SHEET_NAMES.SISWA);
  for (let i = 0; i < siswas.length; i++) {
    const u = String(siswas[i].username || "").trim().toLowerCase();
    const id = String(siswas[i].nis || "").trim().toLowerCase();
    if (u === targetUser && !(ignoreRole === "siswa" && id === ignoreKey)) {
      return true;
    }
  }

  return false;
}

function simpanDataGuru(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const isEdit = Boolean(payload.is_edit);
    const oldNip = String(payload.old_nip_atau_id || payload.nip_atau_id || "").trim();
    const nip = String(payload.nip_atau_id || "").trim();
    const nama = String(payload.nama_lengkap || "").trim();
    const mapel = String(payload.mata_pelajaran || "").trim();
    const username = String(payload.username || "").trim() || nip;
    const password = String(payload.password || "").trim() || nip;

    if (!nip || !nama || !mapel) {
      return { status: "error", message: "NIP/ID Guru, Nama Lengkap, dan Mata Pelajaran wajib diisi!" };
    }

    const ss = getSpreadsheet_(false);
    let sheet = ss.getSheetByName(SHEET_NAMES.GURU);
    if (!sheet) {
      ensureSheetsExist_(ss);
      sheet = ss.getSheetByName(SHEET_NAMES.GURU);
    }

    const lastRow = sheet.getLastRow();
    const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];

    if (isUsernameTaken_(username, "guru", isEdit ? oldNip : "")) {
      return { status: "error", message: "Username \"" + username + "\" sudah digunakan oleh akun lain!" };
    }

    if (isEdit) {
      let targetRowIndex = -1;
      for (let i = 0; i < rows.length; i++) {
        const rowNip = String(rows[i][0]).trim();
        if (rowNip.toLowerCase() === oldNip.toLowerCase()) {
          targetRowIndex = i + 2;
        } else if (rowNip.toLowerCase() === nip.toLowerCase()) {
          return { status: "error", message: "NIP/ID Guru \"" + nip + "\" sudah terdaftar pada guru lain!" };
        }
      }

      if (targetRowIndex === -1) {
        return { status: "error", message: "Data guru yang akan diedit tidak ditemukan." };
      }

      sheet.getRange(targetRowIndex, 1, 1, 5).setValues([[nip, nama, mapel, username, password]]);
      ensureGuruMapelInMasterMapel_(ss, parseGuruMapelList_(mapel));
      return {
        status: "success",
        message: "Data Guru (" + nama + ") berhasil diperbarui!",
        guru: { nip_atau_id: nip, nama_lengkap: nama, mata_pelajaran: mapel, username: username, password: password },
        master_mapel: helperGetNormalizedMasterMapel_()
      };
    } else {
      const nipExists = rows.some(r => String(r[0]).trim().toLowerCase() === nip.toLowerCase());
      if (nipExists) {
        return { status: "error", message: "NIP/ID Guru \"" + nip + "\" sudah terdaftar!" };
      }

      sheet.appendRow([nip, nama, mapel, username, password]);
      ensureGuruMapelInMasterMapel_(ss, parseGuruMapelList_(mapel));
      return {
        status: "success",
        message: "Akun Guru baru (" + nama + ") berhasil ditambahkan!",
        guru: { nip_atau_id: nip, nama_lengkap: nama, mata_pelajaran: mapel, username: username, password: password },
        master_mapel: helperGetNormalizedMasterMapel_()
      };
    }
  } catch (err) {
    return { status: "error", message: "Gagal menyimpan data guru: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function hapusDataGuru(nip_atau_id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const targetNip = String(nip_atau_id || "").trim();
    if (!targetNip) {
      return { status: "error", message: "NIP/ID Guru tidak valid." };
    }

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.GURU);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { status: "error", message: "Data guru kosong." };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).trim().toLowerCase() === targetNip.toLowerCase()) {
        const namaGuru = data[i][1];
        sheet.deleteRow(i + 2);
        return {
          status: "success",
          message: "Data Guru \"" + namaGuru + "\" (" + targetNip + ") berhasil dihapus."
        };
      }
    }

    return { status: "error", message: "Data guru dengan NIP/ID tersebut tidak ditemukan." };
  } catch (err) {
    return { status: "error", message: "Gagal menghapus data guru: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function simpanDataSiswa(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const isEdit = Boolean(payload.is_edit);
    const oldNis = String(payload.old_nis || payload.nis || "").trim();
    const nis = String(payload.nis || "").trim();
    const nama = String(payload.nama_lengkap || "").trim();
    const kelas = String(payload.kelas || "").trim().toUpperCase();
    const username = String(payload.username || "").trim() || nis;
    const password = String(payload.password || "").trim() || nis;

    if (!nis || !nama || !kelas) {
      return { status: "error", message: "NIS, Nama Lengkap, dan Kelas Siswa wajib diisi!" };
    }

    const ss = getSpreadsheet_(false);
    let sheet = ss.getSheetByName(SHEET_NAMES.SISWA);
    if (!sheet) {
      ensureSheetsExist_(ss);
      sheet = ss.getSheetByName(SHEET_NAMES.SISWA);
    }

    const lastRow = sheet.getLastRow();
    const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 5).getValues() : [];

    if (isUsernameTaken_(username, "siswa", isEdit ? oldNis : "")) {
      return { status: "error", message: "Username \"" + username + "\" sudah digunakan oleh akun lain!" };
    }

    if (isEdit) {
      let targetRowIndex = -1;
      for (let i = 0; i < rows.length; i++) {
        const rowNis = String(rows[i][0]).trim();
        if (rowNis.toLowerCase() === oldNis.toLowerCase()) {
          targetRowIndex = i + 2;
        } else if (rowNis.toLowerCase() === nis.toLowerCase()) {
          return { status: "error", message: "NIS \"" + nis + "\" sudah terdaftar pada siswa lain!" };
        }
      }

      if (targetRowIndex === -1) {
        return { status: "error", message: "Data siswa yang akan diedit tidak ditemukan." };
      }

      sheet.getRange(targetRowIndex, 1, 1, 5).setValues([[nis, nama, kelas, username, password]]);
      return {
        status: "success",
        message: "Data Siswa (" + nama + ") berhasil diperbarui!",
        siswa: { nis: nis, nama_lengkap: nama, kelas: kelas, username: username, password: password }
      };
    } else {
      const nisExists = rows.some(r => String(r[0]).trim().toLowerCase() === nis.toLowerCase());
      if (nisExists) {
        return { status: "error", message: "NIS \"" + nis + "\" sudah terdaftar!" };
      }

      sheet.appendRow([nis, nama, kelas, username, password]);
      return {
        status: "success",
        message: "Akun Siswa baru (" + nama + ") berhasil ditambahkan!",
        siswa: { nis: nis, nama_lengkap: nama, kelas: kelas, username: username, password: password }
      };
    }
  } catch (err) {
    return { status: "error", message: "Gagal menyimpan data siswa: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

function hapusDataSiswa(nis) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const targetNis = String(nis || "").trim();
    if (!targetNis) {
      return { status: "error", message: "NIS Siswa tidak valid." };
    }

    const ss = getSpreadsheet_(false);
    const sheet = ss.getSheetByName(SHEET_NAMES.SISWA);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { status: "error", message: "Data siswa kosong." };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]).trim().toLowerCase() === targetNis.toLowerCase()) {
        const namaSiswa = data[i][1];
        sheet.deleteRow(i + 2);
        return {
          status: "success",
          message: "Data Siswa \"" + namaSiswa + "\" (NIS: " + targetNis + ") berhasil dihapus."
        };
      }
    }

    return { status: "error", message: "Data siswa dengan NIS tersebut tidak ditemukan." };
  } catch (err) {
    return { status: "error", message: "Gagal menghapus data siswa: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

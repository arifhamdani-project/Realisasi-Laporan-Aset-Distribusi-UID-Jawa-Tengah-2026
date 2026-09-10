/**
 * ============================================================
 *  APLIKASI REALISASI LAPORAN KINERJA ASET (SE060 / MAXIMO / GIS)
 * ============================================================
 *  Cara pakai:
 *  1. Buka spreadsheet database kamu.
 *  2. Menu Extensions > Apps Script.
 *  3. Hapus isi default, paste isi file ini ke "Code.gs".
 *  4. Buat file HTML baru bernama "Index" (File > New > HTML file),
 *     paste isi "Index.html" ke situ.
 *  5. Klik Deploy > New deployment > pilih tipe "Web app".
 *     - Execute as: Me
 *     - Who has access: sesuaikan (misal "Anyone within organization")
 *  6. Copy URL hasil deploy, itulah link form realisasi kamu.
 * ============================================================
 */

// ------------------------------------------------------------
// KONFIGURASI — SESUAIKAN DENGAN SPREADSHEET KAMU
// ------------------------------------------------------------

// Nama-nama tab/sheet Unit yang bisa dipilih di form.
// Sesuaikan urutan/isinya kalau ada Unit yang kurang/lebih.
const UNITS = [
  "Kudus", "Surakarta", "Magelang", "Purwokerto", "Tegal",
  "Semarang", "Salatiga", "Klaten", "Pekalongan", "Cilacap",
  "Grobogan", "Sukoharjo"
];

// Nama-nama section/blok tabel di dalam tiap sheet Unit.
// Harus PERSIS sama dengan teks judul section di spreadsheet (SE 060 / MAXIMO / GIS).
const SECTIONS = ["SE 060", "MAXIMO", "GIS"];

// Kolom yang berisi nama bulan (baris per bulan) di tiap section.
const MONTH_COLUMN = "B";

// Nama bulan persis seperti yang tertulis di spreadsheet.
const MONTHS = [
  "Januari", "Februari", "Maret", "April", "Mei", "Juni",
  "Juli", "Agustus", "September", "Oktober", "November", "Desember"
];

// Nama sheet untuk mencatat log riwayat pengisian (akan dibuat otomatis kalau belum ada).
const LOG_SHEET_NAME = "Log Pengisian Realisasi";

// Definisi field per grup. `col` = kolom di spreadsheet.
// Kolom "Jumlah" tiap grup TIDAK dimasukkan di sini karena diasumsikan
// hasil RUMUS OTOMATIS (SUM). Kalau ternyata manual, tambahkan sendiri.
const FIELDS = [
  { group: "JTM Switching", key: "jtm_sutm", label: "SUTM", unit: "kms", col: "C" },
  { group: "JTM Switching", key: "jtm_skutm", label: "SKUTM", unit: "kms", col: "D" },
  { group: "JTM Switching", key: "jtm_sktm", label: "SKTM", unit: "kms", col: "E" },
  { group: "JTM Switching", key: "jtm_skltm", label: "SKLTM", unit: "kms", col: "F" },

  { group: "Penyulang", key: "penyulang_jumlah", label: "Jumlah", unit: "buah", col: "H" },

  { group: "JTR", key: "jtr_sutr", label: "SUTR", unit: "kms", col: "I" },
  { group: "JTR", key: "jtr_sktr", label: "SKTR", unit: "kms", col: "J" },
  { group: "JTR", key: "jtr_skutr", label: "SKUTR", unit: "kms", col: "K" },

  { group: "Kubikel", key: "kubikel_jumlah", label: "Jumlah", unit: "buah", col: "M" },

  { group: "MV Cell", key: "mvcell_gi", label: "GI", unit: "buah", col: "N" },
  { group: "MV Cell", key: "mvcell_gd", label: "GD", unit: "buah", col: "O" },
  { group: "MV Cell", key: "mvcell_gh", label: "GH", unit: "buah", col: "P" },

  { group: "Jumlah Gardu", key: "gardu_gd", label: "GD", unit: "buah", col: "R" },
  { group: "Jumlah Gardu", key: "gardu_gh", label: "GH", unit: "buah", col: "S" },

  { group: "Trafo Distribusi", key: "trafo_unit", label: "UNIT", unit: "buah", col: "U" },
  { group: "Trafo Distribusi", key: "trafo_total_kva", label: "TOTAL", unit: "kVA", col: "V" },

  { group: "Jumlah Tiang", key: "tiang_tm", label: "TM", unit: "buah", col: "W" },
  { group: "Jumlah Tiang", key: "tiang_tr", label: "TR", unit: "buah", col: "X" },

  { group: "Pelanggan Tersambung", key: "pelanggan_jumlah", label: "Jumlah", unit: "buah", col: "Y" },

  { group: "SR", key: "sr_kms", label: "SR", unit: "kms", col: "Z" },

  { group: "Asset Register (AR)", key: "ar_switching_motorized", label: "AR Switching Motorized", unit: "buah", col: "AC" },
  { group: "Asset Register (AR)", key: "ar_double_switching", label: "AR Double Switching", unit: "buah", col: "AD" },
  { group: "Asset Register (AR)", key: "ar_switching_manual", label: "AR Switching Manual", unit: "buah", col: "AE" },
  { group: "Asset Register (AR)", key: "ar_phbtr", label: "AR PHBTR", unit: "buah", col: "AF" },
];

// Berapa baris ke bawah dari judul section ("SE 060" dsb) yang boleh dicari
// untuk menemukan baris bulan. Beri margin aman.
const MAX_ROWS_PER_SECTION_BLOCK = 20;

// ------------------------------------------------------------
// WEB APP ENTRY POINT
// ------------------------------------------------------------
function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Form Realisasi Laporan Kinerja Aset")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ------------------------------------------------------------
// UTIL: hitung nama bulan n-1 dari tanggal hari ini
// ------------------------------------------------------------
function getEditableMonth_() {
  const now = new Date();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || "GMT+7";
  const currentMonthIndex = parseInt(Utilities.formatDate(now, tz, "M"), 10); // 1-12
  const currentYear = parseInt(Utilities.formatDate(now, tz, "yyyy"), 10);

  let prevMonthIndex = currentMonthIndex - 1; // n-1
  let prevYear = currentYear;
  if (prevMonthIndex < 1) {
    prevMonthIndex = 12;
    prevYear = currentYear - 1;
  }
  return {
    monthName: MONTHS[prevMonthIndex - 1],
    monthIndex: prevMonthIndex,
    year: prevYear
  };
}

// ------------------------------------------------------------
// Dipanggil oleh front-end saat form dibuka:
// kirim konfigurasi awal (unit list, section list, field list, bulan yang aktif)
// ------------------------------------------------------------
function getFormConfig() {
  const editable = getEditableMonth_();
  return {
    units: UNITS,
    sections: SECTIONS,
    fields: FIELDS,
    editableMonth: editable.monthName,
    editableYear: editable.year
  };
}

// ------------------------------------------------------------
// Cari baris section tertentu di sebuah sheet, lalu cari baris bulan di
// dalam blok section tersebut. Return nomor baris (1-based) atau null.
// ------------------------------------------------------------
function findMonthRow_(sheet, sectionName, monthName) {
  const values = sheet.getDataRange().getValues();
  let sectionRow = -1;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = values[r][c];
      if (typeof cell === "string" && cell.trim().toUpperCase() === sectionName.toUpperCase()) {
        sectionRow = r; // 0-based index dalam array `values`
        break;
      }
    }
    if (sectionRow !== -1) break;
  }

  if (sectionRow === -1) return null;

  const monthColIndex = columnLetterToIndex_(MONTH_COLUMN) - 1; // 0-based
  const searchEnd = Math.min(values.length, sectionRow + MAX_ROWS_PER_SECTION_BLOCK);

  for (let r = sectionRow; r < searchEnd; r++) {
    const cell = values[r][monthColIndex];
    if (typeof cell === "string" && cell.trim().toLowerCase() === monthName.toLowerCase()) {
      return r + 1; // balikin ke 1-based row number untuk Range
    }
  }
  return null;
}

function columnLetterToIndex_(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column = column * 26 + (letter.charCodeAt(i) - 64);
  }
  return column;
}

// ------------------------------------------------------------
// Ambil data yang SUDAH ADA untuk unit + bulan aktif (untuk mode koreksi/edit)
// ------------------------------------------------------------
function getExistingData(unitName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(unitName);
  if (!sheet) throw new Error("Sheet unit '" + unitName + "' tidak ditemukan.");

  const editable = getEditableMonth_();
  const result = {};

  SECTIONS.forEach(function (section) {
    const row = findMonthRow_(sheet, section, editable.monthName);
    const sectionData = {};
    if (row) {
      FIELDS.forEach(function (f) {
        const colIndex = columnLetterToIndex_(f.col);
        const val = sheet.getRange(row, colIndex).getValue();
        sectionData[f.key] = (val === "" || val === "-" ) ? "" : val;
      });
    } else {
      FIELDS.forEach(function (f) { sectionData[f.key] = ""; });
    }
    result[section] = { row: row, data: sectionData };
  });

  return result;
}

// ------------------------------------------------------------
// SIMPAN DATA — dipanggil saat user submit form
// payload = {
//   unit: "Kudus",
//   namaPengisi: "Budi",
//   sections: {
//     "SE 060": { jtm_sutm: 123, ... },
//     "MAXIMO": { ... },
//     "GIS": { ... }
//   }
// }
// ------------------------------------------------------------
function saveRealisasi(payload) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(payload.unit);
  if (!sheet) throw new Error("Sheet unit '" + payload.unit + "' tidak ditemukan.");

  const editable = getEditableMonth_();
  const summary = [];

  SECTIONS.forEach(function (section) {
    const row = findMonthRow_(sheet, section, editable.monthName);
    if (!row) {
      throw new Error(
        "Baris bulan '" + editable.monthName + "' untuk section '" + section +
        "' tidak ditemukan di sheet '" + payload.unit + "'. Cek struktur sheet."
      );
    }

    const sectionValues = (payload.sections && payload.sections[section]) || {};

    FIELDS.forEach(function (f) {
      if (Object.prototype.hasOwnProperty.call(sectionValues, f.key)) {
        const raw = sectionValues[f.key];
        const colIndex = columnLetterToIndex_(f.col);
        // Kosongkan sel kalau input kosong, atau isi angka kalau ada nilai.
        if (raw === "" || raw === null || typeof raw === "undefined") {
          // tidak menimpa data lama kalau memang dikosongkan sengaja bisa diubah sesuai kebutuhan
        } else {
          const num = parseFloat(String(raw).toString().replace(",", "."));
          sheet.getRange(row, colIndex).setValue(isNaN(num) ? raw : num);
        }
      }
    });

    summary.push(section + " -> baris " + row);
  });

  logSubmission_(payload.unit, editable.monthName, editable.year, payload.namaPengisi);

  return {
    ok: true,
    message: "Data realisasi bulan " + editable.monthName + " " + editable.year +
      " untuk Unit " + payload.unit + " berhasil disimpan.\n" + summary.join("\n")
  };
}

// ------------------------------------------------------------
// Catat log setiap kali ada submit (audit trail sederhana)
// ------------------------------------------------------------
function logSubmission_(unit, monthName, year, namaPengisi) {
  const ss = SpreadsheetApp.getActive();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(["Timestamp", "Unit", "Bulan", "Tahun", "Nama Pengisi"]);
  }
  logSheet.appendRow([new Date(), unit, monthName, year, namaPengisi || "-"]);
}

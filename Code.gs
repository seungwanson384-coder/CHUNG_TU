/**
 * Code.gs
 * Entry point, Web App routing, authentication, reference dropdown loading, and draft persistence.
 * Runtime: Google Apps Script (V8)
 */

/**
 * Serves the modular Index.html web application.
 * @param {GoogleAppsScript.Events.DoGet} e
 * @return {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template
    .evaluate()
    .setTitle('Hệ thống Tự động hóa Chứng từ Xuất Nhập Khẩu & Logistics')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper to include modular HTML/CSS/JS template files inside Index.html.
 * Usage in HTML: <?!= include('Stylesheet'); ?>
 * @param {string} filename
 * @return {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Resolves the backing Google Spreadsheet.
 * Uses SPREADSHEET_ID from GS_Config_Templates.gs, or falls back to active bound spreadsheet.
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getAppSpreadsheet_() {
  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID && SPREADSHEET_ID.indexOf('YOUR_') !== 0) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('Không tìm thấy Google Spreadsheet. Vui lòng cấu hình SPREADSHEET_ID trong GS_Config_Templates.gs.');
  }
  return active;
}

/**
 * Authenticates a student/user against tab `Danh_sach_SV`.
 * Columns: [STT, Ma_SV, Ten_SV, Mat_Khau, Draft_JSON]
 * Preserves leading zeros by reading display strings and comparing strict string values.
 *
 * @param {string} maSV
 * @param {string} matKhau
 * @return {{success: boolean, user?: {maSV: string, tenSV: string}, draftJson?: string, message?: string}}
 */
function loginUser(maSV, matKhau) {
  try {
    const inputMaSV = String(maSV == null ? '' : maSV).trim();
    const inputMatKhau = String(matKhau == null ? '' : matKhau).trim();

    if (!inputMaSV || !inputMatKhau) {
      return {
        success: false,
        message: 'Vui lòng nhập đầy đủ Mã sinh viên và Mật khẩu.'
      };
    }

    const ss = getAppSpreadsheet_();
    const sheet = ss.getSheetByName('Danh_sach_SV');
    if (!sheet) {
      return {
        success: false,
        message: 'Không tìm thấy sheet "Danh_sach_SV".'
      };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 4) {
      return {
        success: false,
        message: 'Danh sách sinh viên chưa có dữ liệu.'
      };
    }

    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const displayValues = range.getDisplayValues();
    const rawValues = range.getValues();

    const colIdx = resolveColumnIndices_(displayValues[0], {
      STT: 0,
      Ma_SV: 1,
      Ten_SV: 2,
      Mat_Khau: 3,
      Draft_JSON: 4
    });

    for (let r = 1; r < displayValues.length; r++) {
      const rowDisplay = displayValues[r];
      const cellMaSV = String(rowDisplay[colIdx.Ma_SV] || '').trim();
      const cellMatKhau = String(rowDisplay[colIdx.Mat_Khau] || '').trim();

      if (cellMaSV !== '' && cellMaSV === inputMaSV && cellMatKhau === inputMatKhau) {
        const cellTenSV = String(rowDisplay[colIdx.Ten_SV] || '').trim();
        let draftJson = '';
        if (colIdx.Draft_JSON < rawValues[r].length) {
          const rawDraft = rawValues[r][colIdx.Draft_JSON];
          draftJson = rawDraft == null ? '' : String(rawDraft);
        }

        return {
          success: true,
          user: {
            maSV: cellMaSV,
            tenSV: cellTenSV
          },
          draftJson: draftJson
        };
      }
    }

    return {
      success: false,
      message: 'Mã sinh viên hoặc Mật khẩu không chính xác.'
    };
  } catch (err) {
    return {
      success: false,
      message: 'Lỗi đăng nhập: ' + err.message
    };
  }
}

/**
 * Reads and returns structured JSON from the 3 reference tabs (zero merged cells):
 * - Cuoc_Tau: [POL, POD, Quoc_Gia, Ten_hang_tau, Ten_tau, So_chuyen]
 * - Loai_hinh_XNK: [Phan_loai, Ma_loai_hinh, Ten_loai_hinh]
 * - Danh_Muc_chung: [LOAI_CONT, SERVICE_TYPE, INCOTERM, PAY_TERM, UNIT, CURR, CO_FORM, CO_CRITERIA, TRADE_AGREEMENT]
 *
 * @return {{success: boolean, cuocTau?: Array<Object>, loaiHinhXNK?: Array<Object>, danhMucChung?: Object, message?: string}}
 */
function getInitialDropdownData() {
  try {
    const ss = getAppSpreadsheet_();

    const cuocTau = readTableRecords_(ss, 'Cuoc_Tau', [
      'POL',
      'POD',
      'Quoc_Gia',
      'Ten_hang_tau',
      'Ten_tau',
      'So_chuyen'
    ]);

    const loaiHinhXNK = readTableRecords_(ss, 'Loai_hinh_XNK', [
      'Phan_loai',
      'Ma_loai_hinh',
      'Ten_loai_hinh'
    ]);

    const danhMucChung = readIndependentColumns_(ss, 'Danh_Muc_chung', [
      'LOAI_CONT',
      'SERVICE_TYPE',
      'INCOTERM',
      'PAY_TERM',
      'UNIT',
      'CURR',
      'CO_FORM',
      'CO_CRITERIA',
      'TRADE_AGREEMENT'
    ]);

    return {
      success: true,
      cuocTau: cuocTau,
      loaiHinhXNK: loaiHinhXNK,
      danhMucChung: danhMucChung
    };
  } catch (err) {
    return {
      success: false,
      message: 'Lỗi tải dữ liệu danh mục: ' + err.message
    };
  }
}

/**
 * Finds the user's exact row in `Danh_sach_SV` and overwrites column `Draft_JSON` in-place.
 * Never creates new rows or appends history.
 *
 * @param {string} maSV
 * @param {string} draftJsonString
 * @return {{success: boolean, savedAt?: string, message?: string}}
 */
function saveUserDraft(maSV, draftJsonString) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const targetMaSV = String(maSV == null ? '' : maSV).trim();
    if (!targetMaSV) {
      return {
        success: false,
        message: 'Thiếu Mã sinh viên để lưu bản nháp.'
      };
    }

    const ss = getAppSpreadsheet_();
    const sheet = ss.getSheetByName('Danh_sach_SV');
    if (!sheet) {
      return {
        success: false,
        message: 'Không tìm thấy sheet "Danh_sach_SV".'
      };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) {
      return {
        success: false,
        message: 'Không tìm thấy dữ liệu sinh viên trong sheet "Danh_sach_SV".'
      };
    }

    const displayValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const colIdx = resolveColumnIndices_(displayValues[0], {
      STT: 0,
      Ma_SV: 1,
      Ten_SV: 2,
      Mat_Khau: 3,
      Draft_JSON: 4
    });

    const draftColNumber = colIdx.Draft_JSON + 1;
    const serializedDraft =
      typeof draftJsonString === 'string'
        ? draftJsonString
        : JSON.stringify(draftJsonString || {});

    for (let r = 1; r < displayValues.length; r++) {
      const cellMaSV = String(displayValues[r][colIdx.Ma_SV] || '').trim();
      if (cellMaSV === targetMaSV) {
        const targetCell = sheet.getRange(r + 1, draftColNumber);
        targetCell.setNumberFormat('@');
        targetCell.setValue(serializedDraft);
        SpreadsheetApp.flush();

        return {
          success: true,
          savedAt: new Date().toISOString()
        };
      }
    }

    return {
      success: false,
      message: 'Không tìm thấy sinh viên có Mã SV "' + targetMaSV + '" trong Danh_sach_SV.'
    };
  } catch (err) {
    return {
      success: false,
      message: 'Lỗi lưu bản nháp: ' + err.message
    };
  } finally {
    lock.releaseLock();
  }
}

/* ==========================================================================
 * PRIVATE SHEET PARSING HELPERS
 * ========================================================================== */

/**
 * Resolves 0-based column indices from header row with fallback defaults.
 * @param {Array<string>} headerRow
 * @param {Object<string, number>} defaultsMap
 * @return {Object<string, number>}
 */
function resolveColumnIndices_(headerRow, defaultsMap) {
  const result = {};
  const normalizedHeaders = (headerRow || []).map(function (h) {
    return String(h || '').trim().toLowerCase();
  });

  Object.keys(defaultsMap).forEach(function (key) {
    const foundIdx = normalizedHeaders.indexOf(key.toLowerCase());
    result[key] = foundIdx !== -1 ? foundIdx : defaultsMap[key];
  });

  return result;
}

/**
 * Reads row-based structured objects from a sheet using display strings (preserving leading zeros).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @param {Array<string>} expectedFields
 * @return {Array<Object>}
 */
function readTableRecords_(ss, sheetName, expectedFields) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Không tìm thấy sheet "' + sheetName + '".');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return [];
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const defaults = {};
  expectedFields.forEach(function (field, i) {
    defaults[field] = i;
  });

  const colIdx = resolveColumnIndices_(values[0], defaults);
  const records = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const obj = {};
    let hasAnyValue = false;

    for (let f = 0; f < expectedFields.length; f++) {
      const fieldName = expectedFields[f];
      const idx = colIdx[fieldName];
      const cellVal = idx < row.length ? String(row[idx] || '').trim() : '';
      obj[fieldName] = cellVal;
      if (cellVal !== '') {
        hasAnyValue = true;
      }
    }

    if (hasAnyValue) {
      records.push(obj);
    }
  }

  return records;
}

/**
 * Reads independent dropdown columns from a reference sheet into arrays of unique non-empty strings.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @param {Array<string>} expectedColumns
 * @return {Object<string, Array<string>>}
 */
function readIndependentColumns_(ss, sheetName, expectedColumns) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Không tìm thấy sheet "' + sheetName + '".');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const output = {};

  expectedColumns.forEach(function (col) {
    output[col] = [];
  });

  if (lastRow < 2 || lastCol < 1) {
    return output;
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const defaults = {};
  expectedColumns.forEach(function (col, i) {
    defaults[col] = i;
  });

  const colIdx = resolveColumnIndices_(values[0], defaults);

  expectedColumns.forEach(function (colName) {
    const cIdx = colIdx[colName];
    const seen = {};
    const list = [];

    for (let r = 1; r < values.length; r++) {
      if (cIdx >= values[r].length) continue;
      const val = String(values[r][cIdx] || '').trim();
      if (val !== '' && !seen[val]) {
        seen[val] = true;
        list.push(val);
      }
    }

    output[colName] = list;
  });

  return output;
}

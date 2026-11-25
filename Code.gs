/**
 * Code.gs
 * Entry point, Web App routing, authentication, reference dropdown loading, draft persistence,
 * and Admin Account Management with external sheet scan and safe student bulk delete.
 * Runtime: Google Apps Script (V8)
 */

const NEW_BRAND_LOGO_ID = '1Hjf2eCislKARBB8twtPHunvKviVJS6vq';
const BRAND_LOGO_FALLBACK_URL = 'https://drive.google.com/thumbnail?id=1Hjf2eCislKARBB8twtPHunvKviVJS6vq&sz=w300';

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.logoDataUrl = getBrandLogoBase64();
  return template
    .evaluate()
    .setTitle('Hệ thống xuất chứng từ XNK tự động')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getBrandLogoBase64() {
  const cacheKey = 'BRAND_LOGO_DATA_URI_' + NEW_BRAND_LOGO_ID;
  try {
    const cache = CacheService.getScriptCache();
    const meta = cache.get(cacheKey + '_META');
    if (meta) {
      const parts = parseInt(meta, 10);
      if (parts > 0) {
        let combined = '';
        let ok = true;
        for (let i = 0; i < parts; i++) {
          const chunk = cache.get(cacheKey + '_' + i);
          if (!chunk) {
            ok = false;
            break;
          }
          combined += chunk;
        }
        if (ok && combined.indexOf('data:image/') === 0) {
          return combined;
        }
      }
    }
    const cached = cache.get(cacheKey);
    if (cached && cached.indexOf('data:image/') === 0) {
      return cached;
    }
  } catch (e) {}

  try {
    const file = DriveApp.getFileById(NEW_BRAND_LOGO_ID);
    const blob = file.getBlob();
    const mimeType = blob.getContentType() || 'image/png';
    const base64 = Utilities.base64Encode(blob.getBytes());
    const dataUri = 'data:' + mimeType + ';base64,' + base64;

    try {
      const cache = CacheService.getScriptCache();
      if (dataUri.length <= 90000) {
        cache.put(cacheKey, dataUri, 21600);
      } else {
        const chunkSize = 85000;
        const totalParts = Math.ceil(dataUri.length / chunkSize);
        if (totalParts <= 8) {
          const toPut = {};
          toPut[cacheKey + '_META'] = String(totalParts);
          for (let i = 0; i < totalParts; i++) {
            toPut[cacheKey + '_' + i] = dataUri.substring(i * chunkSize, (i + 1) * chunkSize);
          }
          cache.putAll(toPut, 21600);
        }
      }
    } catch (cacheErr) {}

    return dataUri;
  } catch (err) {
    return BRAND_LOGO_FALLBACK_URL;
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

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

function loginUser(username, password) {
  try {
    const inputUsername = String(username == null ? '' : username).trim().toLowerCase();
    const inputPassword = String(password == null ? '' : password).trim().toLowerCase();

    if (!inputUsername || !inputPassword) {
      return {
        success: false,
        message: 'Vui lòng nhập đầy đủ Tên đăng nhập và Mật khẩu.'
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
    if (lastRow < 2 || lastCol < 3) {
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
      Vai_Tro: 3,
      Draft_JSON: 4
    });

    for (let r = 1; r < displayValues.length; r++) {
      const rowDisplay = displayValues[r];
      const cellMaSV = String(rowDisplay[colIdx.Ma_SV] || '').trim();
      const cellTenSV = String(rowDisplay[colIdx.Ten_SV] || '').trim();
      const cellVaiTro = String(rowDisplay[colIdx.Vai_Tro] || '').trim().toUpperCase();

      if (
        cellTenSV !== '' &&
        cellMaSV !== '' &&
        cellTenSV.toLowerCase() === inputUsername &&
        cellMaSV.toLowerCase() === inputPassword
      ) {
        let draftJson = '';
        if (colIdx.Draft_JSON < rawValues[r].length) {
          const rawDraft = rawValues[r][colIdx.Draft_JSON];
          draftJson = rawDraft == null ? '' : String(rawDraft);
        }

        const role = cellVaiTro === 'ADMIN' ? 'ADMIN' : 'SV';

        return {
          success: true,
          user: {
            maSV: cellMaSV,
            tenSV: cellTenSV,
            role: role
          },
          draftJson: draftJson,
          draft: draftJson
        };
      }
    }

    return {
      success: false,
      message: 'Tên đăng nhập hoặc Mật khẩu không chính xác.'
    };
  } catch (err) {
    return {
      success: false,
      message: 'Lỗi đăng nhập: ' + err.message
    };
  }
}

function verifyAdmin_(adminMaSV) {
  try {
    const target = String(adminMaSV || '').trim().toLowerCase();
    if (!target) return false;
    const ss = getAppSpreadsheet_();
    const sheet = ss.getSheetByName('Danh_sach_SV');
    if (!sheet) return false;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
    const colIdx = resolveColumnIndices_(values[0], { STT: 0, Ma_SV: 1, Ten_SV: 2, Vai_Tro: 3, Draft_JSON: 4 });
    for (let r = 1; r < values.length; r++) {
      const m = String(values[r][colIdx.Ma_SV] || '').trim().toLowerCase();
      if (m === target) {
        const role = String(values[r][colIdx.Vai_Tro] || '').trim().toUpperCase();
        return role === 'ADMIN';
      }
    }
  } catch (e) {}
  return false;
}

function adminGetAccounts(adminMaSV) {
  try {
    if (!verifyAdmin_(adminMaSV)) {
      return { success: false, message: 'Quyền quản trị viên bị từ chối.' };
    }
    const ss = getAppSpreadsheet_();
    const sheet = ss.getSheetByName('Danh_sach_SV');
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { success: true, accounts: [] };

    const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const colIdx = resolveColumnIndices_(values[0], { STT: 0, Ma_SV: 1, Ten_SV: 2, Vai_Tro: 3, Draft_JSON: 4 });
    const accounts = [];

    for (let r = 1; r < values.length; r++) {
      const maSV = String(values[r][colIdx.Ma_SV] || '').trim();
      const tenSV = String(values[r][colIdx.Ten_SV] || '').trim();
      const role = String(values[r][colIdx.Vai_Tro] || '').trim().toUpperCase() === 'ADMIN' ? 'ADMIN' : 'SV';
      if (maSV) {
        accounts.push({ stt: r, maSV: maSV, tenSV: tenSV, vaiTro: role });
      }
    }
    return { success: true, accounts: accounts };
  } catch (err) {
    return { success: false, message: 'Lỗi tải danh sách tài khoản: ' + err.message };
  }
}

function adminAddAccount(adminMaSV, newAccount) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!verifyAdmin_(adminMaSV)) {
      return { success: false, message: 'Quyền quản trị viên bị từ chối.' };
    }

    const maSV = String(newAccount && newAccount.maSV ? newAccount.maSV : '').trim();
    const tenSV = String(newAccount && newAccount.tenSV ? newAccount.tenSV : '').trim();
    const role = String(newAccount && newAccount.vaiTro ? newAccount.vaiTro : 'SV').trim().toUpperCase() === 'ADMIN' ? 'ADMIN' : 'SV';

    if (!maSV || !tenSV) {
      return { success: false, message: 'Vui lòng nhập đầy đủ Mã SV và Tên SV.' };
    }

    const ss = getAppSpreadsheet_();
    const sheet = ss.getSheetByName('Danh_sach_SV');
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 5);
    const values = lastRow >= 2 ? sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues() : [];
    const colIdx = values.length > 0 ? resolveColumnIndices_(values[0], { STT: 0, Ma_SV: 1, Ten_SV: 2, Vai_Tro: 3, Draft_JSON: 4 }) : { STT: 0, Ma_SV: 1, Ten_SV: 2, Vai_Tro: 3, Draft_JSON: 4 };

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][colIdx.Ma_SV] || '').trim().toLowerCase() === maSV.toLowerCase()) {
        return { success: false, message: 'Mã sinh viên "' + maSV + '" đã tồn tại trong hệ thống.' };
      }
    }

    const newRowNumber = sheet.getLastRow() + 1;
    sheet.getRange(newRowNumber, colIdx.STT + 1).setValue(newRowNumber - 1);
    sheet.getRange(newRowNumber, colIdx.Ma_SV + 1).setValue(maSV);
    sheet.getRange(newRowNumber, colIdx.Ten_SV + 1).setValue(tenSV);
    sheet.getRange(newRowNumber, colIdx.Vai_Tro + 1).setValue(role);
    sheet.getRange(newRowNumber, colIdx.Draft_JSON + 1).setValue('');
    SpreadsheetApp.flush();

    return adminGetAccounts(adminMaSV);
  } catch (err) {
    return { success: false, message: 'Lỗi thêm tài khoản: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function adminBulkAddAccounts(adminMaSV, accountsArray) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!verifyAdmin_(adminMaSV)) {
      return { success: false, message: 'Quyền quản trị viên bị từ chối.' };
    }
    if (!Array.isArray(accountsArray) || accountsArray.length === 0) {
      return { success: false, message: 'Danh sách tài khoản trống.' };
    }

    const ss = getAppSpreadsheet_();
    const sheet = ss.getSheetByName('Danh_sach_SV');
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 5);
    const values = lastRow >= 2 ? sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues() : [];
    const colIdx = values.length > 0 ? resolveColumnIndices_(values[0], { STT: 0, Ma_SV: 1, Ten_SV: 2, Vai_Tro: 3, Draft_JSON: 4 }) : { STT: 0, Ma_SV: 1, Ten_SV: 2, Vai_Tro: 3, Draft_JSON: 4 };

    const existingMaSV = {};
    for (let r = 1; r < values.length; r++) {
      const m = String(values[r][colIdx.Ma_SV] || '').trim().toLowerCase();
      if (m) existingMaSV[m] = true;
    }

    let addedCount = 0;
    const rowsToAdd = [];
    accountsArray.forEach(function (acc) {
      const maSV = String(acc && acc.maSV ? acc.maSV : '').trim();
      const tenSV = String(acc && acc.tenSV ? acc.tenSV : '').trim();
      const role = String(acc && acc.vaiTro ? acc.vaiTro : 'SV').trim().toUpperCase() === 'ADMIN' ? 'ADMIN' : 'SV';
      if (maSV && tenSV && !existingMaSV[maSV.toLowerCase()]) {
        existingMaSV[maSV.toLowerCase()] = true;
        rowsToAdd.push([0, maSV, tenSV, role, '']);
        addedCount++;
      }
    });

    if (rowsToAdd.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rowsToAdd.length, 5).setValues(rowsToAdd);
      SpreadsheetApp.flush();
      reindexDanhSachSv_(sheet);
    }

    const res = adminGetAccounts(adminMaSV);
    res.addedCount = addedCount;
    return res;
  } catch (err) {
    return { success: false, message: 'Lỗi nạp hàng loạt tài khoản: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function adminScanExternalSheet(adminMaSV, sheetUrlOrId) {
  try {
    if (!verifyAdmin_(adminMaSV)) {
      return { success: false, message: 'Quyền quản trị viên bị từ chối.' };
    }
    const inputStr = String(sheetUrlOrId || '').trim();
    if (!inputStr) {
      return { success: false, message: 'Vui lòng nhập Link hoặc ID của Google Sheets.' };
    }

    let targetId = inputStr;
    const match = inputStr.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
      targetId = match[1];
    }

    const extSs = SpreadsheetApp.openById(targetId);
    const extSheet = extSs.getSheets()[0];
    const lastRow = extSheet.getLastRow();
    const lastCol = extSheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) {
      return { success: false, message: 'Google Sheets ngoài trống.' };
    }

    const extValues = extSheet.getRange(1, 1, Math.min(lastRow, 100), lastCol).getDisplayValues();
    const headers = extValues[0].map(function (h) { return String(h || '').trim(); });
    const rows = extValues.slice(1);

    return {
      success: true,
      headers: headers,
      rows: rows
    };
  } catch (err) {
    return { success: false, message: 'Lỗi quét Google Sheets: ' + err.message };
  }
}

function adminDeleteAccount(adminMaSV, targetMaSV) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!verifyAdmin_(adminMaSV)) {
      return { success: false, message: 'Quyền quản trị viên bị từ chối.' };
    }
    const target = String(targetMaSV || '').trim().toLowerCase();
    if (!target) {
      return { success: false, message: 'Thiếu Mã SV cần xóa.' };
    }

    const ss = getAppSpreadsheet_();
    const sheet = ss.getSheetByName('Danh_sach_SV');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return adminGetAccounts(adminMaSV);

    const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
    const colIdx = resolveColumnIndices_(values[0], { STT: 0, Ma_SV: 1, Ten_SV: 2, Vai_Tro: 3, Draft_JSON: 4 });

    for (let r = values.length - 1; r >= 1; r--) {
      const m = String(values[r][colIdx.Ma_SV] || '').trim().toLowerCase();
      const role = String(values[r][colIdx.Vai_Tro] || '').trim().toUpperCase();
      if (m === target) {
        if (role === 'ADMIN' && m === String(adminMaSV).trim().toLowerCase()) {
          return { success: false, message: 'Không thể tự xóa tài khoản Quản trị viên đang đăng nhập!' };
        }
        sheet.deleteRow(r + 1);
      }
    }

    SpreadsheetApp.flush();
    reindexDanhSachSv_(sheet);

    return adminGetAccounts(adminMaSV);
  } catch (err) {
    return { success: false, message: 'Lỗi xóa tài khoản: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function adminBulkDeleteStudents(adminMaSV, targetMaSVs) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!verifyAdmin_(adminMaSV)) {
      return { success: false, message: 'Quyền quản trị viên bị từ chối.' };
    }

    const ss = getAppSpreadsheet_();
    const sheet = ss.getSheetByName('Danh_sach_SV');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return adminGetAccounts(adminMaSV);

    const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
    const colIdx = resolveColumnIndices_(values[0], { STT: 0, Ma_SV: 1, Ten_SV: 2, Vai_Tro: 3, Draft_JSON: 4 });

    const isDeleteAll = targetMaSVs === 'ALL' || (Array.isArray(targetMaSVs) && targetMaSVs.length === 0);
    const targetSet = {};
    if (!isDeleteAll && Array.isArray(targetMaSVs)) {
      targetMaSVs.forEach(function (m) { targetSet[String(m).trim().toLowerCase()] = true; });
    }

    for (let r = values.length - 1; r >= 1; r--) {
      const maSV = String(values[r][colIdx.Ma_SV] || '').trim();
      const role = String(values[r][colIdx.Vai_Tro] || '').trim().toUpperCase();

      // STRICTLY PROTECT ADMIN ACCOUNTS
      if (role === 'ADMIN') continue;

      if (isDeleteAll || targetSet[maSV.toLowerCase()]) {
        sheet.deleteRow(r + 1);
      }
    }

    SpreadsheetApp.flush();
    reindexDanhSachSv_(sheet);

    return adminGetAccounts(adminMaSV);
  } catch (err) {
    return { success: false, message: 'Lỗi xóa sinh viên hàng loạt: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function reindexDanhSachSv_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getDisplayValues();
  const colIdx = resolveColumnIndices_(values[0], { STT: 0, Ma_SV: 1, Ten_SV: 2, Vai_Tro: 3, Draft_JSON: 4 });
  for (let r = 1; r < values.length; r++) {
    sheet.getRange(r + 1, colIdx.STT + 1).setValue(r);
  }
  SpreadsheetApp.flush();
}

function adminClearAllStudentsDrafts(adminMaSV) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    if (!verifyAdmin_(adminMaSV)) {
      return { success: false, message: 'Quyền quản trị viên bị từ chối.' };
    }

    const ss = getAppSpreadsheet_();
    const sheet = ss.getSheetByName('Danh_sach_SV');
    if (!sheet) {
      return { success: false, message: 'Không tìm thấy sheet "Danh_sach_SV".' };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 4) {
      return { success: true, clearedCount: 0 };
    }

    const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const colIdx = resolveColumnIndices_(values[0], { STT: 0, Ma_SV: 1, Ten_SV: 2, Vai_Tro: 3, Draft_JSON: 4 });
    const draftColNumber = colIdx.Draft_JSON + 1;
    let clearedCount = 0;

    for (let r = 1; r < values.length; r++) {
      const role = String(values[r][colIdx.Vai_Tro] || '').trim().toUpperCase();
      if (role === 'ADMIN') continue;
      const cell = sheet.getRange(r + 1, draftColNumber);
      cell.setValue('');
      clearedCount++;
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      clearedCount: clearedCount
    };
  } catch (err) {
    return { success: false, message: 'Lỗi xóa dữ liệu nháp sinh viên: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function getInitialDropdownData() {
  try {
    const ss = getAppSpreadsheet_();

    const cuocTauRaw = readTableRecords_(ss, 'Cuoc_Tau', [
      'ID_Lich_Tau',
      'POL',
      'POD',
      'Quoc_Gia',
      'POT',
      'Ten_hang_tau',
      'Ten_tau',
      'So_chuyen',
      'Loai_hinh_van_tai'
    ]);

    const cuocTau = cuocTauRaw.map(function (row) {
      return {
        ID_Lich_Tau: row.ID_Lich_Tau || '',
        idLichTau: row.ID_Lich_Tau || '',
        POL: row.POL || '',
        POD: row.POD || '',
        Quoc_Gia: row.Quoc_Gia || '',
        POT: row.POT || '',
        pot: row.POT || '',
        Ten_hang_tau: row.Ten_hang_tau || '',
        Ten_tau: row.Ten_tau || '',
        So_chuyen: row.So_chuyen || '',
        Loai_hinh_van_tai: row.Loai_hinh_van_tai || '',
        loaiHinhVanTai: row.Loai_hinh_van_tai || ''
      };
    });

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
      'TRADE_AGREEMENT',
      'INSURANCE_TERM',
      'MA_CHU_SO_HUU_CONT',
      'MA_SEAL'
    ]);

    const STANDARD_CURRENCIES_20 = [
      'USD', 'VND', 'EUR', 'JPY', 'CNY', 'GBP', 'SGD', 'KRW', 'AED', 'AUD',
      'BDT', 'BRL', 'CAD', 'IDR', 'INR', 'MYR', 'PHP', 'SAR', 'THB', 'TWD'
    ];
    const rawCurrList = Array.isArray(danhMucChung.CURR) ? danhMucChung.CURR : [];
    const mergedCurrencies = [];
    const seenCurr = {};
    rawCurrList.concat(STANDARD_CURRENCIES_20).forEach(function (code) {
      const clean = String(code || '').trim().toUpperCase();
      if (clean && !seenCurr[clean]) {
        seenCurr[clean] = true;
        mergedCurrencies.push(clean);
      }
    });
    danhMucChung.CURR = mergedCurrencies;

    const contSealList = readContSealPairs_(ss, 'Danh_Muc_chung');

    return {
      success: true,
      logoDataUrl: getBrandLogoBase64(),
      cuocTau: cuocTau,
      loaiHinhXNK: loaiHinhXNK,
      danhMucChung: danhMucChung,
      contSealList: contSealList
    };
  } catch (err) {
    return {
      success: false,
      logoDataUrl: getBrandLogoBase64(),
      message: 'Lỗi tải dữ liệu danh mục: ' + err.message
    };
  }
}

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
      Vai_Tro: 3,
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

function readContSealPairs_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 11) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const colIdx = resolveColumnIndices_(values[0], {
    MA_CHU_SO_HUU_CONT: 10,
    MA_SEAL: 11
  });

  const ownerIdx = colIdx.MA_CHU_SO_HUU_CONT;
  const sealIdx = colIdx.MA_SEAL;
  const pairs = [];
  const seen = {};

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const owner = ownerIdx < row.length ? String(row[ownerIdx] || '').trim() : '';
    const seal = sealIdx < row.length ? String(row[sealIdx] || '').trim() : '';
    if (!owner && !seal) continue;
    const dedupKey = owner + '||' + seal;
    if (!seen[dedupKey]) {
      seen[dedupKey] = true;
      pairs.push({
        owner: owner,
        seal: seal
      });
    }
  }

  return pairs;
}


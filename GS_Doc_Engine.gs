/**
 * GS_Doc_Engine.gs
 * Core Google Docs auto-fill engine supporting BOTH:
 * - Mode A (Folder of 20 Google Docs templates, e.g. folder "CHỨNG TỪ"):
 *   MASTER_TEMPLATE_DOC_ID = Folder ID containing 20 files named by docKey (QUOTATION, BILL_OF_LADING, ...)
 * - Mode B (Single multi-tab Google Doc with 20 Tabs):
 *   MASTER_TEMPLATE_DOC_ID = Google Doc ID containing 20 tabs matching TEMPLATE_MAP[docKey].tabName
 * Runtime: Google Apps Script (V8)
 */

/**
 * Tạo 1 hoặc nhiều chứng từ từ thư mục mẫu 20 file Google Docs hoặc 1 file Google Docs tổng hợp
 * (MASTER_TEMPLATE_DOC_ID) và trả về link Google Docs + link tải PDF.
 *
 * @param {string} maSV
 * @param {Array<string>|string} docKeysArray
 * @param {{shared?: Object, items?: Array<Object>, containers?: Array<Object>}|string} payload
 * @return {{success: boolean, results: Array<{docKey: string, docName: string, tabName: string, docUrl: string, pdfUrl: string}>, errors?: Array<{docKey: string, message: string}>, warning?: string, message?: string}}
 */
function generateDocuments(maSV, docKeysArray, payload) {
  try {
    const rawKeys = Array.isArray(docKeysArray)
      ? docKeysArray
      : (docKeysArray ? [String(docKeysArray)] : []);

    // Loại bỏ key rỗng hoặc trùng lặp nhưng giữ nguyên thứ tự
    const keys = [];
    rawKeys.forEach(function (k) {
      const cleanKey = String(k || '').trim();
      if (cleanKey && keys.indexOf(cleanKey) === -1) {
        keys.push(cleanKey);
      }
    });

    if (keys.length === 0) {
      return {
        success: false,
        results: [],
        message: 'Vui lòng chọn ít nhất một loại chứng từ để xuất.'
      };
    }

    const parsedPayload =
      typeof payload === 'string'
        ? JSON.parse(payload || '{}')
        : (payload || {});

    const shared = buildEnrichedShared_(parsedPayload);
    const items = Array.isArray(parsedPayload.items) ? parsedPayload.items : [];
    const containers = Array.isArray(parsedPayload.containers) ? parsedPayload.containers : [];

    const validated = validateMasterTemplateAndFolder_();
    const outputFolder = validated.outputFolder;
    const suffix = resolveDocFilenameSuffix_(shared, maSV);

    // =========================================================================
    // CHẾ ĐỘ A: MASTER_TEMPLATE_DOC_ID là THƯ MỤC chứa 20 file Google Docs
    // =========================================================================
    if (validated.mode === 'FOLDER') {
      const results = [];
      const errors = [];

      for (let i = 0; i < keys.length; i++) {
        const docKey = keys[i];
        const tplConfig = TEMPLATE_MAP[docKey];

        if (!tplConfig) {
          errors.push({
            docKey: docKey,
            message: 'Không tìm thấy cấu hình cho "' + docKey + '" trong TEMPLATE_MAP.'
          });
          continue;
        }

        let copyFile = null;
        try {
          const tplFile = findTemplateFileInFolder_(validated.templateFolder, docKey, tplConfig);
          if (!tplFile) {
            errors.push({
              docKey: docKey,
              message:
                'Không tìm thấy file Google Docs mẫu "' +
                docKey +
                '" trong thư mục "' +
                validated.templateFolder.getName() +
                '".'
            });
            continue;
          }

          const copyFileName = docKey + '_' + suffix;
          copyFile = tplFile.makeCopy(copyFileName, outputFolder);
          const copyFileId = copyFile.getId();

          const doc = DocumentApp.openById(copyFileId);
          processSingleDocumentObject_(doc, items, containers, shared);
          doc.saveAndClose();

          results.push({
            docKey: docKey,
            docName: copyFileName,
            tabName: tplFile.getName(),
            docUrl: copyFile.getUrl(),
            pdfUrl: 'https://docs.google.com/document/d/' + copyFileId + '/export?format=pdf'
          });
        } catch (fileErr) {
          if (copyFile) {
            try {
              copyFile.setTrashed(true);
            } catch (e) {}
          }
          errors.push({
            docKey: docKey,
            message: 'Lỗi xử lý "' + docKey + '": ' + fileErr.message
          });
        }
      }

      if (results.length === 0) {
        return {
          success: false,
          results: [],
          errors: errors,
          message: errors.length > 0 ? errors[0].message : 'Không tìm thấy file mẫu phù hợp để xuất chứng từ.'
        };
      }

      const folderResponse = {
        success: true,
        results: results
      };
      if (errors.length > 0) {
        folderResponse.errors = errors;
      }
      return folderResponse;
    }

    // =========================================================================
    // CHẾ ĐỘ B: MASTER_TEMPLATE_DOC_ID là 1 FILE GOOGLE DOCS nhiều Tab
    // =========================================================================
    const masterFile = validated.masterFile;
    const copyFileName =
      keys.length === 1
        ? keys[0] + '_' + suffix
        : 'BO_CHUNG_TU_' + suffix;

    const copyFile = masterFile.makeCopy(copyFileName, outputFolder);
    const copyFileId = copyFile.getId();
    const docUrl = copyFile.getUrl();
    const pdfUrl = 'https://docs.google.com/document/d/' + copyFileId + '/export?format=pdf';

    const doc = DocumentApp.openById(copyFileId);
    const topLevelTabs = typeof doc.getTabs === 'function' ? doc.getTabs() : [];
    const allTabs = collectAllTabsFlat_(topLevelTabs);

    const selectedTabIds = [];
    const results = [];
    const errors = [];

    for (let i = 0; i < keys.length; i++) {
      const docKey = keys[i];
      const tplConfig = TEMPLATE_MAP[docKey];

      if (!tplConfig || !tplConfig.tabName) {
        errors.push({
          docKey: docKey,
          message: 'Không tìm thấy cấu hình tabName cho "' + docKey + '" trong TEMPLATE_MAP.'
        });
        continue;
      }

      const targetTab = findTabByTitleInsensitive_(allTabs, tplConfig.tabName);
      if (!targetTab) {
        errors.push({
          docKey: docKey,
          message: 'Không tìm thấy Tab "' + tplConfig.tabName + '" trong file Google Docs mẫu.'
        });
        continue;
      }

      try {
        const docTab = targetTab.asDocumentTab();
        const body = docTab.getBody();
        const sections = collectDocumentTabSections_(docTab);

        processDynamicTableRows(body, items, containers);
        replaceGlobalTags(sections, shared);
        sanitizeRemainingTags_(sections);

        const tabId = targetTab.getId();
        if (tabId && selectedTabIds.indexOf(tabId) === -1) {
          selectedTabIds.push(tabId);
        }

        results.push({
          docKey: docKey,
          docName: keys.length === 1 ? copyFileName : copyFileName + ' [' + tplConfig.tabName + ']',
          tabName: tplConfig.tabName,
          docUrl: docUrl,
          pdfUrl: pdfUrl
        });
      } catch (tabErr) {
        errors.push({
          docKey: docKey,
          message: 'Lỗi xử lý Tab "' + tplConfig.tabName + '": ' + tabErr.message
        });
      }
    }

    doc.saveAndClose();

    if (results.length === 0) {
      try {
        copyFile.setTrashed(true);
      } catch (trashErr) {}

      return {
        success: false,
        results: [],
        errors: errors,
        message: errors.length > 0 ? errors[0].message : 'Không tìm thấy Tab phù hợp để xuất chứng từ.'
      };
    }

    const cleanupStatus = removeUnselectedTabsViaDocsApi_(copyFileId, topLevelTabs, selectedTabIds);

    const response = {
      success: true,
      results: results
    };
    if (errors.length > 0) {
      response.errors = errors;
    }
    if (cleanupStatus && cleanupStatus.warning) {
      response.warning = cleanupStatus.warning;
    }

    return response;
  } catch (err) {
    return {
      success: false,
      results: [],
      message: formatEngineErrorMessage_('Lỗi hệ thống tạo chứng từ', err)
    };
  }
}

/**
 * Tạo bản xem trước (Live Preview) cho 1 chứng từ từ file mẫu (MASTER_TEMPLATE_DOC_ID).
 * Sao chép tạm vào DriveApp.getRootFolder() để tránh lỗi phân quyền trên thư mục chia sẻ OUTPUT_FOLDER_ID,
 * giữ nguyên 100% định dạng gốc, xuất PDF dạng base64, xóa ngay file tạm và trả về { success: true, base64Pdf }.
 *
 * @param {string} maSV
 * @param {string} docKey
 * @param {{shared?: Object, items?: Array<Object>, containers?: Array<Object>}|string} payload
 * @return {{success: boolean, docKey?: string, docName?: string, tabName?: string, base64Pdf?: string, pdfBase64?: string, mimeType?: string, warning?: string, message?: string}}
 */
function previewDocument(maSV, docKey, payload) {
  let tempDocFile = null;

  try {
    const cleanKey = String(docKey || '').trim();
    if (!cleanKey) {
      return {
        success: false,
        message: 'Thiếu mã chứng từ để xem trước.'
      };
    }

    const tplConfig = TEMPLATE_MAP[cleanKey];
    if (!tplConfig) {
      return {
        success: false,
        message: 'Không tìm thấy cấu hình cho "' + cleanKey + '" trong TEMPLATE_MAP.'
      };
    }

    const parsedPayload =
      typeof payload === 'string'
        ? JSON.parse(payload || '{}')
        : (payload || {});

    const shared = buildEnrichedShared_(parsedPayload);
    const items = Array.isArray(parsedPayload.items) ? parsedPayload.items : [];
    const containers = Array.isArray(parsedPayload.containers) ? parsedPayload.containers : [];

    const safeMaSV = String(maSV == null ? 'GUEST' : maSV)
      .trim()
      .replace(/[^A-Za-z0-9_-]/g, '_') || 'GUEST';

    const rootFolder = DriveApp.getRootFolder();
    const tempDocName = '__PREVIEW_TEMP_' + safeMaSV + '_' + Date.now();
    const validated = validateMasterTemplateAndFolder_(false);

    let pdfBlob = null;
    let cleanupStatus = null;
    let resolvedTemplateLabel = tplConfig.tabName || cleanKey;

    if (validated.mode === 'FOLDER') {
      const tplFile = findTemplateFileInFolder_(validated.templateFolder, cleanKey, tplConfig);
      if (!tplFile) {
        return {
          success: false,
          message:
            'Không tìm thấy file Google Docs mẫu "' +
            cleanKey +
            '" trong thư mục "' +
            validated.templateFolder.getName() +
            '".'
        };
      }

      resolvedTemplateLabel = tplFile.getName();
      tempDocFile = tplFile.makeCopy(tempDocName, rootFolder);

      const doc = DocumentApp.openById(tempDocFile.getId());
      processSingleDocumentObject_(doc, items, containers, shared);
      doc.saveAndClose();

      pdfBlob = tempDocFile.getAs('application/pdf');
    } else {
      const masterFile = validated.masterFile;
      tempDocFile = masterFile.makeCopy(tempDocName, rootFolder);
      const tempDocId = tempDocFile.getId();

      const doc = DocumentApp.openById(tempDocId);
      const topLevelTabs = typeof doc.getTabs === 'function' ? doc.getTabs() : [];
      const allTabs = collectAllTabsFlat_(topLevelTabs);
      const targetTab = findTabByTitleInsensitive_(allTabs, tplConfig.tabName);

      if (!targetTab) {
        doc.saveAndClose();
        try {
          tempDocFile.setTrashed(true);
          tempDocFile = null;
        } catch (e) {}

        return {
          success: false,
          message: 'Không tìm thấy Tab "' + tplConfig.tabName + '" trong file Google Docs mẫu.'
        };
      }

      const docTab = targetTab.asDocumentTab();
      const body = docTab.getBody();
      const sections = collectDocumentTabSections_(docTab);

      processDynamicTableRows(body, items, containers);
      replaceGlobalTags(sections, shared);
      sanitizeRemainingTags_(sections);

      const tabId = targetTab.getId();
      doc.saveAndClose();

      cleanupStatus = removeUnselectedTabsViaDocsApi_(
        tempDocId,
        topLevelTabs,
        tabId ? [tabId] : []
      );

      if (cleanupStatus && cleanupStatus.warning && tabId) {
        try {
          const exportUrl =
            'https://docs.google.com/document/d/' +
            tempDocId +
            '/export?format=pdf&tab=' +
            encodeURIComponent(tabId);
          const fetchRes = UrlFetchApp.fetch(exportUrl, {
            headers: {
              Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
            },
            muteHttpExceptions: true
          });
          if (fetchRes.getResponseCode() === 200) {
            pdfBlob = fetchRes.getBlob();
          }
        } catch (fetchErr) {}
      }

      if (!pdfBlob) {
        pdfBlob = tempDocFile.getAs('application/pdf');
      }
    }

    const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());

    if (tempDocFile) {
      try {
        tempDocFile.setTrashed(true);
        tempDocFile = null;
      } catch (trashDocErr) {}
    }

    const response = {
      success: true,
      docKey: cleanKey,
      docName: tplConfig.name || cleanKey,
      tabName: resolvedTemplateLabel,
      base64Pdf: base64Pdf,
      pdfBase64: base64Pdf,
      mimeType: 'application/pdf'
    };

    if (cleanupStatus && cleanupStatus.warning) {
      response.warning = cleanupStatus.warning;
    }

    return response;
  } catch (err) {
    if (tempDocFile) {
      try {
        tempDocFile.setTrashed(true);
        tempDocFile = null;
      } catch (e) {}
    }
    return {
      success: false,
      message: formatEngineErrorMessage_('Lỗi tạo bản xem trước chứng từ', err)
    };
  }
}

/**
 * Kiểm tra MASTER_TEMPLATE_DOC_ID (chấp nhận cả THƯ MỤC chứa 20 file Google Docs
 * lẫn 1 file Google Docs nhiều Tab) và tùy chọn kiểm tra OUTPUT_FOLDER_ID.
 * @param {boolean=} requireOutputFolder
 * @return {{mode: string, masterId: string, templateFolder?: GoogleAppsScript.Drive.Folder, masterFile?: GoogleAppsScript.Drive.File, outputFolderId: string, outputFolder: GoogleAppsScript.Drive.Folder|null}}
 */
function validateMasterTemplateAndFolder_(requireOutputFolder) {
  const shouldCheckOutput = requireOutputFolder !== false;
  const masterId = String(typeof MASTER_TEMPLATE_DOC_ID !== 'undefined' ? MASTER_TEMPLATE_DOC_ID : '').trim();
  const outputFolderId = String(typeof OUTPUT_FOLDER_ID !== 'undefined' ? OUTPUT_FOLDER_ID : '').trim();

  if (!masterId || masterId.indexOf('YOUR_') === 0) {
    throw new Error('Chưa cấu hình MASTER_TEMPLATE_DOC_ID trong GS_Config_Templates.gs.');
  }
  if (shouldCheckOutput && (!outputFolderId || outputFolderId.indexOf('YOUR_') === 0)) {
    throw new Error('Chưa cấu hình OUTPUT_FOLDER_ID trong GS_Config_Templates.gs.');
  }

  const outputFolder = shouldCheckOutput ? DriveApp.getFolderById(outputFolderId) : null;
  let masterItem = null;
  let mimeType = '';
  try {
    masterItem = DriveApp.getFileById(masterId);
    mimeType = masterItem.getMimeType();
  } catch (fileLookupErr) {
    try {
      const folderFallback = DriveApp.getFolderById(masterId);
      return {
        mode: 'FOLDER',
        masterId: masterId,
        templateFolder: folderFallback,
        outputFolderId: outputFolderId,
        outputFolder: outputFolder
      };
    } catch (folderLookupErr) {
      throw fileLookupErr;
    }
  }

  if (mimeType === MimeType.FOLDER) {
    const templateFolder = DriveApp.getFolderById(masterId);
    return {
      mode: 'FOLDER',
      masterId: masterId,
      templateFolder: templateFolder,
      outputFolderId: outputFolderId,
      outputFolder: outputFolder
    };
  }

  if (mimeType === MimeType.GOOGLE_DOCS) {
    return {
      mode: 'SINGLE_MULTI_TAB_DOC',
      masterId: masterId,
      masterFile: masterItem,
      outputFolderId: outputFolderId,
      outputFolder: outputFolder
    };
  }

  if (
    mimeType === MimeType.MICROSOFT_WORD ||
    mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
    mimeType === 'application/msword'
  ) {
    throw new Error(
      'File "' +
        masterItem.getName() +
        '" đang ở định dạng Word (.docx), không phải Google Docs gốc. Hãy mở file trên Google Drive -> Tệp -> Lưu dưới dạng Google Tài liệu (Save as Google Docs).'
    );
  }

  throw new Error(
    'MASTER_TEMPLATE_DOC_ID đang trỏ tới "' +
      masterItem.getName() +
      '" (loại: ' +
      mimeType +
      '). Hãy điền ID của thư mục "CHỨNG TỪ" hoặc file Google Docs tổng hợp.'
  );
}

/**
 * Tìm file Google Docs mẫu tương ứng với docKey bên trong thư mục mẫu (ví dụ thư mục "CHỨNG TỪ").
 * @param {GoogleAppsScript.Drive.Folder} templateFolder
 * @param {string} docKey
 * @param {{name?: string, tabName?: string}} tplConfig
 * @return {GoogleAppsScript.Drive.File|null}
 */
function findTemplateFileInFolder_(templateFolder, docKey, tplConfig) {
  if (!templateFolder) return null;

  const candidates = [docKey];
  if (tplConfig && tplConfig.tabName && candidates.indexOf(tplConfig.tabName) === -1) {
    candidates.push(tplConfig.tabName);
  }

  const normCandidates = candidates.map(function (c) {
    return normalizeTemplateLookupName_(c);
  });

  const filesIter = templateFolder.getFiles();
  let wordFallbackFile = null;

  while (filesIter.hasNext()) {
    const file = filesIter.next();
    const rawName = String(file.getName() || '').trim();
    const strippedName = rawName.replace(/\.(docx?|gdoc)$/i, '').trim();
    const normName = normalizeTemplateLookupName_(strippedName);

    const exactMatch = candidates.some(function (cand) {
      const lowerCand = String(cand || '').trim().toLowerCase();
      return rawName.toLowerCase() === lowerCand || strippedName.toLowerCase() === lowerCand;
    });
    const normMatch = normCandidates.indexOf(normName) !== -1;

    if (exactMatch || normMatch) {
      const mime = file.getMimeType();
      if (mime === MimeType.GOOGLE_DOCS) {
        return file;
      }
      if (
        mime === MimeType.MICROSOFT_WORD ||
        mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
        mime === 'application/msword'
      ) {
        wordFallbackFile = file;
      }
    }
  }

  if (wordFallbackFile) {
    throw new Error(
      'File mẫu "' +
        wordFallbackFile.getName() +
        '" đang ở định dạng Word (.docx). Hãy mở file này trên Drive -> Tệp -> Lưu dưới dạng Google Tài liệu (Google Docs).'
    );
  }

  return null;
}

/**
 * Chuẩn hóa tên file mẫu để so khớp linh hoạt (không dấu, không phân biệt hoa thường, bỏ dấu cách/gạch dưới).
 * @param {string} name
 * @return {string}
 */
function normalizeTemplateLookupName_(name) {
  return String(name || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

/**
 * Xử lý 1 file Google Docs đơn lẻ (giữ nguyên 100% định dạng gốc):
 * hỗ trợ cả tài liệu thường lẫn tài liệu có Tab.
 * @param {GoogleAppsScript.Document.Document} doc
 * @param {Array<Object>} items
 * @param {Array<Object>} containers
 * @param {Object} shared
 */
function processSingleDocumentObject_(doc, items, containers, shared) {
  if (typeof doc.getTabs === 'function') {
    const topTabs = doc.getTabs();
    const allTabs = collectAllTabsFlat_(topTabs);
    if (allTabs.length > 0) {
      for (let t = 0; t < allTabs.length; t++) {
        const docTab = allTabs[t].asDocumentTab();
        const body = docTab.getBody();
        const sections = collectDocumentTabSections_(docTab);
        processDynamicTableRows(body, items, containers);
        replaceGlobalTags(sections, shared);
        sanitizeRemainingTags_(sections);
      }
      return;
    }
  }

  const body = doc.getBody();
  const sections = collectSingleDocSections_(doc);
  processDynamicTableRows(body, items, containers);
  replaceGlobalTags(sections, shared);
  sanitizeRemainingTags_(sections);
}

/**
 * Lấy Body + Header + Footer của 1 file Google Docs đơn lẻ.
 * @param {GoogleAppsScript.Document.Document} doc
 * @return {Array<GoogleAppsScript.Document.Element>}
 */
function collectSingleDocSections_(doc) {
  const sections = [];
  const seen = [];

  function addSection(el) {
    if (!el) return;
    if (seen.indexOf(el) === -1) {
      seen.push(el);
      sections.push(el);
    }
  }

  const body = doc.getBody();
  addSection(body);
  addSection(doc.getHeader());
  addSection(doc.getFooter());

  const root = body && typeof body.getParent === 'function' ? body.getParent() : null;
  if (root && typeof root.getNumChildren === 'function') {
    const numChildren = root.getNumChildren();
    for (let i = 0; i < numChildren; i++) {
      const child = root.getChild(i);
      const type = child.getType();
      if (
        type === DocumentApp.ElementType.HEADER_SECTION ||
        type === DocumentApp.ElementType.FOOTER_SECTION
      ) {
        addSection(child);
      }
    }
  }

  return sections;
}

/**
 * Chạy hàm này 1 lần trực tiếp trong trình soạn thảo Apps Script (Run -> authorizePermissions)
 * để kích hoạt màn hình cấp quyền Drive, Docs, Sheets và UrlFetch cho Web App trước khi Deploy lại.
 */
function authorizePermissions() {
  const ss = getAppSpreadsheet_();
  const sheetTitle = ss.getName();

  DriveApp.getRootFolder().getName();
  ScriptApp.getOAuthToken();
  UrlFetchApp.getRequest('https://docs.google.com');

  const validated = validateMasterTemplateAndFolder_();
  const folderName = validated.outputFolder.getName();

  if (validated.mode === 'FOLDER') {
    const templateFolderName = validated.templateFolder.getName();
    const filesIter = validated.templateFolder.getFiles();
    let googleDocCount = 0;
    let firstDocId = '';

    while (filesIter.hasNext()) {
      const f = filesIter.next();
      if (f.getMimeType() === MimeType.GOOGLE_DOCS) {
        googleDocCount++;
        if (!firstDocId) {
          firstDocId = f.getId();
        }
      }
    }

    if (firstDocId) {
      DocumentApp.openById(firstDocId).getId();
    }

    Logger.log(
      'Authorized OK (FOLDER MODE) -> Sheet: ' +
        sheetTitle +
        ' | Template Folder: ' +
        templateFolderName +
        ' (Google Docs: ' +
        googleDocCount +
        ') | Output Folder: ' +
        folderName
    );

    return {
      success: true,
      mode: 'FOLDER',
      sheetTitle: sheetTitle,
      templateFolderName: templateFolderName,
      googleDocCount: googleDocCount,
      folderName: folderName
    };
  }

  const templateName = validated.masterFile.getName();
  const openedDoc = DocumentApp.openById(validated.masterId);
  const tabCount = typeof openedDoc.getTabs === 'function' ? openedDoc.getTabs().length : 0;
  const docsApiEnabled = typeof Docs !== 'undefined' && Docs.Documents && typeof Docs.Documents.batchUpdate === 'function';

  Logger.log(
    'Authorized OK (SINGLE DOC MODE) -> Sheet: ' +
      sheetTitle +
      ' | Template: ' +
      templateName +
      ' (Tabs: ' +
      tabCount +
      ') | Output Folder: ' +
      folderName +
      ' | Docs API enabled: ' +
      docsApiEnabled
  );

  return {
    success: true,
    mode: 'SINGLE_MULTI_TAB_DOC',
    sheetTitle: sheetTitle,
    templateName: templateName,
    tabCount: tabCount,
    folderName: folderName,
    docsApiEnabled: docsApiEnabled
  };
}

/**
 * Chuẩn hóa thông báo lỗi và gợi ý xử lý khi gặp lỗi phân quyền DriveApp / Web App.
 * @param {string} prefix
 * @param {*} err
 * @return {string}
 */
function formatEngineErrorMessage_(prefix, err) {
  const rawMsg = err && err.message ? String(err.message) : String(err || '');
  const lower = rawMsg.toLowerCase();
  if (
    lower.indexOf('truy cập bị từ chối') !== -1 ||
    lower.indexOf('access denied') !== -1 ||
    lower.indexOf('specified permissions') !== -1 ||
    lower.indexOf('không có quyền') !== -1
  ) {
    return (
      prefix +
      ': ' +
      rawMsg +
      ' — Cách khắc phục: (1) Trong Apps Script Editor, chọn hàm authorizePermissions và bấm Run (Chạy) để cấp quyền Drive/Docs; (2) Khi Deploy Web App, chọn "Execute as: Me (Tôi)"; (3) Tạo New version (Phiên bản mới) sau khi lưu code.'
    );
  }
  return prefix + ': ' + rawMsg;
}

/**
 * Đưa vào thùng rác các file trùng tên trong thư mục đích (dùng để ghi đè file tạm theo maSV).
 * @param {GoogleAppsScript.Drive.Folder} folder
 * @param {string} fileName
 */
function trashFilesByName_(folder, fileName) {
  if (!folder || !fileName) return;
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    try {
      files.next().setTrashed(true);
    } catch (e) {}
  }
}

/**
 * Xóa các Tab không nằm trong danh sách được chọn khỏi file Google Docs bản sao
 * thông qua Google Docs API (Docs.Documents.batchUpdate với request deleteTab).
 *
 * HƯỚNG DẪN BẬT DỊCH VỤ TRONG GOOGLE APPS SCRIPT:
 * 1. Mở trình soạn thảo Apps Script (script.google.com).
 * 2. Ở cột bên trái mục "Dịch vụ (Services)", nhấn dấu "+".
 * 3. Chọn "Google Docs API" (định danh mặc định là "Docs", phiên bản v1) -> nhấn "Thêm (Add)".
 * Nếu chưa bật dịch vụ này, khối try...catch bên dưới sẽ giữ nguyên file kết quả mà không làm lỗi tiến trình xuất.
 *
 * @param {string} documentId
 * @param {Array<GoogleAppsScript.Document.Tab>} topLevelTabs
 * @param {Array<string>} selectedTabIds
 * @return {{deletedCount: number, warning?: string}}
 */
function removeUnselectedTabsViaDocsApi_(documentId, topLevelTabs, selectedTabIds) {
  if (!selectedTabIds || selectedTabIds.length === 0) {
    return { deletedCount: 0 };
  }

  try {
    if (typeof Docs === 'undefined' || !Docs.Documents || typeof Docs.Documents.batchUpdate !== 'function') {
      return {
        deletedCount: 0,
        warning: 'Chưa bật dịch vụ "Google Docs API" (Docs) trong Services của Apps Script nên các Tab khác chưa được xóa tự động.'
      };
    }

    const allTabs = collectAllTabsFlat_(topLevelTabs);
    const deleteRequests = [];

    // Duyệt từ cuối lên đầu để đảm bảo an toàn thứ tự tab và luôn giữ lại ít nhất các tab đã chọn
    for (let i = allTabs.length - 1; i >= 0; i--) {
      const tab = allTabs[i];
      const tabId = tab.getId();
      if (tabId && selectedTabIds.indexOf(tabId) === -1) {
        // Nếu có tab con được chọn nằm bên trong tab cha này thì không xóa tab cha
        if (hasSelectedDescendantTab_(tab, selectedTabIds)) {
          continue;
        }
        deleteRequests.push({
          deleteTab: {
            tabId: tabId
          }
        });
      }
    }

    // Google Docs API không cho phép xóa toàn bộ mọi tab của tài liệu (phải còn ít nhất 1 tab)
    if (deleteRequests.length > 0 && deleteRequests.length < allTabs.length) {
      Docs.Documents.batchUpdate({ requests: deleteRequests }, documentId);
    }

    return { deletedCount: deleteRequests.length };
  } catch (apiErr) {
    return {
      deletedCount: 0,
      warning: 'Không thể xóa các Tab không chọn qua Google Docs API: ' + apiErr.message
    };
  }
}

/**
 * Nhân bản dòng bảng động từ dưới lên trên (Bottom-Up Target-Row Cloning).
 * Chỉ nhân bản TableRow nếu text của dòng chứa ít nhất 1 thẻ thuộc ITEM_TAGS, CONT_TAGS hoặc TAX_TAGS.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @param {Array<Object>} items
 * @param {Array<Object>} containers
 */
function processDynamicTableRows(body, items, containers) {
  const safeItems = Array.isArray(items) ? items : [];
  const safeContainers = Array.isArray(containers) ? containers : [];
  const tables = body.getTables();

  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];

    // Duyệt hàng từ dưới lên trên để việc chèn/xóa không làm lệch chỉ số các hàng phía trên
    for (let r = table.getNumRows() - 1; r >= 0; r--) {
      const row = table.getRow(r);
      const rowText = row.getText();
      const tagsInRow = extractTagNames_(rowText);

      if (!hasDynamicRowTag_(tagsInRow)) {
        continue;
      }

      const isContainerRow = tagsInRow.indexOf('CONT') !== -1;
      const boundArray = isContainerRow ? safeContainers : safeItems;
      const count = boundArray.length;

      if (count === 0) {
        if (table.getNumRows() > 1) {
          table.removeRow(r);
        } else {
          // Bảng chỉ còn 1 dòng duy nhất: xóa sạch thẻ thay vì removeRow để tránh lỗi Google Docs
          row.replaceText('\\{\\{[A-Z0-9_]+\\}\\}', '');
        }
        continue;
      }

      // Chèn N - 1 bản sao ngay bên dưới dòng mẫu r
      for (let i = 1; i < count; i++) {
        table.insertTableRow(r + i, row.copy());
      }

      // Thay thế dữ liệu cho tất cả N dòng (từ r đến r + count - 1)
      for (let idx = 0; idx < count; idx++) {
        const targetRow = table.getRow(r + idx);
        const rowData = boundArray[idx] || {};
        replaceRowTags_(targetRow, rowData, tagsInRow, idx);
      }
    }
  }
}

/**
 * Thay thế toàn bộ biến đơn {{TAG}} trên Body, Header và Footer của Tab
 * bằng dữ liệu từ payload.shared (giữ nguyên số 0 ở đầu dưới dạng chuỗi).
 *
 * @param {Array<GoogleAppsScript.Document.Element>} sections
 * @param {Object} shared
 */
function replaceGlobalTags(sections, shared) {
  const keys = Object.keys(shared).sort(function (a, b) {
    return b.length - a.length;
  });

  for (let s = 0; s < sections.length; s++) {
    const section = sections[s];
    for (let k = 0; k < keys.length; k++) {
      const key = keys[k];
      const rawVal = shared[key];
      const strVal = rawVal == null ? '' : String(rawVal);
      section.replaceText('\\{\\{' + escapeRegex_(key) + '\\}\\}', strVal);
    }
  }
}

/* ==========================================================================
 * CÁC HÀM HỖ TRỢ NỘI BỘ (PRIVATE HELPERS)
 * ========================================================================== */

/**
 * Trải phẳng danh sách Tab (bao gồm cả tab con nếu có).
 * @param {Array<GoogleAppsScript.Document.Tab>} tabs
 * @return {Array<GoogleAppsScript.Document.Tab>}
 */
function collectAllTabsFlat_(tabs) {
  const flat = [];
  if (!Array.isArray(tabs)) return flat;

  for (let i = 0; i < tabs.length; i++) {
    const tab = tabs[i];
    flat.push(tab);
    if (typeof tab.getChildTabs === 'function') {
      const children = tab.getChildTabs();
      if (children && children.length > 0) {
        const flatChildren = collectAllTabsFlat_(children);
        for (let c = 0; c < flatChildren.length; c++) {
          flat.push(flatChildren[c]);
        }
      }
    }
  }
  return flat;
}

/**
 * Kiểm tra xem một Tab cha có chứa Tab con nào đang được chọn hay không.
 * @param {GoogleAppsScript.Document.Tab} tab
 * @param {Array<string>} selectedTabIds
 * @return {boolean}
 */
function hasSelectedDescendantTab_(tab, selectedTabIds) {
  if (typeof tab.getChildTabs !== 'function') return false;
  const children = tab.getChildTabs() || [];
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (selectedTabIds.indexOf(child.getId()) !== -1 || hasSelectedDescendantTab_(child, selectedTabIds)) {
      return true;
    }
  }
  return false;
}

/**
 * Tìm Tab theo tiêu đề tab.getTitle().trim(), so sánh không phân biệt hoa thường.
 * @param {Array<GoogleAppsScript.Document.Tab>} allTabs
 * @param {string} targetTabName
 * @return {GoogleAppsScript.Document.Tab|null}
 */
function findTabByTitleInsensitive_(allTabs, targetTabName) {
  const normalizedTarget = String(targetTabName || '').trim().toLowerCase();
  if (!normalizedTarget) return null;

  for (let i = 0; i < allTabs.length; i++) {
    const tab = allTabs[i];
    const title = String(tab.getTitle() || '').trim().toLowerCase();
    if (title === normalizedTarget) {
      return tab;
    }
  }
  return null;
}

/**
 * Lấy danh sách các vùng nội dung (Body, Header, Footer) thuộc một DocumentTab.
 * @param {GoogleAppsScript.Document.DocumentTab} docTab
 * @return {Array<GoogleAppsScript.Document.Element>}
 */
function collectDocumentTabSections_(docTab) {
  const sections = [];
  const seen = [];

  function addSection(el) {
    if (!el) return;
    if (seen.indexOf(el) === -1) {
      seen.push(el);
      sections.push(el);
    }
  }

  const body = docTab.getBody();
  addSection(body);
  addSection(docTab.getHeader());
  addSection(docTab.getFooter());

  const tabRoot = body && typeof body.getParent === 'function' ? body.getParent() : null;
  if (tabRoot && typeof tabRoot.getNumChildren === 'function') {
    const numChildren = tabRoot.getNumChildren();
    for (let i = 0; i < numChildren; i++) {
      const child = tabRoot.getChild(i);
      const type = child.getType();
      if (
        type === DocumentApp.ElementType.HEADER_SECTION ||
        type === DocumentApp.ElementType.FOOTER_SECTION
      ) {
        addSection(child);
      }
    }
  }

  return sections;
}

/**
 * Tạo bản sao payload.shared và tự động sinh các biến tổng hợp:
 * - VES_VOY: VESSEL + " / " + VOYAGE (nếu cả 2 đều có)
 * - CONTAINER_SEAL_DETAILS: ghép payload.containers dạng "CONT / CTYPE / SEAL" nối bằng \n
 *
 * @param {Object} payload
 * @return {Object}
 */
function buildEnrichedShared_(payload) {
  const sourceShared = (payload && payload.shared && typeof payload.shared === 'object')
    ? payload.shared
    : {};
  const shared = {};

  Object.keys(sourceShared).forEach(function (key) {
    const val = sourceShared[key];
    shared[key] = val == null ? '' : String(val);
  });

  const vessel = String(shared.VESSEL || '').trim();
  const voyage = String(shared.VOYAGE || '').trim();
  if (vessel && voyage) {
    shared.VES_VOY = vessel + ' / ' + voyage;
  }

  const containers = Array.isArray(payload && payload.containers) ? payload.containers : [];
  if (containers.length > 0) {
    const lines = containers
      .map(function (c) {
        if (!c || typeof c !== 'object') return '';
        const cont = String(c.CONT == null ? '' : c.CONT).trim();
        const ctype = String(c.CTYPE == null ? '' : c.CTYPE).trim();
        const seal = String(c.SEAL == null ? '' : c.SEAL).trim();
        return [cont, ctype, seal]
          .filter(function (part) {
            return part !== '';
          })
          .join(' / ');
      })
      .filter(function (line) {
        return line !== '';
      });

    shared.CONTAINER_SEAL_DETAILS = lines.join('\n');
  } else if (shared.CONTAINER_SEAL_DETAILS == null) {
    shared.CONTAINER_SEAL_DETAILS = '';
  }

  return shared;
}

/**
 * Xác định hậu tố tên file [INV_NO hoặc BK_NO hoặc maSV].
 * @param {Object} shared
 * @param {string} maSV
 * @return {string}
 */
function resolveDocFilenameSuffix_(shared, maSV) {
  const invNo = String(shared.INV_NO || '').trim();
  if (invNo) return invNo;

  const bkNo = String(shared.BK_NO || '').trim();
  if (bkNo) return bkNo;

  const studentId = String(maSV == null ? '' : maSV).trim();
  return studentId || 'DOC';
}

/**
 * Trích xuất danh sách tên thẻ {{TAG}} duy nhất trong chuỗi text.
 * @param {string} text
 * @return {Array<string>}
 */
function extractTagNames_(text) {
  const tags = [];
  if (!text) return tags;

  const regex = /\{\{([A-Z0-9_]+)\}\}/g;
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (tags.indexOf(match[1]) === -1) {
      tags.push(match[1]);
    }
  }
  return tags;
}

/**
 * Trả về true nếu dòng chứa ít nhất một thẻ thuộc ITEM_TAGS, CONT_TAGS hoặc TAX_TAGS.
 * @param {Array<string>} tagsInRow
 * @return {boolean}
 */
function hasDynamicRowTag_(tagsInRow) {
  for (let i = 0; i < tagsInRow.length; i++) {
    const tag = tagsInRow[i];
    if (
      ITEM_TAGS.indexOf(tag) !== -1 ||
      CONT_TAGS.indexOf(tag) !== -1 ||
      TAX_TAGS.indexOf(tag) !== -1
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Thay thế các thẻ trong 1 dòng bảng động bằng dữ liệu từ đối tượng container hoặc item tương ứng.
 *
 * @param {GoogleAppsScript.Document.TableRow} row
 * @param {Object} rowObj
 * @param {Array<string>} tagsInRow
 * @param {number} rowIndexZeroBased
 */
function replaceRowTags_(row, rowObj, tagsInRow, rowIndexZeroBased) {
  const rowSpecificTags = ITEM_TAGS
    .concat(CONT_TAGS)
    .concat(TAX_TAGS)
    .concat(['PACK', 'GW', 'CBM']);

  Object.keys(rowObj).forEach(function (key) {
    const val = rowObj[key];
    const strVal = val == null ? '' : String(val);
    if ((key === 'MARKS' || key === 'UNIT') && strVal === '') {
      return;
    }
    row.replaceText('\\{\\{' + escapeRegex_(key) + '\\}\\}', strVal);
  });

  for (let i = 0; i < tagsInRow.length; i++) {
    const tag = tagsInRow[i];
    const hasProp = Object.prototype.hasOwnProperty.call(rowObj, tag);
    const hasNonEmptyVal = hasProp && rowObj[tag] != null && String(rowObj[tag]) !== '';

    if (hasNonEmptyVal) {
      continue;
    }

    if (tag === 'STT' || tag === 'NO' || tag === 'ITEM_NO') {
      row.replaceText('\\{\\{' + escapeRegex_(tag) + '\\}\\}', String(rowIndexZeroBased + 1));
    } else if (rowSpecificTags.indexOf(tag) !== -1 && tag !== 'UNIT' && tag !== 'MARKS') {
      row.replaceText('\\{\\{' + escapeRegex_(tag) + '\\}\\}', '');
    }
  }
}

/**
 * Quét dọn toàn bộ thẻ {{TAG}} còn sót lại trên Body, Header và Footer của Tab.
 *
 * @param {Array<GoogleAppsScript.Document.Element>} sections
 */
function sanitizeRemainingTags_(sections) {
  for (let s = 0; s < sections.length; s++) {
    sections[s].replaceText('\\{\\{[A-Z0-9_]+\\}\\}', '');
  }
}

/**
 * Escape các ký tự đặc biệt của biểu thức chính quy cho DocumentApp.Element.replaceText().
 * @param {string} str
 * @return {string}
 */
function escapeRegex_(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Hàm không tham số dùng để kích hoạt nhanh màn hình cấp quyền Google Drive / Docs trong Apps Script Editor.
 * Chọn authorizeDrive ở thanh công cụ phía trên và bấm Run (Chạy) 1 lần trước khi Deploy.
 */
function authorizeDrive() {
  DriveApp.getRootFolder();
}


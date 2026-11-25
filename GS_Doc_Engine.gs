/**
 * GS_Doc_Engine.gs
 * Core Google Docs auto-fill engine supporting BOTH:
 * - Mode A (Folder of 20 Google Docs templates, e.g. folder "CHỨNG TỪ"):
 *   MASTER_TEMPLATE_DOC_ID = Folder ID containing 20 files named by docKey
 * - Mode B (Single multi-tab Google Doc with 20 Tabs):
 *   MASTER_TEMPLATE_DOC_ID = Google Doc ID containing 20 tabs matching TEMPLATE_MAP[docKey].tabName
 * Runtime: Google Apps Script (V8)
 */

function generateDocuments(maSV, docKeysArray, payload) {
  try {
    const rawKeys = Array.isArray(docKeysArray)
      ? docKeysArray
      : (docKeysArray ? [String(docKeysArray)] : []);

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
          processSingleDocumentObject_(doc, items, containers, shared, docKey);
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
        replaceGlobalTags(sections, shared, docKey);
        sanitizeRemainingTags_(sections);
        cleanupEmptyStaticLabels_(sections);

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
 * Tạo bản xem trước (Live Preview) cho 1 chứng từ.
 * Sao chép tạm vào DriveApp.getRootFolder() (tuyệt đối không lưu vào OUTPUT_FOLDER_ID hay thư mục chia sẻ dự án),
 * giữ nguyên 100% định dạng gốc, xuất PDF dạng base64, và xóa ngay lập tức file tạm trong khối try...finally.
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
      processSingleDocumentObject_(doc, items, containers, shared, cleanKey);
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
        if (tempDocFile) {
          try {
            tempDocFile.setTrashed(true);
            tempDocFile = null;
          } catch (e) {}
        }

        return {
          success: false,
          message: 'Không tìm thấy Tab "' + tplConfig.tabName + '" trong file Google Docs mẫu.'
        };
      }

      const docTab = targetTab.asDocumentTab();
      const body = docTab.getBody();
      const sections = collectDocumentTabSections_(docTab);

      processDynamicTableRows(body, items, containers);
      replaceGlobalTags(sections, shared, cleanKey);
      sanitizeRemainingTags_(sections);
      cleanupEmptyStaticLabels_(sections);

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
    return {
      success: false,
      message: formatEngineErrorMessage_('Lỗi tạo bản xem trước chứng từ', err)
    };
  } finally {
    if (tempDocFile) {
      try {
        tempDocFile.setTrashed(true);
      } catch (e) {}
    }
  }
}

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

function normalizeTemplateLookupName_(name) {
  return String(name || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function processSingleDocumentObject_(doc, items, containers, shared, docKey) {
  if (typeof doc.getTabs === 'function') {
    const topTabs = doc.getTabs();
    const allTabs = collectAllTabsFlat_(topTabs);
    if (allTabs.length > 0) {
      for (let t = 0; t < allTabs.length; t++) {
        const docTab = allTabs[t].asDocumentTab();
        const body = docTab.getBody();
        const sections = collectDocumentTabSections_(docTab);
        processDynamicTableRows(body, items, containers);
        replaceGlobalTags(sections, shared, docKey);
        sanitizeRemainingTags_(sections);
        cleanupEmptyStaticLabels_(sections);
      }
      return;
    }
  }

  const body = doc.getBody();
  const sections = collectSingleDocSections_(doc);
  processDynamicTableRows(body, items, containers);
  replaceGlobalTags(sections, shared, docKey);
  sanitizeRemainingTags_(sections);
  cleanupEmptyStaticLabels_(sections);
}

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

    for (let i = allTabs.length - 1; i >= 0; i--) {
      const tab = allTabs[i];
      const tabId = tab.getId();
      if (tabId && selectedTabIds.indexOf(tabId) === -1) {
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

function processDynamicTableRows(body, items, containers) {
  const safeItems = Array.isArray(items) ? items : [];
  const safeContainers = Array.isArray(containers) ? containers : [];
  const tables = body.getTables();

  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];

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
          row.replaceText('\\{\\{[A-Z0-9_]+\\}\\}', '');
        }
        continue;
      }

      const sttCellIndices = detectTemplateSttCellIndices_(table, r, row);

      for (let i = 1; i < count; i++) {
        table.insertTableRow(r + i, row.copy());
      }

      for (let idx = 0; idx < count; idx++) {
        const targetRow = table.getRow(r + idx);
        const rowData = boundArray[idx] || {};
        replaceRowTags_(targetRow, rowData, tagsInRow, idx);
        for (let sIdx = 0; sIdx < sttCellIndices.length; sIdx++) {
          const cellIndex = sttCellIndices[sIdx];
          if (cellIndex < targetRow.getNumCells()) {
            populateSttCell_(targetRow.getCell(cellIndex), idx + 1);
          }
        }
      }
    }
  }
}

function detectTemplateSttCellIndices_(table, rowIndex, templateRow) {
  const indices = [];
  if (!templateRow) return indices;

  const numCells = templateRow.getNumCells();
  const headerRow = rowIndex > 0 ? table.getRow(0) : null;

  for (let c = 0; c < numCells; c++) {
    const cellText = String(templateRow.getCell(c).getText() || '').trim();
    const isSttLiteralOrTag = /^(?:0?1\.?|\{\{(?:STT|NO|ITEM_NO)\}\})$/.test(cellText);
    if (isSttLiteralOrTag) {
      indices.push(c);
      continue;
    }

    if (c === 0 && cellText === '' && headerRow && headerRow.getNumCells() > 0) {
      const headerText = String(headerRow.getCell(0).getText() || '').trim();
      if (/^(?:STT|SỐ\s*TT|NO\.?|ITEM\s*NO\.?)$/i.test(headerText)) {
        indices.push(0);
      }
    }
  }

  return indices;
}

function populateSttCell_(cell, seqNumber) {
  if (!cell) return;
  const seqStr = String(seqNumber);
  const currentText = String(cell.getText() || '').trim();

  if (currentText !== '') {
    cell.replaceText('^.*$', seqStr);
    return;
  }

  if (cell.getNumChildren() > 0 && cell.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
    cell.getChild(0).asParagraph().setText(seqStr);
  } else {
    cell.setText(seqStr);
  }
}

function replaceGlobalTags(sections, shared, docKey) {
  const keys = Object.keys(shared).sort(function (a, b) {
    return b.length - a.length;
  });

  const coFormVal = String(shared.CO_FORM_NAME || '').trim();
  const coFormNoPrefix = coFormVal.replace(/^[Ff][Oo][Rr][Mm]\s+/, '').trim();
  const hasFormPrefix = coFormNoPrefix !== '' && /^[Ff][Oo][Rr][Mm]\s+/.test(coFormVal);

  for (let s = 0; s < sections.length; s++) {
    const section = sections[s];
    if (hasFormPrefix) {
      section.replaceText('[Ff][Oo][Rr][Mm]\\s+\\{\\{CO_FORM_NAME\\}\\}', 'Form ' + coFormNoPrefix);
    }

    for (let k = 0; k < keys.length; k++) {
      const key = keys[k];
      const rawVal = shared[key];
      let strVal = rawVal == null ? '' : String(rawVal);
      if (key === 'CO_FORM_NAME' && docKey === 'SALES_CONTRACT' && hasFormPrefix) {
        strVal = coFormNoPrefix;
      }
      section.replaceText('\\{\\{' + escapeRegex_(key) + '\\}\\}', strVal);
    }

    if (docKey === 'SALES_CONTRACT') {
      section.replaceText('[Ff][Oo][Rr][Mm]\\s+[Ff][Oo][Rr][Mm]\\b', 'Form');
    }
  }
}

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

function buildEnrichedShared_(payload) {
  const sourceShared = (payload && payload.shared && typeof payload.shared === 'object')
    ? payload.shared
    : {};
  const shared = {};

  const CHECKBOX_KEYS = [
    'DG_YES',
    'DG_NO',
    'FREIGHT_PREPAID',
    'FREIGHT_COLLECT',
    'OUT_OF_GAUGE',
    'NON_STACKABLE',
    'FUMIGATION_NEEDED'
  ];

  Object.keys(sourceShared).forEach(function (key) {
    const val = sourceShared[key];
    if (typeof val === 'boolean') {
      shared[key] = val ? '\u2611' : '\u2610';
    } else {
      shared[key] = val == null ? '' : String(val);
    }
  });

  CHECKBOX_KEYS.forEach(function (cbKey) {
    shared[cbKey] = normalizeCheckboxSymbol_(shared[cbKey]);
  });

  ['INS_TERM', 'CO_FORM_NAME', 'CTYPE', 'UNIT', 'TOT_UNIT'].forEach(function (codeKey) {
    if (shared[codeKey] != null && String(shared[codeKey]).trim() !== '') {
      shared[codeKey] = extractConciseCode_(shared[codeKey], codeKey);
    }
  });

  const docsContact = buildParenContactString_([
    { label: 'Tel', value: shared.DOCS_TEL },
    { label: 'Email', value: shared.DOCS_EMAIL }
  ]);
  shared.DOCS_CONTACT = docsContact;
  shared.DOCS_TEL_EMAIL = docsContact;

  shared.BUYER_CONTACT = buildParenContactString_([
    { label: 'Tel', value: shared.BUYER_TEL }
  ]);

  shared.SHIP_CONTACT = buildParenContactString_([
    { label: 'Tel', value: shared.SHIP_TEL },
    { label: 'Email', value: shared.SHIP_EMAIL }
  ]);

  shared.CNE_CONTACT = buildParenContactString_([
    { label: 'Tel', value: shared.CNE_TEL },
    { label: 'Fax', value: shared.CNE_FAX },
    { label: 'Email', value: shared.CNE_EMAIL }
  ]);

  shared.NTF_CONTACT = buildParenContactString_([
    { label: 'Tel', value: shared.NTF_TEL },
    { label: 'Email', value: shared.NTF_EMAIL }
  ]);

  shared.CAR_CONTACT = buildParenContactString_([
    { label: 'Tel', value: shared.CAR_TEL },
    { label: 'Email', value: shared.CAR_EMAIL }
  ]);

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

function normalizeCheckboxSymbol_(val) {
  if (val === true) return '\u2611';
  if (val === false || val == null) return '\u2610';

  const trimmed = String(val).trim();
  if (!trimmed || trimmed === '\u2610') return '\u2610';

  const lower = trimmed.toLowerCase();
  if (
    trimmed === '\u2611' ||
    lower === 'true' ||
    lower === '1' ||
    lower === 'x' ||
    lower === 'yes' ||
    lower === 'on' ||
    lower === 'checked'
  ) {
    return '\u2611';
  }
  return '\u2610';
}

function buildParenContactString_(pairs) {
  if (!Array.isArray(pairs)) return '';
  const parts = [];
  for (let i = 0; i < pairs.length; i++) {
    const item = pairs[i];
    const val = String(item && item.value != null ? item.value : '').trim();
    if (val) {
      parts.push(item.label + ': ' + val);
    }
  }
  return parts.length > 0 ? '(' + parts.join(', ') + ')' : '';
}

function resolveDocFilenameSuffix_(shared, maSV) {
  const invNo = String(shared.INV_NO || '').trim();
  if (invNo) return invNo;

  const bkNo = String(shared.BK_NO || '').trim();
  if (bkNo) return bkNo;

  const studentId = String(maSV == null ? '' : maSV).trim();
  return studentId || 'DOC';
}

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

function replaceRowTags_(row, rowObj, tagsInRow, rowIndexZeroBased) {
  const rowSpecificTags = ITEM_TAGS
    .concat(CONT_TAGS)
    .concat(TAX_TAGS)
    .concat(['PACK', 'GW', 'CBM']);

  const seqText = String(rowIndexZeroBased + 1);
  row.replaceText('\\{\\{STT\\}\\}', seqText);
  row.replaceText('\\{\\{NO\\}\\}', seqText);
  row.replaceText('\\{\\{ITEM_NO\\}\\}', seqText);

  Object.keys(rowObj).forEach(function (key) {
    if (key === 'STT' || key === 'NO' || key === 'ITEM_NO') {
      return;
    }
    const val = rowObj[key];
    let strVal = val == null ? '' : String(val);
    if ((key === 'MARKS' || key === 'UNIT') && strVal === '') {
      return;
    }
    if (key === 'PREF' || key === 'UNIT' || key === 'UNIT_PACK' || key === 'CTYPE' || key === 'CSIZE') {
      strVal = extractConciseCode_(strVal, key);
    }
    row.replaceText('\\{\\{' + escapeRegex_(key) + '\\}\\}', strVal);
  });

  for (let i = 0; i < tagsInRow.length; i++) {
    const tag = tagsInRow[i];
    if (tag === 'STT' || tag === 'NO' || tag === 'ITEM_NO') {
      continue;
    }
    const hasProp = Object.prototype.hasOwnProperty.call(rowObj, tag);
    const hasNonEmptyVal = hasProp && rowObj[tag] != null && String(rowObj[tag]) !== '';

    if (hasNonEmptyVal) {
      continue;
    }

    if (rowSpecificTags.indexOf(tag) !== -1 && tag !== 'UNIT' && tag !== 'MARKS') {
      row.replaceText('\\{\\{' + escapeRegex_(tag) + '\\}\\}', '');
    }
  }
}

function sanitizeRemainingTags_(sections) {
  for (let s = 0; s < sections.length; s++) {
    sections[s].replaceText('\\{\\{[A-Z0-9_]+\\}\\}', '');
  }
}

function cleanupEmptyStaticLabels_(sections) {
  const telLbl = '(?:[Tt][Ee][Ll]|[Pp][Hh][Oo][Nn][Ee]|[Mm][Oo][Bb][Ii][Ll][E])';
  const emailLbl = '(?:[Ee]-?[Mm][Aa][Ii][Ll])';
  const faxLbl = '(?:[Ff][Aa][Xx])';
  const taxLbl = '(?:[Tt][Aa][Xx](?:\\s+[Cc][Oo][Dd][E])?|[Mm][Ss][Tt])';
  const attnLbl = '(?:[Aa][Tt][Tt][Nn])';
  const anyLbl = '(?:' + telLbl + '|' + emailLbl + '|' + faxLbl + '|' + taxLbl + '|' + attnLbl + ')';

  for (let s = 0; s < sections.length; s++) {
    const section = sections[s];
    section.replaceText('\\([\\s\\p{P}]*\\b' + telLbl + '\\b[\\s\\p{P}]*\\b' + emailLbl + '\\b[\\s\\p{P}]*\\)', '');
    section.replaceText('\\([\\s\\p{P}]*\\b' + anyLbl + '\\b\\s*:\\s*,?\\s*\\)', '');
    section.replaceText('\\b' + anyLbl + '\\s*:\\s*(?=[\\),;\\.\\r\\n]|$)', '');
  }
}

function extractConciseCode_(rawVal, fieldKey) {
  const str = String(rawVal == null ? '' : rawVal).trim();
  if (!str) return '';

  let depth = 0;
  let cutIdx = -1;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charAt(i);
    if (ch === '(') {
      depth++;
    } else if (ch === ')' && depth > 0) {
      depth--;
    } else if (depth === 0) {
      if (ch === ':') {
        cutIdx = i;
        break;
      }
      if ((ch === '-' || ch === '\u2013' || ch === '\u2014') &&
          i > 0 && str.charAt(i - 1) === ' ' &&
          i + 1 < str.length && str.charAt(i + 1) === ' ') {
        cutIdx = i - 1;
        break;
      }
    }
  }

  let code = (cutIdx !== -1 ? str.slice(0, cutIdx) : str).trim();

  if (fieldKey === 'INS_TERM') {
    const m = code.match(/^(.*?)\s*\(([^)]*)\)\s*$/);
    if (m && m[1].trim() && m[2].trim().length > 3) {
      code = m[1].trim();
    }
    return code;
  }

  let prev;
  do {
    prev = code;
    code = code.replace(/\s*\([^)]*\)\s*$/, '').trim();
  } while (code !== prev && code.length > 0);

  return code;
}

function escapeRegex_(str) {
  return String(str || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

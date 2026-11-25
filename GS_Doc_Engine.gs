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

        applyDocumentTagPipeline_(body, sections, items, containers, shared, docKey);

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

      applyDocumentTagPipeline_(body, sections, items, containers, shared, cleanKey);

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
        applyDocumentTagPipeline_(body, sections, items, containers, shared, docKey);
      }
      return;
    }
  }

  const body = doc.getBody();
  const sections = collectSingleDocSections_(doc);
  applyDocumentTagPipeline_(body, sections, items, containers, shared, docKey);
}

function applyDocumentTagPipeline_(body, sections, items, containers, shared, docKey) {
  const rowCandidateParagraphs = processDynamicTableRows(body, items, containers, docKey, shared) || [];
  const sectionCandidateParagraphs = collectPurePlaceholderParagraphs_(sections);

  if (docKey === 'BILL_OF_LADING' || docKey === 'DELIVERY_ORDER') {
    prepareBillOfLadingPackageUnitTags_(sections);
  } else if (docKey === 'SALES_CONTRACT') {
    prepareSalesContractTemplateTags_(sections);
  }

  replaceGlobalTags(sections, shared, docKey, items, containers);
  sanitizeRemainingTags_(sections);
  cleanupEmptyStaticLabels_(sections);

  if (docKey === 'BOOKING_CONFIRMATION') {
    fixBookingConfirmationRemarksColon_(sections);
  } else if (docKey === 'SALES_CONTRACT') {
    cleanupSalesContractUnitPriceSlash_(sections);
  } else if (docKey === 'BILL_OF_LADING' || docKey === 'DELIVERY_ORDER') {
    cleanupBillOfLadingBoxes_(sections, docKey);
  } else if (docKey === 'To_khai_XK' || docKey === 'To_khai_NK') {
    const customsExtras = buildCustomsExtras_(shared, docKey);
    populateCustomsBoxesIfEmpty_(body, docKey, customsExtras);
    cleanupCustomsBlankRows_(body, docKey);
    if (docKey === 'To_khai_XK') {
      normalizeCustomsBoxesColor_(body);
    }
    applyNoWrapFormattingToCustomsTaxCells_(body);
  } else if (docKey === 'VETERINARY_CERTIFICATE') {
    applyVeterinarySectionIIILayout_(body, sections);
  }

  const combinedCandidates = rowCandidateParagraphs.concat(sectionCandidateParagraphs);
  removeResolvedEmptyPlaceholderParagraphs_(combinedCandidates);
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

function processDynamicTableRows(body, items, containers, docKey, shared) {
  if (docKey === 'PHYTOSANITARY_CERTIFICATE' || docKey === 'QUALITY_CERTIFICATE') {
    return [];
  }

  const safeItems = Array.isArray(items) ? items : [];
  const safeContainers = Array.isArray(containers) ? containers : [];
  const tables = body.getTables();
  const candidateParagraphs = [];

  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];

    for (let r = table.getNumRows() - 1; r >= 0; r--) {
      const row = table.getRow(r);
      const rowText = row.getText();
      const tagsInRow = extractTagNames_(rowText);

      if (!hasDynamicRowTag_(tagsInRow, docKey)) {
        continue;
      }

      const isContainerRow =
        docKey !== 'VETERINARY_CERTIFICATE' &&
        docKey !== 'CERTIFICATE_OF_ORIGIN' &&
        tagsInRow.indexOf('CONT') !== -1;
      const boundArray = isContainerRow ? safeContainers : safeItems;
      const count = boundArray.length;

      if (count === 0) {
        if (table.getNumRows() > 1) {
          table.removeRow(r);
        } else {
          collectPurePlaceholderParagraphsInContainer_(row, candidateParagraphs);
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
        collectPurePlaceholderParagraphsInContainer_(targetRow, candidateParagraphs);
        const rowData = boundArray[idx] || {};
        replaceRowTags_(targetRow, rowData, tagsInRow, idx, docKey, shared);
        for (let sIdx = 0; sIdx < sttCellIndices.length; sIdx++) {
          const cellIndex = sttCellIndices[sIdx];
          if (cellIndex < targetRow.getNumCells()) {
            populateSttCell_(targetRow.getCell(cellIndex), idx + 1);
          }
        }
      }
    }
  }

  return candidateParagraphs;
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

function replaceGlobalTags(sections, shared, docKey, items, containers) {
  const effectiveShared = buildDocSpecificShared_(shared, items, containers, docKey);
  if (docKey === 'QUOTATION' && effectiveShared.TOT_AMT_WORDS_VI) {
    effectiveShared.TOT_AMT_WORDS = effectiveShared.TOT_AMT_WORDS_VI;
  } else if (docKey === 'SALES_CONTRACT' && effectiveShared.TOT_AMT_WORDS_SC) {
    effectiveShared.TOT_AMT_WORDS = effectiveShared.TOT_AMT_WORDS_SC;
  } else if (docKey === 'COMMERCIAL_INVOICE' && effectiveShared.TOT_AMT_WORDS_CI) {
    effectiveShared.TOT_AMT_WORDS = effectiveShared.TOT_AMT_WORDS_CI;
  }

  const keys = Object.keys(effectiveShared).sort(function (a, b) {
    return b.length - a.length;
  });

  const coFormVal = String(effectiveShared.CO_FORM_NAME || '').trim();
  const coFormNoPrefix = coFormVal.replace(/^[Ff][Oo][Rr][Mm]\s+/, '').trim();
  const hasFormPrefix = coFormNoPrefix !== '' && /^[Ff][Oo][Rr][Mm]\s+/.test(coFormVal);

  for (let s = 0; s < sections.length; s++) {
    const section = sections[s];
    if (hasFormPrefix) {
      section.replaceText('[Ff][Oo][Rr][Mm]\\s+\\{\\{CO_FORM_NAME\\}\\}', 'Form ' + coFormNoPrefix);
    }

    for (let k = 0; k < keys.length; k++) {
      const key = keys[k];
      const rawVal = effectiveShared[key];
      let strVal = rawVal == null ? '' : String(rawVal);
      if (key === 'CO_FORM_NAME' && docKey === 'SALES_CONTRACT' && hasFormPrefix) {
        strVal = coFormNoPrefix;
      }
      section.replaceText('\\{\\{' + escapeRegex_(key) + '\\}\\}', strVal);
    }

    if (docKey === 'SALES_CONTRACT') {
      section.replaceText('[Ff][Oo][Rr][Mm]\\s+[Ff][Oo][Rr][Mm]', 'Form');
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

  ['INS_TERM', 'CO_FORM_NAME', 'CTYPE', 'UNIT', 'TOT_UNIT', 'CURR'].forEach(function (codeKey) {
    if (shared[codeKey] != null && String(shared[codeKey]).trim() !== '') {
      shared[codeKey] = extractConciseCode_(shared[codeKey], codeKey);
    }
  });

  if (!shared.CURR || String(shared.CURR).trim() === '') {
    shared.CURR = 'USD';
  }

  if (shared.HS != null && String(shared.HS).trim() !== '') {
    shared.HS = formatHsCode_(shared.HS);
  }

  ['NTF_NAME', 'NTF_ADDR', 'NTF_TEL', 'NTF_EMAIL'].forEach(function (ntfKey) {
    if (shared[ntfKey] == null) {
      shared[ntfKey] = '';
    }
  });

  const FIXED2_SHARED_KEYS = {
    TOT_AMT: true,
    VAT_TOTAL: true,
    GRAND_TOTAL: true,
    TOT_NW: true,
    TOT_GW: true,
    TOT_CBM: true,
    TOT_VGM: true,
    TOT_EX_TAX: true,
    TOT_FEE: true,
    TOT_TAX: true,
    IM_VAL: true,
    IM_AMT: true,
    SCT_VAL: true,
    SCT_AMT: true,
    EPT_VAL: true,
    EPT_AMT: true,
    VAT_VAL: true,
    VAT_AMT: true,
    FREIGHT_AMT: true,
    INS_AMT: true
  };

  const NUMERIC_SHARED_KEYS = [
    'TOT_QTY', 'TOT_AMT', 'VAT_GST', 'VAT_TOTAL', 'GRAND_TOTAL',
    'TOT_PACK', 'TOT_NW', 'TOT_GW', 'TOT_CBM', 'TOT_VGM',
    'EX_RATE', 'TOT_EX_TAX', 'TOT_FEE', 'TOT_TAX',
    'IM_VAL', 'IM_RATE', 'IM_AMT',
    'SCT_VAL', 'SCT_RATE', 'SCT_AMT',
    'EPT_VAL', 'EPT_RATE', 'EPT_AMT',
    'VAT_VAL', 'VAT_RATE', 'VAT_AMT',
    'FREIGHT_AMT', 'INS_AMT'
  ];
  NUMERIC_SHARED_KEYS.forEach(function (numKey) {
    if (shared[numKey] != null && String(shared[numKey]).trim() !== '') {
      shared[numKey] = FIXED2_SHARED_KEYS[numKey]
        ? formatDocFixed2_(shared[numKey])
        : formatDocNumber_(shared[numKey]);
    }
  });

  ['TOTAL_CONTAINERS_TEXT', 'TOT_AMT_WORDS', 'TOT_AMT_WORDS_VI', 'TOT_AMT_WORDS_SC', 'TOT_AMT_WORDS_CI', 'TOT_TAX_WORDS'].forEach(function (sayKey) {
    if (shared[sayKey] != null && String(shared[sayKey]).trim() !== '') {
      shared[sayKey] = String(shared[sayKey]).replace(/^\s*SAY\s*:\s*/i, '').trim();
    }
  });

  const items = Array.isArray(payload && payload.items) ? payload.items : [];
  const containers = Array.isArray(payload && payload.containers) ? payload.containers : [];

  let sumItemNw = 0;
  let hasItemNw = false;
  let sumItemGw = 0;
  let hasItemGw = false;
  let sumItemCbm = 0;
  let hasItemCbm = false;
  let sumItemQty = 0;
  let hasItemQty = false;
  let sumItemPack = 0;
  let hasItemPack = false;
  let firstItemUnit = '';
  let firstItemPackUnit = '';

  items.forEach(function (it) {
    if (!it || typeof it !== 'object') return;
    if (!firstItemUnit && it.UNIT) firstItemUnit = extractConciseCode_(it.UNIT, 'UNIT');
    if (!firstItemPackUnit && it.UNIT_PACK) firstItemPackUnit = extractConciseCode_(it.UNIT_PACK, 'UNIT_PACK');

    const nwNum = parseDocNumber_(it.NW);
    if (nwNum !== null) {
      sumItemNw += nwNum;
      hasItemNw = true;
    }
    const gwNum = parseDocNumber_(it.GW);
    if (gwNum !== null) {
      sumItemGw += gwNum;
      hasItemGw = true;
    }
    const cbmNum = parseDocNumber_(it.CBM);
    if (cbmNum !== null) {
      sumItemCbm += cbmNum;
      hasItemCbm = true;
    }
    const qtyNum = parseDocNumber_(it.QTY);
    if (qtyNum !== null) {
      sumItemQty += qtyNum;
      hasItemQty = true;
    }
    const packNum = parseDocNumber_(it.PACK);
    if (packNum !== null) {
      sumItemPack += packNum;
      hasItemPack = true;
    }
  });

  let sumContPack = 0;
  let hasContPack = false;
  let sumContGw = 0;
  let hasContGw = false;
  let sumContCbm = 0;
  let hasContCbm = false;
  let sumContVgm = 0;
  let hasContVgm = false;
  let firstContPackUnit = '';

  containers.forEach(function (c) {
    if (!c || typeof c !== 'object') return;
    if (!firstContPackUnit && c.UNIT_PACK) firstContPackUnit = extractConciseCode_(c.UNIT_PACK, 'UNIT_PACK');
    const packNum = parseDocNumber_(c.PACK);
    if (packNum !== null) {
      sumContPack += packNum;
      hasContPack = true;
    }
    const gwNum = parseDocNumber_(c.GW);
    if (gwNum !== null) {
      sumContGw += gwNum;
      hasContGw = true;
    }
    const cbmNum = parseDocNumber_(c.CBM);
    if (cbmNum !== null) {
      sumContCbm += cbmNum;
      hasContCbm = true;
    }
    const vgmNum = parseDocNumber_(c.VGM_WEI);
    if (vgmNum !== null && vgmNum > 0) {
      sumContVgm += vgmNum;
      hasContVgm = true;
    } else if (gwNum !== null && gwNum > 0) {
      sumContVgm += gwNum;
      hasContVgm = true;
    }
  });

  if (hasItemNw) {
    shared.TOT_NW = formatDocFixed2_(sumItemNw);
  }
  if (hasItemGw && sumItemGw > 0) {
    shared.TOT_GW = formatDocFixed2_(sumItemGw);
  } else if (hasContGw && sumContGw > 0) {
    shared.TOT_GW = formatDocFixed2_(sumContGw);
  }
  if (hasItemCbm && sumItemCbm > 0) {
    shared.TOT_CBM = formatDocFixed2_(sumItemCbm);
  } else if (hasContCbm && sumContCbm > 0) {
    shared.TOT_CBM = formatDocFixed2_(sumContCbm);
  } else if (shared.TOT_CBM) {
    shared.TOT_CBM = formatDocFixed2_(shared.TOT_CBM);
  }
  if (hasContVgm && sumContVgm > 0) {
    shared.TOT_VGM = formatDocFixed2_(sumContVgm);
  } else if (shared.TOT_VGM) {
    shared.TOT_VGM = formatDocFixed2_(shared.TOT_VGM);
  } else if (shared.TOT_GW) {
    shared.TOT_VGM = formatDocFixed2_(shared.TOT_GW);
  }
  if (hasItemQty && (!shared.TOT_QTY || String(shared.TOT_QTY).trim() === '')) {
    shared.TOT_QTY = formatDocNumber_(sumItemQty);
  }

  const effectiveTotPack = hasItemPack && sumItemPack > 0
    ? sumItemPack
    : (hasContPack && sumContPack > 0 ? sumContPack : (parseDocNumber_(shared.TOT_PACK) || 0));
  if (effectiveTotPack > 0) {
    shared.TOT_PACK = formatDocNumber_(effectiveTotPack);
  }
  if ((!shared.TOT_UNIT || String(shared.TOT_UNIT).trim() === '') && firstItemUnit) {
    shared.TOT_UNIT = firstItemUnit;
  }

  const rawPackUnit = String(shared.PACK_UNIT || firstItemPackUnit || firstContPackUnit || '').trim();
  if (rawPackUnit) {
    shared.PACK_UNIT = pluralizePackUnit_(rawPackUnit, effectiveTotPack);
  }

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

function hasDynamicRowTag_(tagsInRow, docKey) {
  if (docKey === 'PHYTOSANITARY_CERTIFICATE' || docKey === 'QUALITY_CERTIFICATE') {
    return false;
  }

  if (docKey === 'CERTIFICATE_OF_ORIGIN') {
    return tagsInRow.indexOf('PROD') !== -1 && tagsInRow.indexOf('HS') !== -1;
  }

  if (docKey === 'VETERINARY_CERTIFICATE') {
    const VET_ITEM_ROW_TAGS = ['PROD', 'PCODE', 'DES', 'LOT_NO', 'QTY', 'NW', 'GW', 'PACK'];
    for (let i = 0; i < tagsInRow.length; i++) {
      if (VET_ITEM_ROW_TAGS.indexOf(tagsInRow[i]) !== -1) {
        return true;
      }
    }
    return false;
  }

  if (docKey === 'BILL_OF_LADING') {
    const hasBlSummaryTag =
      tagsInRow.indexOf('TOT_PACK') !== -1 ||
      tagsInRow.indexOf('TOT_GW') !== -1 ||
      tagsInRow.indexOf('DESCRIPTION_OF_GOODS') !== -1;
    if (hasBlSummaryTag) {
      return false;
    }
  }

  for (let i = 0; i < tagsInRow.length; i++) {
    const tag = tagsInRow[i];
    if (docKey === 'PACKING_LIST' && tag === 'ORG') {
      continue;
    }
    if (docKey === 'To_khai_NK' && TAX_TAGS.indexOf(tag) !== -1) {
      continue;
    }
    if (
      ITEM_TAGS.indexOf(tag) !== -1 ||
      CONT_TAGS.indexOf(tag) !== -1 ||
      (docKey !== 'To_khai_NK' && TAX_TAGS.indexOf(tag) !== -1)
    ) {
      return true;
    }
  }
  return false;
}

function splitContainerSizeAndType_(ctypeVal, csizeVal) {
  const rawType = extractConciseCode_(ctypeVal == null ? '' : String(ctypeVal), 'CTYPE').toUpperCase();
  const rawSize = extractConciseCode_(csizeVal == null ? '' : String(csizeVal), 'CSIZE').toUpperCase();
  const combined = rawSize || rawType;

  let sizeOut = rawSize;
  const sizeMatch = combined.match(/(20|40|45)/);
  if (sizeMatch) {
    sizeOut = sizeMatch[1] + "'";
  } else if (rawSize && !/'$/.test(rawSize) && /^\d+$/.test(rawSize)) {
    sizeOut = rawSize + "'";
  }

  const typeSource = rawType || rawSize;
  let typeOut = typeSource.replace(/^(?:20|40|45)\s*'?\s*[-\/]?\s*/, '').trim();
  if (!typeOut && rawType && !/^\d+'?$/.test(rawType)) {
    typeOut = rawType;
  }

  return {
    CSIZE: sizeOut,
    CTYPE: typeOut || rawType
  };
}

function replaceRowTags_(row, rowObj, tagsInRow, rowIndexZeroBased, docKey, shared) {
  const rowSpecificTags = ITEM_TAGS
    .concat(CONT_TAGS)
    .concat(TAX_TAGS)
    .concat(['PACK', 'GW', 'CBM']);

  const effectiveObj = Object.assign({}, rowObj);

  if (docKey === 'RELEASE_ORDER') {
    const splitCont = splitContainerSizeAndType_(effectiveObj.CTYPE, effectiveObj.CSIZE);
    effectiveObj.CTYPE = splitCont.CTYPE;
    effectiveObj.CSIZE = splitCont.CSIZE;
  }

  if (docKey === 'CERTIFICATE_OF_ORIGIN') {
    if (!effectiveObj.INV_NO) {
      effectiveObj.INV_NO = effectiveObj.INVOICE_REF || (shared && shared.INV_NO) || '';
    }
    if (!effectiveObj.INV_DT) {
      effectiveObj.INV_DT = (shared && shared.INV_DT) || '';
    }
    if (!effectiveObj.ORG && shared && shared.ORG) {
      effectiveObj.ORG = shared.ORG;
    }
    if (!effectiveObj.CURR && shared && shared.CURR) {
      effectiveObj.CURR = shared.CURR;
    }
  }

  const seqText = String(rowIndexZeroBased + 1);
  row.replaceText('\\{\\{STT\\}\\}', seqText);
  row.replaceText('\\{\\{NO\\}\\}', seqText);
  row.replaceText('\\{\\{ITEM_NO\\}\\}', seqText);

  Object.keys(effectiveObj).forEach(function (key) {
    if (key === 'STT' || key === 'NO' || key === 'ITEM_NO') {
      return;
    }
    const val = effectiveObj[key];
    let strVal = val == null ? '' : String(val);
    if ((key === 'MARKS' || key === 'UNIT') && strVal === '') {
      return;
    }
    if (key === 'HS') {
      strVal = formatHsCode_(strVal);
    }
    if (key === 'PREF' || key === 'UNIT' || key === 'UNIT_PACK' || key === 'CTYPE' || key === 'CSIZE') {
      strVal = extractConciseCode_(strVal, key);
    }
    if (
      key === 'NW' || key === 'GW' || key === 'CBM' ||
      key === 'MAX_GW' || key === 'VGM_WEI' ||
      key === 'PRICE' || key === 'AMT' ||
      key === 'EX_VAL' || key === 'EX_AMT' || key === 'FEE_VAL' || key === 'FEE_AMT'
    ) {
      strVal = formatDocFixed2_(strVal);
    } else if (
      key === 'QTY' || key === 'PACK' ||
      key === 'EX_TAX_RATE' || key === 'FEE_RATE'
    ) {
      strVal = formatDocNumber_(strVal);
    }
    row.replaceText('\\{\\{' + escapeRegex_(key) + '\\}\\}', strVal);
  });

  for (let i = 0; i < tagsInRow.length; i++) {
    const tag = tagsInRow[i];
    if (tag === 'STT' || tag === 'NO' || tag === 'ITEM_NO') {
      continue;
    }
    if (docKey === 'CERTIFICATE_OF_ORIGIN' && (tag === 'INV_NO' || tag === 'INV_DT')) {
      continue;
    }
    const hasProp = Object.prototype.hasOwnProperty.call(effectiveObj, tag);
    const hasNonEmptyVal = hasProp && effectiveObj[tag] != null && String(effectiveObj[tag]) !== '';

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

function isPurePlaceholderParagraphText_(text) {
  const str = String(text == null ? '' : text);
  if (!/\{\{[A-Z0-9_]+\}\}/.test(str)) {
    return false;
  }
  return str.replace(/\{\{[A-Z0-9_]+\}\}/g, '').trim() === '';
}

function collectPurePlaceholderParagraphsInContainer_(container, outList) {
  if (!container || typeof container.getNumChildren !== 'function') return;
  const numChildren = container.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const child = container.getChild(i);
    const type = child.getType();

    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const para = child.asParagraph();
      if (isPurePlaceholderParagraphText_(para.getText())) {
        outList.push(para);
      }
    } else if (
      type === DocumentApp.ElementType.TABLE ||
      type === DocumentApp.ElementType.TABLE_ROW ||
      type === DocumentApp.ElementType.TABLE_CELL
    ) {
      collectPurePlaceholderParagraphsInContainer_(child, outList);
    }
  }
}

function collectPurePlaceholderParagraphs_(sections) {
  const outList = [];
  if (!Array.isArray(sections)) return outList;

  for (let s = 0; s < sections.length; s++) {
    collectPurePlaceholderParagraphsInContainer_(sections[s], outList);
  }
  return outList;
}

function canSafelyRemoveParagraph_(para) {
  if (!para) return false;
  let parent = null;
  try {
    parent = para.getParent();
  } catch (e) {
    return false;
  }
  if (!parent || typeof parent.getNumChildren !== 'function') return false;

  const numChildren = parent.getNumChildren();
  if (numChildren <= 1) return false;

  const parentType = parent.getType();
  if (
    parentType === DocumentApp.ElementType.BODY_SECTION ||
    parentType === DocumentApp.ElementType.TABLE_CELL ||
    parentType === DocumentApp.ElementType.HEADER_SECTION ||
    parentType === DocumentApp.ElementType.FOOTER_SECTION
  ) {
    let textBlockCount = 0;
    for (let i = 0; i < numChildren; i++) {
      const childType = parent.getChild(i).getType();
      if (
        childType === DocumentApp.ElementType.PARAGRAPH ||
        childType === DocumentApp.ElementType.LIST_ITEM
      ) {
        textBlockCount++;
        if (textBlockCount > 1) {
          return true;
        }
      }
    }
    return false;
  }

  return true;
}

function removeResolvedEmptyPlaceholderParagraphs_(candidateParagraphs) {
  if (!Array.isArray(candidateParagraphs) || candidateParagraphs.length === 0) return;

  for (let i = candidateParagraphs.length - 1; i >= 0; i--) {
    const para = candidateParagraphs[i];
    if (!para) continue;

    try {
      const currentText = String(para.getText() || '').trim();
      if (currentText !== '') continue;
      if (!canSafelyRemoveParagraph_(para)) continue;
      para.removeFromParent();
    } catch (e) {}
  }
}

function cleanupEmptyStaticLabels_(sections) {
  const emptyPairPatterns = [
    '\\(\\s*[Tt][Ee][Ll]\\s*:\\s*,?\\s*[Ee]-?[Mm][Aa][Ii][Ll]\\s*:\\s*\\)',
    '\\(\\s*[Tt][Ee][Ll]\\s*:\\s*,?\\s*[Ff][Aa][Xx]\\s*:\\s*,?\\s*[Ee]-?[Mm][Aa][Ii][Ll]\\s*:\\s*\\)',
    '\\(\\s*[Tt][Ee][Ll]\\s*:\\s*\\)',
    '\\(\\s*[Pp][Hh][Oo][Nn][Ee]\\s*:\\s*\\)',
    '\\(\\s*[Mm][Oo][Bb][Ii][Ll][Ee]\\s*:\\s*\\)',
    '\\(\\s*[Ee]-?[Mm][Aa][Ii][Ll]\\s*:\\s*\\)',
    '\\(\\s*[Ff][Aa][Xx]\\s*:\\s*\\)',
    '\\(\\s*[Mm][Ss][Tt]\\s*:\\s*\\)',
    '\\(\\s*[Tt][Aa][Xx]\\s*:\\s*\\)',
    '\\(\\s*[Tt][Aa][Xx]\\s+[Cc][Oo][Dd][Ee]\\s*:\\s*\\)',
    '\\(\\s*[Aa][Tt][Tt][Nn]\\s*:\\s*\\)'
  ];

  const trailingEmptyPatterns = [
    '[Tt][Ee][Ll]\\s*:\\s*$',
    '[Pp][Hh][Oo][Nn][Ee]\\s*:\\s*$',
    '[Ee]-?[Mm][Aa][Ii][Ll]\\s*:\\s*$',
    '[Ff][Aa][Xx]\\s*:\\s*$',
    '[Mm][Ss][Tt]\\s*:\\s*$',
    '[Tt][Aa][Xx]\\s+[Cc][Oo][Dd][Ee]\\s*:\\s*$',
    '[Aa][Tt][Tt][Nn]\\s*:\\s*$',
    '[Tt][Ee][Ll]\\s*:\\s*,\\s*',
    '[Ff][Aa][Xx]\\s*:\\s*,\\s*',
    '[Ee]-?[Mm][Aa][Ii][Ll]\\s*:\\s*,\\s*'
  ];

  for (let s = 0; s < sections.length; s++) {
    const section = sections[s];
    for (let i = 0; i < emptyPairPatterns.length; i++) {
      section.replaceText(emptyPairPatterns[i], '');
    }
    for (let j = 0; j < trailingEmptyPatterns.length; j++) {
      section.replaceText(trailingEmptyPatterns[j], '');
    }
    section.replaceText('\\(\\s*,\\s*', '(');
    section.replaceText(',\\s*\\)', ')');
    section.replaceText('\\(\\s*\\)', '');
  }
}

function collectUniqueItemFieldValues_(items, fieldKey, transformer) {
  const list = [];
  if (!Array.isArray(items)) return list;
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (!it || typeof it !== 'object') continue;
    let val = String(it[fieldKey] == null ? '' : it[fieldKey]).trim();
    if (!val) continue;
    if (typeof transformer === 'function') {
      val = transformer(val);
    }
    if (val && list.indexOf(val) === -1) {
      list.push(val);
    }
  }
  return list;
}

function buildDocSpecificShared_(shared, items, containers, docKey) {
  const out = Object.assign({}, shared);
  const safeItems = Array.isArray(items) ? items : [];
  const safeContainers = Array.isArray(containers) ? containers : [];

  const hsJoined = collectUniqueItemFieldValues_(safeItems, 'HS', formatHsCode_).join(', ');
  const contJoined = collectUniqueItemFieldValues_(safeContainers, 'CONT').join(', ');
  const sealJoined = collectUniqueItemFieldValues_(safeContainers, 'SEAL').join(', ');
  const prodJoined = collectUniqueItemFieldValues_(safeItems, 'PROD').join(' & ');

  if (out.HS && String(out.HS).trim() !== '') {
    out.HS = out.HS.split(',').map(function (part) {
      return formatHsCode_(part.trim());
    }).filter(Boolean).join(', ');
  } else if (hsJoined) {
    out.HS = hsJoined;
  }

  if (!out.CONT && contJoined) out.CONT = contJoined;
  if (!out.SEAL && sealJoined) out.SEAL = sealJoined;

  if (!out.CHEMICAL && out.FUM_CHEMICAL) {
    out.CHEMICAL = out.FUM_CHEMICAL;
  }
  if (!out.DOSAGE && out.FUM_DOSAGE) {
    out.DOSAGE = out.FUM_DOSAGE;
  }
  if (!out.EXPOSURE_DURATION && out.FUM_EXPOSURE_TIME) {
    out.EXPOSURE_DURATION = out.FUM_EXPOSURE_TIME;
  }
  if (!out.DESCRIPTION_OF_GOODS && prodJoined) {
    out.DESCRIPTION_OF_GOODS = prodJoined;
  }
  if (!out.FUM_DT && out.FUM_CERT_DT) {
    out.FUM_DT = out.FUM_CERT_DT;
  }

  if (!out.ISS_DT) {
    if (docKey === 'FUMIGATION_CERTIFICATE' && out.FUM_CERT_DT) {
      out.ISS_DT = out.FUM_CERT_DT;
    } else if (docKey === 'PHYTOSANITARY_CERTIFICATE' && out.PHYTO_CERT_DT) {
      out.ISS_DT = out.PHYTO_CERT_DT;
    } else if (docKey === 'QUALITY_CERTIFICATE' && out.QUAL_CERT_DT) {
      out.ISS_DT = out.QUAL_CERT_DT;
    } else if (docKey === 'VETERINARY_CERTIFICATE' && out.VET_CERT_DT) {
      out.ISS_DT = out.VET_CERT_DT;
    } else if (docKey === 'CERTIFICATE_OF_ORIGIN' && out.CO_ISSUE_DT) {
      out.ISS_DT = out.CO_ISSUE_DT;
    }
  }

  ['TOTAL_CONTAINERS_TEXT', 'TOT_AMT_WORDS', 'TOT_AMT_WORDS_VI', 'TOT_AMT_WORDS_SC', 'TOT_AMT_WORDS_CI', 'TOT_TAX_WORDS'].forEach(function (sayKey) {
    if (out[sayKey] != null && String(out[sayKey]).trim() !== '') {
      out[sayKey] = String(out[sayKey]).replace(/^\s*SAY\s*:\s*/i, '').trim();
    }
  });

  if (docKey === 'QUOTATION') {
    if (out.VALID_DAYS != null && String(out.VALID_DAYS).trim() !== '') {
      out.VALID_DAYS = String(out.VALID_DAYS).replace(/\s*(?:days?|ngày)\s*$/i, '').trim();
    }
    if (out.BANK_NAME != null && String(out.BANK_NAME).trim() !== '') {
      out.BANK_NAME = String(out.BANK_NAME).replace(/\s*[-–—]\s*/g, ' - ').replace(/\s{2,}/g, ' ').trim();
    }
  }

  if (docKey === 'SALES_CONTRACT') {
    if (!out.PROD && prodJoined) out.PROD = prodJoined;
    if (!out.SELLER_REP) out.SELLER_REP = out.SALE_NAME || out.DOCS_NAME || '';
    if (!out.BUYER_REP) out.BUYER_REP = out.BUYER_NAME || '';
    if (!out.SHIPMENT_DEADLINE && out.ETD_DT) out.SHIPMENT_DEADLINE = out.ETD_DT;
    if (out.TOT_AMT) out.TOT_AMT = formatDocFixed2_(out.TOT_AMT);
    const firstScUnit = collectUniqueItemFieldValues_(safeItems, 'UNIT', function (u) {
      return extractConciseCode_(u, 'UNIT');
    })[0] || '';
    if (!out.TOT_UNIT && (firstScUnit || out.UNIT)) {
      out.TOT_UNIT = firstScUnit || extractConciseCode_(out.UNIT, 'UNIT');
    }
    if (safeItems.length === 1) {
      const singlePrice = safeItems[0].PRICE != null && String(safeItems[0].PRICE).trim() !== ''
        ? formatDocFixed2_(safeItems[0].PRICE)
        : (out.PRICE ? formatDocFixed2_(out.PRICE) : '');
      const singleUnit = extractConciseCode_(safeItems[0].UNIT || out.TOT_UNIT || out.UNIT || '', 'UNIT');
      if (singlePrice) out.PRICE = singlePrice;
      if (singleUnit) out.UNIT = singleUnit;
    } else if (safeItems.length > 1) {
      const priceParts = safeItems.map(function (it) {
        if (!it || typeof it !== 'object') return '';
        const p = it.PRICE != null && String(it.PRICE).trim() !== '' ? formatDocFixed2_(it.PRICE) : '';
        const u = extractConciseCode_(it.UNIT || out.TOT_UNIT || '', 'UNIT');
        return p ? (u ? (p + '/' + u) : p) : '';
      }).filter(Boolean);
      if (priceParts.length > 0) {
        out.PRICE = priceParts.join('; ');
      }
      out.UNIT = '';
    }
  }

  if (docKey === 'PACKING_LIST') {
    if (!out.FINAL_DESTINATION && out.POD) {
      out.FINAL_DESTINATION = out.POD;
    }
    if (out.TOT_CBM) {
      out.TOT_CBM = formatDocFixed2_(out.TOT_CBM);
    }
  }

  if (docKey === 'FUMIGATION_CERTIFICATE') {
    if (!out.QTY && out.TOT_QTY) out.QTY = out.TOT_QTY;
    if (!out.UNIT && out.TOT_UNIT) out.UNIT = out.TOT_UNIT;
    if (!out.PACK && out.TOT_PACK) out.PACK = out.TOT_PACK;
    if (!out.UNIT_PACK && out.PACK_UNIT) out.UNIT_PACK = out.PACK_UNIT;
    if (!out.PACKAGE_TYPE && out.PACK_UNIT) out.PACKAGE_TYPE = out.PACK_UNIT;
    if (out.TOT_GW) out.TOT_GW = formatDocFixed2_(out.TOT_GW);
    if (out.TOT_NW) out.TOT_NW = formatDocFixed2_(out.TOT_NW);
  }

  if (docKey === 'CERTIFICATE_OF_ORIGIN') {
    if (!out.CO_DT && out.CO_ISSUE_DT) {
      out.CO_DT = out.CO_ISSUE_DT;
    }
    if (!out.CO_PLACE) {
      out.CO_PLACE = out.ISSUE_PLACE || '';
    }
    if (!out.TRADE_PARTNER && out.DEST_CTRY) {
      out.TRADE_PARTNER = out.DEST_CTRY;
    }
  }

  if (docKey === 'PHYTOSANITARY_CERTIFICATE' || docKey === 'QUALITY_CERTIFICATE') {
    if (!out.PROD && prodJoined) out.PROD = prodJoined;
    const pcodeJoined = collectUniqueItemFieldValues_(safeItems, 'PCODE').join(', ');
    const desJoined = collectUniqueItemFieldValues_(safeItems, 'DES').join('; ');
    const lotJoined = collectUniqueItemFieldValues_(safeItems, 'LOT_NO').join(', ');
    const marksJoined = collectUniqueItemFieldValues_(safeItems, 'MARKS').join(', ');
    const moistureJoined = collectUniqueItemFieldValues_(safeItems, 'SPEC_MOISTURE').join(' / ');
    const brokenJoined = collectUniqueItemFieldValues_(safeItems, 'SPEC_BROKEN').join(' / ');

    if (!out.PCODE && pcodeJoined) out.PCODE = pcodeJoined;
    if (!out.DES && desJoined) out.DES = desJoined;
    if (!out.LOT_NO && lotJoined) out.LOT_NO = lotJoined;
    if (!out.MARKS && marksJoined) out.MARKS = marksJoined;
    if (!out.SPEC_MOISTURE && moistureJoined) out.SPEC_MOISTURE = moistureJoined;
    if (!out.SPEC_BROKEN && brokenJoined) out.SPEC_BROKEN = brokenJoined;
    if (!out.UNIT && out.TOT_UNIT) out.UNIT = out.TOT_UNIT;
    if (!out.QTY && out.TOT_QTY) out.QTY = out.TOT_QTY;
    if (!out.PACK && out.TOT_PACK) out.PACK = out.TOT_PACK;
    if (!out.UNIT_PACK && out.PACK_UNIT) out.UNIT_PACK = out.PACK_UNIT;
    if (!out.PACKAGE_TYPE && out.PACK_UNIT) out.PACKAGE_TYPE = out.PACK_UNIT;
    if (!out.NW && out.TOT_NW) out.NW = out.TOT_NW;
    if (!out.GW && out.TOT_GW) out.GW = out.TOT_GW;
    if (!out.CBM && out.TOT_CBM) out.CBM = formatDocFixed2_(out.TOT_CBM);
    if (!out.TREAT_CHEMICAL && out.CHEMICAL_NAME) out.TREAT_CHEMICAL = out.CHEMICAL_NAME;
    if (!out.CHEMICAL_NAME && out.TREAT_CHEMICAL) out.CHEMICAL_NAME = out.TREAT_CHEMICAL;
    if (!out.INSPECTION_CONCLUSION && out.TEST_RESULT) out.INSPECTION_CONCLUSION = out.TEST_RESULT;
    if (!out.TEST_RESULT && out.INSPECTION_CONCLUSION) out.TEST_RESULT = out.INSPECTION_CONCLUSION;
  }

  if (docKey === 'LETTER_OF_CREDIT') {
    let lcItemsSum = 0;
    let hasLcItems = false;
    safeItems.forEach(function (it) {
      if (!it || typeof it !== 'object') return;
      const amtVal = parseDocNumber_(it.AMT);
      if (amtVal !== null) {
        lcItemsSum += amtVal;
        hasLcItems = true;
      } else {
        const q = parseDocNumber_(it.QTY);
        const p = parseDocNumber_(it.PRICE);
        if (q !== null && p !== null) {
          lcItemsSum += q * p;
          hasLcItems = true;
        }
      }
    });
    const baseTotAmt = hasLcItems ? lcItemsSum : (parseDocNumber_(out.TOT_AMT) || 0);
    const freightVal = parseDocNumber_(out.FREIGHT_AMT) || 0;
    const insVal = parseDocNumber_(out.INS_AMT) || 0;
    const lcGrandTotal = baseTotAmt + freightVal + insVal;
    if (hasLcItems || parseDocNumber_(out.TOT_AMT) !== null) {
      out.TOT_AMT = formatDocFixed2_(baseTotAmt);
    }
    out.FREIGHT_AMT = formatDocFixed2_(freightVal);
    out.INS_AMT = formatDocFixed2_(insVal);
    out.GRAND_TOTAL = formatDocFixed2_(lcGrandTotal);
    out.LC_AMOUNT = out.GRAND_TOTAL;
    out.TOTAL_AMOUNT = out.TOT_AMT;
  }

  if (docKey === 'To_khai_XK' || docKey === 'To_khai_NK') {
    let itemsAmtSum = 0;
    let hasItemsAmt = false;
    let sumExTax = 0;
    let sumFee = 0;
    let hasXkTax = false;
    safeItems.forEach(function (it) {
      if (!it || typeof it !== 'object') return;
      const a = parseDocNumber_(it.AMT);
      if (a !== null) {
        itemsAmtSum += a;
        hasItemsAmt = true;
      } else {
        const q = parseDocNumber_(it.QTY);
        const p = parseDocNumber_(it.PRICE);
        if (q !== null && p !== null) {
          itemsAmtSum += q * p;
          hasItemsAmt = true;
        }
      }
      const exVal = parseDocNumber_(it.EX_VAL);
      const exRate = parseDocNumber_(it.EX_TAX_RATE);
      const exAmt = parseDocNumber_(it.EX_AMT);
      if (exAmt !== null) {
        sumExTax += exAmt;
        hasXkTax = true;
      } else if (exVal !== null && exRate !== null) {
        sumExTax += exVal * (exRate / 100);
        hasXkTax = true;
      }
      const feeVal = parseDocNumber_(it.FEE_VAL);
      const feeRate = parseDocNumber_(it.FEE_RATE);
      const feeAmt = parseDocNumber_(it.FEE_AMT);
      if (feeAmt !== null) {
        sumFee += feeAmt;
        hasXkTax = true;
      } else if (feeVal !== null && feeRate !== null) {
        sumFee += feeVal * (feeRate / 100);
        hasXkTax = true;
      }
    });

    if (hasItemsAmt) {
      out.TOT_AMT = formatDocFixed2_(itemsAmtSum);
    } else if (out.TOT_AMT) {
      out.TOT_AMT = formatDocFixed2_(out.TOT_AMT);
    }
    if (out.TOT_GW) out.TOT_GW = formatDocFixed2_(out.TOT_GW);
    if (out.TOT_NW) out.TOT_NW = formatDocFixed2_(out.TOT_NW);

    if (docKey === 'To_khai_XK') {
      if (hasXkTax) {
        out.TOT_EX_TAX = formatDocFixed2_(sumExTax);
        out.TOT_FEE = formatDocFixed2_(sumFee);
        out.TOT_TAX = formatDocFixed2_(sumExTax + sumFee);
      } else {
        if (out.TOT_EX_TAX) out.TOT_EX_TAX = formatDocFixed2_(out.TOT_EX_TAX);
        if (out.TOT_FEE) out.TOT_FEE = formatDocFixed2_(out.TOT_FEE);
        if (out.TOT_TAX) out.TOT_TAX = formatDocFixed2_(out.TOT_TAX);
      }
    } else {
      const exRateNum = parseDocNumber_(out.EX_RATE) || 0;
      const baseTotAmt = hasItemsAmt ? itemsAmtSum : (parseDocNumber_(out.TOT_AMT) || 0);
      let imValNum = parseDocNumber_(out.IM_VAL);
      if (imValNum === null && baseTotAmt > 0 && exRateNum > 0) {
        imValNum = baseTotAmt * exRateNum;
      }
      imValNum = imValNum || 0;
      const imRateNum = parseDocNumber_(out.IM_RATE) || 0;
      const imAmtNum = imValNum * (imRateNum / 100);

      const sctRateNum = parseDocNumber_(out.SCT_RATE) || 0;
      let sctValNum = parseDocNumber_(out.SCT_VAL);
      if (sctValNum === null) {
        sctValNum = sctRateNum > 0 ? (imValNum + imAmtNum) : 0;
      }
      const sctAmtNum = sctValNum * (sctRateNum / 100);

      const eptValNum = parseDocNumber_(out.EPT_VAL) || 0;
      const eptRateNum = parseDocNumber_(out.EPT_RATE) || 0;
      const eptAmtNum = eptValNum * eptRateNum;

      const vatRateNum = parseDocNumber_(out.VAT_RATE) || 0;
      let vatValNum = parseDocNumber_(out.VAT_VAL);
      if (vatValNum === null) {
        vatValNum = vatRateNum > 0 ? (imValNum + imAmtNum + sctAmtNum + eptAmtNum) : 0;
      }
      const vatAmtNum = vatValNum * (vatRateNum / 100);
      const nkTotTax = imAmtNum + sctAmtNum + eptAmtNum + vatAmtNum;

      if (imValNum > 0 || imRateNum > 0) {
        out.IM_VAL = formatDocFixed2_(imValNum);
        out.IM_RATE = formatDocNumber_(imRateNum);
        out.IM_AMT = formatDocFixed2_(imAmtNum);
      }
      if (sctValNum > 0 || sctRateNum > 0) {
        out.SCT_VAL = formatDocFixed2_(sctValNum);
        out.SCT_RATE = formatDocNumber_(sctRateNum);
        out.SCT_AMT = formatDocFixed2_(sctAmtNum);
      }
      if (eptValNum > 0 || eptRateNum > 0) {
        out.EPT_VAL = formatDocFixed2_(eptValNum);
        out.EPT_RATE = formatDocNumber_(eptRateNum);
        out.EPT_AMT = formatDocFixed2_(eptAmtNum);
      }
      if (vatValNum > 0 || vatRateNum > 0) {
        out.VAT_VAL = formatDocFixed2_(vatValNum);
        out.VAT_RATE = formatDocNumber_(vatRateNum);
        out.VAT_AMT = formatDocFixed2_(vatAmtNum);
      }
      if (imValNum > 0 || imRateNum > 0 || sctRateNum > 0 || eptRateNum > 0 || vatRateNum > 0) {
        out.TOT_TAX = formatDocFixed2_(nkTotTax);
      }
    }

    const customsExtras = buildCustomsExtras_(out, docKey);
    if (!out.ATTACHED_DOCS) out.ATTACHED_DOCS = customsExtras.attachedDocsText;
    if (!out.ATTACHED_DOCUMENTS) out.ATTACHED_DOCUMENTS = customsExtras.attachedDocsText;
    if (!out.DOC_LIST) out.DOC_LIST = customsExtras.attachedDocsText;
    if (!out.DOCS_LIST) out.DOCS_LIST = customsExtras.attachedDocsText;
    if (!out.CHUNG_TU_KEM_THEO) out.CHUNG_TU_KEM_THEO = customsExtras.attachedDocsText;
    if (!out.ACCOMPANYING_DOCS) out.ACCOMPANYING_DOCS = customsExtras.attachedDocsText;
    if (!out.DEC_DT) out.DEC_DT = customsExtras.declDateStr;
    if (!out.DEC_DATE) out.DEC_DATE = customsExtras.declDateStr;
    if (!out.DECL_DT) out.DECL_DT = customsExtras.declDateStr;
    if (!out.REG_DT) out.REG_DT = customsExtras.declDateStr;
    if (!out.SIGN_DT) out.SIGN_DT = customsExtras.declDateStr;
    if (!out.ISS_DT) out.ISS_DT = customsExtras.declDateStr;
    if (!out.SIGN_DATE_TEXT) out.SIGN_DATE_TEXT = customsExtras.formalViDate;
    out.DEC_DAY = customsExtras.day;
    out.DEC_MONTH = customsExtras.month;
    out.DEC_YEAR = customsExtras.year;
  }

  return out;
}

function fixBookingConfirmationRemarksColon_(sections) {
  if (!Array.isArray(sections)) return;
  const pattern = ':[^\\s0-9:/\\\\]';
  for (let s = 0; s < sections.length; s++) {
    const sec = sections[s];
    if (!sec || typeof sec.findText !== 'function') continue;
    let rangeEl = sec.findText(pattern);
    let guard = 0;
    while (rangeEl && guard < 300) {
      guard++;
      try {
        const textEl = rangeEl.getElement().asText();
        const colonOffset = rangeEl.getStartOffset();
        const fullText = textEl.getText();
        const beforeSlice = fullText.slice(Math.max(0, colonOffset - 7), colonOffset + 1).toLowerCase();
        const isProtocolOrDrive = /https?:$/.test(beforeSlice) || /mailto:$/.test(beforeSlice);
        if (!isProtocolOrDrive && fullText.charAt(colonOffset + 1) !== ' ') {
          textEl.insertText(colonOffset + 1, ' ');
        }
      } catch (e) {}
      rangeEl = sec.findText(pattern, rangeEl);
    }
  }
}

function walkContainerParagraphsAndCells_(container, onParagraph, onCell) {
  if (!container || typeof container.getNumChildren !== 'function') return;
  const numChildren = container.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = container.getChild(i);
    const type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      if (typeof onParagraph === 'function') onParagraph(child.asParagraph());
    } else if (type === DocumentApp.ElementType.TABLE_CELL) {
      const cell = child.asTableCell();
      if (typeof onCell === 'function') onCell(cell);
      walkContainerParagraphsAndCells_(cell, onParagraph, onCell);
    } else if (
      type === DocumentApp.ElementType.TABLE ||
      type === DocumentApp.ElementType.TABLE_ROW
    ) {
      walkContainerParagraphsAndCells_(child, onParagraph, onCell);
    }
  }
}

function prepareSalesContractTemplateTags_(sections) {
  if (!Array.isArray(sections)) return;
  for (let s = 0; s < sections.length; s++) {
    let inArticle5 = false;
    walkContainerParagraphsAndCells_(sections[s], function (para) {
      const text = String(para.getText() || '');
      if (/ARTICLE\s*5\b/i.test(text)) {
        inArticle5 = true;
      } else if (/ARTICLE\s*[6-9]\b|ARTICLE\s*1\d\b/i.test(text)) {
        inArticle5 = false;
      }

      if ((inArticle5 || /^\s*Product\s*:/i.test(text)) && text.indexOf('{{COMMODITY_GRADE}}') !== -1) {
        para.replaceText('\\{\\{PROD\\}\\}\\s*[-–—,/|]*\\s*\\(?\\{\\{COMMODITY_GRADE\\}\\}\\)?', '{{PROD}}');
        para.replaceText('\\{\\{COMMODITY_GRADE\\}\\}\\s*[-–—,/|]*\\s*\\{\\{PROD\\}\\}', '{{PROD}}');
        para.replaceText('\\s*[-–—,/|]*\\s*\\(?\\{\\{COMMODITY_GRADE\\}\\}\\)?', '');
      }

      if (/Quantity\s*:/i.test(text) && text.indexOf('{{TOT_QTY}}') !== -1 && text.indexOf('{{TOT_UNIT}}') === -1 && text.indexOf('{{UNIT}}') === -1) {
        if (!/\{\{TOT_QTY\}\}\s+[A-Za-z]/.test(text)) {
          para.replaceText('\\{\\{TOT_QTY\\}\\}', '{{TOT_QTY}} {{TOT_UNIT}}');
        }
      }

      if (/Total\s*Amount\s*:/i.test(text) && text.indexOf('{{TOT_AMT}}') !== -1 && text.indexOf('{{CURR}}') === -1) {
        para.replaceText('\\{\\{TOT_AMT\\}\\}', '{{CURR}} {{TOT_AMT}}');
      }
    }, null);
  }
}

function cleanupSalesContractUnitPriceSlash_(sections) {
  if (!Array.isArray(sections)) return;
  for (let s = 0; s < sections.length; s++) {
    walkContainerParagraphsAndCells_(sections[s], function (para) {
      const text = String(para.getText() || '');
      if (!/Unit\s*Price/i.test(text)) return;
      para.replaceText('\\s*/\\s*$', '');
      if (/Unit\s*Price\s*:\s*\/\s*/i.test(String(para.getText() || ''))) {
        para.replaceText('Unit\\s*Price\\s*:\\s*/\\s*', 'Unit Price: ');
      }
    }, null);
  }
}

function prepareBillOfLadingPackageUnitTags_(sections) {
  if (!Array.isArray(sections)) return;
  for (let s = 0; s < sections.length; s++) {
    const sec = sections[s];
    if (!sec || typeof sec.getTables !== 'function') continue;
    const tables = sec.getTables();
    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        for (let c = 0; c < row.getNumCells(); c++) {
          const cell = row.getCell(c);
          const cellText = String(cell.getText() || '');
          if (cellText.indexOf('{{TOT_PACK}}') !== -1 && cellText.indexOf('{{PACK_UNIT}}') === -1) {
            if (!/\{\{TOT_PACK\}\}\s+[A-Za-z]/.test(cellText)) {
              cell.replaceText('\\{\\{TOT_PACK\\}\\}', '{{TOT_PACK}} {{PACK_UNIT}}');
            }
          }
        }
      }
    }
  }
}

function cleanupBillOfLadingBoxes_(sections, docKey) {
  if (!Array.isArray(sections)) return;
  const contSealHeaderPattern = /CONTAINER\s*NO\.?\s*\/?\s*SEAL\s*NO\.?\s*:?/i;
  const contSealHeaderInlineRegex = /(CONTAINER\s*NO\.?\s*\/?\s*SEAL\s*NO\.?\s*:?)/ig;

  for (let s = 0; s < sections.length; s++) {
    const sec = sections[s];
    if (!sec || typeof sec.getTables !== 'function') continue;
    const tables = sec.getTables();

    for (let t = 0; t < tables.length; t++) {
      const table = tables[t];
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        const rowText = String(row.getText() || '');
        const prevRow = r > 0 ? table.getRow(r - 1) : null;
        const prevRowText = prevRow ? String(prevRow.getText() || '') : '';

        const isBox24Row =
          /\(\s*24\s*\)/.test(rowText) ||
          /^\s*SAY\s*:/im.test(rowText) ||
          /CONTAINERS?\s+ONLY\s*$/im.test(rowText);

        if (docKey === 'DELIVERY_ORDER' && isBox24Row) {
          for (let c24 = 0; c24 < row.getNumCells(); c24++) {
            const cell24 = row.getCell(c24);
            const cell24Text = String(cell24.getText() || '');
            if (!/\(\s*24\s*\)|SAY\s*:|CONTAINERS?\s+ONLY/i.test(cell24Text)) continue;
            try {
              cell24.setPaddingTop(2);
              cell24.setPaddingBottom(2);
            } catch (e) {}
            for (let p24 = cell24.getNumChildren() - 1; p24 >= 0; p24--) {
              const ch24 = cell24.getChild(p24);
              if (ch24.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
              const pObj = ch24.asParagraph();
              const pTxt = String(pObj.getText() || '').trim();
              if (pTxt === '' && cell24.getNumChildren() > 1) {
                try {
                  pObj.removeFromParent();
                } catch (e) {}
              } else {
                try {
                  pObj.setKeepLinesTogether(true);
                  pObj.setLineSpacing(1.0);
                  pObj.setSpacingBefore(0);
                  pObj.setSpacingAfter(0);
                } catch (e) {}
              }
            }
          }
        }

        const isBoxes18To21Row =
          contSealHeaderPattern.test(rowText) ||
          contSealHeaderPattern.test(prevRowText) ||
          /\(\s*18\s*\)/.test(prevRowText) ||
          /\(\s*19\s*\)/.test(prevRowText) ||
          /\bKGs\b/i.test(rowText);

        if (!isBoxes18To21Row) continue;

        for (let c = 0; c < row.getNumCells(); c++) {
          const cell = row.getCell(c);
          const prevCellHasHeader =
            prevRow &&
            c < prevRow.getNumCells() &&
            contSealHeaderPattern.test(String(prevRow.getCell(c).getText() || ''));

          let seenHeaderInCell = Boolean(prevCellHasHeader);

          for (let pIdx = cell.getNumChildren() - 1; pIdx >= 0; pIdx--) {
            const child = cell.getChild(pIdx);
            if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
            const para = child.asParagraph();
            const rawText = String(para.getText() || '');
            const trimmed = rawText.trim();

            if (trimmed === '') {
              if (cell.getNumChildren() > 1) {
                try {
                  para.removeFromParent();
                } catch (e) {}
              }
              continue;
            }

            if (contSealHeaderPattern.test(trimmed)) {
              const remainder = trimmed.replace(contSealHeaderInlineRegex, '').trim();
              if (remainder === '') {
                if (seenHeaderInCell && cell.getNumChildren() > 1) {
                  try {
                    para.removeFromParent();
                  } catch (e) {}
                  continue;
                }
                seenHeaderInCell = true;
              } else {
                let firstOccurrenceKept = !seenHeaderInCell;
                const cleanedInline = rawText.replace(contSealHeaderInlineRegex, function (m) {
                  if (!firstOccurrenceKept) {
                    return '';
                  }
                  firstOccurrenceKept = false;
                  return m;
                }).replace(/\s{2,}/g, ' ').trim();
                seenHeaderInCell = true;
                if (cleanedInline !== trimmed) {
                  para.setText(cleanedInline);
                }
              }
            }
          }

          if (c === 0 && !prevCellHasHeader) {
            let firstSeen = false;
            for (let pIdx = 0; pIdx < cell.getNumChildren(); pIdx++) {
              const child = cell.getChild(pIdx);
              if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
              const para = child.asParagraph();
              const tVal = String(para.getText() || '').trim();
              if (contSealHeaderPattern.test(tVal) && tVal.replace(contSealHeaderInlineRegex, '').trim() === '') {
                if (!firstSeen) {
                  firstSeen = true;
                } else if (cell.getNumChildren() > 1) {
                  try {
                    para.removeFromParent();
                    pIdx--;
                  } catch (e) {}
                }
              }
            }
          }
        }
      }
    }
  }
}

function parseCustomsDateParts_(shared) {
  const candidates = [
    shared && shared.DEC_DT,
    shared && shared.REG_DTM,
    shared && shared.SUB_DTM,
    shared && shared.INV_DT,
    shared && shared.CTR_DT
  ];

  for (let i = 0; i < candidates.length; i++) {
    const raw = String(candidates[i] == null ? '' : candidates[i]).trim();
    if (!raw) continue;
    const dmy = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dmy) {
      return {
        day: dmy[1].padStart(2, '0'),
        month: dmy[2].padStart(2, '0'),
        year: dmy[3]
      };
    }
    const ymd = raw.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (ymd) {
      return {
        day: ymd[3].padStart(2, '0'),
        month: ymd[2].padStart(2, '0'),
        year: ymd[1]
      };
    }
  }

  let tz = 'Asia/Ho_Chi_Minh';
  try {
    tz = Session.getScriptTimeZone() || tz;
  } catch (e) {}
  const nowStr = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  const parts = nowStr.split('/');
  return {
    day: parts[0] || '01',
    month: parts[1] || '01',
    year: parts[2] || '2026'
  };
}

function buildCustomsExtras_(shared, docKey) {
  const s = shared || {};
  const docs = [];
  if (s.CTR_NO) {
    docs.push('- Sales Contract No. ' + s.CTR_NO + (s.CTR_DT ? ' (' + s.CTR_DT + ')' : ''));
  }
  if (s.INV_NO) {
    docs.push('- Commercial Invoice No. ' + s.INV_NO + (s.INV_DT ? ' (' + s.INV_DT + ')' : ''));
  }
  if (s.PL_DT || s.INV_NO) {
    docs.push('- Packing List' + (s.PL_DT ? ' (' + s.PL_DT + ')' : ''));
  }
  if (docKey === 'To_khai_NK' && s.BL_NO) {
    docs.push('- Bill of Lading No. ' + s.BL_NO + (s.BL_DT ? ' (' + s.BL_DT + ')' : ''));
  } else if (s.BK_NO) {
    docs.push('- Booking Note No. ' + s.BK_NO + (s.BK_DT ? ' (' + s.BK_DT + ')' : ''));
  }
  if (s.CO_REF_NO) {
    docs.push('- Certificate of Origin No. ' + s.CO_REF_NO + (s.CO_FORM_NAME ? ' (' + s.CO_FORM_NAME + ')' : ''));
  }
  if (s.LIC_NO) {
    docs.push('- License No. ' + s.LIC_NO + (s.LIC_DT ? ' (' + s.LIC_DT + ')' : ''));
  }

  const dateParts = parseCustomsDateParts_(s);
  const declDateStr = dateParts.day + '/' + dateParts.month + '/' + dateParts.year;

  return {
    attachedDocsText: docs.join('\n'),
    day: dateParts.day,
    month: dateParts.month,
    year: dateParts.year,
    declDateStr: declDateStr,
    formalViDate: 'Ngày ' + dateParts.day + ' tháng ' + dateParts.month + ' năm ' + dateParts.year
  };
}

function populateCustomsBoxesIfEmpty_(body, docKey, extras) {
  if (!body || !extras) return;
  const tables = body.getTables();
  const docsBoxPattern = docKey === 'To_khai_XK'
    ? /^\s*26\.\s*Chứng từ/i
    : /^\s*32\.\s*Chứng từ/i;
  const dateBoxPattern = docKey === 'To_khai_XK'
    ? /^\s*27\./
    : /^\s*33\./;

  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        const cellText = String(cell.getText() || '');

        if (extras.attachedDocsText && docsBoxPattern.test(cellText)) {
          const alreadyHasDocs =
            cellText.indexOf('Commercial Invoice') !== -1 ||
            cellText.indexOf('Sales Contract') !== -1;
          if (!alreadyHasDocs) {
            const lines = extras.attachedDocsText.split('\n');
            for (let li = 0; li < lines.length; li++) {
              cell.appendParagraph(lines[li]);
            }
          }
        }

        if (/Ngày\s*\.{2,}\s*tháng\s*\.{2,}\s*năm\s*\.{2,}/i.test(cellText) || /Ngày\s+tháng\s+năm/i.test(cellText)) {
          cell.replaceText('Ngày\\s*\\.{2,}\\s*tháng\\s*\\.{2,}\\s*năm\\s*\\.{2,}', extras.formalViDate);
          cell.replaceText('Ngày\\s+tháng\\s+năm', extras.formalViDate);
        } else if (dateBoxPattern.test(cellText)) {
          const alreadyHasDate =
            /\d{1,2}\/\d{1,2}\/\d{4}/.test(cellText) ||
            /Ngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}/i.test(cellText);
          if (!alreadyHasDate) {
            cell.appendParagraph(extras.formalViDate);
          }
        }
      }
    }
  }
}

function cleanupCustomsBlankRows_(body, docKey) {
  if (!body || typeof body.getTables !== 'function') return;
  const tables = body.getTables();
  const taxTablePattern = /Thuế\s+(?:nhập\s*khẩu|xuất\s*khẩu|TTĐB|BVMT|GTGT)|Sắc\s+thuế|Trị\s+giá\s+tính\s+thuế/i;

  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];
    const tableText = String(table.getText() || '');
    const isTaxTable =
      taxTablePattern.test(tableText) ||
      (docKey === 'To_khai_NK' && /\b26\./.test(tableText) && /\b29\./.test(tableText));
    if (!isTaxTable) continue;

    for (let r = table.getNumRows() - 1; r >= 1; r--) {
      const row = table.getRow(r);
      const compact = String(row.getText() || '').replace(/[\s\u00A0]+/g, '');
      if (compact === '' && table.getNumRows() > 1) {
        try {
          table.removeRow(r);
        } catch (e) {}
      }
    }
  }
}

function normalizeCustomsBoxesColor_(body) {
  if (!body || typeof body.getTables !== 'function') return;
  const tables = body.getTables();
  const box89Pattern = /^\s*(?:8|9)\s*[\.\):\-]|(?:^|\n)\s*(?:8\.|9\.)\s*(?:Người\s+ủy\s+thác|Đại\s+lý|Hợp\s+đồng|Hóa\s+đơn)/i;

  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        const cellText = String(cell.getText() || '');
        if (!box89Pattern.test(cellText)) continue;
        try {
          for (let p = 0; p < cell.getNumChildren(); p++) {
            const child = cell.getChild(p);
            if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
              const txtObj = child.asParagraph().editAsText();
              if (txtObj.getText().length > 0) {
                txtObj.setForegroundColor('#000000');
              }
            }
          }
        } catch (e) {}
      }
    }
  }
}

function applyNoWrapFormattingToCustomsTaxCells_(body) {
  if (!body || typeof body.getTables !== 'function') return;
  const tables = body.getTables();
  const longFormattedNumberRegex = /^-?[\d,]{5,}\.\d{2,}$/;

  for (let t = 0; t < tables.length; t++) {
    const table = tables[t];
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        const trimmed = String(cell.getText() || '').trim();
        if (!longFormattedNumberRegex.test(trimmed)) continue;
        try {
          cell.setPaddingLeft(1);
          cell.setPaddingRight(1);
          if (trimmed.length >= 12) {
            for (let p = 0; p < cell.getNumChildren(); p++) {
              const child = cell.getChild(p);
              if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
                const txt = child.asParagraph().editAsText();
                if (txt.getText().length > 0) {
                  txt.setFontSize(8.5);
                }
              }
            }
          }
        } catch (e) {}
      }
    }
  }
}

function applyKeepLinesTogetherInContainer_(container) {
  if (!container || typeof container.getNumChildren !== 'function') return;
  const numChildren = container.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = container.getChild(i);
    const type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      try {
        child.asParagraph().setKeepLinesTogether(true);
      } catch (e) {}
    } else if (
      type === DocumentApp.ElementType.TABLE ||
      type === DocumentApp.ElementType.TABLE_ROW ||
      type === DocumentApp.ElementType.TABLE_CELL
    ) {
      applyKeepLinesTogetherInContainer_(child);
    }
  }
}

function applyKeepLinesTogetherToDoc_(sections) {
  if (!Array.isArray(sections)) return;
  for (let s = 0; s < sections.length; s++) {
    applyKeepLinesTogetherInContainer_(sections[s]);
  }
}

function applyVeterinarySectionIIILayout_(body, sections) {
  applyKeepLinesTogetherToDoc_(sections);
  if (!body) return;
  let inSection3 = false;
  walkContainerParagraphsAndCells_(
    body,
    function (para) {
      const txt = String(para.getText() || '');
      if (/^III\b/i.test(txt.trim())) {
        inSection3 = true;
      } else if (/^IV\b/i.test(txt.trim())) {
        inSection3 = false;
      }
      if (/Vùng\s+trồng\s*\/\s*Cơ\s+sở/i.test(txt)) {
        para.replaceText('Vùng\\s+trồng\\s*/\\s*Cơ\\s+sở\\s*:', 'Vùng\u00A0trồng/Cơ\u00A0sở\u00A0:');
      }
      if (inSection3 || /Vùng[\s\u00A0]+trồng|Cơ[\s\u00A0]+sở/i.test(txt)) {
        try {
          para.setKeepLinesTogether(true);
          para.setKeepWithNext(true);
        } catch (e) {}
      }
    },
    function (cell) {
      const cTxt = String(cell.getText() || '');
      if (/Vùng[\s\u00A0]+trồng\s*\/?\s*Cơ[\s\u00A0]+sở/i.test(cTxt)) {
        try {
          cell.setPaddingLeft(2);
          cell.setPaddingRight(2);
        } catch (e) {}
      }
    }
  );
}

function formatHsCode_(rawVal) {
  const str = String(rawVal == null ? '' : rawVal).trim();
  if (!str) return '';
  const digits = str.replace(/\D/g, '').slice(0, 10);
  if (!digits) return str;
  if (digits.length <= 4) return digits;
  if (digits.length <= 6) return digits.slice(0, 4) + '.' + digits.slice(4);
  return digits.slice(0, 4) + '.' + digits.slice(4, 6) + '.' + digits.slice(6);
}

function parseDocNumber_(rawVal) {
  if (rawVal == null || rawVal === '') return null;
  const cleaned = String(rawVal).replace(/[,\s]/g, '').trim();
  if (!cleaned || !/^-?\d+(\.\d+)?$/.test(cleaned)) return null;
  const num = Number(cleaned);
  return isNaN(num) ? null : num;
}

function formatDocFixed2_(rawVal) {
  const str = String(rawVal == null ? '' : rawVal).trim();
  if (!str) return '';
  const cleaned = str.replace(/[,\s]/g, '');
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) {
    return str;
  }
  const num = Number(cleaned);
  if (isNaN(num)) return str;
  const fixed = (Math.round(num * 1000) % 10 === 0) ? num.toFixed(2) : num.toFixed(3);
  const parts = fixed.split('.');
  const intPart = parts[0].replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
  return intPart + '.' + parts[1];
}

function pluralizePackUnit_(unitCode, countNum) {
  const clean = extractConciseCode_(unitCode, 'UNIT_PACK').toUpperCase();
  if (!clean) return '';
  if (!(Number(countNum) > 1)) return clean;

  const irregularMap = {
    BAG: 'BAGS',
    CARTON: 'CARTONS',
    CTN: 'CTNS',
    BOX: 'BOXES',
    PACK: 'PACKS',
    PKG: 'PKGS',
    PALLET: 'PALLETS',
    PLT: 'PLTS',
    DRUM: 'DRUMS',
    BALE: 'BALES',
    CASE: 'CASES',
    CRATE: 'CRATES',
    ROLL: 'ROLLS',
    PIECE: 'PIECES',
    SACK: 'SACKS',
    SET: 'SETS',
    UNIT: 'UNITS'
  };
  if (irregularMap[clean]) return irregularMap[clean];
  if (clean.slice(-1) === 'S') return clean;
  if (/(X|CH|SH|Z)$/.test(clean)) return clean + 'ES';
  if (/[^AEIOU]Y$/.test(clean)) return clean.slice(0, -1) + 'IES';
  return clean + 'S';
}

function formatDocNumber_(rawVal) {
  const str = String(rawVal == null ? '' : rawVal).trim();
  if (!str) return '';
  const cleaned = str.replace(/,/g, '');
  if (!/^-?\d+(\.\d+)?$/.test(cleaned)) {
    return str;
  }
  const parts = cleaned.split('.');
  const intPart = parts[0].replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
  if (parts.length > 1) {
    return intPart + '.' + parts[1];
  }
  return intPart;
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

/**
 * GS_Doc_Engine.gs
 * Core Google Docs multi-tab auto-fill engine using 1 master Google Doc ("Chứng từ" - 20 Tabs):
 * 1. Copy MASTER_TEMPLATE_DOC_ID to OUTPUT_FOLDER_ID
 * 2. Open copy via DocumentApp.openById() and read tabs via doc.getTabs()
 * 3. Match tab by tab.getTitle().trim() (case-insensitive) against TEMPLATE_MAP[docKey].tabName
 *    and retrieve Body, Header, Footer via tab.asDocumentTab()
 * 4. Execute 3-step processing per selected Tab:
 *    - Bottom-up target-row cloning (processDynamicTableRows)
 *    - Global single-tag replacement (replaceGlobalTags)
 *    - Sanitizer (/\{\{[A-Z0-9_]+\}\}/g -> "")
 * 5. Save & close, then delete unselected tabs via Google Docs API (Docs.Documents.batchUpdate deleteTab)
 *    with safe try...catch fallback if Advanced Docs Service is not yet enabled.
 * Runtime: Google Apps Script (V8)
 */

/**
 * Tạo 1 hoặc nhiều chứng từ từ file Google Docs tổng hợp (MASTER_TEMPLATE_DOC_ID)
 * và trả về link Google Docs + link tải PDF.
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

    if (!MASTER_TEMPLATE_DOC_ID || MASTER_TEMPLATE_DOC_ID.indexOf('YOUR_') === 0) {
      return {
        success: false,
        results: [],
        message: 'Chưa cấu hình MASTER_TEMPLATE_DOC_ID trong GS_Config_Templates.gs.'
      };
    }

    if (!OUTPUT_FOLDER_ID || OUTPUT_FOLDER_ID.indexOf('YOUR_') === 0) {
      return {
        success: false,
        results: [],
        message: 'Chưa cấu hình OUTPUT_FOLDER_ID trong GS_Config_Templates.gs.'
      };
    }

    const parsedPayload =
      typeof payload === 'string'
        ? JSON.parse(payload || '{}')
        : (payload || {});

    const shared = buildEnrichedShared_(parsedPayload);
    const items = Array.isArray(parsedPayload.items) ? parsedPayload.items : [];
    const containers = Array.isArray(parsedPayload.containers) ? parsedPayload.containers : [];

    // Bước 1: Nhân bản file Google Docs gốc ("Chứng từ") vào OUTPUT_FOLDER_ID
    const suffix = resolveDocFilenameSuffix_(shared, maSV);
    const copyFileName =
      keys.length === 1
        ? keys[0] + '_' + suffix
        : 'BO_CHUNG_TU_' + suffix;

    const masterFile = DriveApp.getFileById(MASTER_TEMPLATE_DOC_ID);
    const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    const copyFile = masterFile.makeCopy(copyFileName, outputFolder);
    const copyFileId = copyFile.getId();
    const docUrl = copyFile.getUrl();
    const pdfUrl = 'https://docs.google.com/document/d/' + copyFileId + '/export?format=pdf';

    // Bước 2: Mở file bản sao và lấy danh sách toàn bộ Tab bằng doc.getTabs()
    const doc = DocumentApp.openById(copyFileId);
    const topLevelTabs = doc.getTabs();
    const allTabs = collectAllTabsFlat_(topLevelTabs);

    const selectedTabIds = [];
    const results = [];
    const errors = [];

    // Bước 3 & 4: Tìm đúng Tab của từng docKey và chạy 3 bước xử lý trên Tab đó
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

        // 4a. Nhân bản dòng bảng động từ dưới lên trên (items[], containers[], tax)
        processDynamicTableRows(body, items, containers);

        // 4b. Thay thế toàn bộ biến đơn {{TAG}} trên Body, Header, Footer của Tab
        replaceGlobalTags(sections, shared);

        // 4c. Quét dọn thẻ sót /\{\{[A-Z0-9_]+\}\}/g -> ""
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

    // Lưu và đóng DocumentApp trước khi gọi Google Docs API để tránh xung đột phiên bản
    doc.saveAndClose();

    if (results.length === 0) {
      // Nếu không có tab nào xử lý thành công, xóa file copy rỗng cho sạch Drive
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

    // Bước 5: Xóa các Tab không được chọn khỏi file bản sao bằng Google Docs API (deleteTab)
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
      message: 'Lỗi hệ thống tạo chứng từ: ' + err.message
    };
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

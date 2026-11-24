import codeGs from '../Code.gs?raw';
import configTemplatesGs from '../GS_Config_Templates.gs?raw';
import docEngineGs from '../GS_Doc_Engine.gs?raw';
import indexHtml from '../Index.html?raw';
import cssStyleHtml from '../CSS_Style.html?raw';
import viewLoginSidebarHtml from '../View_Login_Sidebar.html?raw';
import viewDocumentFormsHtml from '../View_Document_Forms.html?raw';
import jsAppStateHtml from '../JS_App_State.html?raw';
import jsRenderFormsHtml from '../JS_Render_Forms.html?raw';

export interface GasSourceFile {
  id: string;
  filename: string;
  gasType: 'gs' | 'html';
  category: 'Backend (.gs)' | 'Frontend (.html)';
  description: string;
  content: string;
}

export const GAS_FILES: GasSourceFile[] = [
  {
    id: 'Code.gs',
    filename: 'Code.gs',
    gasType: 'gs',
    category: 'Backend (.gs)',
    description: 'doGet, include(), loginUser, getInitialDropdownData, saveUserDraft',
    content: codeGs,
  },
  {
    id: 'GS_Config_Templates.gs',
    filename: 'GS_Config_Templates.gs',
    gasType: 'gs',
    category: 'Backend (.gs)',
    description: 'SPREADSHEET_ID, OUTPUT_FOLDER_ID, TEMPLATE_MAP (20 chứng từ), ITEM/CONT/TAX tags',
    content: configTemplatesGs,
  },
  {
    id: 'GS_Doc_Engine.gs',
    filename: 'GS_Doc_Engine.gs',
    gasType: 'gs',
    category: 'Backend (.gs)',
    description: 'generateDocuments & quy trình 4 bước (Copy, Bottom-Up Clone, Replace, Sanitizer)',
    content: docEngineGs,
  },
  {
    id: 'Index.html',
    filename: 'Index.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Template chính nhúng CSS_Style, View_Login_Sidebar, View_Document_Forms, JS_App_State, JS_Render_Forms',
    content: indexHtml,
  },
  {
    id: 'CSS_Style.html',
    filename: 'CSS_Style.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Giao diện light-mode indigo (#4F46E5), responsive grid, bảng động, badge ISO 6346, toast & spinner',
    content: cssStyleHtml,
  },
  {
    id: 'View_Login_Sidebar.html',
    filename: 'View_Login_Sidebar.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Màn hình đăng nhập, Header cố định và Sidebar 5 nhóm – 20 chứng từ',
    content: viewLoginSidebarHtml,
  },
  {
    id: 'View_Document_Forms.html',
    filename: 'View_Document_Forms.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Khung workspace chính, Modal Xuất hàng loạt và Modal Xác nhận xóa dữ liệu',
    content: viewDocumentFormsHtml,
  },
  {
    id: 'JS_App_State.html',
    filename: 'JS_App_State.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'APP_STATE, đồng bộ 2 chiều, dropdown dây chuyền POL/POD, đọc tiền chữ Anh/Việt, ISO 6346',
    content: jsAppStateHtml,
  },
  {
    id: 'JS_Render_Forms.html',
    filename: 'JS_Render_Forms.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Schema cấu hình & bộ render cho toàn bộ 21 màn hình (MASTER_PARTIES + 20 chứng từ)',
    content: jsRenderFormsHtml,
  },
];

const SANDBOX_SHIM_SCRIPT = `
<script>
(function() {
  var MOCK_USERS = {
    '00123456': { maSV: '00123456', tenSV: 'Nguyễn Văn An', matKhau: '123456' },
    '20210001': { maSV: '20210001', tenSV: 'Trần Thị Bình', matKhau: '123456' }
  };

  var MOCK_DROPDOWNS = {
    success: true,
    cuocTau: [
      { POL: 'CAT LAI, HCMC (VNSGN)', POD: 'LOS ANGELES, US (USLAX)', Quoc_Gia: 'UNITED STATES', Ten_hang_tau: 'MAERSK LINE', Ten_tau: 'MAERSK SENTOSA', So_chuyen: '428E' },
      { POL: 'CAT LAI, HCMC (VNSGN)', POD: 'LONG BEACH, US (USLGB)', Quoc_Gia: 'UNITED STATES', Ten_hang_tau: 'ONE LINE', Ten_tau: 'ONE CONTINUITY', So_chuyen: '064E' },
      { POL: 'CAT LAI, HCMC (VNSGN)', POD: 'ROTTERDAM, NL (NLRTM)', Quoc_Gia: 'NETHERLANDS', Ten_hang_tau: 'MSC', Ten_tau: 'MSC GULSUN', So_chuyen: 'FE412W' },
      { POL: 'CAT LAI, HCMC (VNSGN)', POD: 'TOKYO, JP (JPTYO)', Quoc_Gia: 'JAPAN', Ten_hang_tau: 'WAN HAI', Ten_tau: 'WAN HAI 351', So_chuyen: 'N022' },
      { POL: 'HAI PHONG, VN (VNHPH)', POD: 'BUSAN, KR (KRPUS)', Quoc_Gia: 'SOUTH KOREA', Ten_hang_tau: 'HYUNDAI HMM', Ten_tau: 'HMM ALGECIRAS', So_chuyen: '015N' },
      { POL: 'HAI PHONG, VN (VNHPH)', POD: 'SHANGHAI, CN (CNSHA)', Quoc_Gia: 'CHINA', Ten_hang_tau: 'COSCO SHIPPING', Ten_tau: 'COSCO STAR', So_chuyen: '108E' },
      { POL: 'HAI PHONG, VN (VNHPH)', POD: 'HAMBURG, DE (DEHAM)', Quoc_Gia: 'GERMANY', Ten_hang_tau: 'HAPAG-LLOYD', Ten_tau: 'BERLIN EXPRESS', So_chuyen: '234W' }
    ],
    loaiHinhXNK: [
      { Phan_loai: 'Xuất khẩu', Ma_loai_hinh: 'B11', Ten_loai_hinh: 'Xuất kinh doanh' },
      { Phan_loai: 'Xuất khẩu', Ma_loai_hinh: 'E62', Ten_loai_hinh: 'Xuất sản phẩm sản xuất xuất khẩu' },
      { Phan_loai: 'Xuất khẩu', Ma_loai_hinh: 'B13', Ten_loai_hinh: 'Xuất khẩu hàng đã nhập khẩu' },
      { Phan_loai: 'Nhập khẩu', Ma_loai_hinh: 'A11', Ten_loai_hinh: 'Nhập kinh doanh tiêu dùng' },
      { Phan_loai: 'Nhập khẩu', Ma_loai_hinh: 'A12', Ten_loai_hinh: 'Nhập kinh doanh sản xuất' },
      { Phan_loai: 'Nhập khẩu', Ma_loai_hinh: 'E31', Ten_loai_hinh: 'Nhập nguyên liệu sản xuất xuất khẩu' }
    ],
    danhMucChung: {
      LOAI_CONT: ['20DC', '40DC', '40HC', '45HC', '20RF', '40RH'],
      SERVICE_TYPE: ['CY-CY', 'CFS-CFS', 'CY-CFS', 'CFS-CY', 'DOOR-DOOR'],
      INCOTERM: ['FOB', 'CIF', 'CFR', 'EXW', 'FCA', 'DDP', 'DAP'],
      PAY_TERM: ['T/T 100% IN ADVANCE', 'T/T 30/70', 'L/C AT SIGHT', 'D/P AT SIGHT', 'D/A 60 DAYS'],
      UNIT: ['CTN', 'PCS', 'KGS', 'MTS', 'BAG', 'PALLET', 'SET'],
      CURR: ['USD', 'EUR', 'JPY', 'VND', 'CNY', 'SGD'],
      CO_FORM: ['FORM A', 'FORM B', 'FORM D', 'FORM E', 'FORM AK', 'FORM AJ', 'FORM EUR.1', 'FORM CPTPP'],
      CO_CRITERIA: ['WO', 'PE', 'RVC 40%', 'CTH', 'CTSH'],
      TRADE_AGREEMENT: ['ATIGA', 'EVFTA', 'ACFTA', 'AKFTA', 'VJEPA', 'CPTPP', 'RCEP']
    }
  };

  function createRunner(successCb, failureCb) {
    return {
      withSuccessHandler: function(cb) { return createRunner(cb, failureCb); },
      withFailureHandler: function(cb) { return createRunner(successCb, cb); },
      loginUser: function(maSV, matKhau) {
        setTimeout(function() {
          var key = String(maSV || '').trim();
          var pass = String(matKhau || '').trim();
          var user = MOCK_USERS[key];
          if (!user || user.matKhau !== pass) {
            if (successCb) successCb({ success: false, message: 'Mã SV hoặc Mật khẩu không đúng. (Thử: 00123456 / 123456)' });
            return;
          }
          var saved = '';
          try { saved = localStorage.getItem('gas_preview_draft_' + key) || ''; } catch (e) {}
          if (successCb) successCb({ success: true, user: { maSV: user.maSV, tenSV: user.tenSV }, draftJson: saved });
        }, 250);
      },
      getInitialDropdownData: function() {
        setTimeout(function() {
          if (successCb) successCb(MOCK_DROPDOWNS);
        }, 220);
      },
      saveUserDraft: function(maSV, draftJsonString) {
        setTimeout(function() {
          try {
            localStorage.setItem('gas_preview_draft_' + String(maSV), draftJsonString);
            if (successCb) successCb({ success: true, savedAt: new Date().toISOString() });
          } catch (e) {
            if (failureCb) failureCb(e);
          }
        }, 220);
      },
      generateDocuments: function(maSV, docKeysArray, payload) {
        setTimeout(function() {
          var keys = Array.isArray(docKeysArray) ? docKeysArray : [docKeysArray];
          var shared = Object.assign({}, (payload && payload.shared) || {});
          if (shared.VESSEL && shared.VOYAGE) {
            shared.VES_VOY = shared.VESSEL + ' / ' + shared.VOYAGE;
          }
          var containers = (payload && payload.containers) || [];
          shared.CONTAINER_SEAL_DETAILS = containers
            .map(function(c) { return [c.CONT, c.CTYPE, c.SEAL].filter(Boolean).join(' / '); })
            .filter(Boolean)
            .join('\\n');

          var suffix = (shared.INV_NO || '').trim() || (shared.BK_NO || '').trim() || String(maSV || 'DOC');
          var results = keys.map(function(docKey) {
            var docName = docKey + '_' + suffix;
            return {
              docKey: docKey,
              docName: docName,
              docUrl: '#preview-doc-' + docKey,
              pdfUrl: '#preview-pdf-' + docKey
            };
          });

          window.parent.postMessage({
            type: 'GAS_PREVIEW_GENERATED',
            payload: {
              docKeys: keys,
              results: results,
              shared: shared,
              items: (payload && payload.items) || [],
              containers: containers
            }
          }, '*');

          if (successCb) successCb({ success: true, results: results });
        }, 350);
      }
    };
  }

  window.google = {
    script: {
      run: createRunner(null, null)
    }
  };

  // Helper commands from parent React toolbar (demo login & sample data)
  window.addEventListener('message', function(ev) {
    if (!ev.data || !ev.data.type) return;
    if (ev.data.type === 'GAS_PREVIEW_AUTO_LOGIN') {
      var maEl = document.getElementById('loginMaSV');
      var pwEl = document.getElementById('loginMatKhau');
      var form = document.getElementById('loginForm');
      if (maEl && pwEl && form) {
        maEl.value = '00123456';
        pwEl.value = '123456';
        form.dispatchEvent(new Event('submit', { cancelable: true, bubbles: true }));
      }
    } else if (ev.data.type === 'GAS_PREVIEW_FILL_SAMPLE') {
      if (typeof APP_STATE === 'undefined' || !APP_STATE.user) return;
      APP_STATE.shared = Object.assign({}, APP_STATE.shared, {
        SHIP_NAME: 'VIETNAM MEKONG AGRI EXPORT JSC',
        SELLER_TITLE: 'General Director',
        SHIP_REP: 'Mr. Nguyen Minh Tuan',
        SHIP_TAX: '0314567890',
        SHIP_TEL: '02838123456',
        SHIP_FAX: '02838123457',
        SHIP_EMAIL: 'export@mekongagri.vn',
        SHIP_ADDR: '125 Nguyen Hue Street, District 1, Ho Chi Minh City, Vietnam',
        BANK_NAME: 'VIETCOMBANK - HO CHI MINH BRANCH',
        BANK_ACC: '0071000998877',
        BANK_SWIFT: 'BFTVVNVX007',
        BANK_FAX: '02838297228',
        BANK_ADDR: '5 Me Linh Square, District 1, Ho Chi Minh City, Vietnam',
        CNE_NAME: 'PACIFIC RIM FOODS IMPORTS LLC',
        BUYER_TITLE: 'Purchasing Vice President',
        CNE_REP: 'Ms. Sarah Jenkins',
        CNE_TAX: '95-4821904',
        CNE_TEL: '+1 310 555 0192',
        CNE_FAX: '+1 310 555 0199',
        CNE_EMAIL: 'imports@pacificrimfoods.com',
        CNE_ADDR: '2450 Harbor Blvd, Suite 400, Los Angeles, CA 90012, USA',
        NTF_NAME: 'PACIFIC RIM LOGISTICS BROKERAGE INC.',
        NTF_ADDR: '2450 Harbor Blvd, Suite 405, Los Angeles, CA 90012, USA',
        POL: 'CAT LAI, HCMC (VNSGN)',
        POD: 'LOS ANGELES, US (USLAX)',
        DEST_CTRY: 'UNITED STATES',
        CAR_NAME: 'MAERSK LINE',
        CAR_ADDR: '280 Nguyen Van Troi, Phu Nhuan Dist., HCMC',
        CAR_TEL: '02839978888',
        VESSEL: 'MAERSK SENTOSA',
        VOYAGE: '428E',
        VES_VOY: 'MAERSK SENTOSA / 428E',
        SERVICE_TYPE: 'CY-CY',
        INCOTERM: 'FOB',
        PAY_TERM: 'L/C AT SIGHT',
        CURR: 'USD',
        DOCS_NAME: 'Le Thi Hoa',
        DOCS_TITLE: 'Senior Export Docs Specialist',
        DOCS_TEL: '0909123456',
        DOCS_EMAIL: 'docs@mekongagri.vn',
        QT_NO: 'QT-2026-0088',
        SC_NO: 'SC-2026-0088',
        INV_NO: 'INV-2026-0088',
        PL_NO: 'PL-2026-0088',
        REQ_NO: 'REQ-0088',
        ATT: 'Booking Desk - Maersk HCMC',
        BK_NO: 'MAEU99887766',
        BL_NO: 'MAEU99887766BL',
        DOC_NO: 'DOC-BL-0088',
        REL_NO: 'RO-2026-0088',
        DEPOT_NAME: 'ICD PHUOC LONG 3',
        DEPOT_ADDR: 'Truong Tho Ward, Thu Duc City, HCMC',
        LC_NO: 'LC-US-774411',
        MARKS: 'PACIFIC RIM / LOS ANGELES / CTN NO. 1-1000 / MADE IN VIETNAM',
        PACKING: 'Export Standard 5-ply Carton Boxes on Fumigated Pallets',
        CONT_LOADING: '2 x 40HC Containers FCL/FCL',
        FREIGHT_TERM: 'FREIGHT COLLECT',
        COLLECT: 'AS ARRANGED',
        VAT_RATE: '0',
        EXP_TYPE_CODE: 'B11',
        EXP_TYPE_NAME: 'Xuất kinh doanh',
        IMP_TYPE_CODE: 'A11',
        IMP_TYPE_NAME: 'Nhập kinh doanh tiêu dùng'
      });
      APP_STATE.items = [
        {
          PCODE: 'CAF-01',
          PROD: 'ROBUSTA GREEN COFFEE BEANS G1',
          DES: 'Screen 18, Moisture max 12.5%, Crop 2025/2026',
          HS: '09011110',
          PREF: 'WO',
          MARKS: 'N/M - CROP 2026',
          LOT_NO: 'LOT-CF-01',
          QTY: '18000',
          UNIT: 'KGS',
          PRICE: '4.25',
          AMT: '76500',
          PACK: '300',
          NW: '18000',
          GW: '18300',
          CBM: '24.5',
          EX_VAL: '1912500000',
          EX_TAX_RATE: '0',
          EX_AMT: '0',
          FEE_VAL: '1912500000',
          FEE_RATE: '0.5',
          FEE_AMT: '9562500'
        },
        {
          PCODE: 'CAS-W320',
          PROD: 'VIETNAM CASHEW KERNELS W320',
          DES: 'Vacuum Tin Pack 2x11.34kg/carton',
          HS: '08013200',
          PREF: 'RVC 40%',
          MARKS: 'PACIFIC RIM - W320',
          LOT_NO: 'LOT-CS-09',
          QTY: '10000',
          UNIT: 'KGS',
          PRICE: '6.80',
          AMT: '68000',
          PACK: '441',
          NW: '10000',
          GW: '10550',
          CBM: '18.2',
          EX_VAL: '1700000000',
          EX_TAX_RATE: '0',
          EX_AMT: '0',
          FEE_VAL: '1700000000',
          FEE_RATE: '0.5',
          FEE_AMT: '8500000'
        }
      ];
      APP_STATE.containers = [
        { CONT: 'MSKU1234567', SEAL: 'VN889901', CTYPE: '40HC', CSIZE: '40', PACK: '300', UNIT_PACK: 'BAGS', MAX_GW: '30480', GW: '18300', VGM_WEI: '22200', CBM: '24.5' },
        { CONT: 'MSKU7654321', SEAL: 'VN889902', CTYPE: '40HC', CSIZE: '40', PACK: '441', UNIT_PACK: 'CTNS', MAX_GW: '30480', GW: '10550', VGM_WEI: '14450', CBM: '18.2' }
      ];
      recalculateAllTotals(false);
      renderAllFormViews();
      switchTab(APP_STATE.activeTab || 'MASTER_PARTIES');
      showToast('Đã nạp dữ liệu mẫu lô hàng xuất khẩu để review toàn bộ 20 chứng từ!', 'success');
    }
  });

  // Intercept sandbox preview links (#preview-doc- / #preview-pdf-) to open inspector in parent
  document.addEventListener('click', function(e) {
    var link = e.target && e.target.closest ? e.target.closest('a') : null;
    if (!link) return;
    var href = link.getAttribute('href') || '';
    if (href.indexOf('#preview-doc-') === 0 || href.indexOf('#preview-pdf-') === 0) {
      e.preventDefault();
      var docKey = href.replace('#preview-doc-', '').replace('#preview-pdf-', '');
      window.parent.postMessage({ type: 'GAS_PREVIEW_OPEN_DOC', docKey: docKey }, '*');
    }
  });
})();
</script>
`;

export function buildLivePreviewHtml(): string {
  let html = indexHtml;
  html = html.replace("<?!= include('CSS_Style'); ?>", () => cssStyleHtml);
  html = html.replace("<?!= include('View_Login_Sidebar'); ?>", () => viewLoginSidebarHtml);
  html = html.replace("<?!= include('View_Document_Forms'); ?>", () => viewDocumentFormsHtml);
  html = html.replace(
    "<?!= include('JS_App_State'); ?>",
    () => SANDBOX_SHIM_SCRIPT + '\n' + jsAppStateHtml
  );
  html = html.replace("<?!= include('JS_Render_Forms'); ?>", () => jsRenderFormsHtml);
  return html;
}

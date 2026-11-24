/**
 * GS_Config_Templates.gs
 * Cấu hình ID Google Docs tổng hợp ("Chứng từ" gồm 20 Tab), ID Google Sheet,
 * ID thư mục Drive lưu kết quả, bản đồ 20 Tab chứng từ và mảng phân loại thẻ bảng động.
 * Runtime: Google Apps Script (V8)
 */

/**
 * ID duy nhất của file Google Docs tổng hợp ("Chứng từ") chứa 20 Tab mẫu chứng từ.
 * Thay 'YOUR_MASTER_CHUNG_TU_DOC_ID' bằng ID thực tế của file Google Docs tổng hợp.
 */
const MASTER_TEMPLATE_DOC_ID = 'YOUR_MASTER_CHUNG_TU_DOC_ID';

/**
 * ID Google Sheet dữ liệu chính (chứa Danh_sach_SV, Cuoc_Tau, Loai_hinh_XNK, Danh_Muc_chung).
 */
const SPREADSHEET_ID = '1eJZ685b9Y9pw0xA9Z1yfEM44nzZS3aK5uSCTcfbKhZI';

/**
 * ID thư mục Google Drive lưu các file chứng từ được tạo ra.
 */
const OUTPUT_FOLDER_ID = '10qg61iy-FqOkDKSqkQu_fFnvMcoVbetb';

/**
 * Bản đồ 20 loại chứng từ Xuất Nhập Khẩu & Logistics cố định.
 * Mỗi mục chỉ lưu { name, tabName }, trong đó tabName khớp với tên Tab trong file Google Docs "Chứng từ".
 */
const TEMPLATE_MAP = {
  QUOTATION: {
    name: 'Quotation',
    tabName: 'QUOTATION'
  },
  SALES_CONTRACT: {
    name: 'Sales Contract',
    tabName: 'SALES_CONTRACT'
  },
  COMMERCIAL_INVOICE: {
    name: 'Commercial Invoice',
    tabName: 'COMMERCIAL_INVOICE'
  },
  PACKING_LIST: {
    name: 'Packing List',
    tabName: 'PACKING_LIST'
  },
  BOOKING_REQUEST: {
    name: 'Booking Request',
    tabName: 'BOOKING_REQUEST'
  },
  VGM: {
    name: 'VGM',
    tabName: 'VGM'
  },
  SHIPPING_INSTRUCTION: {
    name: 'Shipping Instruction',
    tabName: 'SHIPPING_INSTRUCTION'
  },
  VETERINARY_CERTIFICATE: {
    name: 'Veterinary Certificate',
    tabName: 'VETERINARY_CERTIFICATE'
  },
  PHYTOSANITARY_CERTIFICATE: {
    name: 'Phytosanitary Certificate',
    tabName: 'PHYTOSANITARY_CERTIFICATE'
  },
  QUALITY_CERTIFICATE: {
    name: 'Quality Certificate',
    tabName: 'QUALITY_CERTIFICATE'
  },
  FUMIGATION_CERTIFICATE: {
    name: 'Fumigation Certificate',
    tabName: 'FUMIGATION_CERTIFICATE'
  },
  CERTIFICATE_OF_ORIGIN: {
    name: 'Certificate of Origin',
    tabName: 'CERTIFICATE_OF_ORIGIN'
  },
  BOOKING_NOTE: {
    name: 'Booking Note',
    tabName: 'BOOKING_NOTE'
  },
  BOOKING_CONFIRMATION: {
    name: 'Booking Confirmation',
    tabName: 'BOOKING_CONFIRMATION'
  },
  RELEASE_ORDER: {
    name: 'Release Order',
    tabName: 'RELEASE_ORDER'
  },
  BILL_OF_LADING: {
    name: 'Bill of Lading',
    tabName: 'BILL_OF_LADING'
  },
  DELIVERY_ORDER: {
    name: 'Delivery Order',
    tabName: 'DELIVERY_ORDER'
  },
  To_khai_XK: {
    name: 'Export Customs Declaration',
    tabName: 'To_khai_XK'
  },
  To_khai_NK: {
    name: 'Import Customs Declaration',
    tabName: 'To_khai_NK'
  },
  LETTER_OF_CREDIT: {
    name: 'Letter of Credit',
    tabName: 'LETTER_OF_CREDIT'
  }
};

/**
 * Các tag nhận diện dòng sản phẩm / hàng hóa trong bảng động.
 */
const ITEM_TAGS = [
  'PROD',
  'PCODE',
  'DES',
  'HS',
  'ORG',
  'PREF',
  'MARKS',
  'LOT_NO',
  'QTY',
  'UNIT',
  'PRICE',
  'AMT',
  'CTYPE',
  'NW'
];

/**
 * Các tag nhận diện dòng container trong bảng động.
 */
const CONT_TAGS = [
  'CONT',
  'SEAL',
  'CTYPE',
  'CSIZE',
  'UNIT_PACK',
  'MAX_GW',
  'VGM_WEI'
];

/**
 * Các tag nhận diện dòng thuế / lệ phí hải quan trong bảng động.
 */
const TAX_TAGS = [
  'EX_VAL',
  'EX_TAX_RATE',
  'EX_AMT',
  'FEE_VAL',
  'FEE_RATE',
  'FEE_AMT'
];

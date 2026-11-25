import { useEffect, useMemo, useRef, useState } from 'react';
import {
  Check,
  Code2,
  Copy,
  Download,
  Eye,
  FileCode2,
  FileText,
  Play,
  Printer,
  Sparkles,
  UserCheck,
  X,
} from 'lucide-react';
import { GAS_FILES, GasSourceFile, buildLivePreviewHtml } from './gasPreviewHelper';
import indexHtml from '../Index.html?raw';
import cssStyleHtml from '../CSS_Style.html?raw';
import viewLoginSidebarHtml from '../View_Login_Sidebar.html?raw';
import viewDocumentFormsHtml from '../View_Document_Forms.html?raw';
import jsRenderFormsHtml from '../JS_Render_Forms.html?raw';
import jsAppStateHtml from '../JS_App_State.html?raw';

const MODIFIED_FILE_IDS = [
  'Index.html',
  'CSS_Style.html',
  'View_Login_Sidebar.html',
  'View_Document_Forms.html',
  'JS_Render_Forms.html',
  'JS_App_State.html',
] as const;

const DISPLAY_GAS_FILES: GasSourceFile[] = (() => {
  const byId = new Map<string, GasSourceFile>();
  GAS_FILES.forEach((f) => byId.set(f.id, f));

  byId.set('Index.html', {
    id: 'Index.html',
    filename: 'Index.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Updated Student Guide & Notes modal to 7 numbered items in exact order',
    content: indexHtml,
  });
  byId.set('CSS_Style.html', {
    id: 'CSS_Style.html',
    filename: 'CSS_Style.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Bare red sequential-lock arrow left of label, HS Code error message & exported-locked batch item styling',
    content: cssStyleHtml,
  });
  byId.set('View_Login_Sidebar.html', {
    id: 'View_Login_Sidebar.html',
    filename: 'View_Login_Sidebar.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'XNK workflow chronology step indices & bare red arrow left of each document label',
    content: viewLoginSidebarHtml,
  });
  byId.set('View_Document_Forms.html', {
    id: 'View_Document_Forms.html',
    filename: 'View_Document_Forms.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Renamed batch export helper button to "Chọn 5 chứng từ"',
    content: viewDocumentFormsHtml,
  });
  byId.set('JS_Render_Forms.html', {
    id: 'JS_Render_Forms.html',
    filename: 'JS_Render_Forms.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: '8-digit HS Code error rendering ("Sai HS Code") & exported-locked state in batch export checklist',
    content: jsRenderFormsHtml,
  });
  byId.set('JS_App_State.html', {
    id: 'JS_App_State.html',
    filename: 'JS_App_State.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: '8-digit HS Code validation, XNK workflow unlock order, "Chọn 5 chứng từ" next-5 selection & post-export lock reset on exit+data change',
    content: jsAppStateHtml,
  });

  return MODIFIED_FILE_IDS.map((id) => byId.get(id)).filter(
    (f): f is GasSourceFile => Boolean(f)
  );
})();

interface GeneratedPreviewPayload {
  docKeys: string[];
  results: Array<{ docKey: string; docName: string; docUrl: string; pdfUrl: string }>;
  shared: Record<string, string>;
  items: Array<Record<string, string>>;
  containers: Array<Record<string, string>>;
}

export default function App() {
  const [activeMode, setActiveMode] = useState<'preview' | 'code'>('preview');
  const [selectedFileId, setSelectedFileId] = useState<string>(DISPLAY_GAS_FILES[0].id);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [lastGenerated, setLastGenerated] = useState<GeneratedPreviewPayload | null>(null);
  const [inspectorOpen, setInspectorOpen] = useState(false);
  const [selectedDocKeys, setSelectedDocKeys] = useState<string[]>([]);
  const [inspectorNotice, setInspectorNotice] = useState<string | null>(null);
  const iframeRef = useRef<HTMLIFrameElement | null>(null);

  const previewHtml = useMemo(() => buildLivePreviewHtml(), []);
  const selectedFile: GasSourceFile =
    DISPLAY_GAS_FILES.find((f) => f.id === selectedFileId) || DISPLAY_GAS_FILES[0];

  useEffect(() => {
    const onMessage = (ev: MessageEvent) => {
      if (!ev.data || typeof ev.data !== 'object') return;
      if (ev.data.type === 'GAS_PREVIEW_GENERATED') {
        const payload = ev.data.payload as GeneratedPreviewPayload;
        setLastGenerated(payload);
        setSelectedDocKeys((payload?.results || []).map((r) => r.docKey));
        setInspectorNotice(null);
      } else if (ev.data.type === 'GAS_PREVIEW_OPEN_DOC') {
        setInspectorOpen(true);
      }
    };
    window.addEventListener('message', onMessage);
    return () => window.removeEventListener('message', onMessage);
  }, []);

  const handleCopyFile = async (file: GasSourceFile) => {
    try {
      await navigator.clipboard.writeText(file.content);
      setCopiedId(file.id);
      setTimeout(() => setCopiedId((prev) => (prev === file.id ? null : prev)), 2000);
    } catch {
      const ta = document.createElement('textarea');
      ta.value = file.content;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      setCopiedId(file.id);
      setTimeout(() => setCopiedId((prev) => (prev === file.id ? null : prev)), 2000);
    }
  };

  const handleDownloadFile = (file: GasSourceFile) => {
    const blob = new Blob([file.content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = file.filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const triggerIframeCommand = (type: string) => {
    iframeRef.current?.contentWindow?.postMessage({ type }, '*');
  };

  const escapeHtml = (value: unknown): string =>
    String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');

  const buildBulkDocumentsHtml = (
    docs: GeneratedPreviewPayload['results'],
    payload: GeneratedPreviewPayload
  ): string => {
    const sharedEntries = Object.entries(payload.shared || {}).filter(
      ([, v]) => String(v ?? '').trim() !== ''
    );
    const items = payload.items || [];
    const containers = payload.containers || [];

    const docSections = docs
      .map((doc) => {
        const sharedRows = sharedEntries
          .map(
            ([k, v]) =>
              `<tr><th>${escapeHtml(k)}</th><td>${escapeHtml(v).replace(/\n/g, '<br/>')}</td></tr>`
          )
          .join('');

        const itemsRows = items
          .map(
            (item, idx) =>
              `<tr>
                <td>${idx + 1}</td>
                <td>${escapeHtml(item.PCODE || '')}</td>
                <td>${escapeHtml(item.PROD || '')}</td>
                <td>${escapeHtml(item.HS || '')}</td>
                <td>${escapeHtml(item.QTY || '')} ${escapeHtml(item.UNIT || '')}</td>
                <td>${escapeHtml(item.PRICE || '')}</td>
                <td>${escapeHtml(item.AMT || '')}</td>
              </tr>`
          )
          .join('');

        const containerRows = containers
          .map(
            (c, idx) =>
              `<tr>
                <td>${idx + 1}</td>
                <td>${escapeHtml(c.CONT || '')}</td>
                <td>${escapeHtml(c.SEAL || '')}</td>
                <td>${escapeHtml(c.CTYPE || '')}</td>
                <td>${escapeHtml(c.PACK || '')} ${escapeHtml(c.UNIT_PACK || '')}</td>
                <td>${escapeHtml(c.GW || '')}</td>
                <td>${escapeHtml(c.CBM || '')}</td>
              </tr>`
          )
          .join('');

        return `
          <section class="doc-page">
            <header class="doc-header">
              <div>
                <div class="doc-key">${escapeHtml(doc.docKey)}</div>
                <h1>${escapeHtml(doc.docName)}</h1>
              </div>
              <div class="doc-meta">
                <div><strong>Shipper:</strong> ${escapeHtml(payload.shared?.SHIP_NAME || '—')}</div>
                <div><strong>Consignee:</strong> ${escapeHtml(payload.shared?.CNE_NAME || '—')}</div>
              </div>
            </header>
            <h2>Thông tin chung lô hàng</h2>
            <table>
              <tbody>${sharedRows || '<tr><td colspan="2">Không có dữ liệu</td></tr>'}</tbody>
            </table>
            ${
              items.length > 0
                ? `<h2>Chi tiết hàng hóa (${items.length})</h2>
                   <table>
                     <thead><tr><th>#</th><th>Mã</th><th>Tên hàng</th><th>HS Code</th><th>Số lượng</th><th>Đơn giá</th><th>Thành tiền</th></tr></thead>
                     <tbody>${itemsRows}</tbody>
                   </table>`
                : ''
            }
            ${
              containers.length > 0
                ? `<h2>Danh sách Container (${containers.length})</h2>
                   <table>
                     <thead><tr><th>#</th><th>Số Container</th><th>Số Seal</th><th>Loại</th><th>Đóng gói</th><th>GW</th><th>CBM</th></tr></thead>
                     <tbody>${containerRows}</tbody>
                   </table>`
                : ''
            }
          </section>
        `;
      })
      .join('\n');

    return `<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8" />
  <title>Chứng từ XNK (${docs.length} tài liệu)</title>
  <style>
    * { box-sizing: border-box; }
    body { font-family: Arial, sans-serif; color: #0f172a; margin: 0; padding: 24px; font-size: 12px; line-height: 1.45; }
    .doc-page { border: 1px solid #cbd5e1; border-radius: 8px; padding: 20px; margin-bottom: 24px; background: #fff; }
    .doc-header { display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; border-bottom: 2px solid #144535; padding-bottom: 10px; margin-bottom: 14px; }
    .doc-key { font-size: 11px; font-weight: 700; color: #144535; text-transform: uppercase; letter-spacing: 0.04em; }
    h1 { margin: 2px 0 0; font-size: 18px; color: #0b261c; }
    h2 { margin: 16px 0 8px; font-size: 13px; color: #144535; border-left: 3px solid #144535; padding-left: 8px; }
    .doc-meta { text-align: right; font-size: 11px; color: #334155; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
    th, td { border: 1px solid #cbd5e1; padding: 5px 8px; text-align: left; vertical-align: top; }
    th { background: #f1f5f9; font-weight: 700; width: 220px; }
    thead th { width: auto; }
    @media print {
      body { padding: 0; }
      .doc-page { border: none; border-radius: 0; padding: 0; margin: 0; page-break-after: always; }
      .doc-page:last-child { page-break-after: auto; }
    }
  </style>
</head>
<body>${docSections}</body>
</html>`;
  };

  const handleToggleSelectAllDocs = () => {
    if (!lastGenerated) return;
    const allKeys = lastGenerated.results.map((r) => r.docKey);
    if (selectedDocKeys.length === allKeys.length) {
      setSelectedDocKeys([]);
    } else {
      setSelectedDocKeys(allKeys);
    }
  };

  const handleToggleDocSelection = (docKey: string) => {
    setSelectedDocKeys((prev) =>
      prev.includes(docKey) ? prev.filter((k) => k !== docKey) : [...prev, docKey]
    );
  };

  const handleBulkDownload = () => {
    if (!lastGenerated) return;
    const selectedDocs = lastGenerated.results.filter((r) =>
      selectedDocKeys.includes(r.docKey)
    );
    if (selectedDocs.length === 0) return;

    const html = buildBulkDocumentsHtml(selectedDocs, lastGenerated);
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download =
      selectedDocs.length === 1
        ? `${selectedDocs[0].docName}.html`
        : `Bo_Chung_Tu_XNK_${selectedDocs.length}_files.html`;
    a.click();
    URL.revokeObjectURL(url);
    setInspectorNotice(`Đã tải xuống ${selectedDocs.length} chứng từ đã chọn.`);
  };

  const handleBulkPrint = () => {
    if (!lastGenerated) return;
    const selectedDocs = lastGenerated.results.filter((r) =>
      selectedDocKeys.includes(r.docKey)
    );
    if (selectedDocs.length === 0) return;

    const html = buildBulkDocumentsHtml(selectedDocs, lastGenerated);
    const printFrame = document.createElement('iframe');
    printFrame.style.position = 'fixed';
    printFrame.style.right = '0';
    printFrame.style.bottom = '0';
    printFrame.style.width = '0';
    printFrame.style.height = '0';
    printFrame.style.border = '0';
    printFrame.setAttribute('aria-hidden', 'true');

    const cleanup = () => {
      setTimeout(() => {
        if (printFrame.parentNode) {
          printFrame.parentNode.removeChild(printFrame);
        }
      }, 1000);
    };

    printFrame.onload = () => {
      try {
        const win = printFrame.contentWindow;
        if (win) {
          win.addEventListener('afterprint', cleanup, { once: true });
          win.focus();
          win.print();
        }
      } finally {
        setTimeout(cleanup, 60000);
      }
    };

    document.body.appendChild(printFrame);
    printFrame.srcdoc = html;
    setInspectorNotice(`Đang mở hộp thoại in cho ${selectedDocs.length} chứng từ đã chọn.`);
  };

  const gsFiles = DISPLAY_GAS_FILES.filter((f) => f.gasType === 'gs');
  const htmlFiles = DISPLAY_GAS_FILES.filter((f) => f.gasType === 'html');

  return (
    <div className="min-h-screen flex flex-col bg-gradient-to-br from-[#6EB592] via-[#96D1B4] to-[#B8E4CE] text-[#0B261C]">
      <header className="sticky top-0 z-30 bg-[#AEE0C6] border-b border-[#7EB89A] px-4 py-2.5 flex flex-wrap items-center justify-between gap-3 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-[#144535] text-white font-bold text-sm flex items-center justify-center">
            GAS
          </div>
          <div>
            <h1 className="text-sm font-bold text-[#0B261C] leading-tight">
              Hệ thống Chứng từ XNK — Review &amp; Copy Source Code
            </h1>
            <p className="text-xs font-medium text-[#0B261C]">
              Hiển thị {DISPLAY_GAS_FILES.length} file đã chỉnh sửa sẵn sàng dán vào Google Apps Script
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          <div className="inline-flex rounded-lg border border-[#7EB89A] bg-[#B8E2CC] p-0.5">
            <button
              type="button"
              onClick={() => setActiveMode('preview')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition cursor-pointer ${
                activeMode === 'preview'
                  ? 'bg-[#144535] text-white shadow-xs'
                  : 'text-[#0B261C] hover:bg-[#A0D4B9]'
              }`}
            >
              <Play className="w-3.5 h-3.5" />
              Chạy thử Live Preview
            </button>
            <button
              type="button"
              onClick={() => setActiveMode('code')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition cursor-pointer ${
                activeMode === 'code'
                  ? 'bg-[#144535] text-white shadow-xs'
                  : 'text-[#0B261C] hover:bg-[#A0D4B9]'
              }`}
            >
              <Code2 className="w-3.5 h-3.5" />
              Xem &amp; Copy Code ({DISPLAY_GAS_FILES.length} files)
            </button>
          </div>
        </div>
      </header>

      <div className={activeMode === 'preview' ? 'flex-1 flex flex-col' : 'hidden'}>
        <div className="bg-[#B8E2CC] border-b border-[#7EB89A] px-4 py-2 flex flex-wrap items-center justify-between gap-2 text-xs">
          <div className="flex items-center gap-2 flex-wrap text-[#0B261C]">
            <span className="font-semibold">Tài khoản test Admin:</span>
            <code className="px-1.5 py-0.5 rounded bg-white border border-[#7EB89A] font-mono font-semibold">
              admin
            </code>
            <span>/</span>
            <code className="px-1.5 py-0.5 rounded bg-white border border-[#7EB89A] font-mono font-semibold">
              AD001
            </code>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              type="button"
              onClick={() => triggerIframeCommand('GAS_PREVIEW_AUTO_LOGIN')}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md bg-[#144535] text-white font-semibold hover:bg-[#0E3327] transition cursor-pointer"
            >
              <UserCheck className="w-3.5 h-3.5" />
              Đăng nhập nhanh (Admin)
            </button>
            <button
              type="button"
              onClick={() => triggerIframeCommand('GAS_PREVIEW_FILL_SAMPLE')}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md bg-white border border-[#7EB89A] text-[#0B261C] font-semibold hover:bg-[#D1FAE5] transition cursor-pointer"
            >
              <Sparkles className="w-3.5 h-3.5" />
              Nạp dữ liệu mẫu lô hàng XK
            </button>
            {lastGenerated && (
              <button
                type="button"
                onClick={() => setInspectorOpen(true)}
                className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md bg-emerald-800 text-white font-semibold hover:bg-emerald-900 transition cursor-pointer"
              >
                <Eye className="w-3.5 h-3.5" />
                Xem Tag Payload đã xuất ({lastGenerated.results.length})
              </button>
            )}
          </div>
        </div>

        <iframe
          ref={iframeRef}
          title="GAS Web App Live Preview"
          srcDoc={previewHtml}
          className="w-full flex-1 border-0 min-h-[calc(100vh-105px)] bg-transparent"
        />
      </div>

      {activeMode === 'code' && (
        <div className="flex-1 flex flex-col lg:flex-row min-h-[calc(100vh-60px)]">
          <aside className="w-full lg:w-80 bg-[#B8E2CC] border-b lg:border-b-0 lg:border-r border-[#7EB89A] p-4 shrink-0 overflow-y-auto">
            {gsFiles.length > 0 && (
              <div className="mb-5">
                <div className="text-[11px] font-bold uppercase tracking-wider text-[#0B261C] mb-2">
                  Backend Google Apps Script ({gsFiles.length} files)
                </div>
                <div className="space-y-1">
                  {gsFiles.map((file) => {
                    const active = file.id === selectedFile.id;
                    return (
                      <button
                        key={file.id}
                        type="button"
                        onClick={() => setSelectedFileId(file.id)}
                        className={`w-full text-left px-3 py-2 rounded-lg border text-xs transition cursor-pointer flex items-center justify-between gap-2 ${
                          active
                            ? 'bg-[#144535] border-[#144535] border-l-4 border-l-[#0B261C] text-white font-semibold'
                            : 'bg-white/80 border-[#7EB89A] text-[#0B261C] font-semibold hover:bg-[#A0D4B9]'
                        }`}
                      >
                        <span className="flex items-center gap-2 truncate">
                          <FileCode2 className="w-4 h-4 shrink-0" />
                          <span className="truncate">{file.filename}</span>
                        </span>
                        <span className={`text-[10px] px-1.5 py-0.5 rounded font-mono ${active ? 'bg-white/20 text-white' : 'bg-white text-[#1E293B]'}`}>
                          .gs
                        </span>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {htmlFiles.length > 0 && (
              <div className="mb-5">
                <div className="text-[11px] font-bold uppercase tracking-wider text-[#0B261C] mb-2">
                  Frontend HTML ({htmlFiles.length} files)
                </div>
                <div className="space-y-1">
                  {htmlFiles.map((file) => {
                    const active = file.id === selectedFile.id;
                    return (
                      <button
                        key={file.id}
                        type="button"
                        onClick={() => setSelectedFileId(file.id)}
                        className={`w-full text-left px-3 py-2 rounded-lg border text-xs transition cursor-pointer flex items-center justify-between gap-2 ${
                          active
                            ? 'bg-[#144535] border-[#144535] border-l-4 border-l-[#0B261C] text-white font-semibold'
                            : 'bg-white/80 border-[#7EB89A] text-[#0B261C] font-semibold hover:bg-[#A0D4B9]'
                        }`}
                      >
                        <span className="flex items-center gap-2 truncate">
                          <FileText className="w-4 h-4 shrink-0" />
                          <span className="truncate">{file.filename}</span>
                        </span>
                        <span className={`text-[10px] px-1.5 py-0.5 rounded font-mono ${active ? 'bg-white/20 text-white' : 'bg-white text-[#1E293B]'}`}>
                          .html
                        </span>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            <div className="mt-6 p-3 rounded-lg bg-white border border-[#7EB89A] shadow-sm text-xs text-[#1E293B] space-y-1.5">
              <div className="font-bold text-[#0B261C]">Hướng dẫn dán vào GAS:</div>
              <p>1. Mở project trên <strong>script.google.com</strong>.</p>
              <p>2. Copy đúng <strong>{DISPLAY_GAS_FILES.length} file</strong> đang hiển thị ở trên vào Google Apps Script.</p>
              <p>3. Bấm <strong>Save</strong> rồi <strong>Deploy &rarr; Manage deployments &rarr; Edit &rarr; New version</strong>.</p>
            </div>
          </aside>

          <section className="flex-1 min-w-0 flex flex-col bg-slate-950 text-slate-100">
            <div className="px-5 py-3 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3">
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-sm font-bold text-white font-mono">
                    {selectedFile.filename}
                  </span>
                  <span className="text-[11px] px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                    {selectedFile.content.split('\n').length} dòng
                  </span>
                </div>
                <p className="text-xs text-slate-400 mt-0.5">{selectedFile.description}</p>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => handleDownloadFile(selectedFile)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold border border-slate-700 transition cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" />
                  Tải file
                </button>
                <button
                  type="button"
                  onClick={() => handleCopyFile(selectedFile)}
                  className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-md text-xs font-semibold transition cursor-pointer ${
                    copiedId === selectedFile.id
                      ? 'bg-emerald-600 text-white'
                      : 'bg-[#144535] hover:bg-[#0E3327] text-white'
                  }`}
                >
                  {copiedId === selectedFile.id ? (
                    <>
                      <Check className="w-3.5 h-3.5" />
                      Đã Copy toàn bộ file!
                    </>
                  ) : (
                    <>
                      <Copy className="w-3.5 h-3.5" />
                      Copy Code
                    </>
                  )}
                </button>
              </div>
            </div>

            <div className="flex-1 overflow-auto p-5">
              <pre className="text-xs leading-relaxed font-mono text-slate-100 whitespace-pre select-all">
                <code>{selectedFile.content}</code>
              </pre>
            </div>
          </section>
        </div>
      )}

      {inspectorOpen && lastGenerated && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl max-h-[88vh] flex flex-col overflow-hidden">
            <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between">
              <div>
                <h2 className="text-sm font-bold text-slate-900">
                  Kiểm tra dữ liệu chứng từ đã xuất
                </h2>
                <p className="text-xs text-slate-500">
                  Khi triển khai trên Google Apps Script thật, dữ liệu này sẽ tự động điền vào Google Docs &amp; xuất PDF.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setInspectorOpen(false)}
                className="p-1.5 rounded-lg border border-slate-200 text-slate-500 hover:bg-slate-100 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-5 overflow-y-auto space-y-4 text-xs">
              <div>
                <div className="flex flex-wrap items-center justify-between gap-2 mb-2.5">
                  <div className="flex items-center gap-3 flex-wrap">
                    <div className="font-bold text-slate-800">
                      Danh sách file sẽ tạo trên Drive ({lastGenerated.results.length}):
                    </div>
                    <label className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md bg-slate-100 hover:bg-slate-200 border border-slate-300 text-slate-800 font-semibold cursor-pointer select-none transition">
                      <input
                        type="checkbox"
                        checked={
                          lastGenerated.results.length > 0 &&
                          selectedDocKeys.length === lastGenerated.results.length
                        }
                        ref={(el) => {
                          if (el) {
                            el.indeterminate =
                              selectedDocKeys.length > 0 &&
                              selectedDocKeys.length < lastGenerated.results.length;
                          }
                        }}
                        onChange={handleToggleSelectAllDocs}
                        className="w-3.5 h-3.5 rounded accent-[#144535] cursor-pointer"
                      />
                      <span>
                        Chọn tất cả ({selectedDocKeys.length}/{lastGenerated.results.length})
                      </span>
                    </label>
                  </div>

                  <div className="flex items-center gap-2 flex-wrap">
                    <button
                      type="button"
                      onClick={handleBulkDownload}
                      disabled={selectedDocKeys.length === 0}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-[#144535] hover:bg-[#0E3327] disabled:bg-slate-300 disabled:text-slate-500 disabled:cursor-not-allowed text-white font-semibold transition cursor-pointer"
                    >
                      <Download className="w-3.5 h-3.5" />
                      Tải hàng loạt ({selectedDocKeys.length})
                    </button>
                    <button
                      type="button"
                      onClick={handleBulkPrint}
                      disabled={selectedDocKeys.length === 0}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-white hover:bg-emerald-50 border border-[#144535] disabled:border-slate-300 disabled:bg-slate-100 disabled:text-slate-400 disabled:cursor-not-allowed text-[#144535] font-semibold transition cursor-pointer"
                    >
                      <Printer className="w-3.5 h-3.5" />
                      In hàng loạt ({selectedDocKeys.length})
                    </button>
                  </div>
                </div>

                {inspectorNotice && (
                  <div className="mb-2.5 px-3 py-2 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-900 font-medium flex items-center justify-between gap-2">
                    <span>{inspectorNotice}</span>
                    <button
                      type="button"
                      onClick={() => setInspectorNotice(null)}
                      className="text-emerald-700 hover:text-emerald-900 font-bold cursor-pointer"
                    >
                      ×
                    </button>
                  </div>
                )}

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {lastGenerated.results.map((r) => {
                    const checked = selectedDocKeys.includes(r.docKey);
                    return (
                      <label
                        key={r.docKey}
                        className={`px-3 py-2 rounded-lg border flex items-center justify-between gap-2 cursor-pointer transition select-none ${
                          checked
                            ? 'bg-emerald-50 border-emerald-400 ring-1 ring-emerald-300'
                            : 'bg-slate-50 border-slate-200 hover:bg-slate-100'
                        }`}
                      >
                        <span className="flex items-center gap-2 min-w-0">
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => handleToggleDocSelection(r.docKey)}
                            className="w-3.5 h-3.5 rounded accent-[#144535] shrink-0 cursor-pointer"
                          />
                          <span className="font-semibold text-emerald-950 truncate">
                            {r.docName}
                          </span>
                        </span>
                        <span
                          className={`text-[11px] font-medium shrink-0 ${
                            checked ? 'text-emerald-700' : 'text-slate-500'
                          }`}
                        >
                          {checked ? 'Đã chọn' : 'Chưa chọn'}
                        </span>
                      </label>
                    );
                  })}
                </div>
              </div>

              <div>
                <div className="font-bold text-slate-800 mb-2">
                  Dữ liệu chung của lô hàng:
                </div>
                <pre className="p-3 rounded-lg bg-slate-950 text-slate-100 font-mono overflow-x-auto max-h-56">
                  {JSON.stringify(lastGenerated.shared, null, 2)}
                </pre>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <div className="font-bold text-slate-800 mb-2">
                    Danh sách dòng hàng hóa ({lastGenerated.items.length} dòng):
                  </div>
                  <pre className="p-3 rounded-lg bg-slate-950 text-slate-100 font-mono overflow-x-auto max-h-48">
                    {JSON.stringify(lastGenerated.items, null, 2)}
                  </pre>
                </div>
                <div>
                  <div className="font-bold text-slate-800 mb-2">
                    Danh sách dòng Container ({lastGenerated.containers.length} dòng):
                  </div>
                  <pre className="p-3 rounded-lg bg-slate-950 text-slate-100 font-mono overflow-x-auto max-h-48">
                    {JSON.stringify(lastGenerated.containers, null, 2)}
                  </pre>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

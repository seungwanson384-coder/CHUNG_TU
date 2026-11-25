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
  Sparkles,
  UserCheck,
  X,
} from 'lucide-react';
import { GAS_FILES, GasSourceFile, buildLivePreviewHtml } from './gasPreviewHelper';
import codeGs from '../Code.gs?raw';
import indexHtml from '../Index.html?raw';
import viewLoginSidebarHtml from '../View_Login_Sidebar.html?raw';
import jsRenderFormsHtml from '../JS_Render_Forms.html?raw';
import jsAppStateHtml from '../JS_App_State.html?raw';
import cssStyleHtml from '../CSS_Style.html?raw';
import gsDocEngineGs from '../GS_Doc_Engine.gs?raw';

const MODIFIED_FILE_IDS = [
  'Code.gs',
  'Index.html',
  'View_Login_Sidebar.html',
  'JS_Render_Forms.html',
  'JS_App_State.html',
  'CSS_Style.html',
  'GS_Doc_Engine.gs',
] as const;

const DISPLAY_GAS_FILES: GasSourceFile[] = (() => {
  const byId = new Map<string, GasSourceFile>();
  GAS_FILES.forEach((f) => byId.set(f.id, f));

  byId.set('Code.gs', {
    id: 'Code.gs',
    filename: 'Code.gs',
    gasType: 'gs',
    category: 'Backend (.gs)',
    description: '20-currency Col F reader, Col K/L Container/Seal prefix lookup & system branding',
    content: codeGs,
  });
  byId.set('Index.html', {
    id: 'Index.html',
    filename: 'Index.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Container/Seal prefix lookup modal with Red-Channel filter & updated Seal guide (6–12 chars)',
    content: indexHtml,
  });
  byId.set('View_Login_Sidebar.html', {
    id: 'View_Login_Sidebar.html',
    filename: 'View_Login_Sidebar.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Progressive 21-step locked sidebar navigation & Admin "Mở khóa tất cả" toggle',
    content: viewLoginSidebarHtml,
  });
  byId.set('JS_Render_Forms.html', {
    id: 'JS_Render_Forms.html',
    filename: 'JS_Render_Forms.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Strict CURR vs UNIT separation, milestone field/table locks, Inspection Code & Cont/Seal lookup button',
    content: jsRenderFormsHtml,
  });
  byId.set('JS_App_State.html', {
    id: 'JS_App_State.html',
    filename: 'JS_App_State.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: '21-step locking engine, destination-smart CURR, 6–12 Seal & milestone rules, EN/VI words (no SAY:)',
    content: jsAppStateHtml,
  });
  byId.set('CSS_Style.html', {
    id: 'CSS_Style.html',
    filename: 'CSS_Style.html',
    gasType: 'html',
    category: 'Frontend (.html)',
    description: 'Locked step/field/table badges, lookup modal styling & fixed Preview modal header/loading box',
    content: cssStyleHtml,
  });
  byId.set('GS_Doc_Engine.gs', {
    id: 'GS_Doc_Engine.gs',
    filename: 'GS_Doc_Engine.gs',
    gasType: 'gs',
    category: 'Backend (.gs)',
    description: 'Per-doc language-accurate TOT_AMT_WORDS (VI for Báo giá, EN+ONLY for SC/CI, no SAY:)',
    content: gsDocEngineGs,
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
  const iframeRef = useRef<HTMLIFrameElement | null>(null);

  const previewHtml = useMemo(() => buildLivePreviewHtml(), []);
  const selectedFile: GasSourceFile =
    DISPLAY_GAS_FILES.find((f) => f.id === selectedFileId) || DISPLAY_GAS_FILES[0];

  useEffect(() => {
    const onMessage = (ev: MessageEvent) => {
      if (!ev.data || typeof ev.data !== 'object') return;
      if (ev.data.type === 'GAS_PREVIEW_GENERATED') {
        setLastGenerated(ev.data.payload);
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
                <div className="font-bold text-slate-800 mb-2">Danh sách file sẽ tạo trên Drive:</div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {lastGenerated.results.map((r) => (
                    <div
                      key={r.docKey}
                      className="px-3 py-2 rounded-lg bg-emerald-50 border border-emerald-200 flex items-center justify-between"
                    >
                      <span className="font-semibold text-emerald-900">{r.docName}</span>
                      <span className="text-[11px] font-medium text-emerald-700">Sẵn sàng xuất</span>
                    </div>
                  ))}
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

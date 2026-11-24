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

interface GeneratedPreviewPayload {
  docKeys: string[];
  results: Array<{ docKey: string; docName: string; docUrl: string; pdfUrl: string }>;
  shared: Record<string, string>;
  items: Array<Record<string, string>>;
  containers: Array<Record<string, string>>;
}

export default function App() {
  const [activeMode, setActiveMode] = useState<'preview' | 'code'>('preview');
  const [selectedFileId, setSelectedFileId] = useState<string>(GAS_FILES[0].id);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [lastGenerated, setLastGenerated] = useState<GeneratedPreviewPayload | null>(null);
  const [inspectorOpen, setInspectorOpen] = useState(false);
  const [inspectedDocKey, setInspectedDocKey] = useState<string | null>(null);
  const iframeRef = useRef<HTMLIFrameElement | null>(null);

  const previewHtml = useMemo(() => buildLivePreviewHtml(), []);
  const selectedFile: GasSourceFile =
    GAS_FILES.find((f) => f.id === selectedFileId) || GAS_FILES[0];

  useEffect(() => {
    const onMessage = (ev: MessageEvent) => {
      if (!ev.data || typeof ev.data !== 'object') return;
      if (ev.data.type === 'GAS_PREVIEW_GENERATED') {
        setLastGenerated(ev.data.payload);
        setInspectedDocKey(ev.data.payload?.docKeys?.[0] || null);
      } else if (ev.data.type === 'GAS_PREVIEW_OPEN_DOC') {
        setInspectedDocKey(ev.data.docKey || null);
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
      // Fallback textarea copy
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

  const gsFiles = GAS_FILES.filter((f) => f.gasType === 'gs');
  const htmlFiles = GAS_FILES.filter((f) => f.gasType === 'html');

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 text-slate-900">
      {/* Top Studio Review Bar */}
      <header className="sticky top-0 z-30 bg-white border-b border-slate-200 px-4 py-2.5 flex flex-wrap items-center justify-between gap-3 shadow-xs">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-indigo-600 text-white font-bold text-sm flex items-center justify-center">
            GAS
          </div>
          <div>
            <h1 className="text-sm font-bold text-slate-900 leading-tight">
              XNK &amp; Logistics Auto-Fill — Review &amp; Copy Source Code
            </h1>
            <p className="text-xs text-slate-500">
              3 file Backend (.gs) + 6 file Frontend (.html) sẵn sàng dán vào Google Apps Script
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          <div className="inline-flex rounded-lg border border-slate-200 bg-slate-100 p-0.5">
            <button
              type="button"
              onClick={() => setActiveMode('preview')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition ${
                activeMode === 'preview'
                  ? 'bg-white text-indigo-600 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Play className="w-3.5 h-3.5" />
              Chạy thử Live Preview
            </button>
            <button
              type="button"
              onClick={() => setActiveMode('code')}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-semibold transition ${
                activeMode === 'code'
                  ? 'bg-white text-indigo-600 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Code2 className="w-3.5 h-3.5" />
              Xem &amp; Copy Code (9 files)
            </button>
          </div>
        </div>
      </header>

      {/* MODE 1: LIVE SANDBOX PREVIEW */}
      <div className={activeMode === 'preview' ? 'flex-1 flex flex-col' : 'hidden'}>
        {/* Quick Test Controls */}
        <div className="bg-indigo-50 border-b border-indigo-100 px-4 py-2 flex flex-wrap items-center justify-between gap-2 text-xs">
          <div className="flex items-center gap-2 flex-wrap text-indigo-950">
            <span className="font-semibold">Tài khoản test:</span>
            <code className="px-1.5 py-0.5 rounded bg-white border border-indigo-200 font-mono">
              00123456
            </code>
            <span>/</span>
            <code className="px-1.5 py-0.5 rounded bg-white border border-indigo-200 font-mono">
              123456
            </code>
            <span className="text-indigo-700 hidden sm:inline">
              (Môi trường chạy thử dùng dữ liệu giả lập Cuoc_Tau, Loai_hinh_XNK, Danh_Muc_chung)
            </span>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <button
              type="button"
              onClick={() => triggerIframeCommand('GAS_PREVIEW_AUTO_LOGIN')}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md bg-indigo-600 text-white font-semibold hover:bg-indigo-700 transition cursor-pointer"
            >
              <UserCheck className="w-3.5 h-3.5" />
              Đăng nhập nhanh (00123456)
            </button>
            <button
              type="button"
              onClick={() => triggerIframeCommand('GAS_PREVIEW_FILL_SAMPLE')}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md bg-white border border-indigo-300 text-indigo-700 font-semibold hover:bg-indigo-100 transition cursor-pointer"
            >
              <Sparkles className="w-3.5 h-3.5" />
              Nạp dữ liệu mẫu lô hàng XK
            </button>
            {lastGenerated && (
              <button
                type="button"
                onClick={() => setInspectorOpen(true)}
                className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-md bg-emerald-600 text-white font-semibold hover:bg-emerald-700 transition cursor-pointer"
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
          className="w-full flex-1 border-0 min-h-[calc(100vh-105px)] bg-white"
        />
      </div>

      {/* MODE 2: COPYABLE CODE VIEWER FOR ALL 9 GAS FILES */}
      {activeMode === 'code' && (
        <div className="flex-1 flex flex-col lg:flex-row min-h-[calc(100vh-60px)]">
          {/* File Explorer Sidebar */}
          <aside className="w-full lg:w-80 bg-white border-b lg:border-b-0 lg:border-r border-slate-200 p-4 shrink-0 overflow-y-auto">
            <div className="mb-5">
              <div className="text-[11px] font-bold uppercase tracking-wider text-slate-400 mb-2">
                Backend Google Apps Script (3 files)
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
                          ? 'bg-indigo-50 border-indigo-200 text-indigo-700 font-semibold'
                          : 'bg-white border-transparent text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      <span className="flex items-center gap-2 truncate">
                        <FileCode2 className="w-4 h-4 shrink-0" />
                        <span className="truncate">{file.filename}</span>
                      </span>
                      <span className="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500 font-mono">
                        .gs
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>

            <div>
              <div className="text-[11px] font-bold uppercase tracking-wider text-slate-400 mb-2">
                Frontend HTML Templates (6 files)
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
                          ? 'bg-indigo-50 border-indigo-200 text-indigo-700 font-semibold'
                          : 'bg-white border-transparent text-slate-700 hover:bg-slate-100'
                      }`}
                    >
                      <span className="flex items-center gap-2 truncate">
                        <FileText className="w-4 h-4 shrink-0" />
                        <span className="truncate">{file.filename}</span>
                      </span>
                      <span className="text-[10px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500 font-mono">
                        .html
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>

            <div className="mt-6 p-3 rounded-lg bg-slate-50 border border-slate-200 text-xs text-slate-600 space-y-1.5">
              <div className="font-bold text-slate-800">Hướng dẫn dán vào GAS:</div>
              <p>1. Mở project trên <strong>script.google.com</strong>.</p>
              <p>2. Tạo đúng tên file (bỏ đuôi <code>.gs</code> / <code>.html</code> khi gõ tên trong GAS).</p>
              <p>3. Điền 20 <code>templateId</code> vào <code>GS_Config_Templates.gs</code>.</p>
              <p>4. Bấm <strong>Deploy &rarr; New deployment &rarr; Web app</strong>.</p>
            </div>
          </aside>

          {/* Code Display Panel */}
          <section className="flex-1 min-w-0 flex flex-col bg-slate-950 text-slate-100">
            <div className="px-5 py-3 bg-slate-900 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3">
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-sm font-bold text-white font-mono">
                    {selectedFile.filename}
                  </span>
                  <span className="text-[11px] px-2 py-0.5 rounded bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
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
                      : 'bg-indigo-600 hover:bg-indigo-500 text-white'
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

      {/* GENERATED TAG PAYLOAD INSPECTOR MODAL */}
      {inspectorOpen && lastGenerated && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl w-full max-w-4xl max-h-[88vh] flex flex-col overflow-hidden">
            <div className="px-5 py-3.5 border-b border-slate-200 flex items-center justify-between">
              <div>
                <h2 className="text-sm font-bold text-slate-900">
                  Kiểm tra dữ liệu Merge Tags gửi sang GS_Doc_Engine.gs
                  {inspectedDocKey ? ` (${inspectedDocKey})` : ''}
                </h2>
                <p className="text-xs text-slate-500">
                  Khi triển khai trên Google Apps Script thật, các tag này sẽ tự động điền vào Google Docs &amp; xuất PDF.
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
                      <span className="font-mono text-[11px] text-emerald-700">{r.docKey}</span>
                    </div>
                  ))}
                </div>
              </div>

              <div>
                <div className="font-bold text-slate-800 mb-2">
                  Tag đơn (payload.shared — bao gồm VES_VOY &amp; CONTAINER_SEAL_DETAILS):
                </div>
                <pre className="p-3 rounded-lg bg-slate-950 text-slate-100 font-mono overflow-x-auto max-h-56">
                  {JSON.stringify(lastGenerated.shared, null, 2)}
                </pre>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <div className="font-bold text-slate-800 mb-2">
                    Dòng hàng hóa nhân bản (items: {lastGenerated.items.length} dòng):
                  </div>
                  <pre className="p-3 rounded-lg bg-slate-950 text-slate-100 font-mono overflow-x-auto max-h-48">
                    {JSON.stringify(lastGenerated.items, null, 2)}
                  </pre>
                </div>
                <div>
                  <div className="font-bold text-slate-800 mb-2">
                    Dòng Container nhân bản (containers: {lastGenerated.containers.length} dòng):
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

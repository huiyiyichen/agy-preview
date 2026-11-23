(() => {
  'use strict';

  const VERSION = String(globalThis.__agyPreviewVersion || '1.0.0');
  const OFFICE_RE = /\.(?:docx?|pptx?)$/i;
  const PPTX_RE = /\.pptx?$/i;
  const HTML_RE = /\.html?$/i;
  const IMAGE_RE = /\.(?:png|jpe?g|webp|gif|bmp|svg)$/i;
  const PREVIEW_LINK_RE = /\.(?:docx?|pptx?|html?)$/i;
  const UNSUPPORTED_TEXT_RE = /Preview is not available for this file type|此文件类型无法预览|不支持预览此文件类型/i;
  const ANALYZED_ROW_RE = /^(?:Analyzed|Viewed|Read|已分析|已查看|查看)\s+(.+?\.(?:png|jpe?g|webp|gif|bmp|svg))$/i;
  const IMAGE_FILENAME_RE = /([^\s\\/:"<>|]+\.(?:png|jpe?g|webp|gif|bmp|svg))/i;
  const HTML_FILENAME_RE = /([^\s\\/:"<>|]+\.html?)$/i;
  const UUID_OR_ID_RE = /^[a-zA-Z0-9][a-zA-Z0-9_-]{7,127}$/;

  let settings = Object.assign({
    enabled: true,
    officePreview: true,
    htmlPreview: true,
    htmlDefaultPreview: true,
    inlineImages: true,
    markdownImages: true,
    chatLinkButtons: true,
    language: 'zh-CN'
  }, globalThis.__agyPreviewInitialSettings || {});

  let disposed = false;
  let reqSeq = 0;
  const pendingRpc = new Map();
  const officeControllers = new WeakMap();
  const htmlControllers = new WeakMap();
  const htmlModeMemory = new Map();
  const inlineStripEntries = new Map(); // anchorEl -> { key, strip, hiddenNodes, dispose }
  const mdGalleryEntries = new Map(); // blockEl -> { key, gallery, dispose }
  const convImageCache = new Map(); // convId -> { fetchedAt, byFileName, images, inflight }
  let activeLightbox = null;
  let activeModal = null;
  let syncTimer = 0;

  /* ── RPC Bridge ──────────────────────────────────────────────────────── */

  globalThis.__agyPreviewReceive = (id, reply) => {
    const entry = pendingRpc.get(id);
    if (!entry) return;
    clearTimeout(entry.timer);
    pendingRpc.delete(id);
    if (reply && typeof reply === 'object' && reply.error) {
      entry.reject(new Error(reply.error));
    } else {
      entry.resolve(reply);
    }
  };

  function rpc(method, payload = {}) {
    if (disposed) return Promise.reject(new Error('插件已卸载'));
    const hostFn = globalThis.__agyPreviewHost;
    if (typeof hostFn !== 'function') {
      return Promise.reject(new Error('预览服务未连接'));
    }
    return new Promise((resolve, reject) => {
      const id = ++reqSeq;
      const timeoutMs = method === 'preview.office' ? 60000
        : method.startsWith('preview.') ? 22000 : 12000;
      const timer = setTimeout(() => {
        pendingRpc.delete(id);
        reject(new Error('预览请求超时'));
      }, timeoutMs);
      pendingRpc.set(id, { resolve, reject, timer });
      try {
        hostFn(JSON.stringify({ id, method, ...payload }));
      } catch (err) {
        clearTimeout(timer);
        pendingRpc.delete(id);
        reject(err);
      }
    });
  }

  /* ── Icons (pure inline SVG, zero external dependency) ───────────────── */

  const ICON_PATHS = {
    eye: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/>',
    code: '<polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/>',
    globe: '<circle cx="12" cy="12" r="10"/><path d="M2 12h20"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>',
    refresh: '<path d="M21 12a9 9 0 1 1-2.64-6.36"/><polyline points="21 3 21 9 15 9"/>',
    external: '<path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/>',
    folder: '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>',
    close: '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>',
    chevLeft: '<polyline points="15 18 9 12 15 6"/>',
    chevRight: '<polyline points="9 18 15 12 9 6"/>',
    chevDown: '<polyline points="6 9 12 15 18 9"/>',
    zoomIn: '<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/>',
    zoomOut: '<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="8" y1="11" x2="14" y2="11"/>',
    notes: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="13" y2="17"/>',
    images: '<rect x="3" y="5" width="14" height="12" rx="2"/><circle cx="7.5" cy="9.5" r="1.2"/><path d="m17 13-3.2-3.2a1 1 0 0 0-1.4 0L7 15"/><path d="M7 17v1a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-2"/>'
  };

  function makeIcon(name, size = 14) {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('viewBox', '0 0 24 24');
    svg.setAttribute('width', String(size));
    svg.setAttribute('height', String(size));
    svg.setAttribute('fill', 'none');
    svg.setAttribute('stroke', 'currentColor');
    svg.setAttribute('stroke-width', '2');
    svg.setAttribute('stroke-linecap', 'round');
    svg.setAttribute('stroke-linejoin', 'round');
    svg.style.flexShrink = '0';
    svg.innerHTML = ICON_PATHS[name] || '';
    return svg;
  }

  function el(tag, attrs = {}, ...children) {
    const node = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === 'className') node.className = v;
      else if (k === 'style' && typeof v === 'string') node.style.cssText = v;
      else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2).toLowerCase(), v);
      else node.setAttribute(k, v === true ? '' : String(v));
    }
    for (const c of children) {
      if (c == null || c === false) continue;
      node.append(typeof c === 'string' ? document.createTextNode(c) : c);
    }
    return node;
  }

  /* ── Scoped Styles ───────────────────────────────────────────────────── */

  function ensureStyles() {
    if (document.querySelector('style[data-agy-preview-style="1"]')) return;
    const style = document.createElement('style');
    style.setAttribute('data-agy-preview-style', '1');
    style.textContent = `
      [data-agy-pv-office-hidden="1"],
      [data-agy-pv-img-hidden="1"] { display: none !important; }

      /* Prevent Antigravity's dimmed placeholder container or dark-mode filters from darkening Office/Image previews */
      [data-agy-pv-office-host="1"] {
        opacity: 1 !important;
        filter: none !important;
      }
      .agy-pv-html-host { position: relative !important; }
      .agy-pv-html-host-active {
        position: relative !important;
        overflow: hidden !important;
      }

      .agy-pv-office-root {
        position: relative;
        display: flex;
        flex-direction: column;
        width: 100%;
        height: 100%;
        min-height: 360px;
        flex: 1 1 auto;
        overflow: hidden;
        opacity: 1 !important;
        filter: none !important;
        background: var(--background, #141417);
        color: var(--foreground, #f4f4f5);
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
        font-size: 13px;
      }
      .agy-pv-toolbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        flex-wrap: wrap;
        padding: 7px 12px;
        border-bottom: 1px solid var(--border, rgba(128,128,128,.22));
        background: var(--card, rgba(24,24,28,.96));
        flex-shrink: 0;
        user-select: none;
      }
      .agy-pv-toolbar-group {
        display: flex;
        align-items: center;
        gap: 6px;
        min-width: 0;
      }
      .agy-pv-badge {
        display: inline-flex;
        align-items: center;
        padding: 1px 7px;
        border-radius: 5px;
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 0.02em;
        background: var(--secondary, rgba(128,128,128,.18));
        color: var(--foreground, #f4f4f5);
        border: 1px solid var(--border, rgba(128,128,128,.24));
      }
      .agy-pv-badge.pptx { color: #fb923c; border-color: rgba(251,146,60,.35); background: rgba(251,146,60,.12); }
      .agy-pv-badge.docx { color: #38bdf8; border-color: rgba(56,189,248,.35); background: rgba(56,189,248,.12); }
      .agy-pv-title {
        font-size: 12px;
        font-weight: 500;
        max-width: 240px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      .agy-pv-muted {
        color: var(--muted-foreground, #a1a1aa);
        font-size: 11.5px;
      }
      .agy-pv-seg {
        display: inline-flex;
        align-items: center;
        padding: 2px;
        border-radius: 7px;
        background: var(--secondary, rgba(128,128,128,.16));
        border: 1px solid var(--border, rgba(128,128,128,.22));
        gap: 2px;
      }
      .agy-pv-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 5px;
        height: 26px;
        padding: 0 9px;
        border-radius: 6px;
        border: 1px solid var(--border, rgba(128,128,128,.24));
        background: var(--secondary, rgba(128,128,128,.14));
        color: var(--foreground, #f4f4f5);
        font: inherit;
        font-size: 12px;
        line-height: 1;
        cursor: pointer;
        white-space: nowrap;
        transition: background .15s ease, border-color .15s ease, opacity .15s ease;
      }
      .agy-pv-btn:hover { background: var(--accent, rgba(128,128,128,.26)); }
      .agy-pv-btn:disabled { opacity: .45; cursor: default; }
      .agy-pv-btn.icon { width: 26px; padding: 0; }
      .agy-pv-seg .agy-pv-btn {
        height: 22px;
        border-color: transparent;
        background: transparent;
        color: var(--muted-foreground, #a1a1aa);
        padding: 0 9px;
      }
      .agy-pv-seg .agy-pv-btn.active {
        background: var(--card, #27272f);
        color: var(--foreground, #fff);
        border-color: var(--border, rgba(128,128,128,.28));
        font-weight: 500;
      }
      .agy-pv-btn.active {
        border-color: rgba(56,189,248,.5);
        background: rgba(56,189,248,.16);
        color: #38bdf8;
      }
      .agy-pv-body {
        flex: 1 1 auto;
        overflow: auto;
        padding: 16px 18px 28px;
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 16px;
        background: var(--background, #141417);
      }
      .agy-pv-body.reader-mode {
        padding: 0;
        overflow: hidden;
      }
      .agy-pv-reader-frame {
        width: 100%;
        height: 100%;
        border: 0;
        display: block;
        background: transparent;
      }
      .agy-pv-page-card {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 8px;
        width: 100%;
        max-width: 100%;
      }
      .agy-pv-page-frame {
        position: relative;
        width: 100%;
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid var(--border, rgba(128,128,128,.24));
        background: #ffffff !important;
        opacity: 1 !important;
        filter: none !important;
        box-shadow: 0 4px 18px rgba(0,0,0,.22);
        cursor: zoom-in;
      }
      .agy-pv-page-frame img,
      .agy-pv-thumb-btn img,
      .agy-pv-lb-stage img {
        display: block;
        opacity: 1 !important;
        filter: none !important;
        mix-blend-mode: normal !important;
      }
      .agy-pv-page-frame img {
        width: 100%;
        height: auto;
      }
      .agy-pv-page-skeleton {
        display: flex;
        align-items: center;
        justify-content: center;
        min-height: 220px;
        color: #71717a;
        font-size: 12px;
        background: linear-gradient(110deg, #f4f4f5 30%, #e4e4e7 50%, #f4f4f5 70%);
        background-size: 200% 100%;
        animation: agyPvShimmer 1.4s infinite linear;
      }
      @keyframes agyPvShimmer {
        0% { background-position: 200% 0; }
        100% { background-position: -200% 0; }
      }
      .agy-pv-page-meta {
        display: flex;
        align-items: center;
        justify-content: space-between;
        width: 100%;
        font-size: 11.5px;
        color: var(--muted-foreground, #a1a1aa);
        padding: 0 2px;
      }
      .agy-pv-notes-box {
        width: 100%;
        padding: 10px 12px;
        border-radius: 8px;
        border: 1px solid var(--border, rgba(128,128,128,.24));
        background: var(--card, rgba(24,24,28,.9));
        color: var(--foreground, #f4f4f5);
        font-size: 12.5px;
        line-height: 1.65;
        white-space: pre-wrap;
        word-break: break-word;
      }
      .agy-pv-notes-label {
        font-size: 11px;
        font-weight: 600;
        color: var(--muted-foreground, #a1a1aa);
        margin-bottom: 4px;
      }
      .agy-pv-state {
        margin: auto;
        max-width: 460px;
        text-align: center;
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 10px;
        padding: 28px 20px;
      }
      .agy-pv-spinner {
        width: 24px;
        height: 24px;
        border-radius: 50%;
        border: 2.5px solid var(--border, rgba(128,128,128,.3));
        border-top-color: #38bdf8;
        animation: agyPvSpin .8s linear infinite;
      }
      @keyframes agyPvSpin { to { transform: rotate(360deg); } }
      .agy-pv-notice {
        width: 100%;
        padding: 6px 12px;
        font-size: 11.5px;
        color: #f59e0b;
        background: rgba(245,158,11,.1);
        border-bottom: 1px solid rgba(245,158,11,.24);
        flex-shrink: 0;
      }

      /* HTML Floating Bar & Overlay */
      .agy-pv-html-bar {
        position: absolute;
        top: 10px;
        right: 16px;
        z-index: 35;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 6px;
        border-radius: 9px;
        background: var(--card, rgba(22,22,26,.94));
        border: 1px solid var(--border, rgba(128,128,128,.28));
        box-shadow: 0 6px 20px rgba(0,0,0,.28);
        backdrop-filter: blur(8px);
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
        font-size: 12px;
        color: var(--foreground, #f4f4f5);
      }
      .agy-pv-html-overlay {
        position: absolute;
        inset: 0;
        z-index: 25;
        background: #ffffff;
        display: flex;
        flex-direction: column;
      }
      .agy-pv-html-overlay.hidden { display: none; }
      .agy-pv-html-overlay iframe {
        width: 100%;
        height: 100%;
        border: 0;
        display: block;
        background: #ffffff;
      }
      .agy-pv-html-err {
        margin: auto;
        padding: 20px;
        color: #ef4444;
        font-size: 13px;
        text-align: center;
      }

      /* Codex-style In-Place Inline Image Strip */
      .agy-pv-img-strip {
        margin: 4px 0 8px;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
      }
      .agy-pv-img-head {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 2px 4px;
        border: 0;
        background: transparent;
        color: var(--muted-foreground, #a1a1aa);
        font: inherit;
        font-size: 12.5px;
        cursor: pointer;
        border-radius: 6px;
        transition: color .15s ease, background .15s ease;
      }
      .agy-pv-img-head:hover {
        color: var(--foreground, #f4f4f5);
        background: var(--secondary, rgba(128,128,128,.12));
      }
      .agy-pv-img-chev {
        transition: transform .16s ease;
      }
      .agy-pv-img-strip.collapsed .agy-pv-img-chev {
        transform: rotate(-90deg);
      }
      .agy-pv-img-strip.collapsed .agy-pv-img-rail {
        display: none;
      }
      .agy-pv-img-rail {
        margin: 8px 0 2px 7px;
        padding: 2px 0 4px 14px;
        border-left: 2px solid var(--border, rgba(128,128,128,.24));
        display: flex;
        flex-wrap: wrap;
        align-items: flex-start;
        gap: 10px;
      }
      .agy-pv-thumb-btn {
        position: relative;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: 0;
        border: 1px solid var(--border, rgba(128,128,128,.28));
        border-radius: 10px;
        overflow: hidden;
        background: var(--card, rgba(28,28,34,.85));
        cursor: zoom-in;
        min-width: 72px;
        min-height: 72px;
        max-width: 340px;
        max-height: 184px;
        transition: transform .14s ease, border-color .14s ease, box-shadow .14s ease;
      }
      .agy-pv-thumb-btn:hover {
        transform: translateY(-1px);
        border-color: rgba(56,189,248,.6);
        box-shadow: 0 6px 18px rgba(0,0,0,.28);
      }
      .agy-pv-thumb-btn img {
        width: auto;
        height: auto;
        max-width: 340px;
        max-height: 182px;
        object-fit: contain;
        background: #ffffff;
      }
      .agy-pv-thumb-cap {
        position: absolute;
        left: 0;
        right: 0;
        bottom: 0;
        padding: 3px 7px;
        font-size: 11px;
        color: #f4f4f5;
        background: linear-gradient(transparent, rgba(0,0,0,.78));
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        opacity: 0;
        transition: opacity .15s ease;
        pointer-events: none;
      }
      .agy-pv-thumb-btn:hover .agy-pv-thumb-cap { opacity: 1; }
      .agy-pv-thumb-ph {
        padding: 18px 14px;
        font-size: 11px;
        color: var(--muted-foreground, #a1a1aa);
      }

      /* Codex-style Inline Markdown Image Gallery (for image links inside agent replies) */
      .agy-pv-md-gallery {
        margin: 10px 0 6px;
        display: flex;
        flex-wrap: wrap;
        align-items: flex-start;
        gap: 10px;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
      }
      .agy-pv-md-card {
        position: relative;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: 0;
        border: 1px solid var(--border, rgba(128,128,128,.28));
        border-radius: 12px;
        overflow: hidden;
        background: var(--card, rgba(28,28,34,.9));
        cursor: zoom-in;
        min-width: 88px;
        min-height: 88px;
        max-width: min(420px, 100%);
        max-height: 240px;
        opacity: 1 !important;
        filter: none !important;
        transition: transform .14s ease, border-color .14s ease, box-shadow .14s ease;
      }
      .agy-pv-md-card:hover {
        transform: translateY(-1px);
        border-color: rgba(56,189,248,.65);
        box-shadow: 0 8px 22px rgba(0,0,0,.3);
      }
      .agy-pv-md-card img {
        display: block;
        width: auto;
        height: auto;
        max-width: min(418px, 100%);
        max-height: 238px;
        object-fit: contain;
        background: #ffffff;
        opacity: 1 !important;
        filter: none !important;
        mix-blend-mode: normal !important;
      }
      .agy-pv-md-card:hover .agy-pv-thumb-cap { opacity: 1; }

      /* Chat Link Quick-Preview Button */
      .agy-pv-chat-btn {
        display: inline-flex;
        align-items: center;
        gap: 3px;
        margin-left: 5px;
        padding: 1px 6px;
        border-radius: 5px;
        border: 1px solid var(--border, rgba(128,128,128,.3));
        background: var(--secondary, rgba(128,128,128,.16));
        color: var(--muted-foreground, #d4d4d8);
        font: inherit;
        font-size: 11px;
        line-height: 1.4;
        vertical-align: middle;
        cursor: pointer;
        transition: background .14s ease, color .14s ease, border-color .14s ease;
      }
      .agy-pv-chat-btn:hover {
        color: #38bdf8;
        border-color: rgba(56,189,248,.5);
        background: rgba(56,189,248,.14);
      }

      /* Fullscreen Lightbox & Preview Modal */
      .agy-pv-backdrop {
        position: fixed;
        inset: 0;
        z-index: 99990;
        background: rgba(0, 0, 0, 0.78);
        backdrop-filter: blur(5px);
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 18px;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
      }
      .agy-pv-modal {
        width: min(88vw, 1280px);
        height: min(88vh, 880px);
        border-radius: 12px;
        overflow: hidden;
        border: 1px solid var(--border, rgba(128,128,128,.3));
        background: var(--background, #141417);
        box-shadow: 0 24px 64px rgba(0,0,0,.55);
        display: flex;
        flex-direction: column;
      }
      .agy-pv-lb-bar {
        width: 100%;
        max-width: 1240px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        padding: 8px 12px;
        border-radius: 10px;
        background: rgba(24,24,28,.92);
        border: 1px solid rgba(255,255,255,.12);
        color: #f4f4f5;
        margin-bottom: 12px;
      }
      .agy-pv-lb-stage {
        position: relative;
        flex: 1 1 auto;
        width: 100%;
        max-width: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
        overflow: auto;
      }
      .agy-pv-lb-stage img {
        max-width: 92vw;
        max-height: 82vh;
        border-radius: 8px;
        box-shadow: 0 12px 40px rgba(0,0,0,.5);
        background: #ffffff;
        cursor: zoom-in;
      }
      .agy-pv-lb-stage img.zoomed {
        max-width: none;
        max-height: none;
        cursor: zoom-out;
      }
      .agy-pv-lb-nav {
        position: fixed;
        top: 50%;
        transform: translateY(-50%);
        width: 40px;
        height: 40px;
        border-radius: 50%;
        border: 1px solid rgba(255,255,255,.18);
        background: rgba(24,24,28,.86);
        color: #fff;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        z-index: 99995;
      }
      .agy-pv-lb-nav:hover { background: rgba(56,189,248,.35); }
      .agy-pv-lb-nav.prev { left: 22px; }
      .agy-pv-lb-nav.next { right: 22px; }
    `;
    document.head.appendChild(style);
  }

  /* ── Helpers: Theme, Conversation ID, Fiber, Path & Breadcrumb Parsing ─ */

  function isDarkTheme() {
    if (document.documentElement.classList.contains('dark')) return true;
    try {
      const bg = getComputedStyle(document.body).backgroundColor || '';
      const m = bg.match(/\d+/g);
      if (m && m.length >= 3) {
        const lum = (0.299 * Number(m[0]) + 0.587 * Number(m[1]) + 0.114 * Number(m[2])) / 255;
        return lum < 0.5;
      }
    } catch {}
    return true;
  }

  function getUrlConversationId() {
    const m = (location.pathname || '').match(/\/c\/([a-zA-Z0-9_-]+)/);
    return m ? m[1] : '';
  }

  function safeDecode(text) {
    try { return decodeURIComponent(String(text || '')); } catch { return String(text || ''); }
  }

  function uriToLocalPath(uri) {
    try {
      const clean = String(uri || '').trim().split('#')[0].split('?')[0];
      if (!/^file:/i.test(clean)) return '';
      const u = new URL(clean);
      let p = decodeURIComponent(u.pathname || '');
      if (/^\/[a-zA-Z]:/.test(p)) p = p.slice(1);
      return p;
    } catch {
      return '';
    }
  }

  function looksLikeLocalPath(str) {
    if (typeof str !== 'string') return false;
    const s = str.trim();
    if (s.length < 3 || s.length > 600) return false;
    return /^file:\/\//i.test(s) || /^[a-zA-Z]:[\\/]/.test(s) || /^\/(?:Users|home|mnt|tmp|var|opt|srv|Volumes|data|workspace|projects)\//i.test(s);
  }

  function getReactFiber(node) {
    if (!node || typeof node !== 'object') return null;
    for (const k of Object.keys(node)) {
      if (k.startsWith('__reactFiber$') || k.startsWith('__reactInternalInstance$')) {
        return node[k];
      }
    }
    return null;
  }

  function getPaneConversationId(node) {
    // First inspect React fiber chain for an explicit cascadeId / conversationId
    let cur = getReactFiber(node);
    for (let i = 0; cur && i < 40; i++, cur = cur.return) {
      const p = cur.memoizedProps;
      if (p && typeof p === 'object') {
        for (const key of ['cascadeId', 'conversationId', 'sessionId', 'threadId']) {
          const val = p[key];
          if (typeof val === 'string' && UUID_OR_ID_RE.test(val)) return val;
        }
        if (p.value && typeof p.value === 'object') {
          for (const key of ['cascadeId', 'conversationId']) {
            const val = p.value[key];
            if (typeof val === 'string' && UUID_OR_ID_RE.test(val)) return val;
          }
        }
      }
    }
    // Never fall back to main URL conversationId if the node sits inside a Side Question pane
    let anc = node instanceof Element ? node : null;
    while (anc && anc !== document.body) {
      const firstHeader = anc.firstElementChild;
      const headTxt = (firstHeader?.textContent || '').trim();
      if (/^(?:Side Question|支线提问)\s*[:：]/i.test(headTxt)) return '';
      anc = anc.parentElement;
    }
    return getUrlConversationId();
  }

  function collectPathsFromObject(obj, extRe, foundFiles, foundRoots, depth = 0, seen = new WeakSet()) {
    if (!obj || depth > 3) return;
    if (typeof obj === 'string') {
      if (looksLikeLocalPath(obj)) {
        const clean = obj.trim().split('#')[0].split('?')[0];
        if (!extRe || extRe.test(clean)) foundFiles.add(clean);
        else if (!/\.[a-zA-Z0-9]{1,8}$/.test(clean)) foundRoots.add(clean);
      }
      return;
    }
    if (typeof obj !== 'object') return;
    if (seen.has(obj)) return;
    seen.add(obj);
    if (Array.isArray(obj)) {
      for (let i = 0; i < Math.min(obj.length, 24); i++) {
        collectPathsFromObject(obj[i], extRe, foundFiles, foundRoots, depth + 1, seen);
      }
      return;
    }
    const keys = Object.keys(obj);
    for (let i = 0; i < Math.min(keys.length, 40); i++) {
      const k = keys[i];
      if (k === 'children' || k === 'ownerDocument' || k === 'window' || k.startsWith('_')) continue;
      let val;
      try { val = obj[k]; } catch { continue; }
      if (typeof val === 'string') {
        if (looksLikeLocalPath(val)) {
          const clean = val.trim().split('#')[0].split('?')[0];
          if (/(?:root|workspace|cwd|project|dir|folder)/i.test(k)) foundRoots.add(clean);
          else if (!extRe || extRe.test(clean)) foundFiles.add(clean);
          else if (!/\.[a-zA-Z0-9]{1,8}$/.test(clean)) foundRoots.add(clean);
        }
      } else if (val && typeof val === 'object' && depth < 2 && /(?:file|path|uri|tab|doc|item|node|workspace|project|entry|source|step|tool|metadata|container)/i.test(k)) {
        collectPathsFromObject(val, extRe, foundFiles, foundRoots, depth + 1, seen);
      }
    }
  }

  function collectWorkspaceRoots() {
    const roots = new Set();
    const dummy = new Set();
    document.querySelectorAll('a[href^="file://"], [data-uri^="file://"]').forEach(elNode => {
      const raw = elNode.getAttribute('data-uri') || elNode.getAttribute('href') || '';
      const p = uriToLocalPath(raw);
      if (!p) return;
      const idx = p.lastIndexOf('/');
      if (idx > 2) roots.add(p.slice(0, idx));
    });
    const sidebar = document.querySelector('aside, nav');
    if (sidebar) {
      let cur = getReactFiber(sidebar);
      for (let i = 0; cur && i < 20; i++, cur = cur.return) {
        collectPathsFromObject(cur.memoizedProps, null, dummy, roots, 0);
      }
    }
    return [...roots];
  }

  // Extracts breadcrumb segments cleanly even when separators are SVG chevrons
  function extractSegmentsFromHeader(headerEl, extRe) {
    if (!headerEl) return [];
    const tokens = [];
    const walker = document.createTreeWalker(headerEl, NodeFilter.SHOW_TEXT);
    let tNode;
    while ((tNode = walker.nextNode())) {
      const p = tNode.parentElement;
      if (!p || p.closest('svg, [aria-hidden="true"], [data-agy-preview-root]')) continue;
      const raw = (tNode.nodeValue || '').trim();
      if (!raw) continue;
      for (const part of raw.split(/\s*[>›/\\]\s*/)) {
        const s = part.trim();
        if (s) tokens.push(s);
      }
    }
    if (!tokens.length) return [];
    if (extRe) {
      let matchIdx = -1;
      for (let i = tokens.length - 1; i >= 0; i--) {
        if (extRe.test(tokens[i])) { matchIdx = i; break; }
      }
      if (matchIdx >= 0) return tokens.slice(0, matchIdx + 1);
      return [];
    }
    return tokens;
  }

  function extractBreadcrumbSegments(pane, extRe) {
    let scope = pane;
    for (let depth = 0; scope && depth < 6; depth++, scope = scope.parentElement) {
      const children = [...scope.children];
      for (const child of children) {
        if (!child.isConnected || child.closest('[data-agy-preview-root]')) continue;
        if (child === pane || child.contains(pane)) continue;
        const text = (child.textContent || '').trim();
        if (!text || text.length > 320) continue;
        const segs = extractSegmentsFromHeader(child, extRe);
        if (segs.length >= 1) return segs;
      }
      const headerNodes = scope.querySelectorAll('[aria-label*="breadcrumb" i], [data-testid*="breadcrumb" i], nav, header');
      for (const node of headerNodes) {
        if (!node.isConnected || node.closest('[data-agy-preview-root]')) continue;
        const segs = extractSegmentsFromHeader(node, extRe);
        if (segs.length >= 1) return segs;
      }
    }
    return [];
  }

  function extractActiveTabName(extRe) {
    const tabs = document.querySelectorAll('[role="tab"][aria-selected="true"], [data-state="active"]');
    for (const tab of tabs) {
      const text = (tab.getAttribute('title') || tab.textContent || '').trim();
      if (text && extRe.test(text)) return text;
    }
    return '';
  }

  function extractFileHintFromPane(pane, extRe, explicitHeader = null) {
    const foundFiles = new Set();
    const foundRoots = new Set();

    const walkFiberFrom = startNode => {
      let cur = getReactFiber(startNode);
      for (let i = 0; cur && i < 36; i++, cur = cur.return) {
        collectPathsFromObject(cur.memoizedProps, extRe, foundFiles, foundRoots, 0);
        if (cur.memoizedState && typeof cur.memoizedState === 'object') {
          let st = cur.memoizedState;
          for (let j = 0; st && j < 10; j++, st = st.next) {
            if (typeof st.memoizedState === 'string') {
              collectPathsFromObject(st.memoizedState, extRe, foundFiles, foundRoots, 0);
            } else if (st.memoizedState && typeof st.memoizedState === 'object' && !st.memoizedState.tag) {
              collectPathsFromObject(st.memoizedState, extRe, foundFiles, foundRoots, 1);
            }
          }
        }
      }
    };

    walkFiberFrom(pane);
    if (explicitHeader) walkFiberFrom(explicitHeader);

    let fileName = '';
    let uri = '';
    const anchors = pane.querySelectorAll('a[download], a[href], [data-uri^="file://"]');
    for (const a of anchors) {
      const dl = safeDecode(a.getAttribute('download') || '').trim();
      if (dl && extRe.test(dl)) fileName = fileName || dl;
      const du = a.getAttribute('data-uri') || a.getAttribute('href') || '';
      if (/^file:/i.test(du) && extRe.test(du.split('#')[0].split('?')[0])) {
        uri = uri || du;
      }
      const txt = safeDecode((a.textContent || '').replace(/^Download\s+/i, '').trim());
      if (txt && extRe.test(txt)) fileName = fileName || txt;
    }

    const segments = explicitHeader
      ? extractSegmentsFromHeader(explicitHeader, extRe)
      : extractBreadcrumbSegments(pane, extRe);
    if (!fileName && segments.length) fileName = segments[segments.length - 1];
    if (!fileName) fileName = extractActiveTabName(extRe);

    const directPath = [...foundFiles].find(p => !/^file:/i.test(p)) || '';
    if (!uri) uri = [...foundFiles].find(p => /^file:/i.test(p)) || '';
    if (!fileName && (directPath || uri)) {
      const raw = safeDecode((directPath || uri).split('#')[0].split('?')[0]);
      fileName = raw.split(/[\\/]/).pop() || '';
    }

    const roots = [...new Set([...foundRoots, ...collectWorkspaceRoots()])];
    return {
      path: directPath,
      uri,
      fileName,
      segments,
      roots,
      conversationId: getUrlConversationId()
    };
  }

  /* ── Lightbox ────────────────────────────────────────────────────────── */

  function closeLightbox() {
    if (!activeLightbox) return;
    activeLightbox.cleanup();
    activeLightbox = null;
  }

  function openLightbox(items, startIndex = 0) {
    closeLightbox();
    if (!items || !items.length) return;
    let index = Math.max(0, Math.min(startIndex, items.length - 1));
    let zoomed = false;

    const titleEl = el('div', { className: 'agy-pv-title', style: 'max-width:50vw' });
    const counterEl = el('span', { className: 'agy-pv-muted' });
    const stage = el('div', { className: 'agy-pv-lb-stage' });
    const imgEl = el('img', { alt: '' });
    const loadingEl = el('div', { className: 'agy-pv-state', style: 'color:#e4e4e7' }, el('div', { className: 'agy-pv-spinner' }), el('span', {}, '正在加载高清图像…'));

    const zoomBtn = el('button', {
      className: 'agy-pv-btn',
      type: 'button',
      onClick: () => {
        zoomed = !zoomed;
        imgEl.classList.toggle('zoomed', zoomed);
        zoomBtn.textContent = zoomed ? '适应窗口' : '100%';
      }
    }, '100%');

    const openBtn = el('button', {
      className: 'agy-pv-btn',
      type: 'button',
      title: '用系统默认程序打开',
      onClick: () => {
        const cur = items[index];
        if (cur?.path) rpc('file.open', { path: cur.path }).catch(() => {});
      }
    }, makeIcon('external', 13), '打开');

    const revealBtn = el('button', {
      className: 'agy-pv-btn',
      type: 'button',
      title: '在文件夹中定位',
      onClick: () => {
        const cur = items[index];
        if (cur?.path) rpc('file.reveal', { path: cur.path }).catch(() => {});
      }
    }, makeIcon('folder', 13), '文件夹');

    const closeBtn = el('button', {
      className: 'agy-pv-btn icon',
      type: 'button',
      title: '关闭 (Esc)',
      onClick: closeLightbox
    }, makeIcon('close', 14));

    const bar = el('div', { className: 'agy-pv-lb-bar', onClick: e => e.stopPropagation() },
      el('div', { className: 'agy-pv-toolbar-group' }, titleEl, counterEl),
      el('div', { className: 'agy-pv-toolbar-group' }, zoomBtn, openBtn, revealBtn, closeBtn)
    );

    const prevBtn = items.length > 1 ? el('button', {
      className: 'agy-pv-lb-nav prev',
      type: 'button',
      onClick: e => { e.stopPropagation(); show(index - 1); }
    }, makeIcon('chevLeft', 20)) : null;

    const nextBtn = items.length > 1 ? el('button', {
      className: 'agy-pv-lb-nav next',
      type: 'button',
      onClick: e => { e.stopPropagation(); show(index + 1); }
    }, makeIcon('chevRight', 20)) : null;

    imgEl.addEventListener('click', e => {
      e.stopPropagation();
      zoomBtn.click();
    });

    const backdrop = el('div', {
      className: 'agy-pv-backdrop',
      'data-agy-preview-root': 'lightbox',
      onClick: closeLightbox
    }, bar, stage, prevBtn, nextBtn);

    let loadToken = 0;
    async function show(nextIdx) {
      index = (nextIdx + items.length) % items.length;
      zoomed = false;
      imgEl.classList.remove('zoomed');
      zoomBtn.textContent = '100%';
      const cur = items[index];
      titleEl.textContent = cur.title || cur.fileName || '图像预览';
      counterEl.textContent = items.length > 1 ? `${index + 1} / ${items.length}` : '';
      openBtn.style.display = cur.path ? '' : 'none';
      revealBtn.style.display = cur.path ? '' : 'none';

      stage.innerHTML = '';
      if (cur.fullDataUrl) {
        imgEl.src = cur.fullDataUrl;
        stage.append(imgEl);
        return;
      }
      stage.append(loadingEl);
      const token = ++loadToken;
      try {
        let dataUrl = '';
        if (typeof cur.loadFull === 'function') {
          dataUrl = await cur.loadFull();
        } else if (cur.path) {
          const res = await rpc('preview.image', { path: cur.path, full: true });
          dataUrl = res.dataUrl;
        }
        if (token !== loadToken) return;
        cur.fullDataUrl = dataUrl;
        imgEl.src = dataUrl;
        stage.innerHTML = '';
        stage.append(imgEl);
      } catch (err) {
        if (token !== loadToken) return;
        stage.innerHTML = '';
        stage.append(el('div', { className: 'agy-pv-state', style: 'color:#f87171' }, err.message || '图片加载失败'));
      }
    }

    const onKey = e => {
      if (e.key === 'Escape') { e.preventDefault(); closeLightbox(); }
      else if (items.length > 1 && e.key === 'ArrowLeft') { e.preventDefault(); show(index - 1); }
      else if (items.length > 1 && e.key === 'ArrowRight') { e.preventDefault(); show(index + 1); }
    };
    window.addEventListener('keydown', onKey, true);
    document.body.append(backdrop);
    void show(index);

    activeLightbox = {
      cleanup() {
        window.removeEventListener('keydown', onKey, true);
        backdrop.remove();
      }
    };
  }

  /* ── Office Viewer Component (Pure Vanilla DOM) ──────────────────────── */

  function buildReaderSrcdoc(innerHtml, dark) {
    const bg = dark ? '#16161a' : '#ffffff';
    const fg = dark ? '#ececef' : '#18181b';
    const muted = dark ? '#a1a1aa' : '#52525b';
    const border = dark ? '#2e2e36' : '#e4e4e7';
    const thBg = dark ? '#222229' : '#f4f4f5';
    return `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8"><style>
      *{box-sizing:border-box}
      body{margin:0;padding:24px 30px 44px;background:${bg};color:${fg};font:14px/1.75 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif;max-width:860px;margin-inline:auto;word-break:break-word}
      h1,h2,h3,h4{margin:1.25em 0 .5em;line-height:1.35}
      h1{font-size:1.55em;padding-bottom:.3em;border-bottom:1px solid ${border}}
      h2{font-size:1.3em}h3{font-size:1.12em}
      p{margin:.55em 0}
      table{border-collapse:collapse;width:100%;margin:1em 0;font-size:13px}
      th,td{border:1px solid ${border};padding:7px 10px;text-align:left;vertical-align:top}
      th{background:${thBg};font-weight:600}
      img{max-width:100%;height:auto;border-radius:6px;margin:8px 0;display:block}
      blockquote{margin:.8em 0;padding:8px 14px;border-left:3px solid #38bdf8;color:${muted}}
      .slide-card{border:1px solid ${border};border-radius:10px;padding:18px 20px;margin:14px 0;background:${thBg}}
      .slide-notes{margin-top:10px;padding:8px 12px;border-radius:6px;border:1px dashed ${border};color:${muted};font-size:12.5px;white-space:pre-wrap}
    </style></head><body>${innerHtml}</body></html>`;
  }

  function createOfficeViewer(fileHint, { onClose } = {}) {
    const isPptx = PPTX_RE.test(fileHint.fileName || fileHint.path || fileHint.uri || '');
    const state = {
      mode: 'layout',
      viewMode: 'scroll',
      showNotes: true,
      slideIndex: 0,
      zoom: 100,
      loading: false,
      error: '',
      notice: '',
      data: null,
      pageCache: new Map(),
      pagePending: new Map(),
      loadSeq: 0,
      io: null
    };

    const root = el('div', {
      className: 'agy-pv-office-root',
      'data-agy-preview-root': 'office',
      tabindex: '0'
    });
    const toolbar = el('div', { className: 'agy-pv-toolbar' });
    const noticeBar = el('div', { className: 'agy-pv-notice', style: 'display:none' });
    const body = el('div', { className: 'agy-pv-body' });
    root.append(toolbar, noticeBar, body);

    function engineLabel(eng) {
      if (!eng) return '';
      if (/powerpoint/i.test(eng)) return 'PowerPoint';
      if (/word/i.test(eng)) return 'Word';
      if (/libreoffice/i.test(eng)) return 'LibreOffice';
      if (/docx/i.test(eng)) return 'python-docx';
      if (/pptx/i.test(eng)) return 'python-pptx';
      return eng;
    }

    async function fetchPageDataUrl(cacheKey, pageIndex) {
      const k = `${cacheKey}:${pageIndex}`;
      if (state.pageCache.has(k)) return state.pageCache.get(k);
      if (state.pagePending.has(k)) return state.pagePending.get(k);
      const p = rpc('preview.officePage', { cacheKey, page: pageIndex })
        .then(res => {
          state.pageCache.set(k, res.dataUrl);
          state.pagePending.delete(k);
          return res.dataUrl;
        })
        .catch(err => {
          state.pagePending.delete(k);
          throw err;
        });
      state.pagePending.set(k, p);
      return p;
    }

    function openPageInLightbox(startIdx) {
      if (!state.data?.pages?.length) return;
      const cacheKey = state.data.cacheKey;
      const items = state.data.pages.map(pg => ({
        title: `${state.data.fileName} - 第 ${pg.index} 页`,
        fileName: state.data.fileName,
        fullDataUrl: state.pageCache.get(`${cacheKey}:${pg.index}`) || '',
        loadFull: () => fetchPageDataUrl(cacheKey, pg.index)
      }));
      openLightbox(items, startIdx);
    }

    function clearObserver() {
      if (state.io) {
        try { state.io.disconnect(); } catch {}
        state.io = null;
      }
    }

    function renderToolbar() {
      toolbar.innerHTML = '';
      const d = state.data;
      const kind = d?.kind || (isPptx ? 'pptx' : 'docx');
      const name = d?.fileName || fileHint.fileName || (kind === 'pptx' ? '演示文稿' : 'Word 文档');

      const left = el('div', { className: 'agy-pv-toolbar-group' },
        el('span', { className: `agy-pv-badge ${kind}` }, kind.toUpperCase()),
        el('span', { className: 'agy-pv-title', title: d?.filePath || name }, name)
      );
      if (d?.pageCount > 0 && d.mode === 'layout') {
        const pageInfo = kind === 'pptx' && state.viewMode === 'slide'
          ? `${state.slideIndex + 1} / ${d.pageCount}`
          : `${d.pageCount} 页`;
        left.append(el('span', { className: 'agy-pv-muted' }, pageInfo));
      }
      if (d?.engine) {
        left.append(el('span', { className: 'agy-pv-muted' }, engineLabel(d.engine)));
      }

      const right = el('div', { className: 'agy-pv-toolbar-group' });

      if (kind === 'docx') {
        const seg = el('div', { className: 'agy-pv-seg' },
          el('button', {
            className: `agy-pv-btn${state.mode === 'layout' ? ' active' : ''}`,
            type: 'button',
            onClick: () => { if (state.mode !== 'layout') { state.mode = 'layout'; void load(false); } }
          }, '原版排版'),
          el('button', {
            className: `agy-pv-btn${state.mode === 'reader' ? ' active' : ''}`,
            type: 'button',
            onClick: () => { if (state.mode !== 'reader') { state.mode = 'reader'; void load(false); } }
          }, '流式阅读')
        );
        right.append(seg);
      }

      if (kind === 'pptx' && d?.mode === 'layout' && d.pageCount > 0) {
        const viewSeg = el('div', { className: 'agy-pv-seg' },
          el('button', {
            className: `agy-pv-btn${state.viewMode === 'scroll' ? ' active' : ''}`,
            type: 'button',
            onClick: () => { state.viewMode = 'scroll'; render(); }
          }, '连续滚动'),
          el('button', {
            className: `agy-pv-btn${state.viewMode === 'slide' ? ' active' : ''}`,
            type: 'button',
            onClick: () => { state.viewMode = 'slide'; render(); }
          }, '单页幻灯')
        );
        right.append(viewSeg);

        const notesCount = (d.pages || []).filter(p => p.notes && p.notes.trim()).length;
        if (notesCount > 0) {
          right.append(el('button', {
            className: `agy-pv-btn${state.showNotes ? ' active' : ''}`,
            type: 'button',
            title: '显示或隐藏演讲者备注',
            onClick: () => { state.showNotes = !state.showNotes; render(); }
          }, makeIcon('notes', 13), `备注 (${notesCount})`));
        }

        if (state.viewMode === 'slide') {
          right.append(
            el('button', {
              className: 'agy-pv-btn icon',
              type: 'button',
              disabled: state.slideIndex <= 0,
              title: '上一页 (左方向键)',
              onClick: () => { if (state.slideIndex > 0) { state.slideIndex--; render(); } }
            }, makeIcon('chevLeft', 14)),
            el('button', {
              className: 'agy-pv-btn icon',
              type: 'button',
              disabled: state.slideIndex >= d.pageCount - 1,
              title: '下一页 (右方向键)',
              onClick: () => { if (state.slideIndex < d.pageCount - 1) { state.slideIndex++; render(); } }
            }, makeIcon('chevRight', 14))
          );
        }
      }

      if (d?.mode === 'layout' && d.pageCount > 0) {
        right.append(
          el('button', {
            className: 'agy-pv-btn icon',
            type: 'button',
            title: '缩小',
            onClick: () => { state.zoom = Math.max(50, state.zoom - 15); updateZoom(); }
          }, makeIcon('zoomOut', 13)),
          el('button', {
            className: 'agy-pv-btn',
            type: 'button',
            title: '重置缩放',
            onClick: () => { state.zoom = 100; updateZoom(); }
          }, `${state.zoom}%`),
          el('button', {
            className: 'agy-pv-btn icon',
            type: 'button',
            title: '放大',
            onClick: () => { state.zoom = Math.min(200, state.zoom + 15); updateZoom(); }
          }, makeIcon('zoomIn', 13))
        );
      }

      right.append(
        el('button', {
          className: 'agy-pv-btn icon',
          type: 'button',
          title: '强制刷新预览',
          disabled: state.loading,
          onClick: () => void load(true)
        }, makeIcon('refresh', 13)),
        el('button', {
          className: 'agy-pv-btn icon',
          type: 'button',
          title: '用系统应用打开',
          onClick: () => {
            const p = state.data?.filePath || fileHint.path || fileHint.uri;
            if (p) rpc('file.open', { path: p }).catch(() => {});
          }
        }, makeIcon('external', 13)),
        el('button', {
          className: 'agy-pv-btn icon',
          type: 'button',
          title: '打开所在文件夹',
          onClick: () => {
            const p = state.data?.filePath || fileHint.path || fileHint.uri;
            if (p) rpc('file.reveal', { path: p }).catch(() => {});
          }
        }, makeIcon('folder', 13))
      );

      if (typeof onClose === 'function') {
        right.append(el('button', {
          className: 'agy-pv-btn icon',
          type: 'button',
          title: '关闭 (Esc)',
          onClick: onClose
        }, makeIcon('close', 14)));
      }

      toolbar.append(left, right);
    }

    function updateZoom() {
      renderToolbar();
      const baseWidth = (state.data?.kind === 'pptx' ? 960 : 820) * (state.zoom / 100);
      body.querySelectorAll('.agy-pv-page-card').forEach(card => {
        card.style.maxWidth = `${Math.round(baseWidth)}px`;
      });
    }

    function buildPageCard(pg, idx, cacheKey) {
      const baseWidth = (state.data?.kind === 'pptx' ? 960 : 820) * (state.zoom / 100);
      const ratio = pg.width && pg.height ? `${pg.width} / ${pg.height}` : (state.data?.kind === 'pptx' ? '16 / 9' : '210 / 297');
      const frame = el('div', {
        className: 'agy-pv-page-frame',
        style: `aspect-ratio:${ratio}`,
        onClick: () => openPageInLightbox(idx)
      });
      const ph = el('div', { className: 'agy-pv-page-skeleton', style: 'width:100%;height:100%' }, `正在加载第 ${pg.index} 页…`);
      frame.append(ph);

      const loadIntoFrame = () => {
        if (frame.dataset.loaded === '1' || frame.dataset.loading === '1') return;
        frame.dataset.loading = '1';
        fetchPageDataUrl(cacheKey, pg.index)
          .then(dataUrl => {
            if (!frame.isConnected) return;
            frame.dataset.loaded = '1';
            frame.style.aspectRatio = '';
            frame.innerHTML = '';
            frame.append(el('img', { src: dataUrl, alt: `第 ${pg.index} 页`, loading: 'lazy' }));
          })
          .catch(err => {
            frame.dataset.loading = '0';
            if (!frame.isConnected) return;
            ph.textContent = `加载失败: ${err.message}`;
          });
      };

      const cached = state.pageCache.get(`${cacheKey}:${pg.index}`);
      if (cached) {
        frame.dataset.loaded = '1';
        frame.style.aspectRatio = '';
        frame.innerHTML = '';
        frame.append(el('img', { src: cached, alt: `第 ${pg.index} 页` }));
      } else {
        frame._loadPage = loadIntoFrame;
      }

      const meta = el('div', { className: 'agy-pv-page-meta' },
        el('span', {}, pg.title ? `第 ${pg.index} 页 - ${pg.title}` : `第 ${pg.index} 页`),
        el('span', {}, '点击放大')
      );

      const card = el('div', {
        className: 'agy-pv-page-card',
        style: `max-width:${Math.round(baseWidth)}px`
      }, frame, meta);

      if (state.showNotes && pg.notes && pg.notes.trim()) {
        card.append(el('div', { className: 'agy-pv-notes-box' },
          el('div', { className: 'agy-pv-notes-label' }, `演讲者备注 (第 ${pg.index} 页)`),
          pg.notes.trim()
        ));
      }
      return card;
    }

    function render() {
      clearObserver();
      renderToolbar();
      if (state.notice) {
        noticeBar.textContent = state.notice;
        noticeBar.style.display = '';
      } else {
        noticeBar.style.display = 'none';
      }

      body.classList.remove('reader-mode');
      body.innerHTML = '';

      if (state.loading) {
        body.append(el('div', { className: 'agy-pv-state' },
          el('div', { className: 'agy-pv-spinner' }),
          el('div', {}, '正在调用本机引擎生成高保真预览…'),
          el('div', { className: 'agy-pv-muted' }, '首次转换需要 2–5 秒，完成后将自动缓存实现秒开')
        ));
        return;
      }

      if (state.error) {
        const actions = el('div', { className: 'agy-pv-toolbar-group', style: 'margin-top:6px' },
          el('button', { className: 'agy-pv-btn', type: 'button', onClick: () => void load(true) }, makeIcon('refresh', 13), '重试'),
          el('button', {
            className: 'agy-pv-btn',
            type: 'button',
            onClick: () => {
              const p = fileHint.path || fileHint.uri;
              if (p) rpc('file.open', { path: p }).catch(() => {});
            }
          }, makeIcon('external', 13), '用系统应用打开')
        );
        body.append(el('div', { className: 'agy-pv-state' },
          el('div', { style: 'color:#f87171;font-weight:500' }, '文档预览生成失败'),
          el('div', { className: 'agy-pv-muted' }, state.error),
          actions
        ));
        return;
      }

      const d = state.data;
      if (!d) return;

      if ((d.mode === 'reader' || d.pageCount === 0) && d.html) {
        body.classList.add('reader-mode');
        const iframe = el('iframe', { className: 'agy-pv-reader-frame', sandbox: '' });
        iframe.srcdoc = buildReaderSrcdoc(d.html, isDarkTheme());
        body.append(iframe);
        return;
      }

      const pages = d.pages || [];
      if (!pages.length) {
        body.append(el('div', { className: 'agy-pv-state' }, el('div', { className: 'agy-pv-muted' }, '未解析出可展示的页面')));
        return;
      }

      if (d.kind === 'pptx' && state.viewMode === 'slide') {
        state.slideIndex = Math.max(0, Math.min(state.slideIndex, pages.length - 1));
        const card = buildPageCard(pages[state.slideIndex], state.slideIndex, d.cacheKey);
        body.append(card);
        const frame = card.querySelector('.agy-pv-page-frame');
        if (frame?._loadPage) frame._loadPage();
        if (pages[state.slideIndex + 1]) void fetchPageDataUrl(d.cacheKey, pages[state.slideIndex + 1].index).catch(() => {});
        if (pages[state.slideIndex - 1]) void fetchPageDataUrl(d.cacheKey, pages[state.slideIndex - 1].index).catch(() => {});
        return;
      }

      state.io = new IntersectionObserver(entries => {
        for (const entry of entries) {
          if (entry.isIntersecting && typeof entry.target._loadPage === 'function') {
            entry.target._loadPage();
            state.io?.unobserve(entry.target);
          }
        }
      }, { root: body, rootMargin: '500px 0px' });

      pages.forEach((pg, idx) => {
        const card = buildPageCard(pg, idx, d.cacheKey);
        body.append(card);
        const frame = card.querySelector('.agy-pv-page-frame');
        if (frame?._loadPage) {
          if (idx < 2) frame._loadPage();
          else state.io.observe(frame);
        }
      });
    }

    async function load(force = false) {
      const seq = ++state.loadSeq;
      state.loading = true;
      state.error = '';
      state.notice = '';
      if (force) state.pageCache.clear();
      render();
      try {
        const res = await rpc('preview.office', {
          ...fileHint,
          mode: state.mode,
          force
        });
        if (seq !== state.loadSeq || disposed) return;
        state.data = res;
        state.mode = res.mode || state.mode;
        state.loading = false;
        render();
      } catch (err) {
        if (seq !== state.loadSeq || disposed) return;
        if (!isPptx && state.mode === 'layout') {
          try {
            const fallback = await rpc('preview.office', { ...fileHint, mode: 'reader', force });
            if (seq !== state.loadSeq || disposed) return;
            state.data = fallback;
            state.mode = 'reader';
            state.notice = '本机未检测到 Word 或 LibreOffice 原版排版引擎，已自动切换为流式阅读模式。';
            state.loading = false;
            render();
            return;
          } catch {}
        }
        state.loading = false;
        state.error = err.message || '文档转换失败';
        render();
      }
    }

    const onKeyDown = e => {
      if (state.data?.kind !== 'pptx' || state.viewMode !== 'slide') return;
      if (e.key === 'ArrowLeft' && state.slideIndex > 0) {
        e.preventDefault();
        state.slideIndex--;
        render();
      } else if (e.key === 'ArrowRight' && state.slideIndex < (state.data.pageCount || 1) - 1) {
        e.preventDefault();
        state.slideIndex++;
        render();
      }
    };
    root.addEventListener('keydown', onKeyDown);

    void load(false);

    return {
      el: root,
      dispose() {
        clearObserver();
        root.removeEventListener('keydown', onKeyDown);
        root.remove();
      }
    };
  }

  /* ── Feature 1: Office File Viewer Takeover (Full-Opacity Host!) ─────── */

  function findOfficePlaceholderContainer(textElem) {
    let block = textElem;
    let chosen = null;
    while (block.parentElement && block.parentElement !== document.body) {
      const parent = block.parentElement;
      if (parent.querySelector('[data-testid="agent-response"], [data-testid="user-input-step"]')) break;
      const rect = parent.getBoundingClientRect();
      const siblings = [...parent.children].filter(c => !c.hasAttribute('data-agy-preview-root'));
      if (rect.height > 180 && rect.width > 180 && siblings.length <= 4) {
        chosen = { host: parent, hiddenTarget: block };
        // Keep climbing past dimmed wrappers (opacity < 0.99) as long as parent doesn't contain tablist/chat
        const op = parseFloat(getComputedStyle(parent).opacity || '1');
        const hasOuterChrome = parent.parentElement?.querySelector('[role="tablist"], [data-testid="agent-response"], [data-testid="user-input-step"]');
        if (op < 0.99 && !hasOuterChrome && parent.parentElement && parent.parentElement !== document.body) {
          block = parent;
          continue;
        }
        break;
      }
      block = parent;
    }
    return chosen;
  }

  function ensureFullOpacityChain(host) {
    const tagged = [];
    let cur = host;
    while (cur && cur !== document.body && cur !== document.documentElement) {
      const cs = getComputedStyle(cur);
      const op = parseFloat(cs.opacity || '1');
      if (op < 0.99 || (cs.filter && cs.filter !== 'none')) {
        cur.setAttribute('data-agy-pv-office-host', '1');
        tagged.push(cur);
      }
      cur = cur.parentElement;
    }
    return tagged;
  }

  function cleanupAllOffice() {
    document.querySelectorAll('[data-agy-preview-root="office"]').forEach(n => {
      if (!n.closest('[data-agy-preview-root="modal"]')) n.remove();
    });
    document.querySelectorAll('[data-agy-pv-office-hidden="1"]').forEach(n => {
      n.removeAttribute('data-agy-pv-office-hidden');
    });
    document.querySelectorAll('[data-agy-pv-office-host="1"]').forEach(n => {
      n.removeAttribute('data-agy-pv-office-host');
    });
  }

  function syncOfficePanes() {
    if (!settings.enabled || !settings.officePreview) {
      cleanupAllOffice();
      return;
    }

    const seenHosts = new Set();
    // Fast targeted check: only inspect leaf text elements outside chat/sidebar
    const candidates = document.querySelectorAll('p, div, span');
    for (let i = 0; i < candidates.length; i++) {
      const elem = candidates[i];
      if (elem.childElementCount > 0) continue;
      if (elem.closest('[data-agy-preview-root], [data-testid="agent-response"], [data-testid="user-input-step"], aside')) continue;
      const txt = elem.textContent;
      if (!txt || txt.length > 140 || !UNSUPPORTED_TEXT_RE.test(txt)) continue;

      const found = findOfficePlaceholderContainer(elem);
      if (!found || !found.host || seenHosts.has(found.host)) continue;

      const hint = extractFileHintFromPane(found.host, OFFICE_RE);
      const targetName = hint.fileName || hint.path || hint.uri || '';
      if (!OFFICE_RE.test(targetName)) continue;

      seenHosts.add(found.host);
      const key = `${hint.path}|${hint.uri}|${hint.fileName}|${hint.segments.join('/')}`;
      const existing = officeControllers.get(found.host);
      if (existing && existing.key === key && existing.viewer.el.isConnected) {
        found.hiddenTarget.setAttribute('data-agy-pv-office-hidden', '1');
        continue;
      }
      if (existing) {
        existing.viewer.dispose();
        existing.hiddenTarget?.removeAttribute('data-agy-pv-office-hidden');
        existing.opacityTagged?.forEach(n => n.removeAttribute('data-agy-pv-office-host'));
      }

      found.hiddenTarget.setAttribute('data-agy-pv-office-hidden', '1');
      const opacityTagged = ensureFullOpacityChain(found.host);
      const viewer = createOfficeViewer(hint);
      found.host.append(viewer.el);
      officeControllers.set(found.host, { key, viewer, hiddenTarget: found.hiddenTarget, opacityTagged });
    }

    document.querySelectorAll('[data-agy-preview-root="office"]').forEach(root => {
      if (root.closest('[data-agy-preview-root="modal"]')) return;
      const host = root.parentElement;
      if (!host || !seenHosts.has(host)) {
        const ctrl = host && officeControllers.get(host);
        if (ctrl) {
          ctrl.viewer.dispose();
          ctrl.hiddenTarget?.removeAttribute('data-agy-pv-office-hidden');
          ctrl.opacityTagged?.forEach(n => n.removeAttribute('data-agy-pv-office-host'));
          officeControllers.delete(host);
        } else {
          root.remove();
        }
      }
    });
  }

  /* ── Feature 2: HTML File Viewer Live Preview Overlay ────────────────── */

  function findHtmlCodeContainers() {
    const results = [];
    const seen = new Set();

    const register = (host, explicitHeader = null) => {
      if (!host || seen.has(host) || host.closest('[data-agy-preview-root]')) return;
      if (host.querySelector('[data-testid="agent-response"], [data-testid="conversation-view"], [data-testid="user-input-step"]')) return;
      if (host.querySelector('[data-agy-pv-office-hidden="1"], [data-agy-preview-root="office"]')) return;
      const rect = host.getBoundingClientRect();
      if (rect.width < 220 || rect.height < 160) return;
      const hint = extractFileHintFromPane(host, HTML_RE, explicitHeader);
      const target = hint.fileName || hint.path || hint.uri || '';
      if (!HTML_RE.test(target)) return;
      seen.add(host);
      results.push({ host, hint });
    };

    // Anchor A: standard code viewer roots
    document.querySelectorAll('.monaco-editor, .cm-editor, [data-testid="file-viewer"], [data-testid="code-viewer"], [data-testid="code-editor"]').forEach(ed => {
      if (ed.closest('[data-agy-preview-root], [data-testid="agent-response"]')) return;
      const host = (ed.classList.contains('monaco-editor') || ed.classList.contains('cm-editor'))
        ? (ed.parentElement || ed)
        : ed;
      register(host);
    });

    // Anchor B: Antigravity File Viewer breadcrumb filename ending in .html/.htm (fast leaf scan, handles SVG chevrons!)
    const leafNodes = document.querySelectorAll('span, a, button, div');
    for (let i = 0; i < leafNodes.length; i++) {
      const node = leafNodes[i];
      if (node.childElementCount > 2) continue;
      if (node.closest('[data-agy-preview-root], [data-testid="agent-response"], [data-testid="user-input-step"], [data-testid="tool-group-collapsible"], aside, [role="tree"], [role="listbox"], [role="menu"]')) continue;
      const text = (node.textContent || '').trim();
      if (!text || text.length > 160 || !HTML_FILENAME_RE.test(text)) continue;
      let headerEl = node;
      let paneContainer = null;
      let bodyEl = null;
      for (let up = 0; headerEl && up < 7; up++, headerEl = headerEl.parentElement) {
        const parent = headerEl.parentElement;
        if (!parent || parent === document.body) break;
        const hRect = headerEl.getBoundingClientRect();
        if (hRect.height < 16 || hRect.height > 120) continue;
        // Look for a large sibling body element after headerEl inside parent
        const siblings = [...parent.children];
        const idx = siblings.indexOf(headerEl);
        for (let s = idx + 1; s < siblings.length; s++) {
          const sib = siblings[s];
          if (sib.hasAttribute('data-agy-preview-root')) continue;
          if (sib.querySelector('[data-testid="agent-response"], [data-testid="user-input-step"]')) continue;
          const sRect = sib.getBoundingClientRect();
          if (sRect.height >= 180 && sRect.width >= 240) {
            paneContainer = parent;
            bodyEl = sib;
            break;
          }
        }
        if (bodyEl) break;
      }
      if (bodyEl && paneContainer) {
        register(bodyEl, headerEl);
      }
    }

    return results;
  }

  function mountHtmlController(host, hint, fileKey) {
    host.classList.add('agy-pv-html-host');

    const remembered = htmlModeMemory.get(fileKey);
    let mode = remembered || (settings.htmlDefaultPreview ? 'preview' : 'source');
    let loadedPath = '';
    let loading = false;

    const previewBtn = el('button', {
      className: 'agy-pv-btn',
      type: 'button',
      onClick: () => setMode('preview', true)
    }, makeIcon('globe', 13), '网页预览');

    const sourceBtn = el('button', {
      className: 'agy-pv-btn',
      type: 'button',
      onClick: () => setMode('source', true)
    }, makeIcon('code', 13), '源码');

    const refreshBtn = el('button', {
      className: 'agy-pv-btn icon',
      type: 'button',
      title: '刷新 HTML 预览',
      onClick: () => void loadHtml(true)
    }, makeIcon('refresh', 13));

    const openBtn = el('button', {
      className: 'agy-pv-btn icon',
      type: 'button',
      title: '在外部浏览器打开',
      onClick: () => {
        const target = loadedPath || hint.path || hint.uri;
        if (target) rpc('file.open', { path: target }).catch(() => {});
      }
    }, makeIcon('external', 13));

    const bar = el('div', { className: 'agy-pv-html-bar', 'data-agy-preview-root': 'html-bar' },
      el('div', { className: 'agy-pv-seg' }, previewBtn, sourceBtn),
      refreshBtn,
      openBtn
    );

    const overlay = el('div', { className: 'agy-pv-html-overlay', 'data-agy-preview-root': 'html-frame' });
    const iframe = el('iframe', {
      sandbox: 'allow-scripts allow-forms allow-popups',
      referrerpolicy: 'no-referrer'
    });
    const errBox = el('div', { className: 'agy-pv-html-err', style: 'display:none' });
    overlay.append(iframe, errBox);

    host.append(overlay, bar);

    async function loadHtml(force = false) {
      if (loading && !force) return;
      loading = true;
      errBox.style.display = 'none';
      iframe.style.display = '';
      try {
        const res = await rpc('preview.html', hint);
        loadedPath = res.filePath || loadedPath;
        iframe.srcdoc = res.html || '';
      } catch (err) {
        iframe.style.display = 'none';
        errBox.style.display = '';
        errBox.textContent = `HTML 预览加载失败：${err.message}`;
      } finally {
        loading = false;
      }
    }

    function setMode(nextMode, remember = false) {
      mode = nextMode;
      if (remember) htmlModeMemory.set(fileKey, mode);
      previewBtn.classList.toggle('active', mode === 'preview');
      sourceBtn.classList.toggle('active', mode === 'source');
      overlay.classList.toggle('hidden', mode !== 'preview');
      host.classList.toggle('agy-pv-html-host-active', mode === 'preview');
      if (mode === 'preview') {
        host.scrollTop = 0;
        if (!iframe.srcdoc) void loadHtml(false);
      }
    }

    setMode(mode, false);

    return {
      fileKey,
      dispose() {
        bar.remove();
        overlay.remove();
        host.classList.remove('agy-pv-html-host', 'agy-pv-html-host-active');
      }
    };
  }

  function cleanupAllHtml() {
    document.querySelectorAll('[data-agy-preview-root="html-bar"], [data-agy-preview-root="html-frame"]').forEach(n => n.remove());
    document.querySelectorAll('.agy-pv-html-host, .agy-pv-html-host-active').forEach(n => {
      n.classList.remove('agy-pv-html-host', 'agy-pv-html-host-active');
    });
  }

  function syncHtmlPanes() {
    if (!settings.enabled || !settings.htmlPreview) {
      cleanupAllHtml();
      return;
    }

    const found = findHtmlCodeContainers();
    const activeHosts = new Set();

    for (const { host, hint } of found) {
      activeHosts.add(host);
      const fileKey = `${hint.path}|${hint.uri}|${hint.fileName}|${hint.segments.join('/')}`;
      const existing = htmlControllers.get(host);
      if (existing && existing.fileKey === fileKey && host.querySelector('[data-agy-preview-root="html-bar"]')) {
        continue;
      }
      if (existing) existing.dispose();
      htmlControllers.set(host, mountHtmlController(host, hint, fileKey));
    }

    document.querySelectorAll('[data-agy-preview-root="html-bar"]').forEach(bar => {
      const host = bar.parentElement;
      if (!host || !activeHosts.has(host)) {
        const ctrl = host && htmlControllers.get(host);
        if (ctrl) {
          ctrl.dispose();
          htmlControllers.delete(host);
        } else {
          bar.remove();
        }
      }
    });
  }

  /* ── Feature 3: Codex-Style In-Place Replacement of Analyzed Images ──── */

  function ensureConversationImages(convId) {
    if (!convId || disposed) return null;
    const entry = convImageCache.get(convId);
    const now = Date.now();
    if (entry && now - entry.fetchedAt < 2500) return entry;
    if (entry?.inflight) return entry;

    const slot = entry || { fetchedAt: 0, byFileName: {}, images: [], inflight: null };
    convImageCache.set(convId, slot);
    slot.inflight = rpc('conversation.images', { conversationId: convId })
      .then(res => {
        slot.fetchedAt = Date.now();
        slot.byFileName = res?.byFileName || {};
        slot.images = Array.isArray(res?.images) ? res.images : [];
        slot.inflight = null;
        scheduleSync();
      })
      .catch(() => {
        slot.fetchedAt = Date.now();
        slot.inflight = null;
      });
    return slot;
  }

  function lookupImagePathByName(convId, fileName, preferredStepIndex = -1) {
    if (!convId || !fileName) return '';
    const slot = ensureConversationImages(convId);
    const list = slot?.byFileName?.[fileName.toLowerCase()];
    if (!Array.isArray(list) || !list.length) return '';
    if (preferredStepIndex >= 0) {
      let best = list[0];
      let bestDist = Math.abs((best.stepIndex ?? 0) - preferredStepIndex);
      for (const item of list) {
        const d = Math.abs((item.stepIndex ?? 0) - preferredStepIndex);
        if (d < bestDist) { bestDist = d; best = item; }
      }
      return best.path;
    }
    return list[list.length - 1].path;
  }

  function extractStepIndexFromFiber(node) {
    let cur = getReactFiber(node);
    for (let i = 0; cur && i < 12; i++, cur = cur.return) {
      const p = cur.memoizedProps;
      if (!p || typeof p !== 'object') continue;
      if (Number.isInteger(p.stepIndex)) return p.stepIndex;
      if (Number.isInteger(p.step?.stepIndex)) return p.step.stepIndex;
      if (Number.isInteger(p.step?.metadata?.stepIndex)) return p.step.metadata.stepIndex;
    }
    return -1;
  }

  // Joins visible leaf text nodes with single spaces (immune to flex-gap spans without whitespace)
  function joinLeafText(root) {
    if (!root) return '';
    const parts = [];
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = w.nextNode())) {
      if (n.parentElement?.closest('svg, [aria-hidden="true"], [data-agy-preview-root]')) continue;
      const s = (n.nodeValue || '').trim();
      if (s) parts.push(s);
    }
    return parts.join(' ').replace(/\s+/g, ' ').trim();
  }

  function extractImageFromRow(rowEl, convId) {
    // 1) Direct data-uri / href on file chip
    const uriEl = rowEl.matches?.('[data-uri^="file://"], a[href^="file://"]')
      ? rowEl
      : rowEl.querySelector?.('[data-uri^="file://"], a[href^="file://"]');
    if (uriEl) {
      const rawUri = uriEl.getAttribute('data-uri') || uriEl.getAttribute('href') || '';
      const local = uriToLocalPath(rawUri);
      if (local && IMAGE_RE.test(local)) {
        return { path: local, fileName: local.split(/[\\/]/).pop() || 'image' };
      }
    }

    // 2) Walk row React fiber (up to 8 levels) for an exact local image path
    const files = new Set();
    const roots = new Set();
    let cur = getReactFiber(rowEl);
    for (let i = 0; cur && i < 8; i++, cur = cur.return) {
      collectPathsFromObject(cur.memoizedProps, IMAGE_RE, files, roots, 0);
      if (files.size > 0) break;
    }
    for (const f of files) {
      const local = /^file:/i.test(f) ? uriToLocalPath(f) : f;
      if (local && IMAGE_RE.test(local)) {
        return { path: local, fileName: local.split(/[\\/]/).pop() || 'image' };
      }
    }

    // 3) Parse filename from "Analyzed <filename.png>" and resolve via THAT pane's transcript.jsonl
    const txt = joinLeafText(rowEl);
    const m = txt.match(ANALYZED_ROW_RE);
    if (!m) return null;
    const fileName = m[1].trim();
    const stepIndex = extractStepIndexFromFiber(rowEl);
    const resolved = lookupImagePathByName(convId, fileName, stepIndex);
    if (resolved) return { path: resolved, fileName };
    return null;
  }

  // Inspects a collapsed tool-group-collapsible's React fiber to see if ALL its steps are image views
  function extractPureImageGroupFromFiber(groupEl, convId) {
    let cur = getReactFiber(groupEl);
    for (let i = 0; cur && i < 10; i++, cur = cur.return) {
      const p = cur.memoizedProps;
      if (!p || typeof p !== 'object') continue;
      const steps = Array.isArray(p.steps) ? p.steps
        : Array.isArray(p.group?.steps) ? p.group.steps
        : Array.isArray(p.items) ? p.items
        : null;
      if (!steps || !steps.length) continue;

      const imgs = [];
      let allImages = true;
      for (const s of steps) {
        const files = new Set();
        const roots = new Set();
        collectPathsFromObject(s, null, files, roots, 0);
        const candidates = [...files].map(f => (/^file:/i.test(f) ? uriToLocalPath(f) : f)).filter(Boolean);
        const imgPath = candidates.find(c => IMAGE_RE.test(c));
        const nonImgPath = candidates.find(c => !IMAGE_RE.test(c) && /\.[a-zA-Z0-9]{1,8}$/.test(c));
        if (imgPath && !nonImgPath) {
          if (!imgs.some(x => x.path.toLowerCase() === imgPath.toLowerCase())) {
            imgs.push({ path: imgPath, fileName: imgPath.split(/[\\/]/).pop() || 'image' });
          }
        } else {
          // Check step case / name if path wasn't directly a string
          const rawJson = (() => { try { return JSON.stringify(s); } catch { return ''; } })();
          const fnMatch = rawJson.match(IMAGE_FILENAME_RE);
          const looksLikeNonImage = /\.(?:py|js|mjs|ts|tsx|jsx|json|md|html?|css|docx?|pptx?|txt|sh|ps1|yaml|yml)\b/i.test(rawJson);
          if (fnMatch && !looksLikeNonImage) {
            const pResolved = lookupImagePathByName(convId, fnMatch[1]);
            if (pResolved) {
              if (!imgs.some(x => x.path.toLowerCase() === pResolved.toLowerCase())) {
                imgs.push({ path: pResolved, fileName: fnMatch[1] });
              }
              continue;
            }
          }
          allImages = false;
          break;
        }
      }
      if (allImages && imgs.length > 0) return imgs;
    }
    return null;
  }

  const TOOL_ROW_VERB_RE = /^(?:Analyzed|Viewed|Read|Listed|Searched|Found|Opened|Inspected|已分析|已查看|已读取|查看|读取|列出|搜索)\s+\S/i;

  function findVisibleToolRowsInGroup(groupEl) {
    // Locate the expanded tool-item rows inside tool-group-collapsible (excluding the toggle button header)
    const toggleBtn = groupEl.querySelector('button[aria-expanded]');
    const candidates = [];

    const nodes = groupEl.querySelectorAll('div, li, button, a');
    for (let i = 0; i < nodes.length; i++) {
      const n = nodes[i];
      if (n === toggleBtn || (toggleBtn && toggleBtn.contains(n))) continue;
      if (n.closest('[data-agy-preview-root]')) continue;
      const txt = joinLeafText(n);
      if (!txt || txt.length > 260) continue;
      if (!TOOL_ROW_VERB_RE.test(txt)) continue;
      candidates.push(n);
    }

    // Keep only innermost (leaf-most) matching rows so list wrappers never shadow individual rows
    const candSet = new Set(candidates);
    const rows = [];
    for (const n of candidates) {
      let hasMatchingDescendant = false;
      for (const other of candSet) {
        if (other !== n && n.contains(other)) {
          hasMatchingDescendant = true;
          break;
        }
      }
      if (!hasMatchingDescendant) rows.push(n);
    }
    return rows;
  }

  function mountInPlaceImageStrip(anchorEl, hiddenNodes, images) {
    const key = images.map(i => i.path.toLowerCase()).join('|');
    const existing = inlineStripEntries.get(anchorEl);
    if (existing && existing.key === key && existing.strip.isConnected) {
      // Ensure hidden nodes stay hidden
      for (const h of hiddenNodes) h.setAttribute('data-agy-pv-img-hidden', '1');
      return;
    }
    if (existing) existing.dispose();

    const chev = makeIcon('chevDown', 13);
    chev.classList.add('agy-pv-img-chev');
    const head = el('button', {
      className: 'agy-pv-img-head',
      type: 'button',
      onClick: () => strip.classList.toggle('collapsed')
    }, makeIcon('images', 14), `已查看 ${images.length} 张图像`, chev);

    const rail = el('div', { className: 'agy-pv-img-rail' });
    const strip = el('div', {
      className: 'agy-pv-img-strip',
      'data-agy-preview-root': 'inline-images'
    }, head, rail);

    const lbItems = images.map(img => ({
      path: img.path,
      fileName: img.fileName,
      title: img.fileName
    }));

    const io = new IntersectionObserver(entries => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const btn = entry.target;
        io.unobserve(btn);
        if (btn.dataset.loaded === '1') continue;
        btn.dataset.loaded = '1';
        const targetPath = btn.dataset.path;
        rpc('preview.image', { path: targetPath, full: false, maxWidth: 520 })
          .then(res => {
            if (!btn.isConnected) return;
            btn.innerHTML = '';
            btn.append(
              el('img', { src: res.dataUrl, alt: res.fileName || '' }),
              el('span', { className: 'agy-pv-thumb-cap' }, res.fileName || '')
            );
          })
          .catch(() => {
            if (!btn.isConnected) return;
            btn.innerHTML = '';
            btn.append(el('span', { className: 'agy-pv-thumb-ph' }, '无法读取图像'));
          });
      }
    }, { rootMargin: '260px 0px' });

    images.forEach((img, idx) => {
      const btn = el('button', {
        className: 'agy-pv-thumb-btn',
        type: 'button',
        title: img.path,
        'data-path': img.path,
        onClick: () => openLightbox(lbItems, idx)
      }, el('span', { className: 'agy-pv-thumb-ph' }, img.fileName || '加载中…'));
      rail.append(btn);
      io.observe(btn);
    });

    // React-safe in-place replacement: hide original React nodes via CSS, insert our strip right after anchorEl
    for (const h of hiddenNodes) h.setAttribute('data-agy-pv-img-hidden', '1');
    anchorEl.insertAdjacentElement('afterend', strip);

    inlineStripEntries.set(anchorEl, {
      key,
      strip,
      hiddenNodes: [...hiddenNodes],
      dispose() {
        try { io.disconnect(); } catch {}
        strip.remove();
        for (const h of hiddenNodes) {
          if (h.isConnected) h.removeAttribute('data-agy-pv-img-hidden');
        }
        inlineStripEntries.delete(anchorEl);
      }
    });
  }

  function cleanupAllInlineImages() {
    for (const entry of [...inlineStripEntries.values()]) {
      entry.dispose();
    }
    inlineStripEntries.clear();
    document.querySelectorAll('[data-agy-preview-root="inline-images"]').forEach(n => n.remove());
    document.querySelectorAll('[data-agy-pv-img-hidden="1"]').forEach(n => n.removeAttribute('data-agy-pv-img-hidden'));
  }

  function syncInlineImages() {
    if (!settings.enabled || !settings.inlineImages) {
      cleanupAllInlineImages();
      return;
    }

    const activeAnchors = new Set();
    const handledRows = new Set();

    // 1) Process Antigravity's Explored file(s) tool groups ([data-testid="tool-group-collapsible"])
    const toolGroups = document.querySelectorAll('[data-testid="tool-group-collapsible"]');
    for (const group of toolGroups) {
      if (!group.isConnected || !group.parentElement || group.closest('[data-agy-preview-root]')) continue;
      const convId = getPaneConversationId(group);
      const rows = findVisibleToolRowsInGroup(group);

      if (rows.length > 0) {
        const classified = rows.map(row => {
          const img = extractImageFromRow(row, convId);
          return { row, img };
        });

        const allImages = classified.every(c => c.img !== null);
        if (allImages) {
          // Pure image Explored files group -> merge all N rows into ONE "已查看 N 张图像" replacing the whole group!
          const merged = [];
          const seen = new Set();
          for (const c of classified) {
            handledRows.add(c.row);
            const low = c.img.path.toLowerCase();
            if (!seen.has(low)) { seen.add(low); merged.push(c.img); }
          }
          if (merged.length > 0) {
            activeAnchors.add(group);
            mountInPlaceImageStrip(group, [group], merged);
          }
        } else {
          // Mixed group (code files + images) -> merge only consecutive Analyzed <image> runs in-place!
          let runRows = [];
          let runImgs = [];
          const flushRun = () => {
            if (!runRows.length) return;
            const deduped = [];
            const seen = new Set();
            for (const im of runImgs) {
              const low = im.path.toLowerCase();
              if (!seen.has(low)) { seen.add(low); deduped.push(im); }
            }
            if (deduped.length > 0) {
              const anchor = runRows[runRows.length - 1];
              activeAnchors.add(anchor);
              mountInPlaceImageStrip(anchor, runRows, deduped);
            }
            runRows = [];
            runImgs = [];
          };
          for (const c of classified) {
            if (c.img) {
              handledRows.add(c.row);
              runRows.push(c.row);
              runImgs.push(c.img);
            } else {
              flushRun();
            }
          }
          flushRun();
        }
      } else {
        // Collapsed group: check React fiber if ALL steps in this collapsed group are image views
        const fiberImgs = extractPureImageGroupFromFiber(group, convId);
        if (fiberImgs && fiberImgs.length > 0) {
          activeAnchors.add(group);
          mountInPlaceImageStrip(group, [group], fiberImgs);
        }
      }
    }

    // 2) Fallback: consecutive standalone "Analyzed <image>" rows outside tool-group-collapsible
    const standaloneRows = [];
    const candidateRows = document.querySelectorAll('[data-testid="agent-response"] div, [data-testid="agent-response"] li');
    for (let i = 0; i < candidateRows.length; i++) {
      const n = candidateRows[i];
      if (handledRows.has(n) || n.closest('[data-testid="tool-group-collapsible"], [data-agy-preview-root], [data-markdown-root]')) continue;
      if (n.childElementCount > 4) continue;
      const txt = joinLeafText(n);
      if (!txt || txt.length > 200 || !ANALYZED_ROW_RE.test(txt)) continue;
      let hasTighter = false;
      for (const ch of n.children) {
        if (ANALYZED_ROW_RE.test(joinLeafText(ch))) { hasTighter = true; break; }
      }
      if (hasTighter) continue;
      const convId = getPaneConversationId(n);
      const img = extractImageFromRow(n, convId);
      if (img) standaloneRows.push({ row: n, img });
    }

    // Group consecutive standalone sibling rows sharing the same parent
    let sRows = [];
    let sImgs = [];
    const flushStandalone = () => {
      if (!sRows.length) return;
      const deduped = [];
      const seen = new Set();
      for (const im of sImgs) {
        const low = im.path.toLowerCase();
        if (!seen.has(low)) { seen.add(low); deduped.push(im); }
      }
      if (deduped.length > 0) {
        // If all items inside an "Explored file(s)" wrapper are these analyzed images, hide the whole wrapper too
        let anchor = sRows[sRows.length - 1];
        let hiddenTargets = sRows;
        const listParent = sRows[0].parentElement;
        const listChildren = listParent
          ? [...listParent.children].filter(c => !c.hasAttribute('data-agy-preview-root'))
          : [];
        if (listParent && listChildren.length === sRows.length) {
          const outer = listParent.parentElement;
          const prevEl = listParent.previousElementSibling;
          const prevTxt = joinLeafText(prevEl);
          if (outer && prevEl && /^(?:Explored files?|已探索文件|已查看文件)$/i.test(prevTxt)) {
            anchor = outer;
            hiddenTargets = [outer];
          }
        }
        activeAnchors.add(anchor);
        mountInPlaceImageStrip(anchor, hiddenTargets, deduped);
      }
      sRows = [];
      sImgs = [];
    };
    for (const item of standaloneRows) {
      if (sRows.length > 0) {
        const prev = sRows[sRows.length - 1];
        const sameParent = prev.parentElement && prev.parentElement === item.row.parentElement;
        let nextSib = prev.nextElementSibling;
        while (nextSib && nextSib.hasAttribute('data-agy-preview-root')) nextSib = nextSib.nextElementSibling;
        if (!sameParent || nextSib !== item.row) flushStandalone();
      }
      sRows.push(item.row);
      sImgs.push(item.img);
    }
    flushStandalone();

    // Clean up strips whose anchor disappeared or no longer qualifies
    for (const [anchor, entry] of [...inlineStripEntries.entries()]) {
      if (!anchor.isConnected || !activeAnchors.has(anchor)) {
        entry.dispose();
      }
    }
  }

  /* ── Feature 4: Chat Link Quick-Preview Button & Modal ───────────────── */

  function closePreviewModal() {
    if (!activeModal) return;
    activeModal.cleanup();
    activeModal = null;
  }

  function openPreviewModal(fileHint) {
    closePreviewModal();
    const target = fileHint.fileName || fileHint.path || fileHint.uri || '';
    const isHtml = HTML_RE.test(target);

    const modal = el('div', {
      className: 'agy-pv-modal',
      onClick: e => e.stopPropagation()
    });

    let innerDispose = () => {};

    if (isHtml) {
      const titleEl = el('span', { className: 'agy-pv-title', style: 'max-width:56vw' }, fileHint.fileName || 'HTML 预览');
      const iframe = el('iframe', {
        style: 'width:100%;flex:1 1 auto;border:0;background:#fff',
        sandbox: 'allow-scripts allow-forms allow-popups',
        referrerpolicy: 'no-referrer'
      });
      const errEl = el('div', { className: 'agy-pv-state', style: 'display:none;color:#f87171' });
      const loadHtml = () => {
        errEl.style.display = 'none';
        iframe.style.display = '';
        rpc('preview.html', fileHint)
          .then(res => {
            titleEl.textContent = res.fileName || titleEl.textContent;
            iframe.srcdoc = res.html || '';
          })
          .catch(err => {
            iframe.style.display = 'none';
            errEl.style.display = '';
            errEl.textContent = `HTML 预览失败: ${err.message}`;
          });
      };
      const bar = el('div', { className: 'agy-pv-toolbar' },
        el('div', { className: 'agy-pv-toolbar-group' },
          el('span', { className: 'agy-pv-badge' }, 'HTML'),
          titleEl
        ),
        el('div', { className: 'agy-pv-toolbar-group' },
          el('button', { className: 'agy-pv-btn icon', type: 'button', title: '刷新', onClick: loadHtml }, makeIcon('refresh', 13)),
          el('button', {
            className: 'agy-pv-btn icon',
            type: 'button',
            title: '在外部浏览器打开',
            onClick: () => rpc('file.open', { path: fileHint.path || fileHint.uri }).catch(() => {})
          }, makeIcon('external', 13)),
          el('button', { className: 'agy-pv-btn icon', type: 'button', title: '关闭 (Esc)', onClick: closePreviewModal }, makeIcon('close', 14))
        )
      );
      modal.append(bar, iframe, errEl);
      loadHtml();
    } else {
      const viewer = createOfficeViewer(fileHint, { onClose: closePreviewModal });
      modal.append(viewer.el);
      innerDispose = viewer.dispose;
    }

    const backdrop = el('div', {
      className: 'agy-pv-backdrop',
      'data-agy-preview-root': 'modal',
      onClick: closePreviewModal
    }, modal);

    const onKey = e => {
      if (e.key === 'Escape' && !activeLightbox) {
        e.preventDefault();
        closePreviewModal();
      }
    };
    window.addEventListener('keydown', onKey, true);
    document.body.append(backdrop);

    activeModal = {
      cleanup() {
        window.removeEventListener('keydown', onKey, true);
        innerDispose();
        backdrop.remove();
      }
    };
  }

  function cleanupChatButtons() {
    document.querySelectorAll('[data-agy-preview-root="chat-btn"]').forEach(n => n.remove());
  }

  function isResponseStreaming(respEl) {
    if (!respEl) return false;
    const txt = (respEl.textContent || '').trim();
    if (!txt || txt === 'Working' || /^Working\b/.test(txt)) return true;
    // Detect active generation via Stop button in the same chat pane (works in both English & localized Chinese)
    const pane = respEl.closest('[data-testid="conversation-view"]') || respEl.parentElement?.parentElement || document.body;
    const hasStopBtn = Boolean(pane.querySelector(
      'button[aria-label*="Stop" i], button[aria-label*="停止"], button[aria-label*="Cancel" i], button[data-testid*="stop" i]'
    ));
    if (hasStopBtn) {
      const all = pane.querySelectorAll('[data-testid="agent-response"]');
      if (all.length && all[all.length - 1] === respEl) return true;
    }
    return false;
  }

  /* ── Feature 3b: Codex-Style Inline Gallery for Reply Image Links ────── */

  function extractImageFromMarkdownChip(chip, convId) {
    const rawUri = chip.getAttribute('data-uri') || chip.getAttribute('href') || '';
    if (rawUri) {
      const local = uriToLocalPath(rawUri);
      if (local && IMAGE_RE.test(local)) {
        return { path: local, fileName: local.split(/[\\/]/).pop() || 'image' };
      }
      return null;
    }
    const label = joinLeafText(chip);
    if (!label || label.length > 180 || !IMAGE_RE.test(label)) return null;
    const fnMatch = label.match(IMAGE_FILENAME_RE);
    if (!fnMatch) return null;
    const files = new Set();
    const roots = new Set();
    let cur = getReactFiber(chip);
    for (let i = 0; cur && i < 6; i++, cur = cur.return) {
      collectPathsFromObject(cur.memoizedProps, IMAGE_RE, files, roots, 0);
      if (files.size > 0) break;
    }
    for (const f of files) {
      const local = /^file:/i.test(f) ? uriToLocalPath(f) : f;
      if (local && IMAGE_RE.test(local)) {
        return { path: local, fileName: local.split(/[\\/]/).pop() || fnMatch[1] };
      }
    }
    const resolved = lookupImagePathByName(convId, fnMatch[1]);
    if (resolved) return { path: resolved, fileName: fnMatch[1] };
    return null;
  }

  function mountMarkdownGallery(blockEl, images) {
    const key = images.map(i => i.path.toLowerCase()).join('|');
    const existing = mdGalleryEntries.get(blockEl);
    if (existing && existing.key === key && existing.gallery.isConnected) return;
    if (existing) existing.dispose();

    const gallery = el('div', {
      className: 'agy-pv-md-gallery',
      'data-agy-preview-root': 'md-images'
    });

    const lbItems = images.map(img => ({
      path: img.path,
      fileName: img.fileName,
      title: img.fileName
    }));

    const io = new IntersectionObserver(entries => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const btn = entry.target;
        io.unobserve(btn);
        if (btn.dataset.loaded === '1') continue;
        btn.dataset.loaded = '1';
        const targetPath = btn.dataset.path;
        rpc('preview.image', { path: targetPath, full: false, maxWidth: 720 })
          .then(res => {
            if (!btn.isConnected) return;
            btn.innerHTML = '';
            btn.append(
              el('img', { src: res.dataUrl, alt: res.fileName || '' }),
              el('span', { className: 'agy-pv-thumb-cap' }, res.fileName || '')
            );
          })
          .catch(() => {
            if (!btn.isConnected) return;
            btn.innerHTML = '';
            btn.append(el('span', { className: 'agy-pv-thumb-ph' }, '无法读取图像'));
          });
      }
    }, { rootMargin: '260px 0px' });

    images.forEach((img, idx) => {
      const card = el('button', {
        className: 'agy-pv-md-card',
        type: 'button',
        title: img.path,
        'data-path': img.path,
        onClick: () => openLightbox(lbItems, idx)
      }, el('span', { className: 'agy-pv-thumb-ph' }, img.fileName || '加载中…'));
      gallery.append(card);
      io.observe(card);
    });

    // Valid HTML placement: inside <li>/<blockquote>/<td>/container, or right after <p>
    if (blockEl.tagName === 'P') {
      blockEl.insertAdjacentElement('afterend', gallery);
    } else {
      blockEl.append(gallery);
    }

    mdGalleryEntries.set(blockEl, {
      key,
      gallery,
      dispose() {
        try { io.disconnect(); } catch {}
        gallery.remove();
        mdGalleryEntries.delete(blockEl);
      }
    });
  }

  function cleanupAllMarkdownImages() {
    for (const entry of [...mdGalleryEntries.values()]) {
      entry.dispose();
    }
    mdGalleryEntries.clear();
    document.querySelectorAll('[data-agy-preview-root="md-images"]').forEach(n => n.remove());
  }

  function syncMarkdownImages() {
    if (!settings.enabled || !settings.markdownImages) {
      cleanupAllMarkdownImages();
      return;
    }

    const activeBlocks = new Set();
    const responses = document.querySelectorAll('[data-testid="agent-response"]');
    for (const resp of responses) {
      if (!resp.isConnected || isResponseStreaming(resp)) continue;
      const convId = getPaneConversationId(resp);
      const mdRoots = [...resp.querySelectorAll('[data-markdown-root]')];
      const scopes = mdRoots.length ? mdRoots : [resp];

      for (const scope of scopes) {
        const byBlock = new Map();
        const chips = scope.querySelectorAll('a[href^="file://"], [data-uri^="file://"], a, button, span[role="button"]');
        for (let i = 0; i < chips.length; i++) {
          const chip = chips[i];
          if (chip.closest('[data-agy-preview-root], [data-testid="tool-group-collapsible"], pre, code')) continue;
          // Skip nested elements if an ancestor chip already carries data-uri/href
          if (!chip.hasAttribute('data-uri') && !chip.hasAttribute('href') && chip.closest('[data-uri^="file://"], a[href^="file://"]')) continue;
          const img = extractImageFromMarkdownChip(chip, convId);
          if (!img) continue;
          const block = chip.closest('li, p, blockquote, td') || scope;
          if (!byBlock.has(block)) byBlock.set(block, []);
          const list = byBlock.get(block);
          if (!list.some(x => x.path.toLowerCase() === img.path.toLowerCase())) {
            list.push(img);
          }
        }

        for (const [block, images] of byBlock.entries()) {
          if (!images.length) continue;
          // Avoid duplicating if the block already renders native <img> elements
          const nativeImgs = [...block.querySelectorAll('img')].filter(im => !im.closest('[data-agy-preview-root]'));
          if (nativeImgs.length >= images.length) continue;
          activeBlocks.add(block);
          mountMarkdownGallery(block, images);
        }
      }
    }

    for (const [block, entry] of [...mdGalleryEntries.entries()]) {
      if (!block.isConnected || !activeBlocks.has(block)) {
        entry.dispose();
      }
    }
  }

  function syncChatLinks() {
    if (!settings.enabled || !settings.chatLinkButtons) {
      cleanupChatButtons();
      return;
    }

    const links = document.querySelectorAll('[data-testid="agent-response"] a[href^="file://"], [data-testid="agent-response"] [data-uri^="file://"]');
    for (const a of links) {
      if (a.closest('[data-agy-preview-root], [data-testid="tool-group-collapsible"]')) continue;
      const resp = a.closest('[data-testid="agent-response"]');
      if (isResponseStreaming(resp)) continue;

      const rawUri = a.getAttribute('data-uri') || a.getAttribute('href') || '';
      const clean = rawUri.split('#')[0].split('?')[0];
      if (!PREVIEW_LINK_RE.test(clean)) continue;
      if (a.nextElementSibling?.getAttribute('data-agy-preview-root') === 'chat-btn') continue;

      const fileName = safeDecode(clean.split('/').pop() || '');
      const btn = el('button', {
        className: 'agy-pv-chat-btn',
        type: 'button',
        'data-agy-preview-root': 'chat-btn',
        title: `在 Antigravity 内预览 ${fileName}`,
        onClick: e => {
          e.preventDefault();
          e.stopPropagation();
          openPreviewModal({
            uri: rawUri,
            fileName,
            conversationId: getPaneConversationId(a)
          });
        }
      }, makeIcon('eye', 11), '预览');
      a.insertAdjacentElement('afterend', btn);
    }
  }

  /* ── Scheduler, Lifecycle & Runtime API ──────────────────────────────── */

  function runSync() {
    if (disposed) return;
    ensureStyles();
    try { syncOfficePanes(); } catch (e) { console.error('[agy-preview] office:', e); }
    try { syncHtmlPanes(); } catch (e) { console.error('[agy-preview] html:', e); }
    try { syncInlineImages(); } catch (e) { console.error('[agy-preview] images:', e); }
    try { syncMarkdownImages(); } catch (e) { console.error('[agy-preview] md-images:', e); }
    try { syncChatLinks(); } catch (e) { console.error('[agy-preview] links:', e); }
  }

  function scheduleSync(delay = 280) {
    if (disposed || syncTimer) return;
    syncTimer = setTimeout(() => {
      syncTimer = 0;
      runSync();
    }, delay);
  }

  const observer = new MutationObserver(mutations => {
    for (const m of mutations) {
      if (m.type !== 'childList') continue;
      if (m.target instanceof Element && m.target.closest('[data-agy-preview-root], [data-agy-powertools-root]')) continue;
      scheduleSync(280);
      return;
    }
  });
  observer.observe(document.body, { childList: true, subtree: true });

  const onRouteChange = () => scheduleSync(150);
  window.addEventListener('popstate', onRouteChange);

  ensureStyles();
  runSync();

  globalThis.__agyPreviewRuntime = {
    version: VERSION,
    applySettings(next) {
      if (disposed || !next || typeof next !== 'object') return;
      settings = Object.assign({}, settings, next);
      if (!settings.enabled) {
        closeLightbox();
        closePreviewModal();
      }
      runSync();
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      clearTimeout(syncTimer);
      observer.disconnect();
      window.removeEventListener('popstate', onRouteChange);
      for (const entry of pendingRpc.values()) {
        clearTimeout(entry.timer);
        entry.reject(new Error('插件已卸载'));
      }
      pendingRpc.clear();
      convImageCache.clear();
      closeLightbox();
      closePreviewModal();
      cleanupAllOffice();
      cleanupAllHtml();
      cleanupAllInlineImages();
      cleanupAllMarkdownImages();
      cleanupChatButtons();
      document.querySelectorAll('style[data-agy-preview-style="1"]').forEach(n => n.remove());
      delete globalThis.__agyPreviewReceive;
      delete globalThis.__agyPreviewRuntime;
    }
  };
})();

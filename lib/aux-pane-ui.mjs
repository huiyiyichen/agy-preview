import { createServer } from 'node:http';
import { readFileSync } from 'node:fs';
import { pluginPaths } from './paths.mjs';
import { openLocalPath } from './file-actions.mjs';
import {
  classifyPreviewExt,
  clearPreviewCache,
  getCacheStats,
  getEnvironmentDiagnostics,
  getOfficePreviewPage,
  renderHtmlPreview,
  renderImagePreview,
  renderOfficePreview
} from './preview-engine.mjs';

const VERSION = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version;

function readJsonBody(req, limit = 128 * 1024) {
  return new Promise((resolve, reject) => {
    let raw = '';
    req.on('data', chunk => {
      raw += chunk.toString('utf8');
      if (raw.length > limit) { req.destroy(); reject(new Error('请求体过大')); }
    });
    req.on('end', () => {
      try { resolve(raw ? JSON.parse(raw) : {}); }
      catch { reject(new Error('无效的 JSON 请求')); }
    });
    req.on('error', reject);
  });
}

function sendJson(res, status, data) {
  const body = JSON.stringify(data);
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
  res.end(body);
}

function renderHtmlShell(token) {
  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>Agy Preview</title>
<style>
  :root{--bg:#131316;--card:#1b1b20;--sec:#26262d;--bd:#2e2e38;--fg:#f4f4f5;--mut:#a1a1aa;--pri:#38bdf8;--ok:#22c55e;--warn:#f59e0b;--err:#ef4444}
  @media(prefers-color-scheme:light){:root{--bg:#fafafa;--card:#ffffff;--sec:#f4f4f5;--bd:#e4e4e7;--fg:#18181b;--mut:#71717a;--pri:#0284c7;--ok:#16a34a;--warn:#d97706;--err:#dc2626}}
  *{box-sizing:border-box}
  body{margin:0;padding:16px;background:var(--bg);color:var(--fg);font:13px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}
  .stack{display:flex;flex-direction:column;gap:14px;max-width:760px;margin:0 auto}
  .head{display:flex;align-items:center;justify-content:space-between;gap:10px}
  .head h1{margin:0;font-size:16px;font-weight:600;display:flex;align-items:center;gap:8px}
  .pill{display:inline-flex;align-items:center;gap:5px;padding:2px 8px;border-radius:999px;font-size:11px;background:var(--sec);border:1px solid var(--bd);color:var(--mut)}
  .pill.ok{color:var(--ok);border-color:color-mix(in srgb,var(--ok) 40%,var(--bd))}
  .pill.warn{color:var(--warn);border-color:color-mix(in srgb,var(--warn) 40%,var(--bd))}
  .card{background:var(--card);border:1px solid var(--bd);border-radius:10px;padding:12px 14px;display:flex;flex-direction:column;gap:10px}
  .card h2{margin:0;font-size:13px;font-weight:600}
  .row{display:flex;align-items:center;justify-content:space-between;gap:10px}
  .mut{color:var(--mut);font-size:12px}
  .toggle{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:6px 0;border-bottom:1px solid var(--bd)}
  .toggle:last-child{border-bottom:0;padding-bottom:0}
  .toggle label{display:flex;flex-direction:column;gap:2px;cursor:pointer;flex:1}
  input[type=checkbox]{width:16px;height:16px;accent-color:var(--pri);cursor:pointer}
  .input-row{display:flex;gap:8px}
  .input-row input{flex:1;min-width:0;height:32px;padding:0 10px;border-radius:7px;border:1px solid var(--bd);background:var(--bg);color:var(--fg);font:inherit;outline:none}
  .input-row input:focus{border-color:var(--pri)}
  button{height:30px;padding:0 12px;border-radius:7px;border:1px solid var(--bd);background:var(--sec);color:var(--fg);font:inherit;font-size:12px;cursor:pointer;white-space:nowrap}
  button.primary{background:var(--pri);color:#fff;border-color:var(--pri);font-weight:500}
  button:disabled{opacity:.5;cursor:default}
  .diag-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}
  .diag-item{display:flex;align-items:center;justify-content:space-between;padding:6px 10px;border-radius:7px;background:var(--sec);font-size:12px}
  .dot{width:7px;height:7px;border-radius:50%;display:inline-block}
  .dot.ok{background:var(--ok)}.dot.no{background:var(--mut)}
  .pv-box{border:1px solid var(--bd);border-radius:8px;overflow:hidden;background:var(--bg)}
  .pv-box iframe{width:100%;height:520px;border:0;display:block;background:#fff}
  .pv-pages{display:flex;flex-direction:column;gap:12px;padding:12px;max-height:560px;overflow:auto}
  .pv-pages img{width:100%;height:auto;border-radius:6px;border:1px solid var(--bd);background:#fff}
  .pv-note{font-size:12px;padding:8px 10px;border-radius:6px;background:var(--card);border:1px solid var(--bd);white-space:pre-wrap}
</style>
</head>
<body>
<div class="stack">
  <div class="head">
    <h1>Agy Preview <span class="pill">v${VERSION}</span></h1>
    <span id="cdp-status" class="pill">检测中…</span>
  </div>

  <section class="card">
    <div class="row"><h2>快速预览本地文件</h2><span class="mut">支持 .docx / .pptx / .html / 图片</span></div>
    <div class="input-row">
      <input id="quick-path" type="text" placeholder="粘贴绝对路径，例如 D:\\docs\\demo.pptx" />
      <button id="quick-go" class="primary" type="button">预览</button>
    </div>
    <div id="quick-out"></div>
  </section>

  <section class="card">
    <h2>功能开关</h2>
    <div id="toggles"></div>
  </section>

  <section class="card">
    <div class="row">
      <h2>本机转换引擎与缓存</h2>
      <div style="display:flex;gap:6px">
        <button id="refresh-diag" type="button">重新检测</button>
        <button id="clear-cache" type="button">清理缓存</button>
      </div>
    </div>
    <div id="diag" class="diag-grid"></div>
    <div id="cache-info" class="mut"></div>
  </section>
</div>
<script>
const TOKEN = ${JSON.stringify(token)};
const headers = TOKEN ? { 'Content-Type': 'application/json', 'x-sidecar-token': TOKEN } : { 'Content-Type': 'application/json' };
async function api(path, body) {
  const res = await fetch(path, body ? { method: 'POST', headers, body: JSON.stringify(body) } : { headers });
  const data = await res.json();
  if (!res.ok || data.error) throw new Error(data.error || ('HTTP ' + res.status));
  return data;
}
const toggleDefs = [
  ['enabled', '启用 Agy Preview', '插件总开关，关闭后立即停用所有界面预览注入'],
  ['officePreview', 'Office 文档预览（DOCX / PPTX）', '在右侧文件标签页直接高保真预览 Word 与 PowerPoint（支持演讲者备注与流式阅读）'],
  ['htmlPreview', 'HTML 网页实时预览', '打开 .html 文件时提供「网页预览 / 源码」一键切换，内联同目录样式与图片'],
  ['htmlDefaultPreview', 'HTML 打开时默认进网页预览', '关闭后默认展示源码，可随时手动切到网页预览'],
  ['inlineImages', '过程看图原位合并展示', '将工具过程里连续的 Analyzed 图片行原位合并为「已查看 N 张图像」缩略图条'],
  ['markdownImages', '回复正文图片链接直接渲染', '在 AI 回复正文引用本地图片链接的段落下方直接渲染横向缩略图卡片（点击灯箱放大）'],
  ['chatLinkButtons', '聊天链接快捷预览按钮', '在回复中的 .docx / .pptx / .html 链接旁显示一键预览图标']
];
async function loadStatus(force = false) {
  const data = await api(force ? '/api/status?force=1' : '/api/status');
  const badge = document.getElementById('cdp-status');
  if (data.cdp.connected) {
    badge.className = 'pill ok';
    badge.textContent = '已挂载主界面 (' + data.cdp.windowCount + ')';
  } else {
    badge.className = 'pill warn';
    badge.textContent = '仅侧栏模式（未检测到 CDP 调试端口）';
  }
  const box = document.getElementById('toggles');
  box.innerHTML = '';
  for (const [key, title, desc] of toggleDefs) {
    const row = document.createElement('div');
    row.className = 'toggle';
    const lbl = document.createElement('label');
    lbl.innerHTML = '<strong>' + title + '</strong><span class="mut">' + desc + '</span>';
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = Boolean(data.settings[key]);
    cb.addEventListener('change', async () => {
      cb.disabled = true;
      try { await api('/api/settings', { [key]: cb.checked }); }
      catch (e) { cb.checked = !cb.checked; alert(e.message); }
      finally { cb.disabled = false; }
    });
    row.append(lbl, cb);
    box.append(row);
  }
  const d = data.diagnostics || {};
  const items = [
    ['PowerPoint COM', d.powerpointCom],
    ['Word COM', d.wordCom],
    ['LibreOffice', d.libreoffice],
    ['PyMuPDF (fitz)', d.pymupdf],
    ['python-docx', d.pythonDocx],
    ['python-pptx', d.pythonPptx],
    ['Pillow (PIL)', d.pillow],
    ['Node.js ' + (d.nodeVersion || ''), true]
  ];
  document.getElementById('diag').innerHTML = items.map(([name, ok]) =>
    '<div class="diag-item"><span>' + name + '</span><span class="dot ' + (ok ? 'ok' : 'no') + '" title="' + (ok ? '可用' : '未安装') + '"></span></div>'
  ).join('');
  const c = data.cache || { entries: 0, bytes: 0 };
  document.getElementById('cache-info').textContent = '本地预览缓存：' + c.entries + ' 个文档，共 ' + (c.bytes / 1048576).toFixed(2) + ' MB';
}
document.getElementById('refresh-diag').addEventListener('click', () => loadStatus(true).catch(e => alert(e.message)));
document.getElementById('clear-cache').addEventListener('click', async () => {
  await api('/api/cache/clear', {});
  await loadStatus();
});
document.getElementById('quick-go').addEventListener('click', async () => {
  const p = document.getElementById('quick-path').value.trim().replace(/^"|"$/g, '');
  if (!p) return;
  const out = document.getElementById('quick-out');
  out.innerHTML = '<div class="mut">正在生成预览…</div>';
  try {
    const res = await api('/api/quick-preview', { path: p });
    if (res.kind === 'html') {
      out.innerHTML = '<div class="pv-box"><iframe sandbox="allow-scripts allow-forms allow-popups"></iframe></div>';
      out.querySelector('iframe').srcdoc = res.html;
    } else if (res.kind === 'image') {
      out.innerHTML = '<div class="pv-box"><div class="pv-pages"><img /></div></div>';
      out.querySelector('img').src = res.dataUrl;
    } else if (res.mode === 'reader' && res.html) {
      out.innerHTML = '<div class="pv-box"><iframe sandbox=""></iframe></div>';
      out.querySelector('iframe').srcdoc = '<!DOCTYPE html><html><body style="font-family:sans-serif;padding:16px;line-height:1.6">' + res.html + '</body></html>';
    } else {
      const wrap = document.createElement('div');
      wrap.className = 'pv-box';
      const list = document.createElement('div');
      list.className = 'pv-pages';
      wrap.append(list);
      out.innerHTML = '';
      out.append(wrap);
      for (const pg of (res.pages || [])) {
        const pageRes = await api('/api/office-page', { cacheKey: res.cacheKey, page: pg.index });
        const img = document.createElement('img');
        img.src = pageRes.dataUrl;
        list.append(img);
        if (pg.notes) {
          const note = document.createElement('div');
          note.className = 'pv-note';
          note.textContent = '备注 (' + pg.index + '): ' + pg.notes;
          list.append(note);
        }
      }
    }
    void loadStatus();
  } catch (e) {
    out.innerHTML = '';
    const errDiv = document.createElement('div');
    errDiv.style.cssText = 'color:var(--err);font-size:12px';
    errDiv.textContent = e.message;
    out.append(errDiv);
  }
});
void loadStatus();
</script>
</body>
</html>`;
}

export function startAuxPaneServer({
  settings,
  cdpBridge,
  dataDir = pluginPaths().dataDir,
  port,
  token = process.env.ANTIGRAVITY_SIDECAR_UI_TOKEN || ''
}) {
  const hasExplicitPort = typeof port === 'number';
  const listenPort = hasExplicitPort ? port : Number(process.env.ANTIGRAVITY_SIDECAR_WEB_PORT || 0);
  if ((!hasExplicitPort && !listenPort) || !Number.isInteger(listenPort) || listenPort < 0 || listenPort > 65535) {
    const noop = async () => {};
    noop.port = 0;
    return Promise.resolve(noop);
  }
  return new Promise((resolve, reject) => {
    const server = createServer(async (req, res) => {
      const activePort = server.address()?.port || listenPort;
      const url = new URL(req.url || '/', `http://127.0.0.1:${activePort}`);
      const authed = !token || (req.headers['x-sidecar-token'] || url.searchParams.get('token')) === token;
      if (url.pathname === '/' && req.method === 'GET') {
        if (!authed) { res.writeHead(403); res.end('Forbidden'); return; }
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
        res.end(renderHtmlShell(token));
        return;
      }
      if (!authed) { sendJson(res, 403, { error: '未授权访问' }); return; }
      try {
        if (url.pathname === '/api/status' && req.method === 'GET') {
          sendJson(res, 200, {
            ok: true,
            version: VERSION,
            settings: settings.get(),
            cdp: cdpBridge.status(),
            diagnostics: await getEnvironmentDiagnostics({ force: url.searchParams.get('force') === '1' }),
            cache: getCacheStats(dataDir)
          });
        } else if (url.pathname === '/api/settings' && req.method === 'POST') {
          const patch = await readJsonBody(req);
          const next = settings.set(patch);
          await cdpBridge.broadcastSettings(next);
          sendJson(res, 200, { ok: true, settings: next });
        } else if (url.pathname === '/api/cache/clear' && req.method === 'POST') {
          sendJson(res, 200, clearPreviewCache(dataDir));
        } else if (url.pathname === '/api/quick-preview' && req.method === 'POST') {
          const { path: targetPath, mode } = await readJsonBody(req);
          const kind = classifyPreviewExt(targetPath);
          if (kind === 'html') sendJson(res, 200, renderHtmlPreview(targetPath));
          else if (kind === 'image') sendJson(res, 200, await renderImagePreview(targetPath, { full: true }));
          else if (kind === 'docx' || kind === 'pptx') sendJson(res, 200, await renderOfficePreview(targetPath, { mode: mode || 'layout', dataDir }));
          else throw new Error('暂不支持此文件类型（仅支持 .docx / .pptx / .html / 常见图片）');
        } else if (url.pathname === '/api/office-page' && req.method === 'POST') {
          const { cacheKey, page } = await readJsonBody(req);
          sendJson(res, 200, getOfficePreviewPage(cacheKey, page, { dataDir }));
        } else if (url.pathname === '/api/file/open' && req.method === 'POST') {
          const { path: targetPath, action } = await readJsonBody(req);
          sendJson(res, 200, openLocalPath(targetPath, { action: action || 'open' }));
        } else {
          sendJson(res, 404, { error: 'Not found' });
        }
      } catch (err) {
        sendJson(res, 400, { error: err.message || '请求失败' });
      }
    });
    server.once('error', reject);
    server.listen(listenPort, '127.0.0.1', () => {
      const close = () => new Promise(done => server.close(() => done()));
      close.port = server.address().port;
      resolve(close);
    });
  });
}

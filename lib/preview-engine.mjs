import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync } from 'node:fs';
import { homedir } from 'node:os';
import { basename, dirname, extname, isAbsolute, join, relative, resolve } from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { pluginPaths } from './paths.mjs';
import { resolveLocalPath } from './file-actions.mjs';
import { scanConversationFiles } from './transcript-images.mjs';

const OFFICE_EXTS = new Set(['.docx', '.doc', '.pptx', '.ppt']);
const HTML_EXTS = new Set(['.html', '.htm']);
const IMAGE_EXTS = new Set(['.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.svg']);
const IMAGE_MIME = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.webp': 'image/webp',
  '.gif': 'image/gif',
  '.bmp': 'image/bmp',
  '.svg': 'image/svg+xml'
};

const MAX_HTML_BYTES = 5 * 1024 * 1024;
const MAX_INLINE_ASSET_BYTES = 3 * 1024 * 1024;
const MAX_DIRECT_IMAGE_BYTES = 350 * 1024;
const MAX_CACHE_ENTRIES = 36;

const HELPER_SCRIPT = resolve(dirname(fileURLToPath(import.meta.url)), '..', 'scripts', 'office-preview.py');
const pendingOffice = new Map();
const imageThumbCache = new Map();
let cachedPython = null;
let cachedDiagnostics = null;

export function classifyPreviewExt(filePath) {
  const ext = extname(String(filePath || '')).toLowerCase();
  if (OFFICE_EXTS.has(ext)) return ext === '.pptx' || ext === '.ppt' ? 'pptx' : 'docx';
  if (HTML_EXTS.has(ext)) return 'html';
  if (IMAGE_EXTS.has(ext)) return 'image';
  return '';
}

function safeDecode(value) {
  try { return decodeURIComponent(String(value || '')); } catch { return String(value || ''); }
}

export function resolvePreviewPath({ path: directPath, uri, fileName, segments = [], roots = [], conversationId = '', appDataDir = '' } = {}) {
  for (const candidate of [directPath, uri]) {
    if (typeof candidate === 'string' && candidate.trim()) {
      try {
        const resolved = resolveLocalPath(candidate);
        if (existsSync(resolved) && statSync(resolved).isFile()) return resolved;
      } catch {}
    }
  }

  const cleanSegments = (Array.isArray(segments) ? segments : [])
    .map(s => safeDecode(String(s || '').trim()))
    .filter(Boolean);
  const targetName = safeDecode(String(fileName || cleanSegments.at(-1) || '').trim());
  const cleanRoots = (Array.isArray(roots) ? roots : [])
    .map(r => {
      try { return resolveLocalPath(r); } catch { return ''; }
    })
    .filter(Boolean);

  if (cleanSegments.length) {
    for (const root of cleanRoots) {
      const rootBase = basename(root).toLowerCase();
      const firstSeg = cleanSegments[0].toLowerCase();
      const tail = firstSeg === rootBase ? cleanSegments.slice(1) : cleanSegments;
      for (let start = 0; start < tail.length; start++) {
        const cand = resolve(root, ...tail.slice(start));
        try {
          if (existsSync(cand) && statSync(cand).isFile()) return cand;
        } catch {}
      }
    }
  }

  if (targetName) {
    const scanned = scanConversationFiles(appDataDir, conversationId);
    const targetLow = targetName.toLowerCase();
    const normSegs = cleanSegments.map(s => s.toLowerCase());
    let best = '';
    let bestScore = -1;
    for (const p of scanned.allPaths) {
      if (basename(p).toLowerCase() !== targetLow) continue;
      try { if (!existsSync(p) || !statSync(p).isFile()) continue; } catch { continue; }
      const parts = p.replace(/\\/g, '/').toLowerCase().split('/');
      let score = 0;
      for (let i = 1; i <= Math.min(normSegs.length, parts.length); i++) {
        if (parts[parts.length - i] === normSegs[normSegs.length - i]) score++;
        else break;
      }
      if (score > bestScore) { bestScore = score; best = p; }
    }
    if (best) return best;
  }

  if (targetName) {
    for (const root of cleanRoots) {
      const cand = resolve(root, targetName);
      try {
        if (existsSync(cand) && statSync(cand).isFile()) return cand;
      } catch {}
    }
  }

  throw new Error('未能定位该本地文件，请确认文件位于当前工作区或对话记录中');
}

function findPython() {
  if (cachedPython && existsSync(cachedPython)) return cachedPython;
  const home = homedir();
  const candidates = [
    process.env.AGY_PREVIEW_PYTHON,
    join(home, 'AppData', 'Local', 'Python', 'bin', 'python.exe'),
    'python',
    'python3',
    join(home, '.cache', 'codex-runtimes', 'codex-primary-runtime', 'dependencies', 'python', 'python.exe')
  ].filter(Boolean);
  for (const cmd of candidates) {
    try {
      const r = spawnSync(cmd, ['-c', 'import sys; print(sys.executable)'], {
        encoding: 'utf8',
        windowsHide: true,
        timeout: 4000
      });
      if (!r.error && r.status === 0 && r.stdout.trim()) {
        cachedPython = r.stdout.trim();
        return cachedPython;
      }
    } catch {}
  }
  return '';
}

function runPythonJson(args, timeoutMs = 55000) {
  const python = findPython();
  if (!python) return Promise.reject(new Error('未找到可用的 Python 运行环境'));
  return new Promise((resolvePromise, rejectPromise) => {
    let stdout = '';
    let stderr = '';
    let settled = false;
    const finish = (err, data) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (err) rejectPromise(err);
      else resolvePromise(data);
    };
    const child = spawn(python, [HELPER_SCRIPT, ...args], {
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
      env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' }
    });
    const timer = setTimeout(() => {
      try { child.kill(); } catch {}
      finish(new Error('文档预览转换超时'));
    }, timeoutMs);
    child.stdout.on('data', chunk => { stdout += chunk.toString('utf8'); });
    child.stderr.on('data', chunk => { if (stderr.length < 4096) stderr += chunk.toString('utf8'); });
    child.on('error', err => finish(err));
    child.on('close', code => {
      const trimmed = stdout.trim();
      let parsed = null;
      if (trimmed) {
        const lastLine = trimmed.split(/\r?\n/).reverse().find(l => l.trim().startsWith('{')) || trimmed;
        try { parsed = JSON.parse(lastLine); } catch {}
      }
      if (parsed && parsed.ok) return finish(null, parsed);
      finish(new Error(parsed?.error || stderr.trim() || `预览转换退出异常 (${code})`));
    });
  });
}

export async function getEnvironmentDiagnostics({ force = false } = {}) {
  if (!force && cachedDiagnostics) return cachedDiagnostics;
  try {
    const info = await runPythonJson(['diagnostics'], 8000);
    cachedDiagnostics = { ...info, nodeVersion: process.versions.node };
    return cachedDiagnostics;
  } catch (err) {
    return {
      ok: false,
      nodeVersion: process.versions.node,
      python: '',
      powerpointCom: false,
      wordCom: false,
      libreoffice: false,
      pymupdf: false,
      pillow: false,
      pythonDocx: false,
      pythonPptx: false,
      error: err.message
    };
  }
}

export function getCacheStats(dataDir = pluginPaths().dataDir) {
  const cacheRoot = join(dataDir, 'preview-cache');
  let entries = 0;
  let bytes = 0;
  const walk = dir => {
    try {
      for (const item of readdirSync(dir, { withFileTypes: true })) {
        const full = join(dir, item.name);
        if (item.isDirectory()) walk(full);
        else if (item.isFile()) {
          try { bytes += statSync(full).size; } catch {}
        }
      }
    } catch {}
  };
  if (existsSync(cacheRoot)) {
    try {
      entries = readdirSync(cacheRoot, { withFileTypes: true }).filter(d => d.isDirectory()).length;
      walk(cacheRoot);
    } catch {}
  }
  return { entries, bytes, imageMemoryEntries: imageThumbCache.size };
}

export function clearPreviewCache(dataDir = pluginPaths().dataDir) {
  const cacheRoot = join(dataDir, 'preview-cache');
  imageThumbCache.clear();
  if (existsSync(cacheRoot)) {
    rmSync(cacheRoot, { recursive: true, force: true });
  }
  return { ok: true, ...getCacheStats(dataDir) };
}

function pruneCacheDir(cacheRoot) {
  try {
    if (!existsSync(cacheRoot)) return;
    const dirs = readdirSync(cacheRoot, { withFileTypes: true })
      .filter(d => d.isDirectory() && /^[a-f0-9]{16,32}$/.test(d.name))
      .map(d => {
        const full = join(cacheRoot, d.name);
        let mtime = 0;
        try { mtime = statSync(full).mtimeMs; } catch {}
        return { full, mtime };
      })
      .sort((a, b) => b.mtime - a.mtime);
    for (const stale of dirs.slice(MAX_CACHE_ENTRIES)) {
      try { rmSync(stale.full, { recursive: true, force: true }); } catch {}
    }
  } catch {}
}

export async function renderOfficePreview(filePathInput, { mode = 'layout', force = false, dataDir = pluginPaths().dataDir } = {}) {
  const filePath = resolveLocalPath(filePathInput);
  if (!existsSync(filePath)) throw new Error('文件不存在');
  const st = statSync(filePath);
  if (!st.isFile()) throw new Error('目标不是文件');
  const kind = classifyPreviewExt(filePath);
  if (kind !== 'docx' && kind !== 'pptx') throw new Error('不支持的 Office 文档格式');

  const safeMode = kind === 'docx' && mode === 'reader' ? 'reader' : 'layout';
  const cacheRoot = join(dataDir, 'preview-cache');
  const profileDir = join(dataDir, 'lo-profile');
  const cacheKey = createHash('sha256')
    .update(`${filePath}|${st.mtimeMs}|${st.size}|${safeMode}`)
    .digest('hex')
    .slice(0, 24);
  const outDir = join(cacheRoot, cacheKey);
  const metaPath = join(outDir, 'meta.json');

  const buildReply = meta => {
    const html = meta.htmlFile && existsSync(join(outDir, meta.htmlFile))
      ? readFileSync(join(outDir, meta.htmlFile), 'utf8')
      : '';
    return {
      ok: true,
      kind: meta.kind || kind,
      mode: meta.mode || safeMode,
      engine: meta.engine || '',
      filePath,
      fileName: basename(filePath),
      mtimeMs: st.mtimeMs,
      size: st.size,
      cacheKey,
      pageCount: Number(meta.pageCount || 0),
      pages: Array.isArray(meta.pages)
        ? meta.pages.map(p => ({ index: p.index, width: p.width, height: p.height, title: p.title || '', notes: p.notes || '' }))
        : [],
      html
    };
  };

  if (!force && existsSync(metaPath)) {
    try {
      const cached = JSON.parse(readFileSync(metaPath, 'utf8'));
      if (cached && cached.ok) return buildReply(cached);
    } catch {}
  }

  const flightKey = `${cacheKey}:${force ? '1' : '0'}`;
  if (pendingOffice.has(flightKey)) return pendingOffice.get(flightKey);

  const work = (async () => {
    mkdirSync(cacheRoot, { recursive: true });
    if (force && existsSync(outDir)) {
      try { rmSync(outDir, { recursive: true, force: true }); } catch {}
    }
    const meta = await runPythonJson(['render-office', filePath, outDir, safeMode, profileDir], 55000);
    pruneCacheDir(cacheRoot);
    return buildReply(meta);
  })();

  pendingOffice.set(flightKey, work);
  try { return await work; }
  finally { pendingOffice.delete(flightKey); }
}

export function getOfficePreviewPage(cacheKey, pageIndex, { dataDir = pluginPaths().dataDir } = {}) {
  if (!/^[a-f0-9]{16,32}$/.test(String(cacheKey || ''))) throw new Error('无效的预览缓存标识');
  const idx = Number(pageIndex);
  if (!Number.isInteger(idx) || idx < 1 || idx > 500) throw new Error('无效的页码');
  const outDir = join(dataDir, 'preview-cache', cacheKey);
  const metaPath = join(outDir, 'meta.json');
  if (!existsSync(metaPath)) throw new Error('预览缓存已过期，请刷新');
  const meta = JSON.parse(readFileSync(metaPath, 'utf8'));
  const entry = (meta.pages || []).find(p => p.index === idx);
  if (!entry || !entry.file || /[\\/]/.test(entry.file)) throw new Error('未找到对应页面');
  const filePath = join(outDir, entry.file);
  if (!existsSync(filePath)) throw new Error('页面图像缺失，请刷新');
  const mime = entry.mime || (entry.file.endsWith('.webp') ? 'image/webp' : 'image/png');
  const dataUrl = `data:${mime};base64,${readFileSync(filePath).toString('base64')}`;
  return { ok: true, index: idx, width: entry.width, height: entry.height, dataUrl };
}

function resolveRelativeAsset(baseDir, ref) {
  const clean = String(ref || '').trim();
  if (!clean || /^(?:https?:|data:|blob:|mailto:|tel:|javascript:|#)/i.test(clean)) return '';
  const withoutHash = clean.split('#')[0].split('?')[0];
  let decoded = withoutHash;
  try { decoded = decodeURIComponent(withoutHash); } catch {}
  if (!decoded) return '';
  const candidate = resolve(baseDir, decoded);
  const rel = relative(baseDir, candidate);
  const upLevels = (rel.match(/(?:^|[\\/])\.\.(?=$|[\\/])/g) || []).length;
  if (upLevels > 2 || isAbsolute(rel)) return '';
  try {
    if (existsSync(candidate) && statSync(candidate).isFile()) return candidate;
  } catch {}
  return '';
}

function assetToDataUri(filePath) {
  try {
    const st = statSync(filePath);
    if (st.size > MAX_INLINE_ASSET_BYTES) return '';
    const ext = extname(filePath).toLowerCase();
    const mime = IMAGE_MIME[ext] || ({
      '.css': 'text/css',
      '.js': 'application/javascript',
      '.woff2': 'font/woff2',
      '.woff': 'font/woff',
      '.ttf': 'font/ttf',
      '.otf': 'font/otf'
    })[ext];
    if (!mime) return '';
    return `data:${mime};base64,${readFileSync(filePath).toString('base64')}`;
  } catch {
    return '';
  }
}

function inlineCssUrls(cssText, cssDir) {
  return cssText.replace(/url\(\s*(['"]?)([^)'"]+)\1\s*\)/gi, (full, quote, ref) => {
    const target = resolveRelativeAsset(cssDir, ref);
    if (!target) return full;
    const dataUri = assetToDataUri(target);
    return dataUri ? `url("${dataUri}")` : full;
  });
}

export function renderHtmlPreview(filePathInput) {
  const filePath = resolveLocalPath(filePathInput);
  if (!existsSync(filePath)) throw new Error('HTML 文件不存在');
  const st = statSync(filePath);
  if (!st.isFile()) throw new Error('目标不是文件');
  if (classifyPreviewExt(filePath) !== 'html') throw new Error('目标不是 HTML 文件');
  if (st.size > MAX_HTML_BYTES) throw new Error('HTML 文件超过 5MB，暂不支持直接渲染');

  const baseDir = dirname(filePath);
  let html = readFileSync(filePath, 'utf8');

  html = html.replace(/<link\b([^>]*?)\bhref=(["'])([^"']+)\2([^>]*?)>/gi, (full, pre, q, ref, post) => {
    if (!/rel\s*=\s*["']?stylesheet/i.test(pre + ' ' + post)) return full;
    const target = resolveRelativeAsset(baseDir, ref);
    if (!target) return full;
    try {
      if (statSync(target).size > MAX_INLINE_ASSET_BYTES) return full;
      const css = inlineCssUrls(readFileSync(target, 'utf8'), dirname(target));
      return `<style data-inlined-from="${basename(target)}">\n${css}\n</style>`;
    } catch { return full; }
  });

  html = html.replace(/<script\b([^>]*?)\bsrc=(["'])([^"']+)\2([^>]*)>\s*<\/script>/gi, (full, pre, q, ref, post) => {
    const target = resolveRelativeAsset(baseDir, ref);
    if (!target) return full;
    try {
      if (statSync(target).size > MAX_INLINE_ASSET_BYTES) return full;
      const js = readFileSync(target, 'utf8').replace(/<\/script>/gi, '<\\/script>');
      return `<script${pre}${post}>\n${js}\n</script>`;
    } catch { return full; }
  });

  html = html.replace(/(<(?:img|source)\b[^>]*?\bsrc=)(["'])([^"']+)\2/gi, (full, prefix, q, ref) => {
    const target = resolveRelativeAsset(baseDir, ref);
    if (!target) return full;
    const dataUri = assetToDataUri(target);
    return dataUri ? `${prefix}"${dataUri}"` : full;
  });

  html = html.replace(/(<style\b[^>]*>)([\s\S]*?)(<\/style>)/gi, (_, open, body, close) => {
    return `${open}${inlineCssUrls(body, baseDir)}${close}`;
  });

  return {
    ok: true,
    kind: 'html',
    filePath,
    fileName: basename(filePath),
    mtimeMs: st.mtimeMs,
    size: st.size,
    html
  };
}

export async function renderImagePreview(filePathInput, { full = false, maxWidth = 520 } = {}) {
  const filePath = resolveLocalPath(filePathInput);
  if (!existsSync(filePath)) throw new Error('图片文件不存在');
  const st = statSync(filePath);
  if (!st.isFile()) throw new Error('目标不是图片文件');
  const ext = extname(filePath).toLowerCase();
  if (!IMAGE_EXTS.has(ext)) throw new Error('不支持的图片格式');

  const targetWidth = full ? 1920 : Math.min(800, Math.max(160, Number(maxWidth) || 520));
  const key = `${filePath}|${st.mtimeMs}|${st.size}|${full ? 'full' : targetWidth}`;
  if (imageThumbCache.has(key)) return imageThumbCache.get(key);

  let reply;
  if (ext === '.svg' || (st.size <= MAX_DIRECT_IMAGE_BYTES && !full) || (full && st.size <= 4 * 1024 * 1024)) {
    const mime = IMAGE_MIME[ext] || 'image/png';
    reply = {
      ok: true,
      kind: 'image',
      filePath,
      fileName: basename(filePath),
      mime,
      size: st.size,
      dataUrl: `data:${mime};base64,${readFileSync(filePath).toString('base64')}`
    };
  } else {
    try {
      const thumb = await runPythonJson(['thumb-image', filePath, String(targetWidth)], 15000);
      reply = {
        ok: true,
        kind: 'image',
        filePath,
        fileName: basename(filePath),
        mime: thumb.mime,
        width: thumb.width,
        height: thumb.height,
        origWidth: thumb.origWidth,
        origHeight: thumb.origHeight,
        size: st.size,
        dataUrl: thumb.dataUrl
      };
    } catch {
      if (st.size > 6 * 1024 * 1024) throw new Error('图片过大且缩略图生成失败');
      const mime = IMAGE_MIME[ext] || 'image/png';
      reply = {
        ok: true,
        kind: 'image',
        filePath,
        fileName: basename(filePath),
        mime,
        size: st.size,
        dataUrl: `data:${mime};base64,${readFileSync(filePath).toString('base64')}`
      };
    }
  }

  imageThumbCache.set(key, reply);
  if (imageThumbCache.size > 80) imageThumbCache.delete(imageThumbCache.keys().next().value);
  return reply;
}

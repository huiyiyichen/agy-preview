import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { appDataDir, devToolsPortFile, pluginPaths } from './paths.mjs';
import { openLocalPath } from './file-actions.mjs';
import { scanConversationFiles } from './transcript-images.mjs';
import {
  clearPreviewCache,
  getCacheStats,
  getEnvironmentDiagnostics,
  getOfficePreviewPage,
  renderHtmlPreview,
  renderImagePreview,
  renderOfficePreview,
  resolvePreviewPath
} from './preview-engine.mjs';

const LIB_DIR = dirname(fileURLToPath(import.meta.url));
const PLUGIN_VERSION = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8')).version;

export function isHubTarget(target) {
  try {
    const url = new URL(target.url);
    return target.type === 'page' &&
      /(?:^| - |\(\d+\) )Antigravity(?:\s*(?:\([^)]+\)|\[[^\]]+\]))?$/.test(target.title || '') &&
      url.protocol === 'https:' &&
      ['127.0.0.1', 'localhost', '[::1]'].includes(url.hostname);
  } catch {
    return false;
  }
}

export function connectCdp(wsUrl) {
  return new Promise((resolvePromise, rejectPromise) => {
    const socket = new WebSocket(wsUrl);
    const pending = new Map();
    let seq = 0;
    const listeners = new Set();
    const timer = setTimeout(() => { socket.close(); rejectPromise(new Error('CDP 连接超时')); }, 4000);

    const failAll = () => {
      for (const entry of pending.values()) {
        clearTimeout(entry.timer);
        entry.reject(new Error('CDP 已断开'));
      }
      pending.clear();
    };

    socket.addEventListener('error', () => { clearTimeout(timer); rejectPromise(new Error('CDP 不可用')); });
    socket.addEventListener('close', failAll);
    socket.addEventListener('message', event => {
      let msg;
      try { msg = JSON.parse(event.data); } catch { return; }
      const entry = pending.get(msg.id);
      if (entry) {
        clearTimeout(entry.timer);
        pending.delete(msg.id);
        if (msg.error) entry.reject(new Error(msg.error.message));
        else entry.resolve(msg.result);
      } else {
        for (const fn of listeners) fn(msg);
      }
    });
    socket.addEventListener('open', () => {
      clearTimeout(timer);
      resolvePromise({
        socket,
        onEvent: fn => listeners.add(fn),
        send(method, params = {}) {
          if (socket.readyState !== WebSocket.OPEN) return Promise.reject(new Error('CDP 已断开'));
          return new Promise((res, rej) => {
            const id = ++seq;
            const timeoutMs = method === 'Runtime.evaluate' ? 20000 : 5000;
            const t = setTimeout(() => { pending.delete(id); rej(new Error(`${method} 超时`)); }, timeoutMs);
            pending.set(id, { resolve: res, reject: rej, timer: t });
            socket.send(JSON.stringify({ id, method, params }));
          });
        }
      });
    });
  });
}

export function buildInjectScript(origin, initialSettings) {
  const clientCode = readFileSync(join(LIB_DIR, 'client-inject.js'), 'utf8');
  return `if (location.origin === ${JSON.stringify(origin)} && window === window.top && globalThis.__agyPreviewRuntime?.version !== ${JSON.stringify(PLUGIN_VERSION)} && !globalThis.__agyPreviewBootPending) {\n` +
    `globalThis.__agyPreviewBootPending = true;\n` +
    `globalThis.__agyPreviewInitialSettings = ${JSON.stringify(initialSettings)};\n` +
    `globalThis.__agyPreviewVersion = ${JSON.stringify(PLUGIN_VERSION)};\n` +
    `(function boot() {\n` +
    `  if (!document.head || !document.body) { setTimeout(boot, 100); return; }\n` +
    `  try {\n` +
    `    globalThis.__agyPreviewRuntime?.dispose?.();\n` +
    `    ${clientCode}\n` +
    `  } catch (err) { console.error('[agy-preview]', err); }\n` +
    `  finally { delete globalThis.__agyPreviewBootPending; }\n` +
    `})();\n` +
    `}`;
}

export function startCdpBridge({ settings, dataDir = pluginPaths().dataDir, portFile, onLog = () => {} }) {
  const activePortFile = devToolsPortFile({ portFile });
  const clients = new Map();
  let stopped = false;
  let polling = false;
  let lastError = '';

  async function handleRpc(req) {
    const current = settings.get();
    const baseAppData = appDataDir();
    if (req.method === 'settings.get') return current;
    if (req.method === 'settings.set') {
      const next = settings.set(req.settings);
      void broadcastSettings(next);
      return next;
    }
    if (!current.enabled && req.method !== 'diagnostics') {
      throw new Error('Agy Preview 已关闭');
    }
    if (req.method === 'diagnostics') {
      return { ...(await getEnvironmentDiagnostics()), cache: getCacheStats(dataDir) };
    }
    if (req.method === 'cache.clear') {
      return clearPreviewCache(dataDir);
    }
    if (req.method === 'conversation.images') {
      const scanned = scanConversationFiles(baseAppData, req.conversationId);
      return { ok: true, turns: scanned.turns, images: scanned.images, byFileName: scanned.byFileName };
    }
    if (req.method === 'preview.office') {
      if (!current.officePreview) throw new Error('Office 文档预览已关闭');
      const resolved = resolvePreviewPath({
        path: req.path, uri: req.uri, fileName: req.fileName,
        segments: req.segments, roots: req.roots,
        conversationId: req.conversationId, appDataDir: baseAppData
      });
      return await renderOfficePreview(resolved, { mode: req.mode === 'reader' ? 'reader' : 'layout', force: req.force === true, dataDir });
    }
    if (req.method === 'preview.officePage') {
      if (!current.officePreview) throw new Error('Office 文档预览已关闭');
      return getOfficePreviewPage(req.cacheKey, req.page, { dataDir });
    }
    if (req.method === 'preview.html') {
      if (!current.htmlPreview) throw new Error('HTML 网页预览已关闭');
      const resolved = resolvePreviewPath({
        path: req.path, uri: req.uri, fileName: req.fileName,
        segments: req.segments, roots: req.roots,
        conversationId: req.conversationId, appDataDir: baseAppData
      });
      return renderHtmlPreview(resolved);
    }
    if (req.method === 'preview.image') {
      if (!current.inlineImages && !current.markdownImages && !current.officePreview) throw new Error('图片预览已关闭');
      const resolved = resolvePreviewPath({
        path: req.path, uri: req.uri, fileName: req.fileName,
        segments: req.segments, roots: req.roots,
        conversationId: req.conversationId, appDataDir: baseAppData
      });
      return await renderImagePreview(resolved, { full: req.full === true, maxWidth: req.maxWidth });
    }
    if (req.method === 'file.open') return openLocalPath(req.path || req.uri, { action: 'open' });
    if (req.method === 'file.reveal') return openLocalPath(req.path || req.uri, { action: 'reveal' });
    throw new Error('不支持的操作');
  }

  async function broadcastSettings(nextSettings) {
    const expr = `globalThis.__agyPreviewRuntime?.applySettings?.(${JSON.stringify(nextSettings)})`;
    for (const client of clients.values()) {
      try { await client.send('Runtime.evaluate', { expression: expr }); } catch {}
    }
  }

  async function attach(target) {
    const client = await connectCdp(target.webSocketDebuggerUrl);
    const origin = new URL(target.url).origin;
    const mainFrame = (await client.send('Page.getFrameTree')).frameTree.frame.id;
    const contexts = new Set();
    client.origin = origin;
    client.endpoint = target.webSocketDebuggerUrl;
    clients.set(target.id, client);
    client.socket.addEventListener('close', () => {
      if (clients.get(target.id) === client) clients.delete(target.id);
    });

    client.onEvent(async event => {
      if (event.method === 'Runtime.executionContextCreated') {
        const ctx = event.params.context;
        if (ctx.origin === origin && ctx.auxData?.isDefault && ctx.auxData.frameId === mainFrame) contexts.add(ctx.id);
        return;
      }
      if (event.method === 'Runtime.executionContextDestroyed') { contexts.delete(event.params.executionContextId); return; }
      if (event.method === 'Runtime.executionContextsCleared') { contexts.clear(); return; }
      if (event.method !== 'Runtime.bindingCalled' || event.params.name !== '__agyPreviewHost') return;
      if (!contexts.has(event.params.executionContextId)) return;
      let req;
      try {
        if (event.params.payload.length > 65536) return;
        req = JSON.parse(event.params.payload);
      } catch { return; }
      if (!Number.isInteger(req.id)) return;
      let reply;
      try { reply = await handleRpc(req); }
      catch (err) { reply = { error: err.message || '操作失败' }; }
      try {
        await client.send('Runtime.evaluate', {
          expression: `globalThis.__agyPreviewReceive?.(${JSON.stringify(req.id)},${JSON.stringify(reply)})`,
          contextId: event.params.executionContextId
        });
      } catch {}
    });

    await client.send('Runtime.enable');
    await client.send('Runtime.addBinding', { name: '__agyPreviewHost' });
    await client.send('Page.enable');
    const source = buildInjectScript(origin, settings.get());
    const installed = await client.send('Page.addScriptToEvaluateOnNewDocument', { source });
    client.scriptId = installed.identifier;
    await client.send('Runtime.evaluate', { expression: source });
    onLog(`已挂载主窗口 (${origin})`);
  }

  async function poll() {
    if (stopped || polling) return;
    polling = true;
    try {
      const port = Number(readFileSync(activePortFile, 'utf8').split('\n')[0]);
      if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error('调试端口未开启');
      const res = await fetch(`http://127.0.0.1:${port}/json/list`, { signal: AbortSignal.timeout(3000) });
      const targets = await res.json();
      const hubs = targets.filter(isHubTarget);
      const liveIds = new Set(hubs.map(t => t.id));
      for (const [id, client] of [...clients.entries()]) {
        if (!liveIds.has(id)) {
          try { client.socket.close(); } catch {}
          clients.delete(id);
        }
      }
      for (const target of hubs) {
        let client = clients.get(target.id);
        if (client && (client.origin !== new URL(target.url).origin || client.endpoint !== target.webSocketDebuggerUrl)) {
          try { client.socket.close(); } catch {}
          clients.delete(target.id);
          client = null;
        }
        if (!client) {
          try { await attach(target); } catch (e) { lastError = e.message; }
        } else {
          try {
            const check = await client.send('Runtime.evaluate', {
              expression: `globalThis.__agyPreviewRuntime?.version === ${JSON.stringify(PLUGIN_VERSION)}`,
              returnByValue: true
            });
            if (!check?.result?.value) {
              await client.send('Runtime.evaluate', { expression: buildInjectScript(client.origin, settings.get()) });
            }
          } catch {
            try { client.socket.close(); } catch {}
            clients.delete(target.id);
          }
        }
      }
      lastError = '';
    } catch (err) {
      lastError = err.message;
    } finally {
      polling = false;
    }
  }

  const timer = setInterval(() => { void poll(); }, 2500);
  void poll();

  return {
    broadcastSettings,
    status() {
      return {
        connected: clients.size > 0,
        windowCount: clients.size,
        version: PLUGIN_VERSION,
        lastError: clients.size > 0 ? '' : lastError
      };
    },
    async stop() {
      stopped = true;
      clearInterval(timer);
      for (const client of clients.values()) {
        try {
          if (client.scriptId) await client.send('Page.removeScriptToEvaluateOnNewDocument', { identifier: client.scriptId });
          await client.send('Runtime.evaluate', { expression: 'globalThis.__agyPreviewRuntime?.dispose?.()' });
        } catch {}
        try { client.socket.close(); } catch {}
      }
      clients.clear();
    }
  };
}

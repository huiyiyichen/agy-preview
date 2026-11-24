import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';

import { DEFAULT_SETTINGS, createSettingsStore } from '../lib/settings.mjs';
import { openLocalPath, resolveLocalPath } from '../lib/file-actions.mjs';
import { cleanPath, scanConversationFiles } from '../lib/transcript-images.mjs';
import {
  classifyPreviewExt,
  clearPreviewCache,
  getCacheStats,
  getEnvironmentDiagnostics,
  getOfficePreviewPage,
  renderHtmlPreview,
  renderImagePreview,
  renderOfficePreview
} from '../lib/preview-engine.mjs';
import { buildInjectScript } from '../lib/cdp-host.mjs';
import { startAuxPaneServer } from '../lib/aux-pane-ui.mjs';
import { collectPackageFiles, installPlugin } from '../scripts/install.mjs';
import { uninstallPlugin } from '../scripts/uninstall.mjs';

const WORK_ROOT = mkdtempSync(join(tmpdir(), 'agy-preview-test-'));

test.after(() => {
  rmSync(WORK_ROOT, { recursive: true, force: true });
});

function createSampleFixtures(dir, pythonBin) {
  mkdirSync(dir, { recursive: true });
  const script = `
import os, sys
from PIL import Image
import docx
from pptx import Presentation

out = sys.argv[1]
png_path = os.path.join(out, 'sample.png')
img = Image.new('RGB', (320, 180), (42, 130, 210))
img.save(png_path, 'PNG')

doc = docx.Document()
doc.add_heading('Agy Preview DOCX Test', level=1)
doc.add_paragraph('这是一段中文测试文本，用于验证原版排版与流式阅读。')
tbl = doc.add_table(rows=2, cols=2)
tbl.cell(0, 0).text = '指标'
tbl.cell(0, 1).text = '状态'
tbl.cell(1, 0).text = '独立插件'
tbl.cell(1, 1).text = '正常'
doc.add_picture(png_path)
doc.save(os.path.join(out, 'sample.docx'))

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = 'Agy Preview PPTX Slide 1'
slide.placeholders[1].text = '支持高保真预览与演讲者备注提取'
slide.notes_slide.notes_text_frame.text = '第一页演讲者备注：独立插件零冲突。'
prs.save(os.path.join(out, 'sample.pptx'))
`;
  const r = spawnSync(pythonBin, ['-c', script, dir], {
    encoding: 'utf8',
    windowsHide: true,
    env: { ...process.env, PYTHONUTF8: '1' }
  });
  assert.equal(r.status, 0, `生成测试样本失败: ${r.stderr}`);

  writeFileSync(join(dir, 'style.css'), 'body { color: #222; background-image: url("./sample.png"); }\n', 'utf8');
  writeFileSync(join(dir, 'app.js'), 'window.__tested = true;\n', 'utf8');
  writeFileSync(
    join(dir, 'page.html'),
    '<!DOCTYPE html><html><head><link rel="stylesheet" href="./style.css"></head><body><h1>Hello HTML</h1><img src="./sample.png"><script src="./app.js"></script></body></html>',
    'utf8'
  );
}

test('inject script compiles cleanly and stays namespace-isolated', () => {
  const src = buildInjectScript('https://127.0.0.1:4242', DEFAULT_SETTINGS);
  assert.doesNotThrow(() => new vm.Script(src));
  assert.match(src, /__agyPreviewRuntime/);
  assert.match(src, /__agyPreviewHost/);
  assert.doesNotMatch(src, /__agyPowertools/);
  assert.doesNotMatch(src, /\bReactDOM\b/);
  assert.doesNotMatch(src, /·/);
});

test('settings store validates and persists updates', () => {
  const storeDir = join(WORK_ROOT, 'settings-store');
  const store = createSettingsStore(storeDir);
  assert.deepEqual(store.get(), DEFAULT_SETTINGS);

  const updated = store.set({ htmlDefaultPreview: false, language: 'en' });
  assert.equal(updated.htmlDefaultPreview, false);
  assert.equal(updated.language, 'en');
  assert.equal(store.get().htmlDefaultPreview, false);

  assert.throws(() => store.set({ unknownKey: true }), /无效的设置项/);
  assert.throws(() => store.set({ enabled: 'yes' }), /无效的设置项/);
});

test('file-actions blocks UNC paths and Windows executables', () => {
  assert.throws(() => resolveLocalPath('\\\\evil-host\\share\\doc.docx'), /仅支持本机文件路径/);
  assert.throws(() => resolveLocalPath('file://evil-host/share/doc.docx'), /仅支持本机文件路径/);

  const exeFile = join(WORK_ROOT, 'dummy.ps1');
  writeFileSync(exeFile, 'Write-Host test', 'utf8');
  assert.throws(
    () => openLocalPath(exeFile, { action: 'open', platform: 'win32' }),
    /可执行脚本或程序/
  );
  assert.equal(resolveLocalPath(pathToFileURL(exeFile).href).toLowerCase(), exeFile.toLowerCase());
});

test('transcript-images extracts double-encoded and normal read_file image paths', () => {
  const fakeAppData = join(WORK_ROOT, 'app-data');
  const convId = 'conv-demo-01';
  const logsDir = join(fakeAppData, 'brain', convId, '.system_generated', 'logs');
  mkdirSync(logsDir, { recursive: true });
  const imgA = join(WORK_ROOT, 'f1_cover.png');
  const imgB = join(WORK_ROOT, 'f2_p11.png');
  const imgC = join(WORK_ROOT, 'f3_thanks.png');
  const htmlDoc = join(WORK_ROOT, 'demo-qa.html');
  // Mimic real Antigravity transcript.jsonl double-encoded JSON string in tool_calls.args.Path
  const doubleEncodedA = JSON.stringify(imgA);
  assert.equal(cleanPath(doubleEncodedA).toLowerCase(), imgA.toLowerCase());

  const lines = [
    JSON.stringify({ step_index: 0, type: 'USER_INPUT', content: '看看这三张图' }),
    JSON.stringify({
      step_index: 1,
      type: 'PLANNER_RESPONSE',
      tool_calls: [
        { name: 'read_file', args: { Path: doubleEncodedA } },
        { name: 'view_file', args: { AbsolutePath: imgB } },
        { name: 'read_file', args: { Path: JSON.stringify(imgC) } },
        { name: 'write_file', args: { Path: JSON.stringify(htmlDoc) } }
      ]
    })
  ];
  writeFileSync(join(logsDir, 'transcript.jsonl'), lines.join('\n') + '\n', 'utf8');

  const scanned = scanConversationFiles(fakeAppData, convId);
  assert.equal(scanned.turns.length, 1);
  assert.equal(scanned.turns[0].images.length, 3);
  assert.equal(scanned.images.length, 3);
  assert.equal(scanned.byFileName['f1_cover.png'][0].path.toLowerCase(), imgA.toLowerCase());
  assert.equal(scanned.byFileName['f2_p11.png'][0].path.toLowerCase(), imgB.toLowerCase());
  assert.equal(scanned.byFileName['f3_thanks.png'][0].path.toLowerCase(), imgC.toLowerCase());
  assert.ok(scanned.allPaths.some(p => p.toLowerCase() === htmlDoc.toLowerCase()));
});

test('end-to-end preview engine (diagnostics, PPTX, DOCX, HTML, Image, cache)', async () => {
  const diag = await getEnvironmentDiagnostics({ force: true });
  assert.equal(diag.ok, true, `Python 环境检测失败: ${diag.error || ''}`);
  assert.ok(diag.python);

  const fixturesDir = join(WORK_ROOT, 'fixtures');
  const cacheDataDir = join(WORK_ROOT, 'cache-data');
  createSampleFixtures(fixturesDir, diag.python);

  assert.equal(classifyPreviewExt(join(fixturesDir, 'sample.pptx')), 'pptx');
  assert.equal(classifyPreviewExt(join(fixturesDir, 'sample.docx')), 'docx');
  assert.equal(classifyPreviewExt(join(fixturesDir, 'page.html')), 'html');
  assert.equal(classifyPreviewExt(join(fixturesDir, 'sample.png')), 'image');

  // HTML preview inlines relative CSS, JS, and PNG
  const htmlRes = renderHtmlPreview(join(fixturesDir, 'page.html'));
  assert.equal(htmlRes.ok, true);
  assert.equal(htmlRes.kind, 'html');
  assert.match(htmlRes.html, /<style data-inlined-from="style\.css">/);
  assert.match(htmlRes.html, /data:image\/png;base64,/);
  assert.match(htmlRes.html, /window\.__tested = true/);

  // Image preview returns dataUrl with kind: 'image'
  const imgThumb = await renderImagePreview(join(fixturesDir, 'sample.png'), { full: false, maxWidth: 240 });
  assert.equal(imgThumb.ok, true);
  assert.equal(imgThumb.kind, 'image');
  assert.match(imgThumb.dataUrl, /^data:image\//);

  const imgFull = await renderImagePreview(join(fixturesDir, 'sample.png'), { full: true });
  assert.equal(imgFull.ok, true);
  assert.equal(imgFull.kind, 'image');

  // PPTX preview (layout + speaker notes + page fetch + cache hit)
  const pptxRes = await renderOfficePreview(join(fixturesDir, 'sample.pptx'), { mode: 'layout', dataDir: cacheDataDir });
  assert.equal(pptxRes.ok, true);
  assert.equal(pptxRes.kind, 'pptx');
  assert.ok(pptxRes.pageCount >= 1);
  assert.match(pptxRes.pages[0].notes, /演讲者备注/);

  const page1 = getOfficePreviewPage(pptxRes.cacheKey, 1, { dataDir: cacheDataDir });
  assert.equal(page1.ok, true);
  assert.match(page1.dataUrl, /^data:image\/(?:webp|png);base64,/);

  const t0 = Date.now();
  const pptxCached = await renderOfficePreview(join(fixturesDir, 'sample.pptx'), { mode: 'layout', dataDir: cacheDataDir });
  assert.equal(pptxCached.cacheKey, pptxRes.cacheKey);
  assert.ok(Date.now() - t0 < 200, 'PPTX 缓存命中应瞬时返回');

  // DOCX preview (layout + reader)
  const docxLayout = await renderOfficePreview(join(fixturesDir, 'sample.docx'), { mode: 'layout', dataDir: cacheDataDir });
  assert.equal(docxLayout.ok, true);
  assert.equal(docxLayout.kind, 'docx');
  assert.ok(docxLayout.pageCount >= 1);

  const docxReader = await renderOfficePreview(join(fixturesDir, 'sample.docx'), { mode: 'reader', dataDir: cacheDataDir });
  assert.equal(docxReader.ok, true);
  assert.equal(docxReader.mode, 'reader');
  assert.match(docxReader.html, /Agy Preview DOCX Test/);
  assert.match(docxReader.html, /<table/);
  assert.match(docxReader.html, /data:image\//);

  const statsBefore = getCacheStats(cacheDataDir);
  assert.ok(statsBefore.entries >= 3);
  const cleared = clearPreviewCache(cacheDataDir);
  assert.equal(cleared.entries, 0);
});

test('Aux Pane HTTP server enforces token auth and serves API endpoints', async () => {
  const auxDir = join(WORK_ROOT, 'aux-pane');
  const settings = createSettingsStore(auxDir);
  const broadcasted = [];
  const fakeCdp = {
    status: () => ({ connected: false, windowCount: 0, version: '1.0.0', lastError: '' }),
    broadcastSettings: async s => { broadcasted.push(s); }
  };
  const token = 'test-secret-token-123';
  const close = await startAuxPaneServer({ settings, cdpBridge: fakeCdp, dataDir: auxDir, port: 0, token });
  assert.ok(close.port > 0);
  const base = `http://127.0.0.1:${close.port}`;
  try {
    const unauth = await fetch(`${base}/api/status`);
    assert.equal(unauth.status, 403);

    const htmlPage = await fetch(`${base}/?token=${token}`);
    assert.equal(htmlPage.status, 200);
    assert.match(await htmlPage.text(), /Agy Preview/);

    const statusRes = await fetch(`${base}/api/status`, { headers: { 'x-sidecar-token': token } });
    assert.equal(statusRes.status, 200);
    const statusJson = await statusRes.json();
    assert.equal(statusJson.ok, true);
    assert.equal(statusJson.settings.enabled, true);

    const setRes = await fetch(`${base}/api/settings`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-sidecar-token': token },
      body: JSON.stringify({ chatLinkButtons: false })
    });
    assert.equal(setRes.status, 200);
    const setJson = await setRes.json();
    assert.equal(setJson.settings.chatLinkButtons, false);
    assert.equal(broadcasted.length, 1);

    const sampleHtml = join(WORK_ROOT, 'aux-sample.html');
    writeFileSync(sampleHtml, '<h1>Aux OK</h1>', 'utf8');
    const qpRes = await fetch(`${base}/api/quick-preview`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-sidecar-token': token },
      body: JSON.stringify({ path: sampleHtml })
    });
    assert.equal(qpRes.status, 200);
    const qpJson = await qpRes.json();
    assert.equal(qpJson.kind, 'html');
    assert.match(qpJson.html, /Aux OK/);

    const clearRes = await fetch(`${base}/api/cache/clear`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-sidecar-token': token },
      body: '{}'
    });
    assert.equal(clearRes.status, 200);
  } finally {
    await close();
  }
});

test('install and uninstall scripts preserve sibling plugins and copy office-preview.py', () => {
  const fakeHome = join(WORK_ROOT, 'fake-home');
  const cfgPath = join(fakeHome, '.gemini', 'config', 'config.json');
  mkdirSync(join(fakeHome, '.gemini', 'config'), { recursive: true });
  writeFileSync(cfgPath, JSON.stringify({
    theme: 'dark',
    plugins: {
      'agy-powertools': { enabled: true },
      'gemini-api': { enabled: true }
    }
  }, null, 2), 'utf8');

  const pkgFiles = collectPackageFiles();
  assert.ok(pkgFiles.includes('scripts/office-preview.py'));
  assert.ok(pkgFiles.includes('lib/client-inject.js'));
  assert.ok(pkgFiles.includes('sidecars/preview/sidecar.json'));

  const installed = installPlugin({ home: fakeHome, noLauncher: true });
  assert.equal(installed.pluginId, 'agy-preview');
  assert.ok(existsSync(join(installed.installedDir, 'plugin.json')));
  assert.ok(existsSync(join(installed.installedDir, 'scripts', 'office-preview.py')));
  assert.ok(existsSync(join(installed.installedDir, '.agy-preview-install.json')));

  const cfgAfterInstall = JSON.parse(readFileSync(cfgPath, 'utf8'));
  assert.equal(cfgAfterInstall.theme, 'dark');
  assert.equal(cfgAfterInstall.plugins['agy-powertools'].enabled, true);
  assert.equal(cfgAfterInstall.plugins['gemini-api'].enabled, true);
  assert.equal(cfgAfterInstall.plugins['agy-preview'].enabled, true);

  // Save a user setting and test uninstall with --keep-settings
  const dataDir = join(fakeHome, '.gemini', 'antigravity', 'plugin_data', 'agy-preview');
  createSettingsStore(dataDir).set({ htmlDefaultPreview: false });

  const unKeep = uninstallPlugin({ home: fakeHome, keepSettings: true });
  assert.equal(unKeep.configUpdated, true);
  assert.equal(existsSync(installed.installedDir), false);
  assert.equal(existsSync(join(dataDir, 'settings.json')), true);

  const cfgAfterUninstall = JSON.parse(readFileSync(cfgPath, 'utf8'));
  assert.equal(cfgAfterUninstall.plugins['agy-preview'], undefined);
  assert.equal(cfgAfterUninstall.plugins['agy-powertools'].enabled, true);
  assert.equal(cfgAfterUninstall.plugins['gemini-api'].enabled, true);

  // Full uninstall cleans settings too
  uninstallPlugin({ home: fakeHome, keepSettings: false });
  assert.equal(existsSync(dataDir), false);
});

test('browser DOM smoke test: in-place merged images, Side Question isolation, SVG-chevron HTML, full-opacity DOCX', async t => {
  const browsers = [
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
  ];
  const browserBin = browsers.find(p => existsSync(p));
  if (!browserBin) {
    t.skip('No local Chromium/Edge binary for DOM smoke test');
    return;
  }

  const clientCode = readFileSync(new URL('../lib/client-inject.js', import.meta.url), 'utf8');
  const tinyPng = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=';
  const iconSvg = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128"><rect width="128" height="128" fill="#a855f7"/></svg>');
  const pageHtml = `<!DOCTYPE html><html class="dark"><head><meta charset="utf-8"><style>
    body{margin:0;width:1280px;height:800px;display:flex;background:#141417;color:#f4f4f5}
    .dim-wrap{opacity:0.72}
    .dark img{filter:brightness(0.75);opacity:0.85}
  </style></head><body>
    <div data-testid="conversation-view" style="display:flex;width:1280px;height:800px">
      <div style="width:440px;height:800px">
        <div data-testid="agent-response" id="resp-pure">
          <div data-testid="tool-group-collapsible" id="grp-pure">
            <button aria-expanded="true"><span>Explored files</span><svg></svg></button>
            <div>
              <div style="display:flex;gap:6px"><span>Analyzed</span><svg></svg><span>f1_cover.png</span></div>
              <div style="display:flex;gap:6px"><span>Analyzed</span><svg></svg><span>f2_p11.png</span></div>
              <div style="display:flex;gap:6px"><span>Analyzed</span><svg></svg><span>f3_thanks.png</span></div>
            </div>
          </div>
          <div data-markdown-root="1">
            <p id="md-p1">导出图见 <img src="${iconSvg}" style="width:16px;height:16px" /><a href="file:///C:/tmp/map-export.png">map-export.png</a>。</p>
            <ul>
              <li id="md-li1">实拍验收：桌面端见 <span style="display:inline-flex;align-items:center"><img src="${iconSvg}" style="width:16px;height:16px" /><span data-uri="file:///C:/tmp/desktop.png">desktop.png</span></span>，手机端见 <img src="${iconSvg}" style="width:16px;height:16px" /><a href="file:///C:/tmp/mobile375.png">mobile375.png</a>，以及无 href 的纯芯片 <span>f3_thanks.png</span>。</li>
            </ul>
          </div>
        </div>
        <div data-testid="agent-response" id="resp-mixed">
          <div data-testid="tool-group-collapsible" id="grp-mixed">
            <button aria-expanded="true"><span>Explored files</span><svg></svg></button>
            <div>
              <div id="m-read1"><span>Read</span><span>app.py</span></div>
              <div id="m-img1"><span>Analyzed</span><span>f1_cover.png</span></div>
              <div id="m-img2"><span>Analyzed</span><span>f2_p11.png</span></div>
              <div id="m-read2"><span>Read</span><span>utils.py</span></div>
            </div>
          </div>
        </div>
      </div>
      <div id="side-question" style="width:260px;height:800px">
        <div>Side Question: test</div>
        <div data-testid="agent-response" id="sq-resp">Working</div>
      </div>
      <div id="html-pane" style="width:280px;height:400px;display:flex;flex-direction:column">
        <div id="html-bc" style="height:32px;display:flex;align-items:center;gap:4px">
          <span>demo-workspace</span><svg></svg><span>docs</span><svg></svg><span>slides</span><svg></svg><span>demo-report-2026.html</span>
        </div>
        <div id="html-body" style="flex:1;height:340px;width:280px;overflow:auto"><pre>&lt;!DOCTYPE html&gt;</pre></div>
      </div>
      <div id="docx-pane" class="dim-wrap" style="width:300px;height:400px;display:flex;flex-direction:column">
        <div id="docx-ph" style="margin:auto">
          <p>Preview is not available for this file type.</p>
          <a download="demo.docx" href="file:///C:/demo.docx">Download demo.docx</a>
        </div>
      </div>
    </div>
    <script>
      history.replaceState({}, '', '/c/demo-conversation-0001');
      globalThis.__agyPreviewVersion = '1.0.3';
      globalThis.__agyPreviewInitialSettings = ${JSON.stringify(DEFAULT_SETTINGS)};
      globalThis.__agyPreviewHost = raw => {
        const req = JSON.parse(raw);
        let reply = { ok: true };
        if (req.method === 'conversation.images') {
          reply = { ok: true, turns: [], images: [], byFileName: {
            'f1_cover.png': [{ path: 'C:/tmp/f1_cover.png', fileName: 'f1_cover.png', stepIndex: 1 }],
            'f2_p11.png': [{ path: 'C:/tmp/f2_p11.png', fileName: 'f2_p11.png', stepIndex: 2 }],
            'f3_thanks.png': [{ path: 'C:/tmp/f3_thanks.png', fileName: 'f3_thanks.png', stepIndex: 3 }]
          }};
        } else if (req.method === 'preview.image') {
          reply = { ok: true, kind: 'image', fileName: 'x.png', dataUrl: ${JSON.stringify(tinyPng)} };
        } else if (req.method === 'preview.html') {
          reply = { ok: true, kind: 'html', filePath: 'D:/demo.html', fileName: req.fileName, html: '<h1>Rendered HTML OK</h1>' };
        } else if (req.method === 'preview.office') {
          reply = { ok: true, kind: 'docx', mode: 'layout', engine: 'word', filePath: 'C:/demo.docx', fileName: 'demo.docx', cacheKey: '0123456789abcdef', pageCount: 1, pages: [{ index: 1, width: 800, height: 1000, title: '', notes: '' }] };
        } else if (req.method === 'preview.officePage') {
          reply = { ok: true, index: 1, width: 800, height: 1000, dataUrl: ${JSON.stringify(tinyPng)} };
        }
        setTimeout(() => globalThis.__agyPreviewReceive(req.id, reply), 10);
      };
    </script>
    <script>${clientCode}</script>
    <script>
      setTimeout(() => {
        const pureGrp = document.getElementById('grp-pure');
        const pureStrip = pureGrp.nextElementSibling;
        const mRead1 = document.getElementById('m-read1');
        const mImg1 = document.getElementById('m-img1');
        const mImg2 = document.getElementById('m-img2');
        const mRead2 = document.getElementById('m-read2');
        const mixedStrip = mImg2.nextElementSibling;
        const sq = document.getElementById('sq-resp');
        const mdP1 = document.getElementById('md-p1');
        const mdP1Gallery = mdP1.nextElementSibling;
        const mdLi1 = document.getElementById('md-li1');
        const mdLi1Gallery = mdLi1.querySelector('[data-agy-preview-root="md-images"]');
        const mdCardImg = mdLi1Gallery?.querySelector('.agy-pv-md-card img');
        const htmlBar = document.querySelector('#html-body [data-agy-preview-root="html-bar"]');
        const htmlFrame = document.querySelector('#html-body [data-agy-preview-root="html-frame"] iframe');
        const docxPane = document.getElementById('docx-pane');
        const docxImg = document.querySelector('#docx-pane .agy-pv-page-frame img');
        const report = {
          pureHidden: pureGrp.getAttribute('data-agy-pv-img-hidden'),
          pureStripText: pureStrip?.querySelector('.agy-pv-img-head')?.textContent,
          pureThumbCount: pureStrip?.querySelectorAll('.agy-pv-thumb-btn')?.length,
          mixedRead1Hidden: mRead1.getAttribute('data-agy-pv-img-hidden'),
          mixedImg1Hidden: mImg1.getAttribute('data-agy-pv-img-hidden'),
          mixedImg2Hidden: mImg2.getAttribute('data-agy-pv-img-hidden'),
          mixedRead2Hidden: mRead2.getAttribute('data-agy-pv-img-hidden'),
          mixedStripText: mixedStrip?.querySelector('.agy-pv-img-head')?.textContent,
          mdP1CardCount: mdP1Gallery?.getAttribute('data-agy-preview-root') === 'md-images' ? mdP1Gallery.querySelectorAll('.agy-pv-md-card').length : 0,
          mdLi1CardCount: mdLi1Gallery ? mdLi1Gallery.querySelectorAll('.agy-pv-md-card').length : 0,
          mdCardImgOpacity: mdCardImg ? getComputedStyle(mdCardImg).opacity : 'none',
          mdCardImgFilter: mdCardImg ? getComputedStyle(mdCardImg).filter : 'none',
          sideQuestionUntouched: sq.children.length === 0 && sq.textContent.trim() === 'Working',
          htmlBarMounted: Boolean(htmlBar),
          htmlSrcdocOk: /Rendered HTML OK/.test(htmlFrame?.srcdoc || ''),
          docxHostOpacity: getComputedStyle(docxPane).opacity,
          docxImgOpacity: docxImg ? getComputedStyle(docxImg).opacity : 'none',
          docxImgFilter: docxImg ? getComputedStyle(docxImg).filter : 'none'
        };
        const out = document.createElement('pre');
        out.id = 'VERIFY_REPORT';
        out.textContent = JSON.stringify(report);
        document.body.append(out);
      }, 900);
    </script>
  </body></html>`;

  const srv = createServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(pageHtml);
  });
  await new Promise(resolve => srv.listen(0, '127.0.0.1', resolve));
  const port = srv.address().port;
  const edgeProfile = join(WORK_ROOT, 'edge-headless-profile');
  mkdirSync(edgeProfile, { recursive: true });
  try {
    const r = await new Promise((resolveRun, rejectRun) => {
      const child = spawn(
        browserBin,
        [
          '--headless=new',
          '--disable-gpu',
          '--window-size=1280,800',
          '--no-first-run',
          '--no-default-browser-check',
          `--user-data-dir=${edgeProfile}`,
          '--virtual-time-budget=2500',
          '--dump-dom',
          `http://127.0.0.1:${port}/`
        ],
        { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] }
      );
      let stdout = '';
      let stderr = '';
      const timer = setTimeout(() => {
        try { child.kill(); } catch {}
        resolveRun({ stdout, stderr, timedOut: true });
      }, 15000);
      child.stdout.on('data', d => { stdout += d.toString('utf8'); });
      child.stderr.on('data', d => { stderr += d.toString('utf8'); });
      child.on('error', err => { clearTimeout(timer); rejectRun(err); });
      child.on('close', code => { clearTimeout(timer); resolveRun({ code, stdout, stderr }); });
    });
    const match = (r.stdout || '').match(/<pre id="VERIFY_REPORT">([\s\S]*?)<\/pre>/);
    assert.ok(match, `未获取到 Chromium 验证结果: ${(r.stderr || '').slice(0, 300)}`);
    const report = JSON.parse(match[1]);

    // 1) Pure image Explored files group replaced in-place by merged "已查看 3 张图像"
    assert.equal(report.pureHidden, '1');
    assert.match(report.pureStripText, /已查看 3 张图像/);
    assert.equal(report.pureThumbCount, 3);

    // 2) Mixed group preserves Read rows and merges only the 2 consecutive Analyzed image rows
    assert.equal(report.mixedRead1Hidden, null);
    assert.equal(report.mixedImg1Hidden, '1');
    assert.equal(report.mixedImg2Hidden, '1');
    assert.equal(report.mixedRead2Hidden, null);
    assert.match(report.mixedStripText, /已查看 2 张图像/);

    // 2b) Markdown reply image links render side-by-side rounded cards even with 16px 128-naturalWidth file-badge <img> icons and plain <span> chips
    assert.equal(report.mdP1CardCount, 1);
    assert.equal(report.mdLi1CardCount, 3);
    assert.equal(report.mdCardImgOpacity, '1');
    assert.equal(report.mdCardImgFilter, 'none');

    // 3) Side Question streaming response completely untouched
    assert.equal(report.sideQuestionUntouched, true);

    // 4) HTML viewer with SVG chevron breadcrumb mounts toggle bar + renders live HTML iframe
    assert.equal(report.htmlBarMounted, true);
    assert.equal(report.htmlSrcdocOk, true);

    // 5) Dimmed DOCX wrapper and dark-mode img filter are fully overridden to bright 1.0 opacity / none filter
    assert.equal(report.docxHostOpacity, '1');
    assert.equal(report.docxImgOpacity, '1');
    assert.equal(report.docxImgFilter, 'none');
  } finally {
    await new Promise(resolve => srv.close(resolve));
  }
});

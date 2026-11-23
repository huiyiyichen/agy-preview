import { createHash } from 'node:crypto';
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync
} from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, relative, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { PLUGIN_ID, pluginPaths } from '../lib/paths.mjs';

const REPO_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const INSTALL_RECEIPT = '.agy-preview-install.json';

const ROOT_FILES = [
  'plugin.json',
  'package.json',
  'LICENSE',
  'README.md',
  'README_EN.md',
  'install.ps1',
  'uninstall.ps1',
  'install.sh',
  'uninstall.sh'
];

const INCLUDE_DIRS = [
  'assets',
  'lib',
  'sidecars',
  'scripts'
];

function parseArgs(argv) {
  let home = homedir();
  let noLauncher = false;
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--home' && argv[i + 1]) {
      home = resolve(argv[++i]);
    } else if (arg.startsWith('--home=')) {
      home = resolve(arg.slice('--home='.length));
    } else if (arg === '--no-launcher') {
      noLauncher = true;
    }
  }
  return { home, noLauncher };
}

export function collectPackageFiles(sourceRoot = REPO_ROOT) {
  const files = [];
  for (const name of ROOT_FILES) {
    const full = join(sourceRoot, name);
    if (existsSync(full)) files.push(name);
  }
  const walk = relDir => {
    const absDir = join(sourceRoot, relDir);
    if (!existsSync(absDir)) return;
    for (const entry of readdirSync(absDir, { withFileTypes: true })) {
      if (entry.name === '__pycache__' || entry.name.startsWith('.')) continue;
      const rel = `${relDir}/${entry.name}`;
      if (entry.isDirectory()) walk(rel);
      else if (entry.isFile()) files.push(rel);
    }
  };
  for (const dir of INCLUDE_DIRS) walk(dir);
  files.sort();
  return files;
}

function sha256File(filePath) {
  return createHash('sha256').update(readFileSync(filePath)).digest('hex');
}

function writeJsonAtomic(filePath, data) {
  mkdirSync(dirname(filePath), { recursive: true });
  const tmp = `${filePath}.${process.pid}.tmp`;
  writeFileSync(tmp, JSON.stringify(data, null, 2) + '\n', 'utf8');
  renameSync(tmp, filePath);
}

function registerPluginConfig(configFile) {
  let config = {};
  if (existsSync(configFile)) {
    const raw = readFileSync(configFile, 'utf8').trim();
    if (raw) config = JSON.parse(raw);
  }
  if (!config || typeof config !== 'object' || Array.isArray(config)) {
    throw new Error(`配置文件格式异常: ${configFile}`);
  }
  if (!config.plugins || typeof config.plugins !== 'object' || Array.isArray(config.plugins)) {
    config.plugins = {};
  }
  const prev = config.plugins[PLUGIN_ID];
  config.plugins[PLUGIN_ID] = {
    ...(prev && typeof prev === 'object' ? prev : {}),
    enabled: true
  };
  writeJsonAtomic(configFile, config);
}

function ensureWindowsShortcuts(dataDir) {
  if (process.platform !== 'win32') return { modified: [], alreadyReady: 0 };
  const psScript = `
$ErrorActionPreference = 'Stop'
$shell = New-Object -ComObject WScript.Shell
$roots = @(
  [Environment]::GetFolderPath('Desktop'),
  [Environment]::GetFolderPath('CommonDesktopDirectory'),
  (Join-Path $env:APPDATA 'Microsoft\\Windows\\Start Menu\\Programs'),
  (Join-Path $env:ProgramData 'Microsoft\\Windows\\Start Menu\\Programs')
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
$modified = @()
$already = 0
foreach ($root in $roots) {
  Get-ChildItem -Path $root -Filter '*.lnk' -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
    try {
      $sc = $shell.CreateShortcut($_.FullName)
      if ($sc.TargetPath -match 'Antigravity\\.exe$') {
        $args = [string]$sc.Arguments
        if ($args -match '--remote-debugging-port=') {
          $already++
        } else {
          $sc.Arguments = (($args.Trim()) + ' --remote-debugging-port=0').Trim()
          $sc.Save()
          $modified += $_.FullName
        }
      }
    } catch {}
  }
}
@{ modified = $modified; alreadyReady = $already } | ConvertTo-Json -Compress
`;
  const res = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', psScript], {
    encoding: 'utf8',
    windowsHide: true,
    timeout: 15000
  });
  if (res.error || res.status !== 0 || !res.stdout.trim()) {
    return { modified: [], alreadyReady: 0 };
  }
  try {
    const parsed = JSON.parse(res.stdout.trim());
    const modified = Array.isArray(parsed.modified) ? parsed.modified : (parsed.modified ? [parsed.modified] : []);
    if (modified.length > 0) {
      const stateFile = join(dataDir, 'launcher-state.json');
      let existing = [];
      if (existsSync(stateFile)) {
        try {
          const prev = JSON.parse(readFileSync(stateFile, 'utf8'));
          if (Array.isArray(prev.modifiedShortcuts)) existing = prev.modifiedShortcuts;
        } catch {}
      }
      writeJsonAtomic(stateFile, {
        flag: '--remote-debugging-port=0',
        modifiedShortcuts: [...new Set([...existing, ...modified])],
        updatedAt: new Date().toISOString()
      });
    }
    return { modified, alreadyReady: Number(parsed.alreadyReady || 0) };
  } catch {
    return { modified: [], alreadyReady: 0 };
  }
}

export function installPlugin({ sourceRoot = REPO_ROOT, home = homedir(), noLauncher = false } = {}) {
  const paths = pluginPaths(home);
  const pkg = JSON.parse(readFileSync(join(sourceRoot, 'package.json'), 'utf8'));
  const files = collectPackageFiles(sourceRoot);
  if (!files.includes('scripts/office-preview.py')) {
    throw new Error('打包清单缺失 scripts/office-preview.py');
  }

  const resolvedSource = resolve(sourceRoot);
  const resolvedDest = resolve(paths.installedDir);
  const sameDir = resolvedSource.toLowerCase() === resolvedDest.toLowerCase();

  const hashes = {};
  if (!sameDir) {
    if (existsSync(resolvedDest)) rmSync(resolvedDest, { recursive: true, force: true });
    mkdirSync(resolvedDest, { recursive: true });
    for (const rel of files) {
      const src = join(resolvedSource, rel);
      const dst = join(resolvedDest, rel);
      mkdirSync(dirname(dst), { recursive: true });
      copyFileSync(src, dst);
      hashes[rel] = sha256File(dst);
    }
  } else {
    for (const rel of files) {
      hashes[rel] = sha256File(join(resolvedDest, rel));
    }
  }

  const receipt = {
    pluginId: PLUGIN_ID,
    version: pkg.version,
    installedAt: new Date().toISOString(),
    files: hashes
  };
  writeJsonAtomic(join(resolvedDest, INSTALL_RECEIPT), receipt);
  registerPluginConfig(paths.configFile);

  let launcher = { modified: [], alreadyReady: 0 };
  const isCustomHome = resolve(home).toLowerCase() !== resolve(homedir()).toLowerCase();
  if (!noLauncher && !isCustomHome && process.platform === 'win32') {
    launcher = ensureWindowsShortcuts(paths.dataDir);
  }

  return {
    pluginId: PLUGIN_ID,
    version: pkg.version,
    installedDir: resolvedDest,
    configFile: paths.configFile,
    fileCount: files.length,
    launcher
  };
}

const isCli = process.argv[1] && resolve(process.argv[1]) === resolve(fileURLToPath(import.meta.url));
if (isCli) {
  const opts = parseArgs(process.argv.slice(2));
  const result = installPlugin(opts);
  console.log(`[agy-preview] 已安装 v${result.version} (${result.fileCount} 个文件)`);
  console.log(`[agy-preview] 插件目录: ${result.installedDir}`);
  console.log(`[agy-preview] 配置文件: ${result.configFile}`);
  if (result.launcher.modified.length > 0) {
    console.log(`[agy-preview] 已为 ${result.launcher.modified.length} 个 Antigravity 快捷方式添加 --remote-debugging-port=0`);
  } else if (result.launcher.alreadyReady > 0) {
    console.log(`[agy-preview] 检测到 ${result.launcher.alreadyReady} 个快捷方式已具备调试端口参数，保持不变`);
  }
  console.log('[agy-preview] 请完全退出并重新启动 Antigravity，即可在侧栏与文件预览中使用。');
}

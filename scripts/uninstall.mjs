import {
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync
} from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { PLUGIN_ID, pluginPaths } from '../lib/paths.mjs';

function parseArgs(argv) {
  let home = homedir();
  let keepSettings = false;
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--home' && argv[i + 1]) {
      home = resolve(argv[++i]);
    } else if (arg.startsWith('--home=')) {
      home = resolve(arg.slice('--home='.length));
    } else if (arg === '--keep-settings') {
      keepSettings = true;
    }
  }
  return { home, keepSettings };
}

function writeJsonAtomic(filePath, data) {
  mkdirSync(dirname(filePath), { recursive: true });
  const tmp = `${filePath}.${process.pid}.tmp`;
  writeFileSync(tmp, JSON.stringify(data, null, 2) + '\n', 'utf8');
  renameSync(tmp, filePath);
}

function unregisterPluginConfig(configFile) {
  if (!existsSync(configFile)) return false;
  const raw = readFileSync(configFile, 'utf8').trim();
  if (!raw) return false;
  const config = JSON.parse(raw);
  if (!config?.plugins || typeof config.plugins !== 'object') return false;
  if (!Object.hasOwn(config.plugins, PLUGIN_ID)) return false;
  delete config.plugins[PLUGIN_ID];
  writeJsonAtomic(configFile, config);
  return true;
}

function revertWindowsShortcuts(dataDir) {
  if (process.platform !== 'win32') return 0;
  const stateFile = join(dataDir, 'launcher-state.json');
  if (!existsSync(stateFile)) return 0;
  let state;
  try { state = JSON.parse(readFileSync(stateFile, 'utf8')); } catch { return 0; }
  const shortcuts = Array.isArray(state.modifiedShortcuts) ? state.modifiedShortcuts : [];
  if (!shortcuts.length) return 0;

  const listJson = JSON.stringify(shortcuts);
  const psScript = `
$ErrorActionPreference = 'SilentlyContinue'
$shell = New-Object -ComObject WScript.Shell
$paths = ConvertFrom-Json ${JSON.stringify(listJson)}
$reverted = 0
foreach ($p in $paths) {
  if ($p -and (Test-Path $p)) {
    $sc = $shell.CreateShortcut($p)
    $args = [string]$sc.Arguments
    if ($args -match '--remote-debugging-port=0') {
      $sc.Arguments = ($args -replace '(?:^|\\s+)--remote-debugging-port=0(?=\\s|$)', '').Trim()
      $sc.Save()
      $reverted++
    }
  }
}
Write-Output $reverted
`;
  const res = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', psScript], {
    encoding: 'utf8',
    windowsHide: true,
    timeout: 15000
  });
  return Number(res.stdout?.trim() || 0) || 0;
}

export function uninstallPlugin({ home = homedir(), keepSettings = false } = {}) {
  const paths = pluginPaths(home);
  const revertedShortcuts = revertWindowsShortcuts(paths.dataDir);
  const configUpdated = unregisterPluginConfig(paths.configFile);

  if (existsSync(paths.installedDir)) {
    rmSync(paths.installedDir, { recursive: true, force: true });
  }
  if (existsSync(paths.sidecarDataDir)) {
    rmSync(paths.sidecarDataDir, { recursive: true, force: true });
  }
  if (existsSync(paths.dataDir)) {
    if (keepSettings) {
      const settingsPath = join(paths.dataDir, 'settings.json');
      const saved = existsSync(settingsPath) ? readFileSync(settingsPath, 'utf8') : null;
      rmSync(paths.dataDir, { recursive: true, force: true });
      if (saved !== null) {
        mkdirSync(paths.dataDir, { recursive: true });
        writeFileSync(settingsPath, saved, 'utf8');
      }
    } else {
      rmSync(paths.dataDir, { recursive: true, force: true });
    }
  }

  return {
    pluginId: PLUGIN_ID,
    removedDir: paths.installedDir,
    configUpdated,
    revertedShortcuts,
    keepSettings
  };
}

const isCli = process.argv[1] && resolve(process.argv[1]) === resolve(fileURLToPath(import.meta.url));
if (isCli) {
  const opts = parseArgs(process.argv.slice(2));
  const res = uninstallPlugin(opts);
  console.log(`[agy-preview] 已卸载插件目录: ${res.removedDir}`);
  if (res.configUpdated) console.log('[agy-preview] 已从 ~/.gemini/config/config.json 移除注册项（未影响其他插件）');
  if (res.revertedShortcuts > 0) console.log(`[agy-preview] 已还原 ${res.revertedShortcuts} 个由本插件修改的快捷方式参数`);
  if (res.keepSettings) console.log('[agy-preview] 已保留用户设置文件 settings.json');
  console.log('[agy-preview] 重启 Antigravity 后卸载生效。');
}

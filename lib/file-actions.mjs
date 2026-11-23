import { existsSync, statSync } from 'node:fs';
import { dirname, extname, isAbsolute, resolve } from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const WINDOWS_RUNNABLE = new Set([
  '.exe', '.com', '.bat', '.cmd', '.ps1', '.psm1', '.psd1',
  '.vbs', '.vbe', '.js', '.jse', '.wsf', '.wsh', '.hta',
  '.msi', '.msp', '.scr', '.pif', '.reg', '.msc', '.cpl', '.lnk', '.url', '.jar'
]);

export function isUncPath(value) {
  return /^\\\\|^\/\//.test(String(value || '').trim());
}

export function resolveLocalPath(input) {
  const raw = String(input || '').trim();
  if (!raw) throw new Error('文件路径无效');
  if (/^file:/i.test(raw)) {
    let parsed;
    try { parsed = new URL(raw); } catch { throw new Error('文件路径格式无效'); }
    if (parsed.protocol !== 'file:') throw new Error('仅支持本地文件链接');
    if (parsed.hostname && parsed.hostname.toLowerCase() !== 'localhost') throw new Error('仅支持本机文件路径');
    parsed.hash = '';
    parsed.search = '';
    return fileURLToPath(parsed);
  }
  if (isUncPath(raw)) throw new Error('仅支持本机文件路径');
  if (!isAbsolute(raw)) throw new Error('需要本机绝对路径');
  return resolve(raw);
}

export function openLocalPath(input, { action = 'open', platform = process.platform } = {}) {
  if (action !== 'open' && action !== 'reveal') throw new Error('不支持的文件操作');
  const filePath = resolveLocalPath(input);
  if (!existsSync(filePath)) throw new Error('文件或文件夹不存在');
  let isDir = false;
  try { isDir = statSync(filePath).isDirectory(); } catch {}
  if (action === 'open' && platform === 'win32' && !isDir && WINDOWS_RUNNABLE.has(extname(filePath).toLowerCase())) {
    throw new Error('该文件类型为 Windows 可执行脚本或程序，已拦截直接打开；请使用“打开所在文件夹”');
  }
  let cmd;
  let args;
  const opts = { detached: true, stdio: 'ignore' };
  if (platform === 'win32') {
    cmd = 'explorer.exe';
    if (action === 'reveal') {
      if (filePath.includes(',')) {
        args = [`/select,"${filePath}"`];
        opts.windowsVerbatimArguments = true;
      } else {
        args = [`/select,${filePath}`];
      }
    } else {
      args = [filePath];
    }
  } else if (platform === 'darwin') {
    cmd = 'open';
    args = action === 'reveal' ? ['-R', filePath] : [filePath];
  } else {
    cmd = 'xdg-open';
    args = [action === 'reveal' ? (isDir ? filePath : dirname(filePath)) : filePath];
  }
  const child = spawn(cmd, args, opts);
  child.unref?.();
  return { ok: true, path: filePath };
}

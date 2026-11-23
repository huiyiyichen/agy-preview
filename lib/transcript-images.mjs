import { existsSync, readFileSync, statSync } from 'node:fs';
import { basename, isAbsolute, join, resolve } from 'node:path';
import { isUncPath, resolveLocalPath } from './file-actions.mjs';

const IMAGE_RE = /\.(?:png|jpe?g|webp|gif|bmp|svg)$/i;
const TRACKED_EXT_RE = /\.(?:docx?|pptx?|html?|png|jpe?g|webp|gif|bmp|svg)$/i;
const VALID_ID = /^[a-zA-Z0-9][a-zA-Z0-9_-]{0,127}$/;
const parseCache = new Map();

export function cleanPath(candidate) {
  if (typeof candidate !== 'string' || !candidate.trim()) return '';
  let raw = candidate.trim();
  // Antigravity transcript.jsonl often stores tool_calls args values as JSON-encoded strings
  // e.g. "\"C:\\\\Users\\\\...\\\\f1_cover.png\""
  if ((raw.startsWith('"') && raw.endsWith('"')) || (raw.startsWith("'") && raw.endsWith("'"))) {
    if (raw.startsWith('"') && raw.endsWith('"')) {
      try {
        const parsed = JSON.parse(raw);
        if (typeof parsed === 'string') raw = parsed.trim();
      } catch {
        raw = raw.slice(1, -1).trim();
      }
    } else {
      raw = raw.slice(1, -1).trim();
    }
  }
  raw = raw.split('#')[0].split('?')[0].trim();
  if (!raw) return '';
  if (/^file:/i.test(raw)) {
    try { raw = resolveLocalPath(raw); } catch { return ''; }
  }
  if (!isAbsolute(raw) || isUncPath(raw)) return '';
  return resolve(raw);
}

function extractInlinePathsFromText(text) {
  if (typeof text !== 'string' || !text) return [];
  const out = [];
  for (const m of text.matchAll(/file:\/\/\/[^\s)"'<>]+/g)) {
    const p = cleanPath(m[0]);
    if (p) out.push(p);
  }
  // Also pick up Windows or POSIX absolute paths to previewable files in tool args/content
  for (const m of text.matchAll(/[a-zA-Z]:[\\/][^"\r\n<>|*?\t]+?\.(?:docx?|pptx?|html?|png|jpe?g|webp|gif|bmp|svg)(?=[\s"'`)\],;]|$)/gi)) {
    const p = cleanPath(m[0]);
    if (p) out.push(p);
  }
  return out;
}

export function scanConversationFiles(baseAppDataDir, conversationId) {
  if (!baseAppDataDir || !VALID_ID.test(conversationId || '')) {
    return { turns: [], images: [], byFileName: {}, allPaths: [] };
  }
  const logPath = join(baseAppDataDir, 'brain', conversationId, '.system_generated', 'logs', 'transcript.jsonl');
  if (!existsSync(logPath)) return { turns: [], images: [], byFileName: {}, allPaths: [] };

  let stamp = '';
  try {
    const st = statSync(logPath);
    stamp = `${st.mtimeMs}:${st.size}`;
  } catch {
    return { turns: [], images: [], byFileName: {}, allPaths: [] };
  }
  const cached = parseCache.get(logPath);
  if (cached?.stamp === stamp) return cached.value;

  const turns = [];
  const images = [];
  const byFileName = {};
  const allPaths = [];
  const allSeen = new Set();
  const pushAll = (p, stepIndex = -1) => {
    if (!p) return;
    const k = p.toLowerCase();
    if (!allSeen.has(k)) { allSeen.add(k); allPaths.push(p); }
    if (IMAGE_RE.test(p)) {
      const fileName = basename(p);
      const lowName = fileName.toLowerCase();
      if (!byFileName[lowName]) byFileName[lowName] = [];
      if (!byFileName[lowName].some(x => x.path.toLowerCase() === k)) {
        byFileName[lowName].push({ stepIndex, path: p, fileName });
      }
    }
  };

  let currentTurn = null;
  let raw = '';
  try { raw = readFileSync(logPath, 'utf8'); } catch { return { turns: [], images: [], byFileName: {}, allPaths: [] }; }

  for (const line of raw.split('\n')) {
    if (!line) continue;
    let step;
    try { step = JSON.parse(line); } catch { continue; }
    if (step.type === 'USER_INPUT') {
      if (/^Error: The stream was interrupted/i.test(step.content || '')) continue;
      currentTurn = { turnNumber: turns.length + 1, stepIndexStart: step.step_index, stepIndexEnd: step.step_index, images: [] };
      turns.push(currentTurn);
      continue;
    }
    if (currentTurn && Number.isInteger(step.step_index)) {
      currentTurn.stepIndexEnd = step.step_index;
    }
    const stepIdx = Number.isInteger(step.step_index) ? step.step_index : -1;
    if (typeof step.content === 'string') {
      for (const p of extractInlinePathsFromText(step.content)) pushAll(p, stepIdx);
    }
    if (Array.isArray(step.tool_calls)) {
      for (let c = 0; c < step.tool_calls.length; c++) {
        const call = step.tool_calls[c];
        const name = String(call?.name || '');
        const args = call?.args || {};
        for (const key of ['Path', 'AbsolutePath', 'TargetFile', 'DirectoryPath', 'path', 'file_path', 'uri']) {
          const val = args[key];
          if (typeof val !== 'string') continue;
          const p = cleanPath(val);
          pushAll(p);
          if (/read_file|view_file|view_image/i.test(name) && p && IMAGE_RE.test(p)) {
            const entry = {
              stepIndex: Number.isInteger(step.step_index) ? step.step_index : -1,
              path: p,
              fileName: basename(p)
            };
            images.push(entry);
            const lowName = entry.fileName.toLowerCase();
            if (!byFileName[lowName]) byFileName[lowName] = [];
            if (!byFileName[lowName].some(x => x.path.toLowerCase() === p.toLowerCase() && x.stepIndex === entry.stepIndex)) {
              byFileName[lowName].push(entry);
            }
            if (currentTurn && !currentTurn.images.some(img => img.path.toLowerCase() === p.toLowerCase())) {
              currentTurn.images.push(entry);
            }
          }
        }
        // Extract previewable paths mentioned in command lines (e.g., scripts generating .docx/.pptx/.html)
        if (typeof args.CommandLine === 'string') {
          let cmdText = args.CommandLine;
          if (cmdText.startsWith('"') && cmdText.endsWith('"')) {
            try { cmdText = JSON.parse(cmdText); } catch {}
          }
          for (const p of extractInlinePathsFromText(cmdText)) {
            if (TRACKED_EXT_RE.test(p)) pushAll(p);
          }
        }
      }
    }
  }

  const value = { turns, images, byFileName, allPaths };
  parseCache.set(logPath, { stamp, value });
  if (parseCache.size > 24) parseCache.delete(parseCache.keys().next().value);
  return value;
}

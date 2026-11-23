import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { pluginPaths } from './paths.mjs';

export const DEFAULT_SETTINGS = Object.freeze({
  enabled: true,
  officePreview: true,
  htmlPreview: true,
  htmlDefaultPreview: true,
  inlineImages: true,
  markdownImages: true,
  chatLinkButtons: true,
  language: 'zh-CN'
});

function isValid(key, value) {
  if (key === 'language') return value === 'zh-CN' || value === 'en';
  return typeof value === 'boolean';
}

export function normalizeSettings(input = {}) {
  const out = {};
  for (const [key, fallback] of Object.entries(DEFAULT_SETTINGS)) {
    out[key] = isValid(key, input?.[key]) ? input[key] : fallback;
  }
  return out;
}

export function createSettingsStore(directory = pluginPaths().dataDir) {
  const filePath = join(directory, 'settings.json');
  return {
    filePath,
    get() {
      try {
        if (!existsSync(filePath)) return { ...DEFAULT_SETTINGS };
        return normalizeSettings(JSON.parse(readFileSync(filePath, 'utf8')));
      } catch {
        return { ...DEFAULT_SETTINGS };
      }
    },
    set(patch) {
      if (!patch || typeof patch !== 'object' || Array.isArray(patch)) {
        throw new Error('设置格式无效');
      }
      for (const [key, value] of Object.entries(patch)) {
        if (!Object.hasOwn(DEFAULT_SETTINGS, key) || !isValid(key, value)) {
          throw new Error(`无效的设置项: ${key}`);
        }
      }
      const next = normalizeSettings({ ...this.get(), ...patch });
      mkdirSync(directory, { recursive: true });
      const tmp = `${filePath}.tmp`;
      writeFileSync(tmp, JSON.stringify(next, null, 2) + '\n', 'utf8');
      renameSync(tmp, filePath);
      return next;
    }
  };
}

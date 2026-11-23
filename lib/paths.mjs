import { homedir } from 'node:os';
import { join, resolve } from 'node:path';

export const PLUGIN_ID = 'agy-preview';

export function antigravityUserDataDir({ platform = process.platform, env = process.env, home = homedir() } = {}) {
  if (env.ANTIGRAVITY_USER_DATA_DIR) return env.ANTIGRAVITY_USER_DATA_DIR;
  if (platform === 'win32') return join(env.APPDATA || join(home, 'AppData', 'Roaming'), 'Antigravity');
  if (platform === 'darwin') return join(home, 'Library', 'Application Support', 'Antigravity');
  return join(env.XDG_CONFIG_HOME || join(home, '.config'), 'Antigravity');
}

export function devToolsPortFile(options = {}) {
  const env = options.env || process.env;
  return options.portFile || env.ANTIGRAVITY_CDP_PORT_FILE || env.AGY_BROWSER_ACTIVE_PORT_FILE ||
    join(antigravityUserDataDir(options), 'DevToolsActivePort');
}

export function appDataDir(env = process.env, home = homedir()) {
  return env.ANTIGRAVITY_APP_DATA_DIR || join(resolve(home), '.gemini', 'antigravity');
}

export function pluginPaths(home = homedir()) {
  const root = join(resolve(home), '.gemini');
  return {
    home: resolve(home),
    root,
    configFile: join(root, 'config', 'config.json'),
    installedDir: join(root, 'config', 'plugins', PLUGIN_ID),
    dataDir: join(root, 'antigravity', 'plugin_data', PLUGIN_ID),
    sidecarDataDir: join(root, 'antigravity', 'sidecar_data', PLUGIN_ID),
    brainDir: join(root, 'antigravity', 'brain')
  };
}

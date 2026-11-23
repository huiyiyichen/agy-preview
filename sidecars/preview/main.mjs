import { createSettingsStore } from '../../lib/settings.mjs';
import { pluginPaths } from '../../lib/paths.mjs';
import { startCdpBridge } from '../../lib/cdp-host.mjs';
import { startAuxPaneServer } from '../../lib/aux-pane-ui.mjs';

const { dataDir } = pluginPaths();
const settings = createSettingsStore(dataDir);
const cdpBridge = startCdpBridge({
  settings,
  dataDir,
  onLog: msg => console.log(`[agy-preview] ${msg}`)
});

const closeWebUi = await startAuxPaneServer({ settings, cdpBridge, dataDir }).catch(err => {
  console.error('[agy-preview] Aux Pane Web UI 启动失败:', err.message);
  return async () => {};
});

async function shutdown() {
  await cdpBridge.stop();
  await closeWebUi();
  process.exit(0);
}

process.on('SIGINT', () => { void shutdown(); });
process.on('SIGTERM', () => { void shutdown(); });

# Agy Preview

[English](./README_EN.md) | 简体中文

**Agy Preview** 是一款面向 **Google Antigravity 2.0** 的独立开源预览增强插件（基于官方 Sidecar + 非侵入式本地桥接架构，不修改 `app.asar`）。

它专注解决 Antigravity 桌面端查看本地产物时的三大痛点：

1. **DOCX / PPTX 高保真原生预览**
   - 直接接管右侧文件标签页中的 `Preview is not available for this file type.` 空白占位。
   - **PPTX / PPT**：静默调用本机 Microsoft PowerPoint COM（或 LibreOffice 无头回退）转 PDF 并由 PyMuPDF 渲染为高清幻灯片；自动提取 **演讲者备注（Speaker Notes）**，支持 **连续滚动 / 单页幻灯（左右方向键翻页）**、缩放、点击全屏灯箱与本地 LRU 缓存秒开。
   - **DOCX / DOC**：支持 **原版排版**（Word COM / LibreOffice 高保真页面）与 **流式阅读**（`python-docx` 自适应深浅色 HTML，文字可选中复制、支持表格与内嵌图片）一键切换；未安装 Office/LibreOffice 时自动降级至流式阅读。
2. **HTML 网页实时渲染**
   - 打开 `.html` / `.htm` 文件时，在文件视图右上角悬浮提供 **`[ 网页预览 | 源码 ]`** 一键切换栏，并安全内联同目录相对路径 CSS、JS 与图片到沙箱 `<iframe srcdoc>`。
   - 可配置打开 HTML 时默认进入网页预览或源码视图，并记住单个文件的手动切换状态。
3. **类 Codex 对话内图片双场景直出与灯箱**
   - **过程看图原位合并**：将工具过程里连续的 `Analyzed xxx.png` 行原位合并替换为折叠式 **`已查看 N 张图像`** 圆角缩略图条（混合组中仅替换连续图片行，保留代码读取记录）。
   - **回复正文图片直接渲染**：当 AI 在回复正文段落或列表中引用本地图片链接（如验收截图）时，直接在对应段落下方横向渲染圆角图片卡片墙，无需点击跳转到右侧才能查看。
   - 点击任意图片均可打开全屏灯箱（支持 `Esc` 关闭、左右切换、`100%` 原图缩放、系统打开与定位所在文件夹）。
4. **官方 Sidecar 侧栏控制台（Aux Pane）**
   - 注册官方 `SIDECAR_UI_ENTRYPOINT_AUX_PANE` 侧栏面板：可实时切换全部功能开关、查看本机渲染引擎诊断状态、一键清理本地预览缓存，以及在未开启 CDP 时直接粘贴本地路径进行快速预览。

---

## 为什么可以独立运行且与 `agy-powertools` 零冲突？

- **独立插件 ID 与数据目录**：注册为 `agy-preview`，配置与缓存位于 `~/.gemini/antigravity/plugin_data/agy-preview/`，与其它插件完全隔离。
- **独立运行命名空间**：使用纯原生 DOM + 内联 SVG + Antigravity 原生 CSS 变量，不依赖 React 内部混淆符号，与 `agy-powertools`（或其他 Sidecar 插件）共用调试端口时互不覆盖、互不干扰。
- **不侵入客户端安装包**：不修补 `app.asar`，不调用受 CSRF 收紧限制的未公开接口；侧栏面板走官方 Sidecar 环境变量鉴权（`ANTIGRAVITY_SIDECAR_WEB_PORT` / `ANTIGRAVITY_SIDECAR_UI_TOKEN`）。

---

## 环境要求与可选渲染引擎

- **必需**：Google Antigravity 2.0+，以及 Node.js `>= 22.13.0`（安装脚本会自动回退使用 Antigravity 内置的 `agy-node`）。
- **HTML 实时预览与聊天内图片展示**：零额外依赖，开箱即用。
- **DOCX / PPTX 高保真预览**（按优先级自动探测与降级）：

| 能力 | 推荐环境 | 降级方案 |
| :--- | :--- | :--- |
| PPTX 高保真页面 | Windows Microsoft PowerPoint (`pywin32`) + `PyMuPDF` | LibreOffice (`soffice`) + `PyMuPDF` -> `python-pptx` 幻灯片卡片 |
| PPTX 演讲者备注 | `python-pptx` | 无备注时自动隐藏备注按钮 |
| DOCX 原版排版 | Windows Microsoft Word (`pywin32`) + `PyMuPDF` | LibreOffice (`soffice`) + `PyMuPDF` -> 自动切至流式阅读 |
| DOCX 流式阅读 | `python-docx` | 无 |
| 大图高速缩略图 | `Pillow` (`PIL`) | 小图直出 Base64，零依赖可用 |

Python 依赖一键安装示例：

```bash
pip install pymupdf python-docx python-pptx pillow pywin32
```

> 可通过环境变量 `AGY_PREVIEW_PYTHON` 手动指定 Python 可执行文件路径。

---

## 安装

### Windows (PowerShell)

请先关闭 Antigravity 客户端，拉取仓库后在仓库根目录执行：

```powershell
git clone https://github.com/huiyiyichen/agy-preview.git
cd agy-preview
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

如不希望自动检查并补充快捷方式的 `--remote-debugging-port=0` 参数（例如仅使用侧栏快速预览，或已有其它工具管理启动参数）：

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1 -NoLauncher
```

### macOS / Linux

```bash
chmod +x ./install.sh ./uninstall.sh
./install.sh
```

> 若在 macOS / Linux 下启用主界面自动接管，请以 `--remote-debugging-port=0` 参数启动 Antigravity。

安装完成后，重新启动 Antigravity 即可生效。

---

## 卸载

### Windows (PowerShell)

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

保留 `settings.json` 个人开关配置：

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1 -KeepSettings
```

### macOS / Linux

```bash
./uninstall.sh
```

卸载只会移除 `agy-preview` 自身目录、`config.json` 中的 `agy-preview` 条目，以及由本插件亲自添加的快捷方式启动参数，不会影响 `agy-powertools` 等其它插件。

---

## 开发与测试

项目零 npm 第三方依赖，直接使用 Node.js 内置测试运行器：

```bash
npm test
```

## 开源协议

[MIT License](./LICENSE)

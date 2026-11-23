# Agy Preview

English | [简体中文](./README.md)

**Agy Preview** is a standalone, open-source preview enhancement plugin for **Google Antigravity 2.0**, built on the official Sidecar architecture with zero `app.asar` patching.

## Features

1. **High-Fidelity DOCX & PPTX Preview**
   - Automatically replaces `"Preview is not available for this file type."` in Antigravity's right-hand File Viewer for `.docx`, `.doc`, `.pptx`, and `.ppt`.
   - **PPTX / PPT**: Silent Microsoft PowerPoint COM (or headless LibreOffice fallback) to PDF + PyMuPDF WebP slide rendering, with automatic **`python-pptx` Speaker Notes extraction**, **Continuous Scroll / Slide View** (`ArrowLeft` / `ArrowRight` navigation), zoom, click-to-lightbox, and disk LRU caching.
   - **DOCX / DOC**: Instant toggle between **Layout View** (Word COM / LibreOffice exact pagination) and **Reader View** (`python-docx` responsive HTML with dark/light theme, tables, embedded images, and selectable text). Automatically falls back to Reader View if neither Word nor LibreOffice is installed.
2. **Live HTML Preview**
   - Adds a floating **`[ Preview | Source ]`** toggle bar to open `.html` / `.htm` files inside Antigravity.
   - Safely inlines same-directory relative CSS, JS, and images into a sandboxed `<iframe srcdoc>`.
3. **Codex-Style Inline Chat Image Rendering & Lightbox**
   - **In-Place Tool Step Merge**: Replaces consecutive `Analyzed <image>` rows inside `Explored files` with a single collapsible **`Viewed N images`** rounded thumbnail strip while preserving non-image file steps.
   - **Direct Reply Image Cards**: Renders horizontal rounded preview cards right beneath markdown paragraphs or list items whenever the assistant references local image files in its reply.
   - Click any thumbnail or card to open the full-screen Lightbox (`Esc`, `ArrowLeft` / `ArrowRight`, `100%` zoom toggle, Open Externally, Reveal in Folder).
4. **Official Sidecar Aux Pane UI**
   - Provides a native side-pane dashboard (`SIDECAR_UI_ENTRYPOINT_AUX_PANE`) for live settings toggles, local rendering engine diagnostics, cache cleanup, and quick local file preview (works even without CDP).

## Zero-Conflict Isolation

- Uses its own plugin ID (`agy-preview`), isolated data directory (`~/.gemini/antigravity/plugin_data/agy-preview`), and isolated DOM/global namespaces (`__agyPreview*`, `data-agy-preview-*`).
- Pure vanilla DOM + inline SVG + Antigravity CSS custom properties — zero dependency on React internals or minified symbols.
- Runs cleanly standalone or alongside `agy-powertools` and other Antigravity plugins.

## Requirements

- Google Antigravity 2.0+ and Node.js `>= 22.13.0` (installers automatically fall back to Antigravity's bundled `agy-node`).
- HTML preview and inline chat image display require no extra dependencies.
- Optional Python packages for Office document rendering:

```bash
pip install pymupdf python-docx python-pptx pillow pywin32
```

Set `AGY_PREVIEW_PYTHON` to point to a specific Python binary if needed.

## Installation

### Windows (PowerShell)

Close Antigravity first, clone the repository, and run the installer:

```powershell
git clone https://github.com/huiyiyichen/agy-preview.git
cd agy-preview
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

Pass `-NoLauncher` to skip checking Antigravity `.lnk` shortcuts for `--remote-debugging-port=0`.

### macOS / Linux

```bash
chmod +x ./install.sh ./uninstall.sh
./install.sh
```

Restart Antigravity after installation.

## Uninstallation

```powershell
# Windows
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

```bash
# macOS / Linux
./uninstall.sh
```

Pass `-KeepSettings` (PowerShell) or `--keep-settings` (Node/Bash) to preserve `settings.json`.

## Testing

Zero npm dependencies; run the built-in Node test suite:

```bash
npm test
```

## License

[MIT](./LICENSE)

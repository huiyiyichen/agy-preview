import base64
import html
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')


def find_soffice():
    for name in ('soffice.com', 'soffice.exe', 'soffice', 'libreoffice'):
        hit = shutil.which(name)
        if hit:
            return hit
    candidates = [
        r'C:\Program Files\LibreOffice\program\soffice.com',
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'D:\Program Files\LibreOffice\program\soffice.com',
        r'D:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.com',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
        '/usr/bin/soffice',
        '/usr/bin/libreoffice',
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def has_module(name):
    try:
        __import__(name)
        return True
    except Exception:
        return False


def has_com_progid(progid):
    if os.name != 'nt':
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
            return True
    except Exception:
        return False


def cmd_diagnostics():
    soffice = find_soffice()
    print(
        json.dumps(
            {
                'ok': True,
                'python': sys.executable,
                'pythonVersion': sys.version.split()[0],
                'powerpointCom': has_com_progid('PowerPoint.Application') and has_module('win32com'),
                'wordCom': has_com_progid('Word.Application') and has_module('win32com'),
                'libreoffice': bool(soffice),
                'libreofficePath': soffice or '',
                'pymupdf': has_module('fitz'),
                'pillow': has_module('PIL'),
                'pythonDocx': has_module('docx'),
                'pythonPptx': has_module('pptx'),
            },
            ensure_ascii=False,
        )
    )


def convert_pdf_via_com(src_path, pdf_path, kind):
    if os.name != 'nt':
        return None
    try:
        import pythoncom
        import win32com.client
    except Exception:
        return None
    abs_src = os.path.abspath(src_path)
    abs_pdf = os.path.abspath(pdf_path)
    pythoncom.CoInitialize()
    try:
        if kind == 'pptx':
            reused = False
            try:
                app = win32com.client.GetActiveObject('PowerPoint.Application')
                reused = True
            except Exception:
                app = win32com.client.DispatchEx('PowerPoint.Application')
            pres = None
            try:
                pres = app.Presentations.Open(abs_src, True, False, False)
                pres.SaveAs(abs_pdf, 32)
            finally:
                if pres is not None:
                    try:
                        pres.Close()
                    except Exception:
                        pass
                if not reused:
                    try:
                        app.Quit()
                    except Exception:
                        pass
            if os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 0:
                return 'powerpoint'
            return None
        if kind == 'docx':
            reused = False
            try:
                app = win32com.client.GetActiveObject('Word.Application')
                reused = True
            except Exception:
                app = win32com.client.DispatchEx('Word.Application')
                try:
                    app.Visible = False
                except Exception:
                    pass
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            doc = None
            try:
                doc = app.Documents.Open(
                    abs_src,
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False,
                )
                doc.SaveAs2(abs_pdf, FileFormat=17)
            finally:
                if doc is not None:
                    try:
                        doc.Close(SaveChanges=0)
                    except Exception:
                        pass
                if not reused:
                    try:
                        app.Quit()
                    except Exception:
                        pass
            if os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 0:
                return 'word'
            return None
    except Exception:
        return None
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return None


def convert_pdf_via_libreoffice(src_path, pdf_path, profile_dir):
    soffice = find_soffice()
    if not soffice:
        return None
    os.makedirs(profile_dir, exist_ok=True)
    work_dir = tempfile.mkdtemp(prefix='agy-pv-lo-')
    try:
        stem = 'input' + os.path.splitext(src_path)[1].lower()
        temp_src = os.path.join(work_dir, stem)
        shutil.copy2(src_path, temp_src)
        profile_uri = 'file:///' + os.path.abspath(profile_dir).replace('\\', '/')
        creationflags = 0x08000000 if os.name == 'nt' else 0
        res = subprocess.run(
            [
                soffice,
                f'-env:UserInstallation={profile_uri}',
                '--headless',
                '--norestore',
                '--nolockcheck',
                '--convert-to',
                'pdf',
                '--outdir',
                work_dir,
                temp_src,
            ],
            capture_output=True,
            text=True,
            timeout=50,
            creationflags=creationflags,
        )
        produced = os.path.join(work_dir, 'input.pdf')
        if res.returncode == 0 and os.path.exists(produced) and os.path.getsize(produced) > 0:
            shutil.copy2(produced, pdf_path)
            return 'libreoffice'
    except Exception:
        return None
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
    return None


def extract_pptx_metadata(src_path):
    notes_map = {}
    if not src_path.lower().endswith('.pptx'):
        return notes_map
    try:
        from pptx import Presentation

        prs = Presentation(src_path)
        for idx, slide in enumerate(prs.slides, start=1):
            title = ''
            try:
                if slide.shapes.title and slide.shapes.title.text:
                    title = slide.shapes.title.text.strip().splitlines()[0][:120]
            except Exception:
                pass
            notes = ''
            try:
                if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                    notes = slide.notes_slide.notes_text_frame.text.strip()
            except Exception:
                pass
            notes_map[idx] = {'title': title, 'notes': notes}
    except Exception:
        pass
    return notes_map


def render_pdf_pages(pdf_path, out_dir, notes_map=None, max_dim=1600):
    import fitz
    from PIL import Image

    os.makedirs(out_dir, exist_ok=True)
    pages = []
    doc = fitz.open(pdf_path)
    try:
        for i, page in enumerate(doc, start=1):
            rect = page.rect
            longest = max(rect.width, rect.height, 1.0)
            scale = min(2.4, max(1.2, max_dim / longest))
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
            out_name = f'page-{i:04d}.webp'
            out_path = os.path.join(out_dir, out_name)
            try:
                img.save(out_path, 'WEBP', quality=86, method=4)
                mime = 'image/webp'
            except Exception:
                out_name = f'page-{i:04d}.png'
                out_path = os.path.join(out_dir, out_name)
                img.save(out_path, 'PNG', optimize=True)
                mime = 'image/png'
            meta = (notes_map or {}).get(i, {})
            pages.append(
                {
                    'index': i,
                    'file': out_name,
                    'mime': mime,
                    'width': pix.width,
                    'height': pix.height,
                    'title': meta.get('title', ''),
                    'notes': meta.get('notes', ''),
                }
            )
    finally:
        doc.close()
    return pages


def render_docx_reader_html(src_path):
    from docx import Document
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(src_path)
    rel_images = {}
    for r_id, rel in doc.part.rels.items():
        try:
            if 'image' in str(rel.reltype):
                blob = rel.target_part.blob
                ctype = rel.target_part.content_type or 'image/png'
                if len(blob) <= 4 * 1024 * 1024:
                    rel_images[r_id] = f'data:{ctype};base64,{base64.b64encode(blob).decode("ascii")}'
        except Exception:
            pass

    def render_run(run):
        chunks = []
        try:
            for blip in run._r.xpath('.//*[local-name()="blip"]'):
                for attr_k, attr_v in blip.attrib.items():
                    if attr_k.endswith('}embed') and attr_v in rel_images:
                        chunks.append(f'<img class="doc-img" src="{rel_images[attr_v]}" alt="" />')
        except Exception:
            pass
        txt = run.text or ''
        if txt:
            esc = html.escape(txt).replace('\n', '<br/>')
            if run.bold:
                esc = f'<strong>{esc}</strong>'
            if run.italic:
                esc = f'<em>{esc}</em>'
            if run.underline:
                esc = f'<u>{esc}</u>'
            chunks.append(esc)
        return ''.join(chunks)

    def render_paragraph(p):
        inner = ''.join(render_run(r) for r in p.runs)
        if not inner.strip():
            return ''
        style_name = (p.style.name if p.style else '') or ''
        low = style_name.lower()
        if low.startswith('heading '):
            lvl = low.replace('heading ', '').strip()
            tag = f'h{lvl}' if lvl in ('1', '2', '3', '4', '5', '6') else 'h3'
            return f'<{tag}>{inner}</{tag}>'
        if 'title' in low and 'sub' not in low:
            return f'<h1 class="doc-title">{inner}</h1>'
        if 'list' in low:
            return f'<p class="doc-list-item">{inner}</p>'
        return f'<p>{inner}</p>'

    def render_table(tbl):
        rows_html = []
        for r_idx, row in enumerate(tbl.rows):
            tag = 'th' if r_idx == 0 else 'td'
            cells = []
            for cell in row.cells:
                body = ''.join(render_paragraph(p) for p in cell.paragraphs) or html.escape(cell.text.strip())
                cells.append(f'<{tag}>{body}</{tag}>')
            rows_html.append(f'<tr>{"".join(cells)}</tr>')
        return f'<div class="doc-table-wrap"><table>{"".join(rows_html)}</table></div>'

    parts = []
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            frag = render_paragraph(Paragraph(child, doc))
            if frag:
                parts.append(frag)
        elif isinstance(child, CT_Tbl):
            parts.append(render_table(Table(child, doc)))
    return '\n'.join(parts) if parts else '<p class="doc-empty">(空文档)</p>'


def render_pptx_fallback_html(src_path):
    from pptx import Presentation

    prs = Presentation(src_path)
    cards = []
    for idx, slide in enumerate(prs.slides, start=1):
        lines = []
        for shape in slide.shapes:
            if getattr(shape, 'has_text_frame', False):
                txt = shape.text_frame.text.strip()
                if txt:
                    lines.append(f'<p>{html.escape(txt).replace(chr(10), "<br/>")}</p>')
            elif getattr(shape, 'has_table', False):
                rows = []
                for r in shape.table.rows:
                    rows.append('<tr>' + ''.join(f'<td>{html.escape(c.text.strip())}</td>' for c in r.cells) + '</tr>')
                lines.append(f'<div class="doc-table-wrap"><table>{"".join(rows)}</table></div>')
        notes = ''
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                nt = slide.notes_slide.notes_text_frame.text.strip()
                if nt:
                    notes = f'<div class="slide-notes"><strong>演讲者备注：</strong>{html.escape(nt)}</div>'
        except Exception:
            pass
        cards.append(f'<section class="slide-card"><h3>Slide {idx}</h3>{"".join(lines)}{notes}</section>')
    return '\n'.join(cards) if cards else '<p class="doc-empty">(空演示文稿)</p>'


def cmd_render_office(src_path, out_dir, mode, profile_dir):
    t0 = time.time()
    ext = os.path.splitext(src_path)[1].lower()
    kind = 'pptx' if ext in ('.pptx', '.ppt') else 'docx'
    os.makedirs(out_dir, exist_ok=True)

    if kind == 'docx' and mode == 'reader' and ext == '.docx':
        body_html = render_docx_reader_html(src_path)
        with open(os.path.join(out_dir, 'reader.html'), 'w', encoding='utf-8') as f:
            f.write(body_html)
        meta = {
            'ok': True,
            'kind': 'docx',
            'mode': 'reader',
            'engine': 'python-docx',
            'pageCount': 0,
            'pages': [],
            'htmlFile': 'reader.html',
            'elapsedMs': int((time.time() - t0) * 1000),
        }
        with open(os.path.join(out_dir, 'meta.json'), 'w', encoding='utf-8') as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        print(json.dumps(meta, ensure_ascii=False))
        return

    pdf_path = os.path.join(out_dir, 'converted.pdf')
    engine = convert_pdf_via_com(src_path, pdf_path, kind)
    if not engine:
        engine = convert_pdf_via_libreoffice(src_path, pdf_path, profile_dir)

    if engine and os.path.exists(pdf_path):
        notes_map = extract_pptx_metadata(src_path) if kind == 'pptx' else None
        pages = render_pdf_pages(pdf_path, out_dir, notes_map=notes_map)
        try:
            os.remove(pdf_path)
        except Exception:
            pass
        meta = {
            'ok': True,
            'kind': kind,
            'mode': 'layout',
            'engine': engine,
            'pageCount': len(pages),
            'pages': pages,
            'elapsedMs': int((time.time() - t0) * 1000),
        }
        with open(os.path.join(out_dir, 'meta.json'), 'w', encoding='utf-8') as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        print(json.dumps(meta, ensure_ascii=False))
        return

    if kind == 'docx' and ext == '.docx':
        body_html = render_docx_reader_html(src_path)
        with open(os.path.join(out_dir, 'reader.html'), 'w', encoding='utf-8') as f:
            f.write(body_html)
        meta = {
            'ok': True,
            'kind': 'docx',
            'mode': 'reader',
            'engine': 'python-docx-fallback',
            'pageCount': 0,
            'pages': [],
            'htmlFile': 'reader.html',
            'elapsedMs': int((time.time() - t0) * 1000),
        }
        with open(os.path.join(out_dir, 'meta.json'), 'w', encoding='utf-8') as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        print(json.dumps(meta, ensure_ascii=False))
        return

    if kind == 'pptx' and ext == '.pptx':
        body_html = render_pptx_fallback_html(src_path)
        with open(os.path.join(out_dir, 'reader.html'), 'w', encoding='utf-8') as f:
            f.write(body_html)
        meta = {
            'ok': True,
            'kind': 'pptx',
            'mode': 'reader',
            'engine': 'python-pptx-fallback',
            'pageCount': 0,
            'pages': [],
            'htmlFile': 'reader.html',
            'elapsedMs': int((time.time() - t0) * 1000),
        }
        with open(os.path.join(out_dir, 'meta.json'), 'w', encoding='utf-8') as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        print(json.dumps(meta, ensure_ascii=False))
        return

    raise RuntimeError('未找到可用的 Office 或 LibreOffice 转换引擎，也无法以 Python 回退解析该文件')


def cmd_thumb_image(src_path, max_width):
    from PIL import Image, ImageOps

    with Image.open(src_path) as im:
        im = ImageOps.exif_transpose(im)
        orig_w, orig_h = im.size
        if max_width > 0 and orig_w > max_width:
            ratio = max_width / float(orig_w)
            new_size = (max(1, int(round(orig_w * ratio))), max(1, int(round(orig_h * ratio))))
            im = im.resize(new_size, Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        mime = 'image/webp'
        try:
            if im.mode not in ('RGB', 'RGBA'):
                im = im.convert('RGBA' if 'A' in im.getbands() else 'RGB')
            im.save(buf, format='WEBP', quality=86, method=4)
        except Exception:
            buf = io.BytesIO()
            mime = 'image/png'
            im.save(buf, format='PNG', optimize=True)
        b64 = base64.b64encode(buf.getvalue()).decode('ascii')
        print(
            json.dumps(
                {
                    'ok': True,
                    'mime': mime,
                    'width': im.size[0],
                    'height': im.size[1],
                    'origWidth': orig_w,
                    'origHeight': orig_h,
                    'dataUrl': f'data:{mime};base64,{b64}',
                },
                ensure_ascii=False,
            )
        )


def main():
    if len(sys.argv) < 2:
        raise SystemExit('Usage: office-preview.py <diagnostics|render-office|thumb-image> ...')
    cmd = sys.argv[1]
    if cmd == 'diagnostics':
        cmd_diagnostics()
    elif cmd == 'render-office':
        if len(sys.argv) < 6:
            raise SystemExit('Usage: office-preview.py render-office <src> <outDir> <mode> <profileDir>')
        cmd_render_office(sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
    elif cmd == 'thumb-image':
        if len(sys.argv) < 4:
            raise SystemExit('Usage: office-preview.py thumb-image <src> <maxWidth>')
        cmd_thumb_image(sys.argv[2], int(sys.argv[3]))
    else:
        raise SystemExit(f'Unknown command: {cmd}')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(json.dumps({'ok': False, 'error': str(exc)}, ensure_ascii=False))
        sys.exit(1)

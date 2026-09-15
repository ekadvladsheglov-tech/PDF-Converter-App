# app_web.py
import os, io, sys, json
import urllib.request
import tempfile
import subprocess
import webview
from pdf2docx import Converter
from docx import Document
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from PIL import Image
try:
    import pymupdf as fitz
except ImportError:
    import fitz

APP_VERSION = "3.1.0"
UPDATE_URL = "https://raw.githubusercontent.com/ekadvladsheglov-tech/PDF-Converter-App/main/update/version.json"

def res(p):
    return os.path.join(getattr(sys, '_MEIPASS', os.path.abspath('.')), p)

_font_done = False
def pdf_font():
    global _font_done
    if not _font_done:
        fp = res('DejaVuSans.ttf')
        if os.path.exists(fp):
            try:
                pdfmetrics.registerFont(TTFont('DejaVu', fp))
            except Exception:
                pass
        _font_done = True
    return 'DejaVu' if 'DejaVu' in pdfmetrics.getRegisteredFontNames() else 'Helvetica'

def parse_range(text, total):
    out = set()
    for part in text.replace(' ', '').split(','):
        if not part:
            continue
        try:
            if '-' in part:
                a, b = part.split('-', 1)
                a, b = int(a), int(b)
                if a > b:
                    a, b = b, a
                out |= {p - 1 for p in range(a, b + 1) if 1 <= p <= total}
            else:
                p = int(part)
                if 1 <= p <= total:
                    out.add(p - 1)
        except ValueError:
            continue
    return sorted(out)

WIN = None

def js(code):
    try:
        WIN.evaluate_js(code)
    except Exception:
        pass


class Api:
    # ---------- диалоги (pywebview 4.x / 5.x / 6.x) ----------
    def browse(self, kind):
        FD = getattr(webview, 'FileDialog', None)
        if kind == 'folder':
            r = WIN.create_file_dialog(FD.FOLDER) if FD else WIN.create_file_dialog(webview.FOLDER_DIALOG)
            return r[0] if r else None
        ft_map = {
            'pdf': ('PDF файлы', '*.pdf'),
            'docx': ('Word файлы', '*.docx'),
            'img': ('Изображения', '*.jpg;*.png;*.bmp;*.gif;*.tiff'),
            'multipdf': ('PDF файлы', '*.pdf'),
            'multiimg': ('Изображения', '*.jpg;*.png;*.bmp;*.gif;*.tiff'),
        }
        name, mask = ft_map[kind]
        multi = kind.startswith('multi')
        if FD:
            try:
                r = WIN.create_file_dialog(FD.OPEN, file_types=(f'{name} ({mask})',), allow_multiple=multi)
            except TypeError:
                r = WIN.create_file_dialog(FD.OPEN, file_types=(f'{name} ({mask})',))
        else:
            r = WIN.create_file_dialog(webview.OPEN_DIALOG, multiple=multi, file_types=(f'{name} ({mask})',))
        if not r:
            return None
        return list(r) if multi else r[0]

    def _save(self, name, ext):
        FD = getattr(webview, 'FileDialog', None)
        if FD:
            r = WIN.create_file_dialog(FD.SAVE, save_filename=name + ext)
        else:
            r = WIN.create_file_dialog(webview.SAVE_DIALOG, save_filename=name + ext)
        if not r:
            return None
        p = r if isinstance(r, str) else r[0]
        return p if p.lower().endswith(ext) else p + ext

    # ---------- обновления ----------
    def get_version(self):
        return APP_VERSION

    def check_update(self):
        try:
            req = urllib.request.Request(UPDATE_URL, headers={'User-Agent': 'PDFConverter'})
            with urllib.request.urlopen(req, timeout=6) as r:
                data = json.loads(r.read().decode('utf-8'))
            if data.get('version') and data['version'] != APP_VERSION:
                return {'ok': True, 'new': data['version'], 'url': data['url'], 'notes': data.get('notes', '')}
            return {'ok': True, 'new': None}
        except Exception:
            return {'ok': False}

    def do_update(self, url):
        try:
            tmp = os.path.join(tempfile.gettempdir(), 'PDF_Converter_Update.exe')
            req = urllib.request.Request(url, headers={'User-Agent': 'PDFConverter'})
            with urllib.request.urlopen(req, timeout=20) as r, open(tmp, 'wb') as f:
                total = int(r.headers.get('Content-Length') or 0)
                done = 0
                while True:
                    chunk = r.read(256 * 1024)
                    if not chunk:
                        break
                    f.write(chunk)
                    done += len(chunk)
                    if total:
                        js(f'UI.prog({done/total}, "Загрузка обновления… {done//1048576} МБ")')
            js('UI.prog(1, "Запуск установщика…")')
            subprocess.Popen([tmp], shell=False)
            WIN.destroy()
            return {'ok': True}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    # ---------- PDF -> Word ----------
    def run_p2w(self, d):
        try:
            if d.get('mode') == 1:
                FD = getattr(webview, 'FileDialog', None)
                out_dir = WIN.create_file_dialog(FD.FOLDER) if FD else WIN.create_file_dialog(webview.FOLDER_DIALOG)
                if not out_dir:
                    return {'ok': False, 'silent': True}
                out_dir = out_dir[0]
                files = [f for f in os.listdir(d['path']) if f.lower().endswith('.pdf')]
                n = len(files)
                ok = 0
                for i, f in enumerate(files):
                    try:
                        cv = Converter(os.path.join(d['path'], f))
                        cv.convert(os.path.join(out_dir, os.path.splitext(f)[0] + '.docx'))
                        cv.close()
                        ok += 1
                    except Exception:
                        pass
                    js(f'UI.prog({(i+1)/n}, "Файл {i+1}/{n}: {f}")')
                return {'ok': True, 'msg': f'Готово: {ok}/{n}', 'file': out_dir}
            out = self._save('converted', '.docx')
            if not out:
                return {'ok': False, 'silent': True}
            js('UI.prog(0.4, "Конвертация…")')
            cv = Converter(d['path'])
            cv.convert(out)
            cv.close()
            js('UI.prog(1, "")')
            return {'ok': True, 'msg': 'Сохранено', 'file': os.path.basename(out)}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    # ---------- Word -> PDF ----------
    def run_w2p(self, d):
        try:
            if d.get('mode') == 1:
                FD = getattr(webview, 'FileDialog', None)
                out_dir = WIN.create_file_dialog(FD.FOLDER) if FD else WIN.create_file_dialog(webview.FOLDER_DIALOG)
                if not out_dir:
                    return {'ok': False, 'silent': True}
                out_dir = out_dir[0]
                files = [f for f in os.listdir(d['path']) if f.lower().endswith('.docx')]
                n = len(files)
                ok = 0
                for i, f in enumerate(files):
                    try:
                        self._w2p(os.path.join(d['path'], f), os.path.join(out_dir, os.path.splitext(f)[0] + '.pdf'))
                        ok += 1
                    except Exception:
                        pass
                    js(f'UI.prog({(i+1)/n}, "Файл {i+1}/{n}: {f}")')
                return {'ok': True, 'msg': f'Готово: {ok}/{n}', 'file': out_dir}
            out = self._save('converted', '.pdf')
            if not out:
                return {'ok': False, 'silent': True}
            js('UI.prog(0.4, "Конвертация…")')
            self._w2p(d['path'], out)
            js('UI.prog(1, "")')
            return {'ok': True, 'msg': 'Сохранено', 'file': os.path.basename(out)}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    def _w2p(self, word_path, pdf_path):
        doc = Document(word_path)
        c = canvas.Canvas(pdf_path, pagesize=A4)
        w, h = A4
        fn = pdf_font()
        fs = 12
        x, y = 20 * mm, h - 20 * mm
        lh, mw, bm = 6 * mm, w - 40 * mm, 20 * mm
        c.setFont(fn, fs)
        for para in doc.paragraphs:
            t = para.text
            if not t.strip():
                y -= lh
                if y < bm:
                    c.showPage(); c.setFont(fn, fs); y = h - 20 * mm
                continue
            words = []
            for word in t.split(' '):
                if not word:
                    continue
                while word and c.stringWidth(word, fn, fs) > mw:
                    lo, hi = 1, len(word)
                    while lo < hi:
                        mid = (lo + hi + 1) // 2
                        if c.stringWidth(word[:mid], fn, fs) <= mw:
                            lo = mid
                        else:
                            hi = mid - 1
                    words.append(word[:lo])
                    word = word[lo:]
                words.append(word)
            line = ''
            for word in words:
                test = line + ' ' + word if line else word
                if c.stringWidth(test, fn, fs) <= mw:
                    line = test
                else:
                    c.drawString(x, y, line)
                    y -= lh
                    if y < bm:
                        c.showPage(); c.setFont(fn, fs); y = h - 20 * mm
                    line = word
            if line:
                c.drawString(x, y, line)
                y -= lh
            if y < bm:
                c.showPage(); c.setFont(fn, fs); y = h - 20 * mm
        c.save()

    # ---------- Фото -> PDF ----------
    def run_img(self, d):
        try:
            out = self._save('images', '.pdf')
            if not out:
                return {'ok': False, 'silent': True}
            imgs = []
            for p in d['files']:
                im = Image.open(p)
                imgs.append(im.convert('RGB') if im.mode != 'RGB' else im)
            imgs[0].save(out, 'PDF', save_all=True, append_images=imgs[1:], resolution=150)
            return {'ok': True, 'msg': f'PDF из {len(imgs)} изображений', 'file': os.path.basename(out)}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    # ---------- Разделить ----------
    def run_split(self, d):
        try:
            FD = getattr(webview, 'FileDialog', None)
            out_dir = WIN.create_file_dialog(FD.FOLDER) if FD else WIN.create_file_dialog(webview.FOLDER_DIALOG)
            if not out_dir:
                return {'ok': False, 'silent': True}
            out_dir = out_dir[0]
            rd = PdfReader(d['path'])
            base = os.path.splitext(os.path.basename(d['path']))[0]
            for i, pg in enumerate(rd.pages):
                wr = PdfWriter()
                wr.add_page(pg)
                with open(os.path.join(out_dir, f'{base}_page_{i+1}.pdf'), 'wb') as f:
                    wr.write(f)
            return {'ok': True, 'msg': f'Страниц: {len(rd.pages)}', 'file': base}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    # ---------- Объединить ----------
    def run_merge(self, d):
        try:
            out = self._save('merged', '.pdf')
            if not out:
                return {'ok': False, 'silent': True}
            wr = PdfWriter()
            for p in d['files']:
                for pg in PdfReader(p).pages:
                    wr.add_page(pg)
            with open(out, 'wb') as f:
                wr.write(f)
            return {'ok': True, 'msg': f'Объединено файлов: {len(d["files"])}', 'file': os.path.basename(out)}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    # ---------- Сжать ----------
    def run_compress(self, d):
        try:
            if d.get('mode') == 1:
                FD = getattr(webview, 'FileDialog', None)
                out_dir = WIN.create_file_dialog(FD.FOLDER) if FD else WIN.create_file_dialog(webview.FOLDER_DIALOG)
                if not out_dir:
                    return {'ok': False, 'silent': True}
                out_dir = out_dir[0]
                files = [f for f in os.listdir(d['path']) if f.lower().endswith('.pdf')]
                n = len(files)
                ok = 0
                for i, f in enumerate(files):
                    try:
                        doc = fitz.open(os.path.join(d['path'], f))
                        doc.save(os.path.join(out_dir, f), garbage=4, deflate=True, clean=True)
                        doc.close()
                        ok += 1
                    except Exception:
                        pass
                    js(f'UI.prog({(i+1)/n}, "Файл {i+1}/{n}: {f}")')
                return {'ok': True, 'msg': f'Готово: {ok}/{n}', 'file': out_dir}
            out = self._save('compressed', '.pdf')
            if not out:
                return {'ok': False, 'silent': True}
            js('UI.prog(0.4, "Сжатие…")')
            doc = fitz.open(d['path'])
            doc.save(out, garbage=4, deflate=True, clean=True)
            doc.close()
            o = os.path.getsize(d['path']) / 1024
            c2 = os.path.getsize(out) / 1024
            pct = (1 - c2 / o) * 100 if o else 0
            js('UI.prog(1, "")')
            return {'ok': True, 'msg': f'Экономия {pct:.1f}% ({o:.0f}→{c2:.0f} КБ)', 'file': os.path.basename(out)}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    # ---------- Пароль ----------
    def run_protect(self, d):
        try:
            if not d.get('pwd'):
                return {'ok': False, 'msg': 'Введите пароль'}
            out = self._save('protected', '.pdf')
            if not out:
                return {'ok': False, 'silent': True}
            wr = PdfWriter()
            for pg in PdfReader(d['path']).pages:
                wr.add_page(pg)
            wr.encrypt(user_password=d['pwd'], owner_password=d['pwd'], use_128bit=True)
            with open(out, 'wb') as f:
                wr.write(f)
            return {'ok': True, 'msg': 'Защищено паролем', 'file': os.path.basename(out)}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    # ---------- Извлечь ----------
    def run_extract(self, d):
        try:
            rd = PdfReader(d['path'])
            total = len(rd.pages)
            idx = parse_range(d.get('range', ''), total) or list(range(total))
            out = self._save('extracted', '.pdf')
            if not out:
                return {'ok': False, 'silent': True}
            wr = PdfWriter()
            for i in idx:
                wr.add_page(rd.pages[i])
            with open(out, 'wb') as f:
                wr.write(f)
            return {'ok': True, 'msg': f'Извлечено страниц: {len(idx)}', 'file': os.path.basename(out)}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    # ---------- Повернуть ----------
    def run_rotate(self, d):
        try:
            rd = PdfReader(d['path'])
            total = len(rd.pages)
            idx = parse_range(d.get('range', ''), total) or list(range(total))
            out = self._save('rotated', '.pdf')
            if not out:
                return {'ok': False, 'silent': True}
            wr = PdfWriter()
            for i, pg in enumerate(rd.pages):
                if i in idx:
                    pg.rotate(int(d.get('angle', 90)))
                wr.add_page(pg)
            with open(out, 'wb') as f:
                wr.write(f)
            return {'ok': True, 'msg': f'Повернуто: {len(idx)}', 'file': os.path.basename(out)}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    # ---------- Удалить ----------
    def run_delete(self, d):
        try:
            rd = PdfReader(d['path'])
            total = len(rd.pages)
            idx = parse_range(d.get('range', ''), total)
            if not idx:
                return {'ok': False, 'msg': 'Укажите страницы'}
            out = self._save('cleaned', '.pdf')
            if not out:
                return {'ok': False, 'silent': True}
            wr = PdfWriter()
            deln = 0
            for i, pg in enumerate(rd.pages):
                if i in idx:
                    deln += 1
                else:
                    wr.add_page(pg)
            if len(wr.pages) == 0:
                return {'ok': False, 'msg': 'Нельзя удалить все страницы'}
            with open(out, 'wb') as f:
                wr.write(f)
            return {'ok': True, 'msg': f'Удалено: {deln}, осталось: {len(wr.pages)}', 'file': os.path.basename(out)}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}

    # ---------- Номера ----------
    def run_numbers(self, d):
        try:
            out = self._save('numbered', '.pdf')
            if not out:
                return {'ok': False, 'silent': True}
            rd = PdfReader(d['path'])
            wr = PdfWriter()
            for i, pg in enumerate(rd.pages):
                w, h = float(pg.mediabox.width), float(pg.mediabox.height)
                pkt = io.BytesIO()
                cn = canvas.Canvas(pkt, pagesize=(w, h))
                cn.setFont('Helvetica', 10)
                cn.drawCentredString(w / 2, 20, str(i + 1))
                cn.save()
                pkt.seek(0)
                pg.merge_page(PdfReader(pkt).pages[0])
                wr.add_page(pg)
            with open(out, 'wb') as f:
                wr.write(f)
            return {'ok': True, 'msg': 'Нумерация добавлена', 'file': os.path.basename(out)}
        except Exception as e:
            return {'ok': False, 'msg': str(e)}


HTML = r'''<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><style>
:root{--bg:#F2F1EE;--card:#fff;--ink:#171717;--muted:#8a8a86;--line:#e7e5e0;--coral:#F2694B;--yellow:#F2E500;--green:#2EBD59;--dark:#171717;--dark3:#2e2e2e;--r:24px;--pill:999px;
--sh:0 12px 32px rgba(20,20,20,.07),0 2px 6px rgba(20,20,20,.05);--shc:0 6px 16px rgba(242,105,75,.35)}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--ink);height:100vh;display:flex;gap:16px;padding:16px;overflow:hidden;user-select:none}
aside{width:232px;background:var(--dark);border-radius:var(--r);padding:22px 12px;display:flex;flex-direction:column;color:#fff;flex:none}
.logo{display:flex;align-items:center;gap:10px;padding:0 12px 18px}
.logo i{width:38px;height:38px;border-radius:12px;background:var(--coral);display:grid;place-items:center;font-style:normal;font-size:16px;color:#fff;font-weight:700}
.logo b{font-size:16px}.logo span{display:block;font-size:10px;color:#8b8b88;font-weight:600}
.grp{font-size:10px;font-weight:700;letter-spacing:.14em;color:#757570;margin:14px 14px 6px}
.nav button{display:flex;align-items:center;gap:10px;width:100%;padding:9px 12px;border:0;border-radius:12px;background:transparent;color:#c9c9c6;font:600 13px 'Segoe UI';cursor:pointer;transition:.15s}
.nav button:hover{background:var(--dark3);color:#fff}
.nav button.on{background:var(--coral);color:#fff;box-shadow:var(--shc)}
.foot{margin-top:auto;padding:12px 14px;font-size:11px;color:#757570}
.foot em{color:var(--green);font-style:normal}
#upd{margin:0 12px 10px;width:calc(100% - 24px);height:36px;font-size:12px}
main{flex:1;display:flex;flex-direction:column;gap:14px;min-width:0}
header{display:flex;align-items:baseline;gap:12px;padding:6px 8px}
header h1{font-size:22px}header span{font-size:12px;color:var(--muted);font-weight:600}
.grid{flex:1;display:grid;grid-template-columns:1fr 300px;gap:16px;min-height:0}
.card{background:var(--card);border-radius:var(--r);box-shadow:var(--sh);padding:26px 28px;animation:rise .3s ease both}
@keyframes rise{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
.panel{overflow:auto;display:flex;flex-direction:column;gap:14px}
.right{display:flex;flex-direction:column;gap:16px;overflow:auto}
.right .card{padding:20px 22px}
.lbl{font-size:12px;font-weight:600;color:var(--muted)}
.row{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.seg{display:inline-flex;background:#EDECE8;border-radius:var(--pill);padding:4px}
.seg button{border:0;background:transparent;padding:8px 18px;border-radius:var(--pill);color:var(--muted);font:600 13px 'Segoe UI';cursor:pointer;transition:.15s}
.seg button.on{background:var(--coral);color:#fff;box-shadow:var(--shc)}
input,select{width:100%;height:46px;border:1px solid var(--line);border-radius:var(--pill);padding:0 20px;font:400 13px 'Segoe UI';background:#FBFBFA;color:var(--ink);transition:.15s}
input:focus,select:focus{outline:none;border-color:var(--coral);box-shadow:0 0 0 4px rgba(242,105,75,.15)}
select{width:auto;padding-right:36px}
.btn{border:0;border-radius:var(--pill);height:42px;padding:0 26px;font:700 13px 'Segoe UI';cursor:pointer;transition:.15s;display:inline-flex;align-items:center;gap:8px}
.btn:hover{transform:translateY(-1px)}.btn:active{transform:scale(.97)}
.btn:disabled{opacity:.5;pointer-events:none}
.btn.primary{background:var(--coral);color:#fff;box-shadow:var(--shc)}
.btn.primary:hover{box-shadow:0 10px 24px rgba(242,105,75,.45)}
.btn.dark{background:var(--dark);color:#fff}
.btn.ghost{background:#EDECE8;color:var(--ink)}
.btn.yellow{background:var(--yellow);color:var(--ink)}
.progress{height:8px;background:#EDECE8;border-radius:var(--pill);overflow:hidden}
.progress i{display:block;height:100%;width:0;border-radius:var(--pill);background:linear-gradient(90deg,var(--coral),#ff9068);transition:width .25s}
.status{font-size:12px;color:var(--muted);font-weight:600}
.status.ok{color:var(--green)}.status.err{color:var(--coral)}
.listbox{min-height:120px;max-height:200px;overflow:auto;border:1px dashed var(--line);border-radius:16px;padding:12px 16px;font-size:12px;color:var(--muted);background:#FBFBFA}
.stat{display:flex;justify-content:space-between;align-items:baseline;padding:6px 0}
.stat b{font-size:22px}.stat span{font-size:11px;color:var(--muted);font-weight:600}
.tips{font-size:12px;color:var(--muted);line-height:1.5}
.recent{font-size:12px;color:var(--muted);display:flex;flex-direction:column;gap:6px}
.recent div{display:flex;gap:8px;align-items:center}
.recent div::before{content:'';width:6px;height:6px;border-radius:50%;background:var(--coral);flex:none}
.toast{position:fixed;right:24px;bottom:24px;background:var(--dark);color:#fff;border-radius:16px;padding:14px 20px;font:600 13px 'Segoe UI';box-shadow:0 16px 40px rgba(0,0,0,.25);animation:rise .25s ease both;display:flex;gap:10px;align-items:center}
.toast.err{background:var(--coral)}
</style></head><body>
<aside>
  <div class="logo"><i>P</i><div><b>PDF Tools</b><span id="ver">v3.1.0</span></div></div>
  <div class="nav" id="nav"></div>
  <button class="btn ghost" id="upd">🔄 Проверить обновления</button>
  <div class="foot"><em>✓</em> Локально &nbsp; <em>✓</em> Безопасно</div>
</aside>
<main>
  <header><h1 id="ttl"></h1><span id="sub"></span></header>
  <div class="grid">
    <section class="card panel" id="panel"></section>
    <div class="right">
      <div class="card"><div class="stat"><b id="s1">0</b><span>операций</span></div><div class="stat"><b id="s2">0</b><span>успешных</span></div></div>
      <div class="card"><div class="lbl" style="margin-bottom:8px">Подсказка</div><div class="tips" id="tips"></div></div>
      <div class="card"><div class="lbl" style="margin-bottom:8px">Недавние</div><div class="recent" id="rec"><div>пока пусто</div></div></div>
    </div>
  </div>
</main>
<script>
const S=[
 {id:'p2w',g:'Конвертация',ic:'🔄',t:'PDF → Word',sub:'Конвертация документов',mode:1,ftype:'pdf',prog:1,act:'run_p2w',actLabel:'🚀 Конвертировать',info:'Режим «Вся папка» обработает все PDF пакетно. Прогресс виден в реальном времени.',ph:'Путь к файлу или папке…'},
 {id:'w2p',g:'Конвертация',ic:'🔄',t:'Word → PDF',sub:'Конвертация документов',mode:1,ftype:'docx',prog:1,act:'run_w2p',actLabel:'🚀 Конвертировать',info:'Кириллица поддерживается (шрифт DejaVu встроен).',ph:'Путь к файлу или папке…'},
 {id:'img',g:'Конвертация',ic:'🖼',t:'Фото → PDF',sub:'Сборка из изображений',list:'multiimg',act:'run_img',actLabel:'🚀 Создать PDF',info:'JPG, PNG, BMP, GIF, TIFF. Порядок файлов = порядок страниц.'},
 {id:'split',g:'Инструменты',ic:'✂️',t:'Разделить PDF',sub:'На отдельные страницы',ftype:'pdf',act:'run_split',actLabel:'✂️ Разделить',info:'Каждая страница сохранится отдельным файлом в выбранную папку.'},
 {id:'merge',g:'Инструменты',ic:'🔗',t:'Объединить PDF',sub:'Склейка файлов',list:'multipdf',act:'run_merge',actLabel:'🔗 Объединить',info:'Файлы склеиваются в порядке добавления.'},
 {id:'comp',g:'Инструменты',ic:'📉',t:'Сжать PDF',sub:'Уменьшение размера',mode:1,ftype:'pdf',prog:1,act:'run_compress',actLabel:'📉 Сжать',info:'Без потери качества текста: сжатие потоков и очистка мусора.'},
 {id:'prot',g:'Инструменты',ic:'🛡',t:'Защита паролем',sub:'Шифрование 128 бит',ftype:'pdf',pass:1,act:'run_protect',actLabel:'🛡 Защитить',info:'Пароль потребуется для открытия файла.'},
 {id:'extr',g:'Страницы',ic:'📑',t:'Извлечь страницы',sub:'По диапазону',ftype:'pdf',range:1,act:'run_extract',actLabel:'📑 Извлечь',info:'Пустое поле = извлечь все страницы.'},
 {id:'rot',g:'Страницы',ic:'📐',t:'Повернуть страницы',sub:'Ориентация',ftype:'pdf',range:1,angle:1,act:'run_rotate',actLabel:'📐 Повернуть',info:'Пустое поле страниц = повернуть все.'},
 {id:'del',g:'Страницы',ic:'🗑',t:'Удалить страницы',sub:'Чистка документа',ftype:'pdf',range:1,act:'run_delete',actLabel:'🗑 Удалить',info:'Удалить все страницы нельзя — сработает защита.'},
 {id:'num',g:'Страницы',ic:'🔢',t:'Номера страниц',sub:'Нумерация',ftype:'pdf',act:'run_numbers',actLabel:'🔢 Добавить номера',info:'Номера ставятся внизу по центру.'}];
let cur=null,modeVal=0,files=[],ops=0,oks=0;
const $=id=>document.getElementById(id);
function nav(){let h='',g='';S.forEach(s=>{if(s.g!==g){g=s.g;h+=`<div class="grp">${g.toUpperCase()}</div>`}h+=`<button data-id="${s.id}">${s.ic} ${s.t}</button>`});$('nav').innerHTML=h;
 $('nav').querySelectorAll('button').forEach(b=>b.onclick=()=>render(b.dataset.id));}
function render(id){cur=S.find(s=>s.id===id);modeVal=0;files=[];
 $('ttl').textContent=cur.t;$('sub').textContent=cur.sub;$('tips').textContent=cur.info;
 $('nav').querySelectorAll('button').forEach(b=>b.classList.toggle('on',b.dataset.id===id));
 let h='';
 if(cur.mode)h+=`<div class="row"><span class="lbl">Режим:</span><div class="seg" id="seg"><button class="on" data-m="0">Один файл</button><button data-m="1">Вся папка</button></div></div>`;
 if(cur.list)h+=`<div class="listbox" id="lb">Список пуст — добавьте файлы</div><div class="row"><button class="btn yellow" id="add">＋ Добавить</button><button class="btn ghost" id="clr">Очистить</button></div>`;
 else h+=`<input id="path" placeholder="${cur.ph||'Путь к файлу…'}"><div class="row"><button class="btn dark" id="brw">📁 Обзор</button></div>`;
 if(cur.range)h+=`<div><div class="lbl" style="margin-bottom:6px">Страницы (напр. 1-3, 7)</div><input id="rng" placeholder="пусто = все"></div>`;
 if(cur.angle)h+=`<div class="row"><span class="lbl">Угол:</span><select id="ang"><option value="90">90°</option><option value="180">180°</option><option value="270">270°</option></select></div>`;
 if(cur.pass)h+=`<div><div class="lbl" style="margin-bottom:6px">Пароль</div><input id="pwd" type="password" placeholder="Придумайте пароль"></div>`;
 h+=`<div class="row" style="margin-top:4px"><button class="btn primary" id="go">${cur.actLabel}</button></div>`;
 if(cur.prog)h+=`<div class="progress"><i id="bar"></i></div>`;
 h+=`<div class="status" id="st">Готов к работе</div>`;
 $('panel').innerHTML=h;wire();}
function wire(){
 if($('seg'))$('seg').querySelectorAll('button').forEach(b=>b.onclick=()=>{modeVal=+b.dataset.m;$('seg').querySelectorAll('button').forEach(x=>x.classList.toggle('on',x===b));});
 if($('brw'))$('brw').onclick=async()=>{const k=(cur.mode&&modeVal===1)?'folder':cur.ftype;const p=await pywebview.api.browse(k);if(p){$('path').value=p;st('Файл выбран','ok');}};
 if($('add'))$('add').onclick=async()=>{const r=await pywebview.api.browse(cur.list);if(r){files=files.concat(r);draw();}};
 if($('clr'))$('clr').onclick=()=>{files=[];draw();};
 $('go').onclick=go;}
function draw(){$('lb').innerHTML=files.length?files.map(f=>f.split(/[\\/]/).pop()).join('<br>'):'Список пуст — добавьте файлы';}
function st(t,k){const e=$('st');if(e){e.textContent=t;e.className='status '+(k||'');}}
async function go(){const b=$('go');b.disabled=true;st('Выполняется…');if($('bar'))$('bar').style.width='0%';
 const d={path:$('path')?$('path').value:'',mode:modeVal,files:files,range:$('rng')?$('rng').value:'',angle:$('ang')?$('ang').value:90,pwd:$('pwd')?$('pwd').value:''};
 try{const r=await pywebview.api[cur.act](d);ops++;if(r.ok)oks++;
  $('s1').textContent=ops;$('s2').textContent=oks;
  if(r.ok){st('✓ '+r.msg,'ok');toast('✓ '+r.msg,false);if(r.file)rec(r.file);}
  else if(!r.silent){st('✗ '+r.msg,'err');toast('✗ '+r.msg,true);}
  else st('Отменено');
 }catch(e){st('✗ '+e,'err');}
 b.disabled=false;}
function rec(f){const e=$('rec');if(e.querySelector('div')&&!e.querySelector('div+div'))e.innerHTML='';e.insertAdjacentHTML('afterbegin',`<div>${f}</div>`);while(e.children.length>6)e.lastChild.remove();}
function toast(m,err){const t=document.createElement('div');t.className='toast'+(err?' err':'');t.textContent=m;document.body.appendChild(t);setTimeout(()=>t.remove(),4500);}
window.UI={prog:(p,l)=>{const b=$('bar');if(b)b.style.width=(p*100)+'%';if(l)st(l);},};
async function askUpdate(){
 const r=await pywebview.api.check_update();
 if(!r.ok){toast('Нет связи с сервером обновлений',true);return;}
 if(!r.new){toast('У вас последняя версия ✓',false);return;}
 if(confirm(`Доступна версия ${r.new}\n${r.notes||''}\n\nСкачать и установить сейчас?`)){
   st('Обновление…');
   const res=await pywebview.api.do_update(r.url);
   if(!res.ok)toast('Ошибка загрузки: '+res.msg,true);
 }}
window.addEventListener('pywebviewready',async()=>{
 nav();render('p2w');
 pywebview.api.get_version().then(v=>{$('ver').textContent='v'+v;});
 $('upd').onclick=askUpdate;
 const r=await pywebview.api.check_update();
 if(r.ok&&r.new)toast(`⬆ Доступна версия ${r.new} — нажмите «Проверить обновления»`,false);
});
</script></body></html>'''

if __name__ == '__main__':
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('ekadvladsheglov.pdfconverter.3.1')
    except Exception:
        pass
    api = Api()
    WIN = webview.create_window(
        'PDF Конвертер & Инструменты v3.1',
        html=HTML,
        js_api=api,
        width=1200,
        height=780,
        min_size=(1000, 660),
        background_color='#F2F1EE'
    )
    webview.start()
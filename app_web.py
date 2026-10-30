# app_web.py
import os, io, sys, json
import urllib.request
import tempfile
import subprocess
import logging
from pathlib import Path
import webview
from pdf2docx import Converter
from docx import Document
from pypdf import PdfReader, PdfWriter
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

# --- НАСТРОЙКА ЛОГИРОВАНИЯ ---
log_file = Path(os.getenv('LOCALAPPDATA', tempfile.gettempdir())) / 'PDF_Converter' / 'app.log'
log_file.parent.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    filename=log_file, level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

APP_VERSION = "3.1.1"
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
            except Exception as e:
                logging.error(f"Font registration failed: {e}")
        _font_done = True
    return 'DejaVu' if 'DejaVu' in pdfmetrics.getRegisteredFontNames() else 'Helvetica'


def parse_range(text, total):
    out = set()
    for part in text.replace(' ', '').split(','):
        if not part: continue
        try:
            if '-' in part:
                a, b = part.split('-', 1)
                a, b = int(a), int(b)
                if a > b: a, b = b, a
                out |= {p - 1 for p in range(a, b + 1) if 1 <= p <= total}
            else:
                p = int(part)
                if 1 <= p <= total: out.add(p - 1)
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
    def browse(self, kind):
        try:
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
            name, mask = ft_map.get(kind, ('Все файлы', '*.*'))
            multi = kind.startswith('multi')
            
            if FD:
                try:
                    r = WIN.create_file_dialog(FD.OPEN, file_types=(f'{name} ({mask})',), allow_multiple=multi)
                except TypeError:
                    r = WIN.create_file_dialog(FD.OPEN, file_types=(f'{name} ({mask})',))
            else:
                r = WIN.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=multi, file_types=(f'{name} ({mask})',))
                
            if not r: 
                return None
            return list(r) if multi else r[0]
        except Exception as e:
            logging.error(f"Ошибка в методе browse: {e}")
            return f"ERROR: {e}"

    def _save(self, name, ext):
        FD = getattr(webview, 'FileDialog', None)
        if FD:
            r = WIN.create_file_dialog(FD.SAVE, save_filename=name + ext)
        else:
            r = WIN.create_file_dialog(webview.SAVE_DIALOG, save_filename=name + ext)
        if not r: return None
        p = r if isinstance(r, str) else r[0]
        return str(Path(p).with_suffix(ext)) if not p.lower().endswith(ext) else p

    def open_folder(self, filepath):
        try:
            folder = str(Path(filepath).parent)
            os.startfile(folder)
            return True
        except Exception as e:
            logging.error(f"Open folder failed: {e}")
            return False

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
            tmp_exe = os.path.join(tempfile.gettempdir(), 'PDF_Converter_Update.exe')
            req = urllib.request.Request(url, headers={'User-Agent': 'PDFConverter'})
            with urllib.request.urlopen(req, timeout=20) as r, open(tmp_exe, 'wb') as f:
                total = int(r.headers.get('Content-Length') or 0)
                done = 0
                while True:
                    chunk = r.read(256 * 1024)
                    if not chunk: break
                    f.write(chunk)
                    done += len(chunk)
                    if total: js(f'UI.prog({done/total}, "Загрузка… {done//1048576} МБ")')
            
            js('UI.prog(1, "Подготовка к установке…")')
            
            # ИСПРАВЛЕНИЕ: Используем bat-файл как прослойку для безопасного обновления
            bat_path = os.path.join(tempfile.gettempdir(), 'run_update.bat')
            with open(bat_path, 'w', encoding='utf-8') as b:
                b.write('@echo off\n')
                b.write('timeout /t 2 /nobreak > NUL\n')   # Ждем 2 секунды, чтобы старый _MEI освободился
                b.write(f'start "" "{tmp_exe}"\n')        # Запускаем установщик
                b.write('del "%~f0"\n')                   # Удаляем bat-файл после запуска
            
            # Запускаем bat-файл скрыто (без черного окна консоли)
            subprocess.Popen(['cmd.exe', '/c', bat_path], creationflags=subprocess.CREATE_NO_WINDOW)
            
            WIN.destroy()
            return {'ok': True}
        except Exception as e:
            logging.error(f"Update failed: {e}")
            return {'ok': False, 'msg': str(e)}

    def run_p2w(self, d):
        try:
            if d.get('mode') == 1:
                FD = getattr(webview, 'FileDialog', None)
                out_dir = WIN.create_file_dialog(FD.FOLDER) if FD else WIN.create_file_dialog(webview.FOLDER_DIALOG)
                if not out_dir: return {'ok': False, 'silent': True}
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
                    except Exception as e:
                        logging.error(f"Batch p2w error {f}: {e}")
                    js(f'UI.prog({(i+1)/n}, "Файл {i+1}/{n}: {f}")')
                return {'ok': True, 'msg': f'Готово: {ok}/{n}', 'file': out_dir, 'is_dir': True}
            out = self._save('converted', '.docx')
            if not out: return {'ok': False, 'silent': True}
            js('UI.prog(0.4, "Конвертация…")')
            cv = Converter(d['path'])
            cv.convert(out)
            cv.close()
            js('UI.prog(1, "")')
            return {'ok': True, 'msg': 'Сохранено', 'file': out}
        except Exception as e:
            logging.error(f"run_p2w error: {e}")
            return {'ok': False, 'msg': str(e)}

    def run_w2p(self, d):
        try:
            if d.get('mode') == 1:
                FD = getattr(webview, 'FileDialog', None)
                out_dir = WIN.create_file_dialog(FD.FOLDER) if FD else WIN.create_file_dialog(webview.FOLDER_DIALOG)
                if not out_dir: return {'ok': False, 'silent': True}
                out_dir = out_dir[0]
                files = [f for f in os.listdir(d['path']) if f.lower().endswith('.docx')]
                n = len(files)
                ok = 0
                for i, f in enumerate(files):
                    try:
                        self._w2p(os.path.join(d['path'], f), os.path.join(out_dir, os.path.splitext(f)[0] + '.pdf'))
                        ok += 1
                    except Exception as e:
                        logging.error(f"Batch w2p error {f}: {e}")
                    js(f'UI.prog({(i+1)/n}, "Файл {i+1}/{n}: {f}")')
                return {'ok': True, 'msg': f'Готово: {ok}/{n}', 'file': out_dir, 'is_dir': True}
            out = self._save('converted', '.pdf')
            if not out: return {'ok': False, 'silent': True}
            js('UI.prog(0.4, "Конвертация…")')
            self._w2p(d['path'], out)
            js('UI.prog(1, "")')
            return {'ok': True, 'msg': 'Сохранено', 'file': out}
        except Exception as e:
            logging.error(f"run_w2p error: {e}")
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
                if not word: continue
                while word and c.stringWidth(word, fn, fs) > mw:
                    lo, hi = 1, len(word)
                    while lo < hi:
                        mid = (lo + hi + 1) // 2
                        if c.stringWidth(word[:mid], fn, fs) <= mw: lo = mid
                        else: hi = mid - 1
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

    def run_img(self, d):
        try:
            out = self._save('images', '.pdf')
            if not out: return {'ok': False, 'silent': True}
            imgs = []
            for p in d['files']:
                im = Image.open(p)
                imgs.append(im.convert('RGB') if im.mode != 'RGB' else im)
            imgs[0].save(out, 'PDF', save_all=True, append_images=imgs[1:], resolution=150)
            return {'ok': True, 'msg': f'PDF из {len(imgs)} изображений', 'file': out}
        except Exception as e:
            logging.error(f"run_img error: {e}")
            return {'ok': False, 'msg': str(e)}

    def run_split(self, d):
        try:
            FD = getattr(webview, 'FileDialog', None)
            out_dir = WIN.create_file_dialog(FD.FOLDER) if FD else WIN.create_file_dialog(webview.FOLDER_DIALOG)
            if not out_dir: return {'ok': False, 'silent': True}
            out_dir = out_dir[0]
            rd = PdfReader(d['path'])
            base = Path(d['path']).stem
            for i, pg in enumerate(rd.pages):
                wr = PdfWriter()
                wr.add_page(pg)
                with open(os.path.join(out_dir, f'{base}_page_{i+1}.pdf'), 'wb') as f:
                    wr.write(f)
            return {'ok': True, 'msg': f'Страниц: {len(rd.pages)}', 'file': out_dir, 'is_dir': True}
        except Exception as e:
            logging.error(f"run_split error: {e}")
            return {'ok': False, 'msg': str(e)}

    def run_merge(self, d):
        try:
            out = self._save('merged', '.pdf')
            if not out: return {'ok': False, 'silent': True}
            wr = PdfWriter()
            for p in d['files']:
                for pg in PdfReader(p).pages:
                    wr.add_page(pg)
            with open(out, 'wb') as f:
                wr.write(f)
            return {'ok': True, 'msg': f'Объединено: {len(d["files"])}', 'file': out}
        except Exception as e:
            logging.error(f"run_merge error: {e}")
            return {'ok': False, 'msg': str(e)}

    def run_compress(self, d):
        try:
            if d.get('mode') == 1:
                FD = getattr(webview, 'FileDialog', None)
                out_dir = WIN.create_file_dialog(FD.FOLDER) if FD else WIN.create_file_dialog(webview.FOLDER_DIALOG)
                if not out_dir: return {'ok': False, 'silent': True}
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
                    except Exception as e:
                        logging.error(f"Batch compress error {f}: {e}")
                    js(f'UI.prog({(i+1)/n}, "Файл {i+1}/{n}: {f}")')
                return {'ok': True, 'msg': f'Готово: {ok}/{n}', 'file': out_dir, 'is_dir': True}
            out = self._save('compressed', '.pdf')
            if not out: return {'ok': False, 'silent': True}
            js('UI.prog(0.4, "Сжатие…")')
            doc = fitz.open(d['path'])
            doc.save(out, garbage=4, deflate=True, clean=True)
            doc.close()
            o, c2 = os.path.getsize(d['path']) / 1024, os.path.getsize(out) / 1024
            pct = (1 - c2 / o) * 100 if o else 0
            js('UI.prog(1, "")')
            return {'ok': True, 'msg': f'Экономия {pct:.1f}% ({o:.0f}→{c2:.0f} КБ)', 'file': out}
        except Exception as e:
            logging.error(f"run_compress error: {e}")
            return {'ok': False, 'msg': str(e)}

    def run_protect(self, d):
        try:
            if not d.get('pwd'): return {'ok': False, 'msg': 'Введите пароль'}
            out = self._save('protected', '.pdf')
            if not out: return {'ok': False, 'silent': True}
            wr = PdfWriter()
            for pg in PdfReader(d['path']).pages:
                wr.add_page(pg)
            wr.encrypt(user_password=d['pwd'], owner_password=d['pwd'], use_128bit=True)
            with open(out, 'wb') as f:
                wr.write(f)
            return {'ok': True, 'msg': 'Защищено паролем', 'file': out}
        except Exception as e:
            logging.error(f"run_protect error: {e}")
            return {'ok': False, 'msg': str(e)}

    def run_extract(self, d):
        try:
            rd = PdfReader(d['path'])
            total = len(rd.pages)
            idx = parse_range(d.get('range', ''), total) or list(range(total))
            out = self._save('extracted', '.pdf')
            if not out: return {'ok': False, 'silent': True}
            wr = PdfWriter()
            for i in idx: wr.add_page(rd.pages[i])
            with open(out, 'wb') as f: wr.write(f)
            return {'ok': True, 'msg': f'Извлечено: {len(idx)}', 'file': out}
        except Exception as e:
            logging.error(f"run_extract error: {e}")
            return {'ok': False, 'msg': str(e)}

    def run_rotate(self, d):
        try:
            rd = PdfReader(d['path'])
            total = len(rd.pages)
            idx = parse_range(d.get('range', ''), total) or list(range(total))
            out = self._save('rotated', '.pdf')
            if not out: return {'ok': False, 'silent': True}
            wr = PdfWriter()
            for i, pg in enumerate(rd.pages):
                if i in idx: pg.rotate(int(d.get('angle', 90)))
                wr.add_page(pg)
            with open(out, 'wb') as f: wr.write(f)
            return {'ok': True, 'msg': f'Повернуто: {len(idx)}', 'file': out}
        except Exception as e:
            logging.error(f"run_rotate error: {e}")
            return {'ok': False, 'msg': str(e)}

    def run_delete(self, d):
        try:
            rd = PdfReader(d['path'])
            total = len(rd.pages)
            idx = parse_range(d.get('range', ''), total)
            if not idx: return {'ok': False, 'msg': 'Укажите страницы'}
            out = self._save('cleaned', '.pdf')
            if not out: return {'ok': False, 'silent': True}
            wr = PdfWriter()
            deln = 0
            for i, pg in enumerate(rd.pages):
                if i in idx: deln += 1
                else: wr.add_page(pg)
            if len(wr.pages) == 0: return {'ok': False, 'msg': 'Нельзя удалить все страницы'}
            with open(out, 'wb') as f: wr.write(f)
            return {'ok': True, 'msg': f'Удалено: {deln}, осталось: {len(wr.pages)}', 'file': out}
        except Exception as e:
            logging.error(f"run_delete error: {e}")
            return {'ok': False, 'msg': str(e)}

    def run_numbers(self, d):
        try:
            out = self._save('numbered', '.pdf')
            if not out: return {'ok': False, 'silent': True}
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
            with open(out, 'wb') as f: wr.write(f)
            return {'ok': True, 'msg': 'Нумерация добавлена', 'file': out}
        except Exception as e:
            logging.error(f"run_numbers error: {e}")
            return {'ok': False, 'msg': str(e)}


if __name__ == '__main__':
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('ekadvladsheglov.pdfconverter.3.1')
    except Exception:
        pass
    
    try:
        api = Api()
        WIN = webview.create_window(
            'PDF Конвертер & Инструменты v3.1.1',
            url=res('web/index.html'),
            js_api=api,
            width=1200,
            height=780,
            min_size=(1000, 660),
            background_color='#F2F1EE'
        )
        webview.start(debug=True)
    except Exception as e:
        import traceback
        print("\n" + "="*50)
        print("КРИТИЧЕСКАЯ ОШИБКА ПРИ ЗАПУСКЕ:")
        traceback.print_exc()
        print("="*50 + "\n")
        input("Нажмите Enter, чтобы закрыть окно...") 
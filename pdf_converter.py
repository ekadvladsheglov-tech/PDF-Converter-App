import os
import sys
import io
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
from pdf2docx import Converter
from docx import Document
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from PIL import Image
import fitz

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False


def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ==================== ДИЗАЙН-ТОКЕНЫ ====================
COLORS = {
    "bg_page": "#FFFFFF",
    "bg_surface": "#F5F5F3",
    "bg_card_light": "#ECECE9",
    "bg_dark": "#171717",
    "bg_dark_elev": "#2A2A2A",
    "bg_dark_hover": "#3A3A3A",
    "accent_coral": "#F2694B",
    "accent_yellow": "#F2E500",
    "success": "#2EBD59",
    "text_primary": "#171717",
    "text_secondary": "#8A8A86",
    "text_on_dark": "#FFFFFF",
    "text_on_dark_muted": "#9C9C99",
}

RADIUS = {
    "container": 20,
    "card": 16,
    "pill": 999,
    "button": 10,
}

FONT_FAMILY = "Segoe UI"
FONT = (FONT_FAMILY, 10)
FONT_BOLD = (FONT_FAMILY, 10, "bold")
FONT_TITLE = (FONT_FAMILY, 16, "bold")
FONT_SECTION = (FONT_FAMILY, 9, "bold")
FONT_SMALL = (FONT_FAMILY, 9)


# ==================== ВСПОМОГАТЕЛЬНЫЕ КЛАССЫ ====================
class PageParser:
    @staticmethod
    def parse(text_input, total_pages):
        pages = set()
        parts = text_input.replace(" ", "").split(",")
        for part in parts:
            if not part:
                continue
            if "-" in part:
                try:
                    start, end = part.split("-", 1)
                    start, end = int(start), int(end)
                    if start > end:
                        start, end = end, start
                    for p in range(start, end + 1):
                        if 1 <= p <= total_pages:
                            pages.add(p - 1)
                except ValueError:
                    continue
            else:
                try:
                    p = int(part)
                    if 1 <= p <= total_pages:
                        pages.add(p - 1)
                except ValueError:
                    continue
        return sorted(list(pages))


class BatchProcessor:
    @staticmethod
    def process_folder(input_folder, output_folder, file_ext, out_ext,
                       process_func, progress_callback):
        """ИСПРАВЛЕНО: расширение результата передаётся явно (out_ext),
        вместо гадания по имени функции."""
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        files = [f for f in os.listdir(input_folder)
                 if f.lower().endswith(file_ext) and not f.startswith("~$")]
        total = len(files)
        if total == 0:
            return 0, 0
        success_count = 0
        error_count = 0
        for i, filename in enumerate(files):
            in_path = os.path.join(input_folder, filename)
            base_name = os.path.splitext(filename)[0]
            out_path = os.path.join(output_folder, base_name + out_ext)
            try:
                process_func(in_path, out_path)
                success_count += 1
            except Exception:
                error_count += 1
            progress_callback(i + 1, total, filename)
        return success_count, error_count


def ensure_pdf_font():
    """ИСПРАВЛЕНО: шрифт регистрируется ОДИН раз, иначе краш на пакете."""
    try:
        if 'DejaVu' in pdfmetrics.getRegisteredFontNames():
            return 'DejaVu'
        font_path = get_resource_path("DejaVuSans.ttf")
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('DejaVu', font_path))
            return 'DejaVu'
    except Exception:
        pass
    return 'Helvetica'


# ==================== ВЫБОР БАЗОВОГО ОКНА ====================
ctk.set_appearance_mode("Light")

# ИСПРАВЛЕНО: никакого множественного наследования.
# Сначала "пробник": если tkdnd не загрузится (частый случай в PyInstaller),
# молча откатываемся на чистый ctk.CTk без краха приложения.
USE_DND = False
BaseWindow = ctk.CTk
if HAS_DND:
    _probe = None
    try:
        _probe = TkinterDnD.Tk()
        _probe.withdraw()
        _probe.destroy()
        BaseWindow = TkinterDnD.Tk
        USE_DND = True
    except Exception:
        HAS_DND = False
        if _probe is not None:
            try:
                _probe.destroy()
            except Exception:
                pass


# ==================== ГЛАВНОЕ ОКНО ====================
class PDFConverterApp(BaseWindow):
    def __init__(self):
        super().__init__()
        self.title("PDF Конвертер & Инструменты v3.0")
        self.resizable(False, False)

        # ИСПРАВЛЕНО: у TkinterDnD.Tk нет fg_color, используем bg.
        # geometry() работает у обеих баз, вызываем один раз.
        if USE_DND:
            self.configure(bg=COLORS["bg_page"])
        else:
            self.configure(fg_color=COLORS["bg_page"])
        self.geometry("1000x650")

        if USE_DND:
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind('<<Drop>>', self.on_files_dropped)
            except Exception:
                pass

        main_container = ctk.CTkFrame(self, fg_color=COLORS["bg_surface"],
                                      corner_radius=RADIUS["container"])
        main_container.pack(fill="both", expand=True, padx=8, pady=8)
        main_container.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(1, weight=1)

        self.frames = {}
        self.menu_buttons = {}

        self.create_sidebar(main_container)

        self.content_area = ctk.CTkFrame(main_container, fg_color="transparent", corner_radius=0)
        self.content_area.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        self.create_pdf_to_word_frame()
        self.create_word_to_pdf_frame()
        self.create_images_to_pdf_frame()
        self.create_split_frame()
        self.create_merge_frame()
        self.create_compress_frame()
        self.create_protect_frame()
        self.create_extract_frame()
        self.create_rotate_frame()
        self.create_delete_frame()
        self.create_numbers_frame()

        for frame in self.frames.values():
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("pdf_to_word")

    # ==================== БОКОВОЕ МЕНЮ ====================
    def create_sidebar(self, parent):
        sidebar = ctk.CTkFrame(parent, width=200, fg_color=COLORS["bg_dark"],
                               corner_radius=RADIUS["container"])
        sidebar.grid(row=0, column=0, sticky="nsw", padx=8, pady=8)
        sidebar.grid_propagate(False)

        header_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        header_frame.pack(fill="x", padx=15, pady=(15, 10))

        ctk.CTkLabel(header_frame, text="📄", font=(FONT_FAMILY, 22)).pack(side="left", padx=(0, 10))

        title_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_frame.pack(side="left")
        ctk.CTkLabel(title_frame, text="PDF Tools", font=FONT_TITLE,
                     text_color=COLORS["text_on_dark"]).pack(anchor="w")
        ctk.CTkLabel(title_frame, text="v3.0", font=FONT_SMALL,
                     text_color=COLORS["text_on_dark_muted"]).pack(anchor="w")

        self._add_section_label(sidebar, "КОНВЕРТАЦИЯ")
        self._add_menu_button(sidebar, "pdf_to_word", "PDF → Word", "🔄")
        self._add_menu_button(sidebar, "word_to_pdf", "Word → PDF", "🔄")
        self._add_menu_button(sidebar, "images_to_pdf", "Фото → PDF", "🖼")

        self._add_section_label(sidebar, "ИНСТРУМЕНТЫ")
        self._add_menu_button(sidebar, "split", "Разделить", "✂️")
        self._add_menu_button(sidebar, "merge", "Объединить", "🔗")
        self._add_menu_button(sidebar, "compress", "Сжать", "📉")
        self._add_menu_button(sidebar, "protect", "Пароль", "🛡")

        self._add_section_label(sidebar, "СТРАНИЦЫ")
        self._add_menu_button(sidebar, "extract", "Извлечь", "📑")
        self._add_menu_button(sidebar, "rotate", "Повернуть", "📐")
        self._add_menu_button(sidebar, "delete", "Удалить", "🗑")
        self._add_menu_button(sidebar, "numbers", "Номера", "🔢")

        bottom_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        bottom_frame.pack(side="bottom", fill="x", padx=15, pady=15)
        ctk.CTkLabel(bottom_frame, text="✓ Локально", font=FONT_SMALL,
                     text_color=COLORS["success"]).pack(anchor="w")
        ctk.CTkLabel(bottom_frame, text="✓ Безопасно", font=FONT_SMALL,
                     text_color=COLORS["success"]).pack(anchor="w")

    def _add_section_label(self, parent, text):
        ctk.CTkLabel(parent, text=text, font=FONT_SECTION,
                     text_color=COLORS["text_on_dark_muted"]).pack(anchor="w", padx=15, pady=(12, 4))

    def _add_menu_button(self, parent, frame_name, text, icon):
        btn = ctk.CTkButton(
            parent,
            text=f"{icon}  {text}",
            anchor="w",
            fg_color="transparent",
            text_color=COLORS["text_on_dark"],
            hover_color=COLORS["bg_dark_hover"],
            height=32,
            corner_radius=RADIUS["button"],
            font=FONT,
            command=lambda fn=frame_name: self.show_frame(fn)
        )
        btn.pack(fill="x", padx=8, pady=1)
        self.menu_buttons[frame_name] = btn

    def show_frame(self, frame_name):
        for btn in self.menu_buttons.values():
            btn.configure(fg_color="transparent", text_color=COLORS["text_on_dark"])
        self.menu_buttons[frame_name].configure(fg_color=COLORS["accent_coral"], text_color="white")
        self.frames[frame_name].tkraise()

    # ==================== ОБЩИЕ ВИДЖЕТЫ ====================
    def _create_content_frame(self, title, subtitle=""):
        frame = ctk.CTkFrame(self.content_area, fg_color=COLORS["bg_card_light"],
                             corner_radius=RADIUS["card"])
        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(15, 12))
        ctk.CTkLabel(header, text=title, font=FONT_TITLE,
                     text_color=COLORS["text_primary"]).pack(side="left")
        if subtitle:
            ctk.CTkLabel(header, text=subtitle, font=FONT_SMALL,
                         text_color=COLORS["text_secondary"]).pack(side="left", padx=(10, 0))
        return frame

    def _create_pill_button(self, parent, text, command,
                            color=COLORS["accent_coral"], width=180, height=34):
        return ctk.CTkButton(
            parent, text=text, command=command,
            fg_color=color, hover_color=self._darken_color(color),
            height=height, width=width,
            corner_radius=RADIUS["pill"], font=FONT_BOLD
        )

    def _create_entry(self, parent, width=500, placeholder="", height=36):
        return ctk.CTkEntry(
            parent, width=width, height=height,
            corner_radius=RADIUS["pill"],
            fg_color=COLORS["bg_page"],
            border_color=COLORS["text_secondary"],
            border_width=1, font=FONT,
            placeholder_text=placeholder,
            text_color=COLORS["text_primary"]
        )

    def _create_segmented_button(self, parent, values, command):
        return ctk.CTkSegmentedButton(
            parent, values=values, command=command,
            fg_color=COLORS["bg_page"],
            selected_color=COLORS["accent_coral"],
            selected_hover_color=self._darken_color(COLORS["accent_coral"]),
            unselected_color=COLORS["bg_card_light"],
            height=32, corner_radius=RADIUS["pill"], font=FONT
        )

    def _create_info_bar(self, frame, text):
        info_frame = ctk.CTkFrame(frame, fg_color=COLORS["bg_page"], corner_radius=RADIUS["button"])
        info_frame.pack(fill="x", padx=20, pady=(15, 15))
        ctk.CTkLabel(info_frame, text=text, font=FONT_SMALL,
                     text_color=COLORS["text_secondary"]).pack(padx=12, pady=8, anchor="w")

    def _darken_color(self, hex_color, factor=0.85):
        hex_color = hex_color.lstrip('#')
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        r, g, b = int(r * factor), int(g * factor), int(b * factor)
        return f"#{r:02x}{g:02x}{b:02x}"

    # ==================== PDF -> WORD ====================
    def create_pdf_to_word_frame(self):
        frame = self._create_content_frame("PDF → Word", "Конвертация документов")
        self.frames["pdf_to_word"] = frame

        top_bar = ctk.CTkFrame(frame, fg_color="transparent")
        top_bar.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(top_bar, text="Режим:", font=FONT,
                     text_color=COLORS["text_secondary"]).pack(side="left", padx=(0, 10))
        self.mode_p2w = self._create_segmented_button(top_bar, ["Один файл", "Вся папка"], self.toggle_p2w_mode)
        self.mode_p2w.pack(side="left")
        self.mode_p2w.set("Один файл")

        self.entry_p2w = self._create_entry(frame, placeholder="Выберите файл или папку...")
        self.entry_p2w.pack(pady=(0, 8), padx=20, fill="x")

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(pady=(0, 12))
        self._create_pill_button(btn_frame, "📁 Обзор", self.browse_p2w, COLORS["bg_dark"], 120).pack(side="left", padx=5)
        self._create_pill_button(btn_frame, "🚀 Конвертировать", self.start_p2w, COLORS["accent_coral"], 170).pack(side="left", padx=5)

        self.progress_p2w = ctk.CTkProgressBar(frame, width=460, height=6,
                                               corner_radius=RADIUS["pill"],
                                               progress_color=COLORS["accent_coral"])
        self.progress_p2w.pack(pady=(0, 6))
        self.progress_p2w.set(0)

        self.label_status_p2w = ctk.CTkLabel(frame, text="Готов к работе", font=FONT_SMALL,
                                             text_color=COLORS["text_secondary"])
        self.label_status_p2w.pack()

        self._create_info_bar(frame, "💡 Поддерживает пакетную обработку нескольких файлов")

    def toggle_p2w_mode(self, value):
        self.entry_p2w.delete(0, "end")
        self.entry_p2w.configure(placeholder_text="Выберите папку..." if value == "Вся папка" else "Выберите файл...")

    def browse_p2w(self):
        if self.mode_p2w.get() == "Один файл":
            path = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")])
        else:
            path = filedialog.askdirectory(title="Выберите папку с PDF файлами")
        if path:
            self.entry_p2w.delete(0, "end")
            self.entry_p2w.insert(0, path)

    def start_p2w(self):
        path = self.entry_p2w.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Ошибка", "Путь не выбран!")
            return
        if self.mode_p2w.get() == "Один файл":
            out_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word файлы", "*.docx")])
            if not out_path:
                return
            self.label_status_p2w.configure(text="⏳ Конвертация...")
            self.progress_p2w.set(0.5)
            threading.Thread(target=self._run_single_p2w, args=(path, out_path), daemon=True).start()
        else:
            out_dir = filedialog.askdirectory(title="Куда сохранить результаты?")
            if not out_dir:
                return
            self.label_status_p2w.configure(text="⏳ Пакетная обработка...")
            self.progress_p2w.set(0)
            threading.Thread(target=self._run_batch_p2w, args=(path, out_dir), daemon=True).start()

    def _single_p2w_logic(self, in_path, out_path):
        cv = Converter(in_path)
        cv.convert(out_path)
        cv.close()

    def _run_single_p2w(self, in_path, out_path):
        try:
            self._single_p2w_logic(in_path, out_path)
            self.after(0, lambda: self.progress_p2w.set(1.0))
            self.after(0, lambda: self.label_status_p2w.configure(text="✓ Успешно!", text_color=COLORS["success"]))
            self.after(0, lambda p=out_path: messagebox.showinfo("Успех", f"Сохранено:\n{p}"))
        except Exception as e:
            err_text = str(e)  # ИСПРАВЛЕНО: копируем до выхода из except
            self.after(0, lambda: self.label_status_p2w.configure(text="✗ Ошибка!", text_color=COLORS["accent_coral"]))
            self.after(0, lambda t=err_text: messagebox.showerror("Ошибка", t))

    def _run_batch_p2w(self, in_dir, out_dir):
        def update_progress(current, total, filename):
            self.after(0, lambda p=current/total, c=current, t=total, f=filename:
                       self._update_batch_ui_p2w(p, c, t, f))
        success, errors = BatchProcessor.process_folder(in_dir, out_dir, ".pdf", ".docx",
                                                        self._single_p2w_logic, update_progress)
        self.after(0, lambda s=success, e=errors, d=out_dir: self._finish_batch_p2w(s, e, d))

    def _update_batch_ui_p2w(self, percent, current, total, filename):
        self.progress_p2w.set(percent)
        self.label_status_p2w.configure(text=f"📄 Файл {current}/{total}: {filename}")

    def _finish_batch_p2w(self, success, errors, out_dir):
        self.progress_p2w.set(1.0)
        self.label_status_p2w.configure(text=f"✓ Готово! Успешно: {success}, Ошибок: {errors}",
                                        text_color=COLORS["success"])
        messagebox.showinfo("Пакетная обработка",
                            f"Готово!\nУспешно: {success}\nС ошибками: {errors}\n\nПапка: {out_dir}")

    # ==================== WORD -> PDF ====================
    def create_word_to_pdf_frame(self):
        frame = self._create_content_frame("Word → PDF", "Конвертация документов")
        self.frames["word_to_pdf"] = frame

        top_bar = ctk.CTkFrame(frame, fg_color="transparent")
        top_bar.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(top_bar, text="Режим:", font=FONT,
                     text_color=COLORS["text_secondary"]).pack(side="left", padx=(0, 10))
        self.mode_w2p = self._create_segmented_button(top_bar, ["Один файл", "Вся папка"], self.toggle_w2p_mode)
        self.mode_w2p.pack(side="left")
        self.mode_w2p.set("Один файл")

        self.entry_w2p = self._create_entry(frame, placeholder="Выберите файл или папку...")
        self.entry_w2p.pack(pady=(0, 8), padx=20, fill="x")

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(pady=(0, 12))
        self._create_pill_button(btn_frame, "📁 Обзор", self.browse_w2p, COLORS["bg_dark"], 120).pack(side="left", padx=5)
        self._create_pill_button(btn_frame, "🚀 Конвертировать", self.start_w2p, COLORS["accent_coral"], 170).pack(side="left", padx=5)

        self.progress_w2p = ctk.CTkProgressBar(frame, width=460, height=6,
                                               corner_radius=RADIUS["pill"],
                                               progress_color=COLORS["accent_coral"])
        self.progress_w2p.pack(pady=(0, 6))
        self.progress_w2w = self.progress_w2p
        self.progress_w2p.set(0)

        self.label_status_w2p = ctk.CTkLabel(frame, text="Готов к работе", font=FONT_SMALL,
                                             text_color=COLORS["text_secondary"])
        self.label_status_w2p.pack()

        self._create_info_bar(frame, "💡 Сохраняет структуру текста документа")

    def toggle_w2p_mode(self, value):
        self.entry_w2p.delete(0, "end")
        self.entry_w2p.configure(placeholder_text="Выберите папку..." if value == "Вся папка" else "Выберите файл...")

    def browse_w2p(self):
        if self.mode_w2p.get() == "Один файл":
            path = filedialog.askopenfilename(filetypes=[("Word файлы", "*.docx")])
        else:
            path = filedialog.askdirectory(title="Выберите папку с Word файлами")
        if path:
            self.entry_w2p.delete(0, "end")
            self.entry_w2w.insert(0, path)

    def start_w2p(self):
        path = self.entry_w2w.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Ошибка", "Путь не выбран!")
            return
        if self.mode_w2p.get() == "Один файл":
            out_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
            if not out_path:
                return
            self.label_status_w2p.configure(text="⏳ Конвертация...")
            self.progress_w2p.set(0.5)
            threading.Thread(target=self._run_single_w2p, args=(path, out_path), daemon=True).start()
        else:
            out_dir = filedialog.askdirectory(title="Куда сохранить результаты?")
            if not out_dir:
                return
            self.label_status_w2p.configure(text="⏳ Пакетная обработка...")
            self.progress_w2p.set(0)
            threading.Thread(target=self._run_batch_w2p, args=(path, out_dir), daemon=True).start()

    def _single_w2p_logic(self, word_path, pdf_path):
        doc = Document(word_path)
        c = canvas.Canvas(pdf_path, pagesize=A4)
        width, height = A4
        font_name = ensure_pdf_font()
        font_size = 12
        x = 20 * mm
        y = height - (20 * mm)
        line_height = 6 * mm
        max_width = width - (40 * mm)
        bottom_margin = 20 * mm
        c.setFont(font_name, font_size)
        for para in doc.paragraphs:
            text = para.text
            if not text.strip():
                y -= line_height
                if y < bottom_margin:
                    c.showPage()
                    c.setFont(font_name, font_size)
                    y = height - (20 * mm)
                continue
            words = [w for w in text.split(' ') if w]
            safe_words = []
            for word in words:
                safe_words.extend(self._split_long_word(word, c, font_name, font_size, max_width))
            current_line = ""
            for word in safe_words:
                test_line = current_line + " " + word if current_line else word
                if c.stringWidth(test_line, font_name, font_size) <= max_width:
                    current_line = test_line
                else:
                    c.drawString(x, y, current_line)
                    y -= line_height
                    if y < bottom_margin:
                        c.showPage()
                        c.setFont(font_name, font_size)
                        y = height - (20 * mm)
                    current_line = word
            if current_line:
                c.drawString(x, y, current_line)
                y -= line_height
            if y < bottom_margin:
                c.showPage()
                c.setFont(font_name, font_size)
                y = height - (20 * mm)
        c.save()

    def _split_long_word(self, word, c, font_name, font_size, max_width):
        parts = []
        while word and c.stringWidth(word, font_name, font_size) > max_width:
            low, high = 1, len(word)
            while low < high:
                mid = (low + high + 1) // 2
                if c.stringWidth(word[:mid], font_name, font_size) <= max_width:
                    low = mid
                else:
                    high = mid - 1
            parts.append(word[:low])
            word = word[low:]
        if word:
            parts.append(word)
        return parts

    def _run_single_w2p(self, in_path, out_path):
        try:
            self._single_w2p_logic(in_path, out_path)
            self.after(0, lambda: self.progress_w2p.set(1.0))
            self.after(0, lambda: self.label_status_w2p.configure(text="✓ Успешно!", text_color=COLORS["success"]))
            self.after(0, lambda p=out_path: messagebox.showinfo("Успех", f"Сохранено:\n{p}"))
        except Exception as e:
            err_text = str(e)
            self.after(0, lambda: self.label_status_w2p.configure(text="✗ Ошибка!", text_color=COLORS["accent_coral"]))
            self.after(0, lambda t=err_text: messagebox.showerror("Ошибка", t))

    def _run_batch_w2p(self, in_dir, out_dir):
        def update_progress(current, total, filename):
            self.after(0, lambda p=current/total, c=current, t=total, f=filename:
                       self._update_batch_ui_w2p(p, c, t, f))
        success, errors = BatchProcessor.process_folder(in_dir, out_dir, ".docx", ".pdf",
                                                        self._single_w2p_logic, update_progress)
        self.after(0, lambda s=success, e=errors, d=out_dir: self._finish_batch_w2p(s, e, d))

    def _update_batch_ui_w2p(self, percent, current, total, filename):
        self.progress_w2p.set(percent)
        self.label_status_w2p.configure(text=f"📄 Файл {current}/{total}: {filename}")

    def _finish_batch_w2p(self, success, errors, out_dir):
        self.progress_w2p.set(1.0)
        self.label_status_w2w.configure(text=f"✓ Готово! Успешно: {success}, Ошибок: {errors}",
                                        text_color=COLORS["success"])
        messagebox.showinfo("Пакетная обработка",
                            f"Готово!\nУспешно: {success}\nС ошибками: {errors}\n\nПапка: {out_dir}")

    # ==================== ФОТО -> PDF ====================
    def create_images_to_pdf_frame(self):
        frame = self._create_content_frame("Фото → PDF", "Создание PDF из изображений")
        self.frames["images_to_pdf"] = frame

        self.image_files = []
        self.listbox_images = ctk.CTkTextbox(
            frame, height=180,
            fg_color=COLORS["bg_page"],
            border_color=COLORS["text_secondary"],
            border_width=1,
            corner_radius=RADIUS["card"],
            font=FONT, text_color=COLORS["text_primary"]
        )
        self.listbox_images.pack(pady=(0, 10), padx=20, fill="x")
        self.listbox_images.configure(state="disabled")

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(pady=(0, 10))
        self._create_pill_button(btn_frame, "➕ Добавить", self.add_image_files, COLORS["accent_yellow"], 140).pack(side="left", padx=5)
        self._create_pill_button(btn_frame, "🗑 Очистить", self.clear_image_files, COLORS["bg_dark"], 120).pack(side="left", padx=5)
        self._create_pill_button(btn_frame, "🚀 Создать PDF", self.images_to_pdf, COLORS["accent_coral"], 150).pack(side="left", padx=5)

        self.label_status_images = ctk.CTkLabel(frame, text="Добавьте изображения для создания PDF",
                                                font=FONT_SMALL, text_color=COLORS["text_secondary"])
        self.label_status_images.pack(pady=(5, 15))

    def add_image_files(self):
        filenames = filedialog.askopenfilenames(filetypes=[("Изображения", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff")])
        if filenames:
            self.image_files.extend(filenames)
            self.update_image_listbox()

    def clear_image_files(self):
        self.image_files = []
        self.update_image_listbox()

    def update_image_listbox(self):
        self.listbox_images.configure(state="normal")
        self.listbox_images.delete("1.0", "end")
        for f in self.image_files:
            self.listbox_images.insert("end", f + "\n")
        self.listbox_images.configure(state="disabled")
        count = len(self.image_files)
        self.label_status_images.configure(
            text=f"Добавлено файлов: {count}" if count > 0 else "Добавьте изображения для создания PDF")

    def images_to_pdf(self):
        if not self.image_files:
            messagebox.showwarning("Внимание", "Добавьте хотя бы одно изображение!")
            return
        pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
        if not pdf_path:
            return
        try:
            images = []
            for img_path in self.image_files:
                img = Image.open(img_path)
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                images.append(img)
            images[0].save(pdf_path, "PDF", save_all=True, append_images=images[1:], resolution=150)
            self.label_status_images.configure(text=f"✓ Создан PDF из {len(images)} изображений!",
                                               text_color=COLORS["success"])
            messagebox.showinfo("Успех", f"Создан PDF из {len(images)} изображений!")
            self.clear_image_files()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== РАЗДЕЛИТЬ ====================
    def create_split_frame(self):
        frame = self._create_content_frame("Разделить PDF", "Разделение на страницы")
        self.frames["split"] = frame

        self.entry_split = self._create_entry(frame)
        self.entry_split.pack(pady=(0, 8), padx=20, fill="x")

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(pady=(0, 12))
        self._create_pill_button(btn_frame, "📁 Обзор", lambda: self._browse_generic(self.entry_split), COLORS["bg_dark"], 120).pack(side="left", padx=5)
        self._create_pill_button(btn_frame, "✂️ Разделить", self.split_pdf, COLORS["accent_yellow"], 140).pack(side="left", padx=5)

        self.label_status_3 = ctk.CTkLabel(frame, text="Готов к работе", font=FONT_SMALL,
                                           text_color=COLORS["text_secondary"])
        self.label_status_3.pack(pady=(5, 15))
        self._create_info_bar(frame, "💡 Каждая страница сохранится отдельным файлом")

    def split_pdf(self):
        pdf_path = self.entry_split.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Ошибка", "Файл не выбран!")
            return
        output_dir = filedialog.askdirectory(title="Выберите папку для сохранения страниц")
        if not output_dir:
            return
        try:
            reader = PdfReader(pdf_path)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                with open(os.path.join(output_dir, f"{base_name}_page_{i + 1}.pdf"), "wb") as f:
                    writer.write(f)
            self.label_status_3.configure(text=f"✓ Создано {len(reader.pages)} файлов.",
                                          text_color=COLORS["success"])
            messagebox.showinfo("Успех", f"PDF разделен на {len(reader.pages)} страниц.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ОБЪЕДИНИТЬ ====================
    def create_merge_frame(self):
        frame = self._create_content_frame("Объединить PDF", "Объединение файлов")
        self.frames["merge"] = frame

        self.merge_files = []
        self.listbox_merge = ctk.CTkTextbox(
            frame, height=180,
            fg_color=COLORS["bg_page"],
            border_color=COLORS["text_secondary"],
            border_width=1,
            corner_radius=RADIUS["card"],
            font=FONT, text_color=COLORS["text_primary"]
        )
        self.listbox_merge.pack(pady=(0, 10), padx=20, fill="x")
        self.listbox_merge.configure(state="disabled")

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(pady=(0, 10))
        self._create_pill_button(btn_frame, "➕ Добавить", self.add_merge_files, COLORS["accent_yellow"], 140).pack(side="left", padx=5)
        self._create_pill_button(btn_frame, "🗑 Очистить", self.clear_merge_files, COLORS["bg_dark"], 120).pack(side="left", padx=5)
        self._create_pill_button(btn_frame, "🔗 Объединить", self.merge_pdfs, COLORS["accent_coral"], 150).pack(side="left", padx=5)

        self.label_status_4 = ctk.CTkLabel(frame, text="Добавьте минимум 2 файла",
                                           font=FONT_SMALL, text_color=COLORS["text_secondary"])
        self.label_status_4.pack(pady=(5, 15))

    def add_merge_files(self):
        filenames = filedialog.askopenfilenames(filetypes=[("PDF файлы", "*.pdf")])
        if filenames:
            self.merge_files.extend(filenames)
            self.update_merge_listbox()

    def clear_merge_files(self):
        self.merge_files = []
        self.update_merge_listbox()

    def update_merge_listbox(self):
        self.listbox_merge.configure(state="normal")
        self.listbox_merge.delete("1.0", "end")
        for f in self.merge_files:
            self.listbox_merge.insert("end", f + "\n")
        self.listbox_merge.configure(state="disabled")
        count = len(self.merge_files)
        self.label_status_4.configure(text=f"Добавлено файлов: {count}" if count > 0 else "Добавьте минимум 2 файла")

    def merge_pdfs(self):
        if len(self.merge_files) < 2:
            messagebox.showwarning("Внимание", "Нужно выбрать минимум 2 файла!")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
        if not output_path:
            return
        try:
            writer = PdfWriter()
            for pdf in self.merge_files:
                for page in PdfReader(pdf).pages:
                    writer.add_page(page)
            with open(output_path, "wb") as f:
                writer.write(f)
            self.label_status_4.configure(text="✓ Успешно объединены!", text_color=COLORS["success"])
            messagebox.showinfo("Успех", f"Файлы объединены в:\n{output_path}")
            self.clear_merge_files()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== СЖАТЬ ====================
    def create_compress_frame(self):
        frame = self._create_content_frame("Сжать PDF", "Уменьшение размера файла")
        self.frames["compress"] = frame

        top_bar = ctk.CTkFrame(frame, fg_color="transparent")
        top_bar.pack(fill="x", padx=20, pady=(0, 10))
        ctk.CTkLabel(top_bar, text="Режим:", font=FONT,
                     text_color=COLORS["text_secondary"]).pack(side="left", padx=(0, 10))
        self.mode_comp = self._create_segmented_button(top_bar, ["Один файл", "Вся папка"], self.toggle_comp_mode)
        self.mode_comp.pack(side="left")
        self.mode_comp.set("Один файл")

        self.entry_compress = self._create_entry(frame, placeholder="Выберите файл или папку...")
        self.entry_compress.pack(pady=(0, 8), padx=20, fill="x")

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(pady=(0, 12))
        self._create_pill_button(btn_frame, "📁 Обзор", self.browse_compress, COLORS["bg_dark"], 120).pack(side="left", padx=5)
        self._create_pill_button(btn_frame, "📉 Сжать", self.start_compress, COLORS["accent_coral"], 120).pack(side="left", padx=5)

        self.progress_comp = ctk.CTkProgressBar(frame, width=460, height=6,
                                                corner_radius=RADIUS["pill"],
                                                progress_color=COLORS["accent_coral"])
        self.progress_comp.pack(pady=(0, 6))
        self.progress_comp.set(0)

        self.label_status_compress = ctk.CTkLabel(frame, text="Готов к работе", font=FONT_SMALL,
                                                  text_color=COLORS["text_secondary"])
        self.label_status_compress.pack()

        self._create_info_bar(frame, "💡 Оптимизирует структуру файла и удаляет мусор")

    def toggle_comp_mode(self, value):
        self.entry_compress.delete(0, "end")
        self.entry_compress.configure(placeholder_text="Выберите папку..." if value == "Вся папка" else "Выберите файл...")

    def browse_compress(self):
        if self.mode_comp.get() == "Один файл":
            path = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")])
        else:
            path = filedialog.askdirectory(title="Выберите папку с PDF файлами")
        if path:
            self.entry_compress.delete(0, "end")
            self.entry_compress.insert(0, path)

    def start_compress(self):
        path = self.entry_compress.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Ошибка", "Путь не выбран!")
            return
        if self.mode_comp.get() == "Один файл":
            out_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
            if not out_path:
                return
            self.label_status_compress.configure(text="⏳ Сжатие...")
            self.progress_comp.set(0.5)
            threading.Thread(target=self._run_single_compress, args=(path, out_path), daemon=True).start()
        else:
            out_dir = filedialog.askdirectory(title="Куда сохранить сжатые файлы?")
            if not out_dir:
                return
            self.label_status_compress.configure(text="⏳ Пакетная обработка...")
            self.progress_comp.set(0)
            threading.Thread(target=self._run_batch_compress, args=(path, out_dir), daemon=True).start()

    def _single_compress_logic(self, in_path, out_path):
        doc = fitz.open(in_path)
        doc.save(out_path, garbage=4, deflate=True, clean=True)
        doc.close()

    def _run_single_compress(self, in_path, out_path):
        try:
            self._single_compress_logic(in_path, out_path)
            orig = os.path.getsize(in_path) / 1024
            comp = os.path.getsize(out_path) / 1024
            pct = (1 - comp / orig) * 100 if orig > 0 else 0
            self.after(0, lambda: self.progress_comp.set(1.0))
            self.after(0, lambda p=pct: self.label_status_compress.configure(
                text=f"✓ Сжато! Экономия: {p:.1f}%", text_color=COLORS["success"]))
            self.after(0, lambda o=orig, c=comp, p=pct: messagebox.showinfo(
                "Успех", f"Было: {o:.1f} КБ\nСтало: {c:.1f} КБ\nЭкономия: {p:.1f}%"))
        except Exception as e:
            err_text = str(e)
            self.after(0, lambda: self.label_status_compress.configure(text="✗ Ошибка!", text_color=COLORS["accent_coral"]))
            self.after(0, lambda t=err_text: messagebox.showerror("Ошибка", t))

    def _run_batch_compress(self, in_dir, out_dir):
        def update_progress(current, total, filename):
            self.after(0, lambda p=current/total, c=current, t=total, f=filename:
                       self._update_batch_ui_comp(p, c, t, f))
        success, errors = BatchProcessor.process_folder(in_dir, out_dir, ".pdf", ".pdf",
                                                        self._single_compress_logic, update_progress)
        self.after(0, lambda s=success, e=errors, d=out_dir: self._finish_batch_comp(s, e, d))

    def _update_batch_ui_comp(self, percent, current, total, filename):
        self.progress_comp.set(percent)
        self.label_status_compress.configure(text=f"📄 Файл {current}/{total}: {filename}")

    def _finish_batch_comp(self, success, errors, out_dir):
        self.progress_comp.set(1.0)
        self.label_status_compress.configure(text=f"✓ Готово! Успешно: {success}, Ошибок: {errors}",
                                             text_color=COLORS["success"])
        messagebox.showinfo("Пакетная обработка",
                            f"Готово!\nУспешно: {success}\nС ошибками: {errors}\n\nПапка: {out_dir}")

    # ==================== ПАРОЛЬ ====================
    def create_protect_frame(self):
        frame = self._create_content_frame("Защита паролем", "Шифрование PDF")
        self.frames["protect"] = frame

        self.entry_protect = self._create_entry(frame)
        self.entry_protect.pack(pady=(0, 8), padx=20, fill="x")

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(pady=(0, 12))
        self._create_pill_button(btn_frame, "📁 Обзор", lambda: self._browse_generic(self.entry_protect), COLORS["bg_dark"], 120).pack(side="left", padx=5)
        self._create_pill_button(btn_frame, "🛡 Защитить", self.protect_pdf, COLORS["accent_coral"], 140).pack(side="left", padx=5)

        self.label_status_protect = ctk.CTkLabel(frame, text="Готов к работе", font=FONT_SMALL,
                                                 text_color=COLORS["text_secondary"])
        self.label_status_protect.pack(pady=(5, 15))

        self._create_info_bar(frame, "🔒 Использует 128-битное шифрование")

    def protect_pdf(self):
        pdf_path = self.entry_protect.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Ошибка", "Файл не выбран!")
            return
        password = simpledialog.askstring("Пароль", "Введите пароль:", show="*", parent=self)
        if not password:
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
        if not output_path:
            return
        try:
            writer = PdfWriter()
            for page in PdfReader(pdf_path).pages:
                writer.add_page(page)
            writer.encrypt(user_password=password, owner_password=password, use_128bit=True)
            with open(output_path, "wb") as f:
                writer.write(f)
            self.label_status_protect.configure(text="✓ Защищено!", text_color=COLORS["success"])
            messagebox.showinfo("Успех", "PDF защищён паролем!")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ИЗВЛЕЧЬ ====================
    def create_extract_frame(self):
        frame = self._create_content_frame("Извлечь страницы", "По диапазону")
        self.frames["extract"] = frame

        self.entry_extract = self._create_entry(frame)
        self.entry_extract.pack(pady=(0, 8), padx=20, fill="x")

        btn_frame1 = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame1.pack(pady=(0, 8))
        self._create_pill_button(btn_frame1, "📁 Обзор", lambda: self._browse_generic(self.entry_extract), COLORS["bg_dark"], 120).pack(side="left", padx=5)

        ctk.CTkLabel(frame, text="Какие страницы извлечь? (Например: 1-5, 8, 12-15)",
                     font=FONT_SMALL, text_color=COLORS["text_secondary"]).pack(pady=(5, 3))
        self.entry_range_extract = self._create_entry(frame, width=300, placeholder="1-3, 5, 8-10")
        self.entry_range_extract.pack(pady=(0, 8))

        self._create_pill_button(frame, "📑 Извлечь страницы", self.extract_pages, COLORS["accent_yellow"], 180).pack(pady=(0, 12))

        self.label_status_extract = ctk.CTkLabel(frame, text="Готов к работе", font=FONT_SMALL,
                                                 text_color=COLORS["text_secondary"])
        self.label_status_extract.pack(pady=(0, 15))

    def extract_pages(self):
        pdf_path = self.entry_extract.get()
        range_text = self.entry_range_extract.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Ошибка", "Файл не выбран!")
            return
        if not range_text:
            messagebox.showwarning("Внимание", "Введите диапазон страниц!")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
        if not output_path:
            return
        try:
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            indices = PageParser.parse(range_text, total_pages)
            if not indices:
                messagebox.showwarning("Внимание", f"Не найдено подходящих страниц!\nВ документе всего страниц: {total_pages}")
                return
            writer = PdfWriter()
            for idx in indices:
                writer.add_page(reader.pages[idx])
            with open(output_path, "wb") as f:
                writer.write(f)
            self.label_status_extract.configure(text=f"✓ Извлечено {len(indices)} страниц!",
                                                text_color=COLORS["success"])
            messagebox.showinfo("Успех", f"Извлечено {len(indices)} страниц из {total_pages}.\n\nСохранено:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ПОВЕРНУТЬ ====================
    def create_rotate_frame(self):
        frame = self._create_content_frame("Повернуть страницы", "Изменение ориентации")
        self.frames["rotate"] = frame

        self.entry_rotate = self._create_entry(frame)
        self.entry_rotate.pack(pady=(0, 8), padx=20, fill="x")

        btn_frame1 = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame1.pack(pady=(0, 8))
        self._create_pill_button(btn_frame1, "📁 Обзор", lambda: self._browse_generic(self.entry_rotate), COLORS["bg_dark"], 120).pack(side="left", padx=5)

        opts_frame = ctk.CTkFrame(frame, fg_color="transparent")
        opts_frame.pack(pady=(5, 5))
        ctk.CTkLabel(opts_frame, text="Угол:", font=FONT,
                     text_color=COLORS["text_secondary"]).pack(side="left", padx=(0, 10))
        self.combo_angle = ctk.CTkComboBox(
            opts_frame,
            values=["90°", "180°", "270°"],
            fg_color=COLORS["bg_page"],
            border_color=COLORS["text_secondary"],
            button_color=COLORS["accent_coral"],
            button_hover_color=self._darken_color(COLORS["accent_coral"]),
            dropdown_fg_color=COLORS["bg_page"],
            dropdown_hover_color=COLORS["bg_card_light"],
            text_color=COLORS["text_primary"],
            font=FONT, dropdown_font=FONT,
            width=100, height=32
        )
        self.combo_angle.pack(side="left")
        self.combo_angle.set("90°")

        ctk.CTkLabel(frame, text="Какие страницы? (пусто = все)",
                     font=FONT_SMALL, text_color=COLORS["text_secondary"]).pack(pady=(8, 3))
        self.entry_range_rotate = self._create_entry(frame, width=300, placeholder="Все, или 1-3, 5")
        self.entry_range_rotate.pack(pady=(0, 8))

        self._create_pill_button(frame, "📐 Повернуть", self.rotate_pages, COLORS["accent_coral"], 140).pack(pady=(0, 12))

        self.label_status_rotate = ctk.CTkLabel(frame, text="Готов к работе", font=FONT_SMALL,
                                                text_color=COLORS["text_secondary"])
        self.label_status_rotate.pack(pady=(0, 15))

    def rotate_pages(self):
        pdf_path = self.entry_rotate.get()
        range_text = self.entry_range_rotate.get()
        angle_str = self.combo_angle.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Ошибка", "Файл не выбран!")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
        if not output_path:
            return
        try:
            angle = int(angle_str.replace("°", ""))
        except ValueError:
            angle = 90
        try:
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            total_pages = len(reader.pages)
            indices = list(range(total_pages)) if not range_text.strip() else PageParser.parse(range_text, total_pages)
            for i, page in enumerate(reader.pages):
                if i in indices:
                    page.rotate(angle)
                writer.add_page(page)
            with open(output_path, "wb") as f:
                writer.write(f)
            self.label_status_rotate.configure(text=f"✓ Повернуто {len(indices)} страниц на {angle}°!",
                                               text_color=COLORS["success"])
            messagebox.showinfo("Успех", f"Повернуто {len(indices)} страниц на {angle}°.\n\nСохранено:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== УДАЛИТЬ ====================
    def create_delete_frame(self):
        frame = self._create_content_frame("Удалить страницы", "Удаление из PDF")
        self.frames["delete"] = frame

        self.entry_delete = self._create_entry(frame)
        self.entry_delete.pack(pady=(0, 8), padx=20, fill="x")

        btn_frame1 = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame1.pack(pady=(0, 8))
        self._create_pill_button(btn_frame1, "📁 Обзор", lambda: self._browse_generic(self.entry_delete), COLORS["bg_dark"], 120).pack(side="left", padx=5)

        ctk.CTkLabel(frame, text="Какие страницы УДАЛИТЬ?",
                     font=FONT_SMALL, text_color=COLORS["accent_coral"]).pack(pady=(5, 3))
        self.entry_range_delete = self._create_entry(frame, width=300, placeholder="2, 4, 10-15")
        self.entry_range_delete.pack(pady=(0, 8))

        self._create_pill_button(frame, "🗑 Удалить страницы", self.delete_pages, COLORS["accent_coral"], 180).pack(pady=(0, 12))

        self.label_status_delete = ctk.CTkLabel(frame, text="Готов к работе", font=FONT_SMALL,
                                                text_color=COLORS["text_secondary"])
        self.label_status_delete.pack(pady=(0, 15))

    def delete_pages(self):
        pdf_path = self.entry_delete.get()
        range_text = self.entry_range_delete.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Ошибка", "Файл не выбран!")
            return
        if not range_text:
            messagebox.showwarning("Внимание", "Введите номера страниц для удаления!")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
        if not output_path:
            return
        try:
            reader = PdfReader(pdf_path)
            total_pages = len(reader.pages)
            indices_to_delete = PageParser.parse(range_text, total_pages)
            if not indices_to_delete:
                messagebox.showwarning("Внимание", "Таких страниц нет в документе!")
                return
            writer = PdfWriter()
            deleted_count = 0
            for i, page in enumerate(reader.pages):
                if i not in indices_to_delete:
                    writer.add_page(page)
                else:
                    deleted_count += 1
            if len(writer.pages) == 0:
                messagebox.showwarning("Внимание", "Вы пытаетесь удалить ВСЕ страницы!")
                return
            with open(output_path, "wb") as f:
                writer.write(f)
            self.label_status_delete.configure(text=f"✓ Удалено {deleted_count} страниц!",
                                               text_color=COLORS["success"])
            messagebox.showinfo("Успех", f"Удалено {deleted_count} страниц.\nОсталось: {len(writer.pages)}\n\nСохранено:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== НОМЕРА ====================
    def create_numbers_frame(self):
        frame = self._create_content_frame("Номера страниц", "Нумерация документа")
        self.frames["numbers"] = frame

        self.entry_numbers = self._create_entry(frame)
        self.entry_numbers.pack(pady=(0, 8), padx=20, fill="x")

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(pady=(0, 12))
        self._create_pill_button(btn_frame, "📁 Обзор", lambda: self._browse_generic(self.entry_numbers), COLORS["bg_dark"], 120).pack(side="left", padx=5)
        self._create_pill_button(btn_frame, "🔢 Добавить номера", self.add_page_numbers, COLORS["accent_yellow"], 180).pack(side="left", padx=5)

        self.label_status_numbers = ctk.CTkLabel(frame, text="Готов к работе", font=FONT_SMALL,
                                                 text_color=COLORS["text_secondary"])
        self.label_status_numbers.pack(pady=(5, 15))

        self._create_info_bar(frame, "💡 Номера добавляются внизу по центру каждой страницы")

    def add_page_numbers(self):
        pdf_path = self.entry_numbers.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Ошибка", "Файл не выбран!")
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
        if not output_path:
            return
        try:
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            for i, page in enumerate(reader.pages):
                w, h = float(page.mediabox.width), float(page.mediabox.height)
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=(w, h))
                can.setFont("Helvetica", 10)
                can.drawCentredString(w / 2, 20, str(i + 1))
                can.save()
                packet.seek(0)
                page.merge_page(PdfReader(packet).pages[0])
                writer.add_page(page)
            with open(output_path, "wb") as f:
                writer.write(f)
            self.label_status_numbers.configure(text="✓ Номера добавлены!", text_color=COLORS["success"])
            messagebox.showinfo("Успех", "Номера страниц добавлены!")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== СЛУЖЕБНЫЕ ====================
    def _browse_generic(self, entry_widget):
        filename = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")])
        if filename:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, filename)

    def on_files_dropped(self, event):
        try:
            files = self.tk.splitlist(event.data)
        except Exception:
            return
        image_exts = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'}
        images = [f for f in files if os.path.splitext(f)[1].lower() in image_exts]
        pdfs = [f for f in files if os.path.splitext(f)[1].lower() == '.pdf']
        if images:
            self.image_files.extend(images)
            self.update_image_listbox()
        if pdfs:
            self.merge_files.extend(pdfs)
            self.update_merge_listbox()
        if images and not pdfs:
            self.show_frame("images_to_pdf")
        elif pdfs and not images:
            self.show_frame("merge")


if __name__ == "__main__":
    app = PDFConverterApp()
    app.mainloop()
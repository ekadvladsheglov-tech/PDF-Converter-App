import os
import sys
import io
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
import threading

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


ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")
BaseWindow = TkinterDnD.Tk if HAS_DND else ctk.CTk


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
    def process_folder(input_folder, output_folder, file_ext, process_func, progress_callback):
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        files = [f for f in os.listdir(input_folder) if f.lower().endswith(file_ext)]
        total = len(files)
        if total == 0:
            return 0, 0
        success_count = 0
        error_count = 0
        for i, filename in enumerate(files):
            in_path = os.path.join(input_folder, filename)
            base_name = os.path.splitext(filename)[0]
            out_ext = ".docx" if file_ext == ".pdf" and "word" in str(process_func.__name__).lower() else ".pdf"
            out_path = os.path.join(output_folder, f"{base_name}{out_ext}")
            try:
                process_func(in_path, out_path)
                success_count += 1
            except Exception:
                error_count += 1
            progress_callback(i + 1, total, filename)
        return success_count, error_count


class PDFConverterApp(BaseWindow):
    def __init__(self):
        super().__init__()
        self.title("PDF Конвертер & Инструменты v3.0")
        self.geometry("1050x720")
        self.resizable(False, False)

        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self.on_files_dropped)

        # Основной контейнер
        main_container = ctk.CTkFrame(self, corner_radius=0)
        main_container.pack(fill="both", expand=True)
        main_container.grid_rowconfigure(0, weight=1)
        main_container.grid_columnconfigure(1, weight=1)

        # Словари для хранения фреймов и кнопок меню
        self.frames = {}
        self.menu_buttons = {}

        # Боковое меню
        self.create_sidebar(main_container)

        # Область контента
        self.content_area = ctk.CTkFrame(main_container, corner_radius=0)
        self.content_area.grid(row=0, column=1, sticky="nsew")
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        # Создаём все фреймы инструментов
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

        # Размещаем все фреймы в одной позиции (для переключения через tkraise)
        for frame in self.frames.values():
            frame.grid(row=0, column=0, sticky="nsew")

        # Показываем первый фрейм по умолчанию
        self.show_frame("pdf_to_word")

    # ==================== БОКОВОЕ МЕНЮ ====================
    def create_sidebar(self, parent):
        sidebar = ctk.CTkFrame(parent, width=230, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)  # Фиксируем ширину

        # Логотип / Заголовок
        ctk.CTkLabel(sidebar, text="📄 PDF Конвертер", font=("Arial", 18, "bold")).pack(pady=(20, 5))
        ctk.CTkLabel(sidebar, text="v3.0", font=("Arial", 11), text_color="gray").pack(pady=(0, 20))

        # --- КАТЕГОРИЯ: КОНВЕРТАЦИЯ ---
        ctk.CTkLabel(sidebar, text="КОНВЕРТАЦИЯ", font=("Arial", 11, "bold"), text_color="gray").pack(anchor="w", padx=15, pady=(10, 5))
        
        self._add_menu_button(sidebar, "pdf_to_word", "PDF → Word", "🔄")
        self._add_menu_button(sidebar, "word_to_pdf", "Word → PDF", "🔄")
        self._add_menu_button(sidebar, "images_to_pdf", "Фото → PDF", "🖼")

        # --- КАТЕГОРИЯ: ИНСТРУМЕНТЫ ---
        ctk.CTkLabel(sidebar, text="ИНСТРУМЕНТЫ", font=("Arial", 11, "bold"), text_color="gray").pack(anchor="w", padx=15, pady=(15, 5))
        
        self._add_menu_button(sidebar, "split", "Разделить", "✂️")
        self._add_menu_button(sidebar, "merge", "Объединить", "🔗")
        self._add_menu_button(sidebar, "compress", "Сжать", "📉")
        self._add_menu_button(sidebar, "protect", "Пароль", "🛡")

        # --- КАТЕГОРИЯ: СТРАНИЦЫ ---
        ctk.CTkLabel(sidebar, text="СТРАНИЦЫ", font=("Arial", 11, "bold"), text_color="gray").pack(anchor="w", padx=15, pady=(15, 5))
        
        self._add_menu_button(sidebar, "extract", "Извлечь", "📑")
        self._add_menu_button(sidebar, "rotate", "Повернуть", "📐")
        self._add_menu_button(sidebar, "delete", "Удалить", "🗑")
        self._add_menu_button(sidebar, "numbers", "Номера", "🔢")

        # Нижняя часть (инфо)
        ctk.CTkLabel(sidebar, text="Локально • Безопасно", font=("Arial", 10), text_color="gray").pack(side="bottom", pady=15)

    def _add_menu_button(self, parent, frame_name, text, icon):
        btn = ctk.CTkButton(
            parent,
            text=f"{icon}  {text}",
            anchor="w",
            fg_color="transparent",
            text_color=("gray10", "gray90"),
            hover_color=("gray75", "gray25"),
            height=35,
            command=lambda fn=frame_name: self.show_frame(fn)
        )
        btn.pack(fill="x", padx=10, pady=1)
        self.menu_buttons[frame_name] = btn

    def show_frame(self, frame_name):
        """Переключает видимый фрейм и подсвечивает активную кнопку."""
        # Сбрасываем подсветку всех кнопок
        for btn in self.menu_buttons.values():
            btn.configure(fg_color="transparent", text_color=("gray10", "gray90"))
        
        # Подсвечиваем активную кнопку
        self.menu_buttons[frame_name].configure(fg_color="#3B8ED0", text_color="white")
        
        # Показываем нужный фрейм
        self.frames[frame_name].tkraise()

    # ==================== ФРЕЙМ: PDF -> WORD ====================
    def create_pdf_to_word_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["pdf_to_word"] = frame

        ctk.CTkLabel(frame, text="Конвертация PDF в Word", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        
        self.mode_p2w = ctk.CTkSegmentedButton(frame, values=["Один файл", "Вся папка"], command=self.toggle_p2w_mode)
        self.mode_p2w.pack(pady=5)
        self.mode_p2w.set("Один файл")

        self.entry_p2w = ctk.CTkEntry(frame, width=600, placeholder_text="Выберите файл или папку...")
        self.entry_p2w.pack(pady=5)
        ctk.CTkButton(frame, text="Обзор...", command=self.browse_p2w).pack(pady=5)
        ctk.CTkButton(frame, text="Конвертировать", command=self.start_p2w, fg_color="green", hover_color="darkgreen", height=40, width=200).pack(pady=15)

        self.progress_p2w = ctk.CTkProgressBar(frame, width=600)
        self.progress_p2w.pack(pady=5)
        self.progress_p2w.set(0)
        self.label_status_p2w = ctk.CTkLabel(frame, text="")
        self.label_status_p2w.pack(pady=5)

    def toggle_p2w_mode(self, value):
        self.entry_p2w.delete(0, "end")
        self.entry_p2w.configure(placeholder_text="Выберите папку..." if value == "Вся папка" else "Выберите файл...")

    def browse_p2w(self):
        path = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")]) if self.mode_p2w.get() == "Один файл" else filedialog.askdirectory(title="Выберите папку с PDF файлами")
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
            if not out_path: return
            self.label_status_p2w.configure(text="Конвертация...")
            self.progress_p2w.set(0.5)
            threading.Thread(target=self._run_single_p2w, args=(path, out_path), daemon=True).start()
        else:
            out_dir = filedialog.askdirectory(title="Куда сохранить результаты?")
            if not out_dir: return
            self.label_status_p2w.configure(text="Пакетная обработка...")
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
            self.after(0, lambda: self.label_status_p2w.configure(text="Успешно!"))
            self.after(0, lambda: messagebox.showinfo("Успех", f"Сохранено:\n{out_path}"))
        except Exception as e:
            self.after(0, lambda: self.label_status_p2w.configure(text="Ошибка!"))
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))

    def _run_batch_p2w(self, in_dir, out_dir):
        def update_progress(current, total, filename):
            self.after(0, lambda p=current/total, c=current, t=total, f=filename: self._update_batch_ui_p2w(p, c, t, f))
        success, errors = BatchProcessor.process_folder(in_dir, out_dir, ".pdf", self._single_p2w_logic, update_progress)
        self.after(0, lambda s=success, e=errors, d=out_dir: self._finish_batch_p2w(s, e, d))

    def _update_batch_ui_p2w(self, percent, current, total, filename):
        self.progress_p2w.set(percent)
        self.label_status_p2w.configure(text=f"Файл {current} из {total}: {filename}")

    def _finish_batch_p2w(self, success, errors, out_dir):
        self.progress_p2w.set(1.0)
        self.label_status_p2w.configure(text=f"Готово! Успешно: {success}, Ошибок: {errors}")
        messagebox.showinfo("Пакетная обработка", f"Готово!\nУспешно: {success}\nС ошибками: {errors}\n\nПапка: {out_dir}")

    # ==================== ФРЕЙМ: WORD -> PDF ====================
    def create_word_to_pdf_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["word_to_pdf"] = frame

        ctk.CTkLabel(frame, text="Конвертация Word в PDF", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        
        self.mode_w2p = ctk.CTkSegmentedButton(frame, values=["Один файл", "Вся папка"], command=self.toggle_w2p_mode)
        self.mode_w2p.pack(pady=5)
        self.mode_w2p.set("Один файл")

        self.entry_w2p = ctk.CTkEntry(frame, width=600, placeholder_text="Выберите файл или папку...")
        self.entry_w2p.pack(pady=5)
        ctk.CTkButton(frame, text="Обзор...", command=self.browse_w2p).pack(pady=5)
        ctk.CTkButton(frame, text="Конвертировать", command=self.start_w2p, fg_color="green", hover_color="darkgreen", height=40, width=200).pack(pady=15)

        self.progress_w2p = ctk.CTkProgressBar(frame, width=600)
        self.progress_w2p.pack(pady=5)
        self.progress_w2p.set(0)
        self.label_status_w2p = ctk.CTkLabel(frame, text="")
        self.label_status_w2p.pack(pady=5)

    def toggle_w2p_mode(self, value):
        self.entry_w2p.delete(0, "end")
        self.entry_w2p.configure(placeholder_text="Выберите папку..." if value == "Вся папка" else "Выберите файл...")

    def browse_w2p(self):
        path = filedialog.askopenfilename(filetypes=[("Word файлы", "*.docx")]) if self.mode_w2p.get() == "Один файл" else filedialog.askdirectory(title="Выберите папку с Word файлами")
        if path:
            self.entry_w2p.delete(0, "end")
            self.entry_w2p.insert(0, path)

    def start_w2p(self):
        path = self.entry_w2p.get()
        if not path or not os.path.exists(path):
            messagebox.showerror("Ошибка", "Путь не выбран!")
            return
        if self.mode_w2p.get() == "Один файл":
            out_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
            if not out_path: return
            threading.Thread(target=self._run_single_w2p, args=(path, out_path), daemon=True).start()
        else:
            out_dir = filedialog.askdirectory(title="Куда сохранить результаты?")
            if not out_dir: return
            threading.Thread(target=self._run_batch_w2p, args=(path, out_dir), daemon=True).start()

    def _single_w2p_logic(self, word_path, pdf_path):
        doc = Document(word_path)
        c = canvas.Canvas(pdf_path, pagesize=A4)
        width, height = A4
        font_path = get_resource_path("DejaVuSans.ttf")
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('DejaVu', font_path))
            font_name = 'DejaVu'
        else:
            font_name = 'Helvetica'
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
            self.after(0, lambda: self.label_status_w2p.configure(text="Конвертация..."))
            self.after(0, lambda: self.progress_w2p.set(0.5))
            self._single_w2p_logic(in_path, out_path)
            self.after(0, lambda: self.progress_w2p.set(1.0))
            self.after(0, lambda: self.label_status_w2p.configure(text="Успешно!"))
            self.after(0, lambda: messagebox.showinfo("Успех", f"Сохранено:\n{out_path}"))
        except Exception as e:
            self.after(0, lambda: self.label_status_w2p.configure(text="Ошибка!"))
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))

    def _run_batch_w2p(self, in_dir, out_dir):
        def update_progress(current, total, filename):
            self.after(0, lambda p=current/total, c=current, t=total, f=filename: self._update_batch_ui_w2p(p, c, t, f))
        success, errors = BatchProcessor.process_folder(in_dir, out_dir, ".docx", self._single_w2p_logic, update_progress)
        self.after(0, lambda s=success, e=errors, d=out_dir: self._finish_batch_w2p(s, e, d))

    def _update_batch_ui_w2p(self, percent, current, total, filename):
        self.progress_w2p.set(percent)
        self.label_status_w2p.configure(text=f"Файл {current} из {total}: {filename}")

    def _finish_batch_w2p(self, success, errors, out_dir):
        self.progress_w2p.set(1.0)
        self.label_status_w2p.configure(text=f"Готово! Успешно: {success}, Ошибок: {errors}")
        messagebox.showinfo("Пакетная обработка", f"Готово!\nУспешно: {success}\nС ошибками: {errors}\n\nПапка: {out_dir}")

    # ==================== ФРЕЙМ: ФОТО -> PDF ====================
    def create_images_to_pdf_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["images_to_pdf"] = frame

        ctk.CTkLabel(frame, text="Создание PDF из изображений", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        
        self.listbox_images = ctk.CTkTextbox(frame, width=600, height=220)
        self.listbox_images.pack(pady=5)
        self.listbox_images.configure(state="disabled")
        self.image_files = []

        frame_btns = ctk.CTkFrame(frame, fg_color="transparent")
        frame_btns.pack(pady=10)
        ctk.CTkButton(frame_btns, text="Добавить изображения", command=self.add_image_files).grid(row=0, column=0, padx=5)
        ctk.CTkButton(frame_btns, text="Очистить список", command=self.clear_image_files, fg_color="red").grid(row=0, column=1, padx=5)

        ctk.CTkButton(frame, text="Создать PDF", command=self.images_to_pdf, fg_color="blue", hover_color="darkblue", height=40, width=200).pack(pady=15)
        self.label_status_images = ctk.CTkLabel(frame, text="")
        self.label_status_images.pack(pady=5)

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
            self.label_status_images.configure(text=f"Создан PDF из {len(images)} изображений!")
            messagebox.showinfo("Успех", f"Создан PDF из {len(images)} изображений!")
            self.clear_image_files()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ФРЕЙМ: РАЗДЕЛИТЬ ====================
    def create_split_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["split"] = frame

        ctk.CTkLabel(frame, text="Разделение PDF на страницы", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        self.entry_split = ctk.CTkEntry(frame, width=600)
        self.entry_split.pack(pady=5)
        ctk.CTkButton(frame, text="Обзор...", command=lambda: self._browse_generic(self.entry_split)).pack(pady=5)
        ctk.CTkButton(frame, text="Разделить", command=self.split_pdf, fg_color="orange", hover_color="darkorange", height=40, width=200).pack(pady=15)
        self.label_status_3 = ctk.CTkLabel(frame, text="")
        self.label_status_3.pack(pady=5)

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
            self.label_status_3.configure(text=f"Создано {len(reader.pages)} файлов.")
            messagebox.showinfo("Успех", f"PDF разделен на {len(reader.pages)} страниц.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ФРЕЙМ: ОБЪЕДИНИТЬ ====================
    def create_merge_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["merge"] = frame

        ctk.CTkLabel(frame, text="Объединение PDF файлов", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        self.listbox_merge = ctk.CTkTextbox(frame, width=600, height=220)
        self.listbox_merge.pack(pady=5)
        self.listbox_merge.configure(state="disabled")
        self.merge_files = []

        frame_btns = ctk.CTkFrame(frame, fg_color="transparent")
        frame_btns.pack(pady=10)
        ctk.CTkButton(frame_btns, text="Добавить файлы", command=self.add_merge_files).grid(row=0, column=0, padx=5)
        ctk.CTkButton(frame_btns, text="Очистить список", command=self.clear_merge_files, fg_color="red").grid(row=0, column=1, padx=5)

        ctk.CTkButton(frame, text="Объединить", command=self.merge_pdfs, fg_color="purple", hover_color="darkviolet", height=40, width=200).pack(pady=15)
        self.label_status_4 = ctk.CTkLabel(frame, text="")
        self.label_status_4.pack(pady=5)

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
            self.label_status_4.configure(text="Успешно объединены!")
            messagebox.showinfo("Успех", f"Файлы объединены в:\n{output_path}")
            self.clear_merge_files()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ФРЕЙМ: СЖАТЬ ====================
    def create_compress_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["compress"] = frame

        ctk.CTkLabel(frame, text="Сжатие PDF файлов", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        
        self.mode_comp = ctk.CTkSegmentedButton(frame, values=["Один файл", "Вся папка"], command=self.toggle_comp_mode)
        self.mode_comp.pack(pady=5)
        self.mode_comp.set("Один файл")

        self.entry_compress = ctk.CTkEntry(frame, width=600, placeholder_text="Выберите файл или папку...")
        self.entry_compress.pack(pady=5)
        ctk.CTkButton(frame, text="Обзор...", command=self.browse_compress).pack(pady=5)
        ctk.CTkButton(frame, text="Сжать", command=self.start_compress, fg_color="teal", hover_color="darkgreen", height=40, width=200).pack(pady=15)

        self.progress_comp = ctk.CTkProgressBar(frame, width=600)
        self.progress_comp.pack(pady=5)
        self.progress_comp.set(0)
        self.label_status_compress = ctk.CTkLabel(frame, text="")
        self.label_status_compress.pack(pady=5)

    def toggle_comp_mode(self, value):
        self.entry_compress.delete(0, "end")
        self.entry_compress.configure(placeholder_text="Выберите папку..." if value == "Вся папка" else "Выберите файл...")

    def browse_compress(self):
        path = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")]) if self.mode_comp.get() == "Один файл" else filedialog.askdirectory(title="Выберите папку с PDF файлами")
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
            if not out_path: return
            threading.Thread(target=self._run_single_compress, args=(path, out_path), daemon=True).start()
        else:
            out_dir = filedialog.askdirectory(title="Куда сохранить сжатые файлы?")
            if not out_dir: return
            threading.Thread(target=self._run_batch_compress, args=(path, out_dir), daemon=True).start()

    def _single_compress_logic(self, in_path, out_path):
        doc = fitz.open(in_path)
        doc.save(out_path, garbage=4, deflate=True, clean=True)
        doc.close()

    def _run_single_compress(self, in_path, out_path):
        try:
            self.after(0, lambda: self.label_status_compress.configure(text="Сжатие..."))
            self.after(0, lambda: self.progress_comp.set(0.5))
            self._single_compress_logic(in_path, out_path)
            orig = os.path.getsize(in_path) / 1024
            comp = os.path.getsize(out_path) / 1024
            pct = (1 - comp / orig) * 100 if orig > 0 else 0
            self.after(0, lambda: self.progress_comp.set(1.0))
            self.after(0, lambda: self.label_status_compress.configure(text=f"Сжато! Экономия: {pct:.1f}%"))
            self.after(0, lambda: messagebox.showinfo("Успех", f"Было: {orig:.1f} КБ\nСтало: {comp:.1f} КБ\nЭкономия: {pct:.1f}%"))
        except Exception as e:
            self.after(0, lambda: self.label_status_compress.configure(text="Ошибка!"))
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))

    def _run_batch_compress(self, in_dir, out_dir):
        def update_progress(current, total, filename):
            self.after(0, lambda p=current/total, c=current, t=total, f=filename: self._update_batch_ui_comp(p, c, t, f))
        success, errors = BatchProcessor.process_folder(in_dir, out_dir, ".pdf", self._single_compress_logic, update_progress)
        self.after(0, lambda s=success, e=errors, d=out_dir: self._finish_batch_comp(s, e, d))

    def _update_batch_ui_comp(self, percent, current, total, filename):
        self.progress_comp.set(percent)
        self.label_status_compress.configure(text=f"Файл {current} из {total}: {filename}")

    def _finish_batch_comp(self, success, errors, out_dir):
        self.progress_comp.set(1.0)
        self.label_status_compress.configure(text=f"Готово! Успешно: {success}, Ошибок: {errors}")
        messagebox.showinfo("Пакетная обработка", f"Готово!\nУспешно: {success}\nС ошибками: {errors}\n\nПапка: {out_dir}")

    # ==================== ФРЕЙМ: ПАРОЛЬ ====================
    def create_protect_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["protect"] = frame

        ctk.CTkLabel(frame, text="Защита PDF паролем", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        self.entry_protect = ctk.CTkEntry(frame, width=600)
        self.entry_protect.pack(pady=5)
        ctk.CTkButton(frame, text="Обзор...", command=lambda: self._browse_generic(self.entry_protect)).pack(pady=5)
        ctk.CTkButton(frame, text="Защитить", command=self.protect_pdf, fg_color="darkred", hover_color="red", height=40, width=200).pack(pady=15)
        self.label_status_protect = ctk.CTkLabel(frame, text="")
        self.label_status_protect.pack(pady=5)

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
            self.label_status_protect.configure(text="Защищено!")
            messagebox.showinfo("Успех", "PDF защищён паролем!")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ФРЕЙМ: ИЗВЛЕЧЬ ====================
    def create_extract_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["extract"] = frame

        ctk.CTkLabel(frame, text="Извлечение страниц по диапазону", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        self.entry_extract = ctk.CTkEntry(frame, width=600)
        self.entry_extract.pack(pady=5)
        ctk.CTkButton(frame, text="Обзор...", command=lambda: self._browse_generic(self.entry_extract)).pack(pady=5)
        
        ctk.CTkLabel(frame, text="Какие страницы извлечь? (Например: 1-5, 8, 12-15)", font=("Arial", 12)).pack(pady=(10, 0))
        self.entry_range_extract = ctk.CTkEntry(frame, width=300, placeholder_text="1-3, 5, 8-10")
        self.entry_range_extract.pack(pady=5)
        
        ctk.CTkButton(frame, text="Извлечь страницы", command=self.extract_pages, fg_color="#FF8C00", hover_color="#CC7000", height=40, width=200).pack(pady=15)
        self.label_status_extract = ctk.CTkLabel(frame, text="")
        self.label_status_extract.pack(pady=5)

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
            self.label_status_extract.configure(text=f"Извлечено {len(indices)} страниц!")
            messagebox.showinfo("Успех", f"Извлечено {len(indices)} страниц из {total_pages}.\n\nСохранено:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ФРЕЙМ: ПОВЕРНУТЬ ====================
    def create_rotate_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["rotate"] = frame

        ctk.CTkLabel(frame, text="Поворот страниц PDF", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        self.entry_rotate = ctk.CTkEntry(frame, width=600)
        self.entry_rotate.pack(pady=5)
        ctk.CTkButton(frame, text="Обзор...", command=lambda: self._browse_generic(self.entry_rotate)).pack(pady=5)
        
        frame_opts = ctk.CTkFrame(frame, fg_color="transparent")
        frame_opts.pack(pady=10)
        ctk.CTkLabel(frame_opts, text="Угол поворота:").grid(row=0, column=0, padx=5)
        self.combo_angle = ctk.CTkComboBox(frame_opts, values=["90° (по часовой)", "180°", "270° (против часовой)"])
        self.combo_angle.grid(row=0, column=1, padx=5)
        self.combo_angle.set("90° (по часовой)")
        
        ctk.CTkLabel(frame, text="Какие страницы повернуть? (Оставьте пустым для ВСЕХ)", font=("Arial", 12)).pack(pady=(10, 0))
        self.entry_range_rotate = ctk.CTkEntry(frame, width=300, placeholder_text="Все, или 1-3, 5")
        self.entry_range_rotate.pack(pady=5)
        
        ctk.CTkButton(frame, text="Повернуть", command=self.rotate_pages, fg_color="#2E8B57", hover_color="#246B43", height=40, width=200).pack(pady=15)
        self.label_status_rotate = ctk.CTkLabel(frame, text="")
        self.label_status_rotate.pack(pady=5)

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
        if "90" in angle_str:
            angle = 90
        elif "180" in angle_str:
            angle = 180
        else:
            angle = 270
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
            self.label_status_rotate.configure(text=f"Повернуто {len(indices)} страниц на {angle}°!")
            messagebox.showinfo("Успех", f"Повернуто {len(indices)} страниц на {angle}°.\n\nСохранено:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ФРЕЙМ: УДАЛИТЬ ====================
    def create_delete_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["delete"] = frame

        ctk.CTkLabel(frame, text="Удаление страниц из PDF", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        self.entry_delete = ctk.CTkEntry(frame, width=600)
        self.entry_delete.pack(pady=5)
        ctk.CTkButton(frame, text="Обзор...", command=lambda: self._browse_generic(self.entry_delete)).pack(pady=5)
        
        ctk.CTkLabel(frame, text="Какие страницы УДАЛИТЬ? (Например: 2, 4, 10-15)", font=("Arial", 12), text_color="#FF6347").pack(pady=(10, 0))
        self.entry_range_delete = ctk.CTkEntry(frame, width=300, placeholder_text="2, 4, 10-15")
        self.entry_range_delete.pack(pady=5)
        
        ctk.CTkButton(frame, text="Удалить страницы", command=self.delete_pages, fg_color="#8B0000", hover_color="#5C0000", height=40, width=200).pack(pady=15)
        self.label_status_delete = ctk.CTkLabel(frame, text="")
        self.label_status_delete.pack(pady=5)

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
            self.label_status_delete.configure(text=f"Удалено {deleted_count} страниц!")
            messagebox.showinfo("Успех", f"Удалено {deleted_count} страниц.\nОсталось: {len(writer.pages)}\n\nСохранено:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ФРЕЙМ: НОМЕРА СТРАНИЦ ====================
    def create_numbers_frame(self):
        frame = ctk.CTkFrame(self.content_area, corner_radius=0)
        self.frames["numbers"] = frame

        ctk.CTkLabel(frame, text="Добавление номеров страниц", font=("Arial", 18, "bold")).pack(pady=(20, 10))
        self.entry_numbers = ctk.CTkEntry(frame, width=600)
        self.entry_numbers.pack(pady=5)
        ctk.CTkButton(frame, text="Обзор...", command=lambda: self._browse_generic(self.entry_numbers)).pack(pady=5)
        ctk.CTkButton(frame, text="Добавить номера", command=self.add_page_numbers, fg_color="darkblue", hover_color="navy", height=40, width=200).pack(pady=15)
        self.label_status_numbers = ctk.CTkLabel(frame, text="")
        self.label_status_numbers.pack(pady=5)

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
            self.label_status_numbers.configure(text="Номера добавлены!")
            messagebox.showinfo("Успех", "Номера страниц добавлены!")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ====================
    def _browse_generic(self, entry_widget):
        filename = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")])
        if filename:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, filename)

    # ==================== DRAG & DROP ====================
    def on_files_dropped(self, event):
        try:
            files = self.tk.splitlist(event.data)
        except Exception:
            return
        image_exts = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'}
        pdf_exts = {'.pdf'}
        images = [f for f in files if os.path.splitext(f)[1].lower() in image_exts]
        pdfs = [f for f in files if os.path.splitext(f)[1].lower() in pdf_exts]
        if images and not pdfs:
            self.image_files.extend(images)
            self.update_image_listbox()
            self.show_frame("images_to_pdf")
        elif pdfs and not images:
            self.merge_files.extend(pdfs)
            self.update_merge_listbox()
            self.show_frame("merge")
        elif images and pdfs:
            self.image_files.extend(images)
            self.update_image_listbox()
            self.merge_files.extend(pdfs)
            self.update_merge_listbox()


if __name__ == "__main__":
    app = PDFConverterApp()
    app.mainloop()
import os
import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
from docx import Document
from PyPDF2 import PdfReader, PdfWriter
from fpdf import FPDF
import threading

def get_resource_path(relative_path):
    """Получает абсолютный путь к ресурсам, работает для dev и для PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PDFConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PDF Конвертер & Инструменты v1.0")
        self.geometry("750x650")
        self.resizable(False, False)
        
        self.tabview = ctk.CTkTabview(self, width=730, height=600)
        self.tabview.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.tab_pdf_to_word = self.tabview.add("PDF в Word")
        self.tab_word_to_pdf = self.tabview.add("Word в PDF")
        self.tab_split = self.tabview.add("Разделить PDF")
        self.tab_merge = self.tabview.add("Объединить PDF")

        self.setup_pdf_to_word_tab()
        self.setup_word_to_pdf_tab()
        self.setup_split_tab()
        self.setup_merge_tab()

    # ==================== ВКЛАДКА 1: PDF -> WORD ====================
    def setup_pdf_to_word_tab(self):
        ctk.CTkLabel(self.tab_pdf_to_word, text="Выберите PDF файл для конвертации в Word:", font=("Arial", 14)).pack(pady=(20, 10))
        self.entry_pdf_to_word = ctk.CTkEntry(self.tab_pdf_to_word, width=550)
        self.entry_pdf_to_word.pack(pady=5)
        ctk.CTkButton(self.tab_pdf_to_word, text="Обзор...", command=self.browse_pdf_to_word).pack(pady=5)
        ctk.CTkButton(self.tab_pdf_to_word, text="Конвертировать PDF в Word", 
                      command=self.convert_pdf_to_word, fg_color="green", hover_color="darkgreen", height=40).pack(pady=20)
        self.progress_pdf_to_word = ctk.CTkProgressBar(self.tab_pdf_to_word, width=550)
        self.progress_pdf_to_word.pack(pady=10)
        self.progress_pdf_to_word.set(0)
        self.label_status_1 = ctk.CTkLabel(self.tab_pdf_to_word, text="")
        self.label_status_1.pack(pady=5)

    def browse_pdf_to_word(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")])
        if filename:
            self.entry_pdf_to_word.delete(0, "end")
            self.entry_pdf_to_word.insert(0, filename)

    def convert_pdf_to_word(self):
        pdf_path = self.entry_pdf_to_word.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Ошибка", "Файл не выбран или не существует!")
            return
        docx_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word файлы", "*.docx")])
        if not docx_path: return

        self.label_status_1.configure(text="Конвертация... Пожалуйста, подождите.")
        self.progress_pdf_to_word.set(0.5)
        threading.Thread(target=self._run_pdf_to_word, args=(pdf_path, docx_path), daemon=True).start()

    def _run_pdf_to_word(self, pdf_path, docx_path):
        try:
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()
            self.after(0, lambda: self.progress_pdf_to_word.set(1.0))
            self.after(0, lambda: self.label_status_1.configure(text="Успешно сохранено!"))
            self.after(0, lambda: messagebox.showinfo("Успех", f"Файл сохранен:\n{docx_path}"))
        except Exception as e:
            self.after(0, lambda: self.label_status_1.configure(text="Ошибка!"))
            self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))

    # ==================== ВКЛАДКА 2: WORD -> PDF ====================
    def setup_word_to_pdf_tab(self):
        ctk.CTkLabel(self.tab_word_to_pdf, text="Выберите Word файл (.docx) для конвертации в PDF:", font=("Arial", 14)).pack(pady=(20, 10))
        self.entry_word_to_pdf = ctk.CTkEntry(self.tab_word_to_pdf, width=550)
        self.entry_word_to_pdf.pack(pady=5)
        ctk.CTkButton(self.tab_word_to_pdf, text="Обзор...", command=self.browse_word_to_pdf).pack(pady=5)
        ctk.CTkButton(self.tab_word_to_pdf, text="Конвертировать Word в PDF", 
                      command=self.convert_word_to_pdf, fg_color="green", hover_color="darkgreen", height=40).pack(pady=20)
        ctk.CTkLabel(self.tab_word_to_pdf, text="Примечание: сохраняется текст и базовое форматирование.", text_color="gray").pack(pady=5)

    def browse_word_to_pdf(self):
        filename = filedialog.askopenfilename(filetypes=[("Word файлы", "*.docx")])
        if filename:
            self.entry_word_to_pdf.delete(0, "end")
            self.entry_word_to_pdf.insert(0, filename)

    def convert_word_to_pdf(self):
        word_path = self.entry_word_to_pdf.get()
        if not word_path or not os.path.exists(word_path):
            messagebox.showerror("Ошибка", "Файл не выбран!")
            return
        pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF файлы", "*.pdf")])
        if not pdf_path: return

        try:
            doc = Document(word_path)
            pdf = FPDF()
            pdf.add_page()
            # Пытаемся загрузить шрифт с кириллицей, если он лежит рядом
            font_path = get_resource_path("DejaVuSans.ttf")
            if os.path.exists(font_path):
                pdf.add_font("DejaVu", "", font_path, uni=True)
                pdf.set_font("DejaVu", size=12)
            else:
                pdf.set_font("Arial", size=12) # Без кириллицы
            
            for para in doc.paragraphs:
                pdf.multi_cell(0, 10, txt=para.text)
            pdf.output(pdf_path)
            messagebox.showinfo("Успех", f"Текст из Word сохранен в PDF:\n{pdf_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ВКЛАДКА 3: РАЗДЕЛИТЬ PDF ====================
    def setup_split_tab(self):
        ctk.CTkLabel(self.tab_split, text="Выберите PDF для разделения на отдельные страницы:", font=("Arial", 14)).pack(pady=(20, 10))
        self.entry_split = ctk.CTkEntry(self.tab_split, width=550)
        self.entry_split.pack(pady=5)
        ctk.CTkButton(self.tab_split, text="Обзор...", command=self.browse_split).pack(pady=5)
        ctk.CTkButton(self.tab_split, text="Разделить на страницы", 
                      command=self.split_pdf, fg_color="orange", hover_color="darkorange", height=40).pack(pady=20)
        self.label_status_3 = ctk.CTkLabel(self.tab_split, text="")
        self.label_status_3.pack(pady=5)

    def browse_split(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF файлы", "*.pdf")])
        if filename:
            self.entry_split.delete(0, "end")
            self.entry_split.insert(0, filename)

    def split_pdf(self):
        pdf_path = self.entry_split.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Ошибка", "Файл не выбран!")
            return
        output_dir = filedialog.askdirectory(title="Выберите папку для сохранения страниц")
        if not output_dir: return

        try:
            reader = PdfReader(pdf_path)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            for i, page in enumerate(reader.pages):
                writer = PdfWriter()
                writer.add_page(page)
                output_filename = os.path.join(output_dir, f"{base_name}_page_{i+1}.pdf")
                with open(output_filename, "wb") as out_file:
                    writer.write(out_file)
            self.label_status_3.configure(text=f"Успешно! Создано {len(reader.pages)} файлов.")
            messagebox.showinfo("Успех", f"PDF разделен на {len(reader.pages)} страниц.")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # ==================== ВКЛАДКА 4: ОБЪЕДИНИТЬ PDF ====================
    def setup_merge_tab(self):
        ctk.CTkLabel(self.tab_merge, text="Выберите несколько PDF файлов для объединения:", font=("Arial", 14)).pack(pady=(20, 10))
        self.listbox_merge = ctk.CTkTextbox(self.tab_merge, width=550, height=200)
        self.listbox_merge.pack(pady=5)
        self.listbox_merge.configure(state="disabled")
        self.merge_files = []

        frame_btns = ctk.CTkFrame(self.tab_merge, fg_color="transparent")
        frame_btns.pack(pady=10)
        ctk.CTkButton(frame_btns, text="Добавить файлы", command=self.add_merge_files).grid(row=0, column=0, padx=5)
        ctk.CTkButton(frame_btns, text="Очистить список", command=self.clear_merge_files, fg_color="red").grid(row=0, column=1, padx=5)
        ctk.CTkButton(self.tab_merge, text="Объединить PDF", 
                      command=self.merge_pdfs, fg_color="purple", hover_color="darkviolet", height=40).pack(pady=20)
        self.label_status_4 = ctk.CTkLabel(self.tab_merge, text="")
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
        if not output_path: return

        try:
            writer = PdfWriter()
            for pdf in self.merge_files:
                reader = PdfReader(pdf)
                for page in reader.pages:
                    writer.add_page(page)
            with open(output_path, "wb") as out_file:
                writer.write(out_file)
            self.label_status_4.configure(text="Файлы успешно объединены!")
            messagebox.showinfo("Успех", f"Файлы объединены в:\n{output_path}")
            self.clear_merge_files()
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

if __name__ == "__main__":
    app = PDFConverterApp()
    app.mainloop()
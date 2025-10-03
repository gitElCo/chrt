import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from .base_tab import BaseTab
from utils.file_utils import add_files, clear_files
from utils.error_utils import handle_error, log_info
from com_utils.kompas_utils import analyze_drawings
import openpyxl

class AnalysisTab(BaseTab):
    def __init__(self, notebook, kompas_app, root):
        self.tab_title = "Анализ трудоемкости"
        super().__init__(notebook, kompas_app, root)

    def setup_ui(self):
        self.setup_file_ui()

        self.progress_bar = ttk.Progressbar(self.tab, mode="determinate")
        self.progress_bar.pack(pady=10)

        ttk.Button(self.tab, text="Проанализировать", command=self.process_files).pack(pady=20)

    def process_files(self):
        files = self.file_listbox.get(0, tk.END)

        if not files:
            handle_error("Выберите файлы для анализа!")
            return

        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Файл", "Формат", "Масштаб", "Количество размеров", "Технические требования", "Количество видов", "Трудоемкость"])

        total = len(files)
        self.progress_bar["maximum"] = total
        for i, file in enumerate(files):
            if not os.path.exists(file):
                handle_error(f"Файл {file} не существует")
                continue

            try:
                result = analyze_drawings(self.kompas_app, file)
                sheet.append([
                    result["filename"],
                    result["format"],
                    result["scale"],
                    result["dimensions"],
                    result["technical_requirements"],
                    result["views"],
                    result["labor_intensity"]
                ])
            except Exception as e:
                handle_error(f"Ошибка при анализе файла {file}: {e}")
            finally:
                self.progress_bar.step(1)
                self.root.update_idletasks()

        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if output_file:
            workbook.save(output_file)
            log_info(f"Результаты сохранены в {output_file}")
            messagebox.showinfo("Успех", f"Результаты анализа сохранены в {output_file}")

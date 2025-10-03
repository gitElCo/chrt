import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from utils.file_utils import add_files, clear_files, get_output_filename, create_output_folder
from utils.error_utils import handle_error, log_info
from com_utils.kompas_utils import export_to_step
from .base_tab import BaseTab

class STEPTab(BaseTab):
    def __init__(self, notebook, kompas_app, root):  # ← Добавлен root
        self.tab_title = "Экспорт в STEP"
        super().__init__(notebook, kompas_app, root)

    def setup_ui(self):
        self.setup_file_ui()

        ttk.Label(self.tab, text="Версия STEP:").pack(pady=5)
        self.step_version = tk.StringVar(value="AP203")
        versions = ["AP203", "AP214", "AP242"]
        version_combobox = ttk.Combobox(self.tab, textvariable=self.step_version, values=versions, state="readonly")
        version_combobox.pack(pady=5)

        ttk.Label(self.tab, text="Папка для STEP:").pack(pady=5)
        self.output_path = tk.StringVar()
        ttk.Entry(self.tab, textvariable=self.output_path, width=80).pack(pady=5)
        ttk.Button(self.tab, text="Выбрать папку", command=self.select_output_folder).pack(pady=5)

        self.progress_bar = ttk.Progressbar(self.tab, mode="determinate")
        self.progress_bar.pack(pady=10)

        ttk.Button(self.tab, text="Экспортировать в STEP", command=self.process_files).pack(pady=20)

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_path.set(create_output_folder(folder))

    def process_files(self):
        files = self.file_listbox.get(0, tk.END)
        output_folder = self.output_path.get()
        step_version = self.step_version.get()

        if not files or not output_folder:
            handle_error("Выберите файлы и папку для сохранения!")
            return

        total = len(files)
        self.progress_bar["maximum"] = total
        for i, file in enumerate(files):
            if not os.path.exists(file):
                handle_error(f"Файл {file} не существует")
                continue

            output_file = os.path.join(output_folder, get_output_filename(file, ".step"))
            try:
                success = export_to_step(self.kompas_app, file, output_file, step_version)
                if success:
                    log_info(f"Файл {file} успешно экспортирован в {output_file}")
                else:
                    handle_error(f"Не удалось экспортировать файл: {file}")
            except Exception as e:
                handle_error(f"Ошибка при экспорте файла {file}: {e}")
            finally:
                self.progress_bar.step(1)
                self.root.update_idletasks()

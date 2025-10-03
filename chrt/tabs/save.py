import tkinter as tk
from tkinter import ttk
from .base_tab import BaseTab
from utils.file_utils import add_files, clear_files, get_output_filename, create_output_folder
from utils.error_utils import handle_error, log_info
from com_utils.kompas_utils import save_to_version

class SaveTab(BaseTab):
    def __init__(self, notebook, kompas_app, root):  # ✅ 3 параметра
        self.tab_title = "Сохранение в версии"
        super().__init__(notebook, kompas_app, root)  # ✅ Передаем root

    def setup_ui(self):
        self.setup_file_ui()

        ttk.Label(self.tab, text="Версия для сохранения:").pack(pady=5)
        self.version_var = tk.StringVar(value="Текущая версия")
        versions = ["Текущая версия", "Предыдущая версия", "Версия 5.11"]
        version_combobox = ttk.Combobox(self.tab, textvariable=self.version_var, values=versions, state="readonly")
        version_combobox.pack(pady=5)

        ttk.Label(self.tab, text="Папка для сохранения:").pack(pady=5)
        self.output_path = tk.StringVar()
        ttk.Entry(self.tab, textvariable=self.output_path, width=80).pack(pady=5)
        ttk.Button(self.tab, text="Выбрать папку", command=self.select_output_folder).pack(pady=5)

        self.progress_bar = ttk.Progressbar(self.tab, mode="determinate")
        self.progress_bar.pack(pady=10)

        ttk.Button(self.tab, text="Сохранить", command=self.process_files).pack(pady=20)

    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_path.set(create_output_folder(folder))

    def process_files(self):
        files = self.file_listbox.get(0, tk.END)
        output_folder = self.output_path.get()
        version = self.version_var.get()

        if not files or not output_folder:
            handle_error("Выберите файлы и папку для сохранения!")
            return

        version_map = {
            "Текущая версия": 0,
            "Предыдущая версия": -1,
            "Версия 5.11": 1,
        }

        total = len(files)
        self.progress_bar["maximum"] = total
        for i, file in enumerate(files):
            output_file = os.path.join(output_folder, get_output_filename(file, os.path.splitext(file)[1]))
            try:
                success = save_to_version(self.kompas_app, file, output_file, version_map[version])
                if success:
                    log_info(f"Файл {file} успешно сохранен в {output_file}")
                else:
                    handle_error(f"Не удалось сохранить файл: {file}")
            except Exception as e:
                handle_error(f"Ошибка при сохранении файла {file}: {e}")
            finally:
                self.progress_bar.step(1)
                self.root.update_idletasks()

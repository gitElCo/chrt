import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from utils.file_utils import add_files, clear_files

class BaseTab:
    def __init__(self, notebook, kompas_app, root):  # ✅ 3 параметра
        self.tab = ttk.Frame(notebook)
        notebook.add(self.tab, text=self.tab_title)
        self.kompas_app = kompas_app
        self.root = root  # ✅ Сохраняем корневое окно
        self.setup_ui()

    def setup_file_ui(self):
        self.file_listbox = tk.Listbox(self.tab, height=10, width=80)
        self.file_listbox.pack(pady=5)

        file_button_frame = ttk.Frame(self.tab)
        file_button_frame.pack(pady=5)

        ttk.Button(
            file_button_frame,
            text="Добавить файлы",
            command=lambda: add_files(
                self.file_listbox,
                [("KOMPAS Files", "*.cdw;*.frw;*.m3d;*.a3d")]
            )
        ).pack(side="left", padx=5)

        ttk.Button(
            file_button_frame,
            text="Очистить список",
            command=lambda: clear_files(self.file_listbox)
        ).pack(side="left", padx=5)

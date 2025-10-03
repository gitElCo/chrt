import os
import tkinter as tk
from tkinter import ttk, messagebox
from ttkthemes import ThemedTk
from com_utils.kompas_utils import initialize_kompas
from tabs.save import SaveTab
from tabs.pdf import PDFTab
from tabs.step import STEPTab
from tabs.analysis import AnalysisTab
from theme_manager import ThemeManager
import pythoncom

class KompasBatchProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("КОМПАС Batch Processor")
        self.root.geometry("1000x700")

        pythoncom.CoInitialize()  # ← Теперь pythoncom определён

        self.kompas_app = initialize_kompas()

        self.theme_manager = ThemeManager(self.root)

        # Создаем notebook до инициализации вкладок
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Инициализируем вкладки
        self.save_tab = SaveTab(self.notebook, self.kompas_app, self.root)
        self.pdf_tab = PDFTab(self.notebook, self.kompas_app, self.root)
        self.step_tab = STEPTab(self.notebook, self.kompas_app, self.root)
        self.analysis_tab = AnalysisTab(self.notebook, self.kompas_app, self.root)

        self.create_menu()

    def create_menu(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Выход", command=self.root.quit)
        menubar.add_cascade(label="Файл", menu=file_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        self.theme_manager.create_theme_menu(view_menu)
        menubar.add_cascade(label="Вид", menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="О программе", command=self.show_about)
        menubar.add_cascade(label="Справка", menu=help_menu)

        self.root.config(menu=menubar)

    def show_about(self):
        messagebox.showinfo(
            "О программе",
            "КОМПАС Batch Processor\nВерсия 1.0\nАвтор: Николай Тарейкин"
        )

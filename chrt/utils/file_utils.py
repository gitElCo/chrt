import os
import tkinter as tk
from tkinter import filedialog

def add_files(listbox, filetypes):
    filenames = filedialog.askopenfilenames(
        title="Выберите файлы",
        filetypes=filetypes
    )
    if filenames:
        for filename in filenames:
            if filename not in listbox.get(0, tk.END):
                listbox.insert(tk.END, filename)

def clear_files(listbox):
    listbox.delete(0, tk.END)

def get_output_filename(input_path, output_ext):
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    return f"{base_name}{output_ext}"

def create_output_folder(folder_path):
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

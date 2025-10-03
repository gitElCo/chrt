import tkinter as tk

class ThemeManager:
    def __init__(self, root):
        self.root = root
        self.theme_mode = tk.StringVar(value="plastik")

    def create_theme_menu(self, menu):
        menu.add_radiobutton(
            label="Светлая тема (Plastik)",
            variable=self.theme_mode,
            value="plastik",
            command=self.change_theme
        )
        menu.add_radiobutton(
            label="Темная тема (Equilux)",
            variable=self.theme_mode,
            value="equilux",
            command=self.change_theme
        )

    def change_theme(self):
        try:
            self.root.set_theme(self.theme_mode.get())
        except Exception as e:
            print(f"Ошибка смены темы: {e}")

import os
import sys
import ctypes
from ttkthemes import ThemedTk
from app import KompasBatchProcessor
import pythoncom

if __name__ == "__main__":
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
 
    pythoncom.CoInitialize() 

    root = ThemedTk(theme="plastik")
    app = KompasBatchProcessor(root)
    root.mainloop()

    pythoncom.CoUninitialize()

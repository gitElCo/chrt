import logging
from tkinter import messagebox

logging.basicConfig(
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filename="kompas_batch_errors.log"
)

def handle_error(error, title="Ошибка"):
    logging.error(str(error))
    messagebox.showerror(title, str(error))

def log_info(message):
    logging.info(message)

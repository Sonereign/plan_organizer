import tkinter as tk
from gui import PlanoKratiseonApp
from logger import logger


if __name__ == "__main__":
    logger.info("Starting PlanoKratiseonApp...")

    root = tk.Tk()
    app = PlanoKratiseonApp(root)

    logger.info("Application is running.")
    try:
        root.mainloop()
    except Exception as e:
        logger.exception(f"An unexpected error occurred {e}")
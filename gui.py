import tkinter as tk
import threading
from datetime import datetime
from logger import logger  # Import the global logger

from file_handler import select_file
from processing import process_files


class PlanoKratiseonApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Πλάνο Κρατήσεων")
        self.root.geometry("600x700")
        self.root.configure(bg="#f0f0f0")

        self.cleanup_outputs = False
        self.availability_per_zone_path = None
        self.availability_per_type_path = None
        self.availability_per_nationality_path = None
        self.previous_years_nationality_paths = {}
        self.previous_years_zone_paths = {}

        self.create_widgets()
        logger.info("Application UI initialized.")

    def create_widgets(self):
        title_label = tk.Label(self.root, text="Πλάνο Κρατήσεων", font=("Arial", 18, "bold"), bg="#f0f0f0")
        title_label.pack(pady=10)

        file_frame = tk.Frame(self.root, bg="#f0f0f0")
        file_frame.pack(pady=10, fill="x", padx=20)

        zone_frame = tk.LabelFrame(file_frame, text="Availability Per Zone & Previous Years",
                                   font=("Arial", 12, "bold"), bg="#f0f0f0", padx=10, pady=10)
        zone_frame.pack(fill="x", pady=10)

        self.availability_per_zone_text = self.create_file_section(zone_frame, "Availability per Zone Current Year")
        self.previous_years_zone_frame = tk.Frame(zone_frame, bg="#f0f0f0")
        self.previous_years_zone_frame.pack(fill="x", pady=5)

        add_zone_year_button = tk.Button(
            zone_frame, text="+ Add Year", command=self.add_previous_year_per_zone,
            font=("Arial", 10), bg="#008CBA", fg="white"
        )
        add_zone_year_button.pack(pady=5)

        type_frame = tk.LabelFrame(file_frame, text="Availability Per Type",
                                   font=("Arial", 12, "bold"), bg="#f0f0f0", padx=10, pady=10)
        type_frame.pack(fill="x", pady=10)

        self.availability_per_type_text = self.create_file_section(type_frame, "Availability per Type Current Year")

        nationality_frame = tk.LabelFrame(file_frame, text="Availability Per Nationality & Previous Years",
                                          font=("Arial", 12, "bold"), bg="#f0f0f0", padx=10, pady=10)
        nationality_frame.pack(fill="x", pady=10)

        self.availability_per_nationality_text = self.create_file_section(nationality_frame,
                                                                          "Availability per Nationality Current Year")
        self.previous_years_nationality_frame = tk.Frame(nationality_frame, bg="#f0f0f0")
        self.previous_years_nationality_frame.pack(fill="x", pady=5)

        add_year_button = tk.Button(
            nationality_frame, text="+ Add Year", command=self.add_previous_year_per_nationality,
            font=("Arial", 10), bg="#008CBA", fg="white"
        )
        add_year_button.pack(pady=5)

        self.process_button = tk.Button(
            self.root, text="Process Files", command=self.start_processing,
            font=("Arial", 12), bg="#4CAF50", fg="white", padx=10, pady=5
        )
        self.process_button.pack(pady=20)

        self.cleanup_var = tk.BooleanVar(value=True)
        self.cleanup_checkbox = tk.Checkbutton(self.root, text="Enable Cleanup", variable=self.cleanup_var,
                                               command=self.toggle_cleanup)
        self.cleanup_checkbox.pack()

        self.load_logo()

        self.status_label = tk.Label(self.root, text="", fg="blue", bg="#f0f0f0", font=("Arial", 10))
        self.status_label.pack(pady=10)

    def load_logo(self):
        try:
            self.logo_image = tk.PhotoImage(file="logo.png")  # Only supports PNG
            self.logo_label = tk.Label(self.root, image=self.logo_image, bg="#f0f0f0")
            self.logo_label.pack(pady=10)
            logger.info("Logo loaded successfully.")
        except Exception as e:
            logger.error(f"Error loading logo: {e}")

    def create_file_section(self, parent, label_text):
        frame = tk.Frame(parent, bg="#f0f0f0")
        frame.pack(fill="x", pady=5)

        label = tk.Label(frame, text=label_text, font=("Arial", 12), bg="#f0f0f0")
        label.pack(side="left", padx=5)

        text_widget = tk.Entry(frame, width=40, font=("Arial", 10))
        text_widget.pack(side="left", padx=5, pady=5)

        button = tk.Button(
            frame, text="Browse", command=lambda: select_file(label_text, text_widget, self),
            font=("Arial", 10), bg="#008CBA", fg="white"
        )
        button.pack(side="right", padx=5)

        return text_widget

    def add_previous_year_per_zone(self):
        current_year = datetime.now().year
        year = current_year - 1 - len(self.previous_years_zone_paths)
        frame = tk.Frame(self.previous_years_zone_frame, bg="#f0f0f0")
        frame.pack(fill="x", pady=2)

        label = tk.Label(frame, text=f"Year {year}", font=("Arial", 12), bg="#f0f0f0")
        label.pack(side="left", padx=5)

        text_widget = tk.Entry(frame, width=35, font=("Arial", 10))
        text_widget.pack(side="left", padx=5, pady=5)

        button = tk.Button(
            frame, text="Browse", command=lambda: select_file(f"Availability per Zone Year {year}", text_widget, self),
            font=("Arial", 10), bg="#008CBA", fg="white"
        )
        button.pack(side="right", padx=5)

        self.previous_years_zone_paths[year] = text_widget
        logger.info(f"Added input field for previous zone year: {year}")

    def add_previous_year_per_nationality(self):
        current_year = datetime.now().year
        year = current_year - 1 - len(self.previous_years_nationality_paths)
        frame = tk.Frame(self.previous_years_nationality_frame, bg="#f0f0f0")
        frame.pack(fill="x", pady=2)

        label = tk.Label(frame, text=f"Year {year}", font=("Arial", 12), bg="#f0f0f0")
        label.pack(side="left", padx=5)

        text_widget = tk.Entry(frame, width=35, font=("Arial", 10))
        text_widget.pack(side="left", padx=5, pady=5)

        button = tk.Button(
            frame, text="Browse",
            command=lambda: select_file(f"Availability per Nationality Year {year}", text_widget, self),
            font=("Arial", 10), bg="#008CBA", fg="white"
        )
        button.pack(side="right", padx=5)

        self.previous_years_nationality_paths[year] = text_widget
        logger.info(f"Added input field for previous year: {year}")

    def start_processing(self):
        self.process_button.config(state=tk.DISABLED)
        logger.info("Processing started.")

        try:
            threading.Thread(target=process_files, args=(self,), daemon=True).start()
            logger.info("Processing thread started successfully.")
        except Exception as e:
            logger.exception(f"Error starting processing thread. {e}")

    def toggle_cleanup(self):
        if self.cleanup_var.get():
            logger.info("Temporary files will be removed.")
        else:
            logger.info("Temporary files will stay.")

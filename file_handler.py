import re
from tkinter import filedialog
from logger import logger


def select_file(file_description, text_widget, app):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        text_widget.delete(0, 'end')
        text_widget.insert(0, file_path)

        print(f"Selected {file_description}: {file_path}")  # Debug print

        # Extract year from description (e.g., "Availability per Zone Year 2024")
        match = re.search(r'\b(\d{4})\b', file_description)
        if match:
            year = int(match.group(1))
            if "Availability per Nationality Year" in file_description:
                app.previous_years_nationality_paths[year] = file_path
            elif "Availability per Zone Year" in file_description:
                app.previous_years_zone_paths[year] = file_path
        else:
            if file_description == "Availability per Zone Current Year":
                app.availability_per_zone_path = file_path
            elif file_description == "Availability per Type Current Year":
                app.availability_per_type_path = file_path
            elif file_description == "Availability per Nationality Current Year":
                app.availability_per_nationality_path = file_path

        logger.debug(f"app.availability_per_zone_path: {app.availability_per_zone_path}")
        logger.debug(f"app.availability_per_type_path: {app.availability_per_type_path}")
        logger.debug(f"app.availability_per_nationality_path: {app.availability_per_nationality_path}")
        logger.debug(f"app.previous_years_nationality_paths: {app.previous_years_nationality_paths}")
        logger.debug(f"app.previous_years_zone_paths: {app.previous_years_zone_paths}")
        logger.debug('=====================================')
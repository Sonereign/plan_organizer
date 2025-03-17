from tkinter import filedialog

def select_file(file_description, text_widget, app):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        text_widget.delete(0, 'end')
        text_widget.insert(0, file_path)

        if file_description == "Availability Per Zone":
            app.availability_per_zone_path = file_path
        elif file_description == "Availability Per Type":
            app.availability_per_type_path = file_path
        elif file_description == "Current Year":
            app.availability_per_nationality_path = file_path
        elif "Year" in file_description:
            year = int(file_description.split()[1])
            app.previous_years_paths[year] = file_path

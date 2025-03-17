import os
from copy import copy
from datetime import datetime
from tkinter import messagebox

import openpyxl
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import get_column_letter

from stage1 import stage1
from stage2 import stage2
from stage3 import stage3
from stage4 import stage4
from stage5 import stage5
from stage6 import stage6
from stage7 import stage7
from stage8 import stage8
from stage9 import stage9
from stage10 import stage10


def process_files(app):
    try:
        # Generate final output file name and sheet names
        today = datetime.today().strftime("%d-%m-%y")
        final_output = f"{today}_availabilityPerZone&Nationality.xlsx"
        sheet1_name = f"{today}-πληρότητα-units"
        sheet2_name = "εθνικότητες"

        stage1_output = "stage1_output.xlsx"
        stage2_output = "stage2_output.xlsx"
        stage3_output = "stage3_output.xlsx"
        stage4_output = "stage4_output.xlsx"
        stage5_output = "stage5_output.xlsx"
        stage6_output_filenames = []
        stage7_output = "stage7_output.xlsx"
        stage8_output = "stage8_output.xlsx"
        stage9_output = "stage9_output.xlsx"
        stage10_output = "stage10_output.xlsx"

        stage1(app.availability_per_zone_path, stage1_output)
        stage2(stage1_output, app.availability_per_type_path, stage2_output)
        stage3(stage2_output, stage3_output)
        stage4(stage3_output, stage4_output)

        # Run stage5 if nationality file is provided
        if app.availability_per_nationality_path:
            stage5(app.availability_per_nationality_path, stage5_output)

        # Run Stage 6 for previous years
        for year, file_path in app.previous_years_paths.items():
            output_file = f"stage6_Output_{year}.xlsx"
            stage6_output_filenames.append(output_file)  # Append file name to list
            stage6(input_file=file_path, output_file=output_file, year=year)

        previous_years = list(app.previous_years_paths.keys())  # Extract years from dictionary keys
        number_of_previous_year_data = len(stage6_output_filenames)

        if not stage6_output_filenames:
            # Create final output by combining sheets from stage4 and stage5 outputs
            combine_sheets(stage4_output=stage4_output, stage5_output=stage5_output, final_output=final_output,
                               sheet1_name=sheet1_name, sheet2_name=sheet2_name, app=app)
        else:
            # Create final output by combining sheets from stage4 and stage10 outputs
            stage7(stage5_output, stage6_output_filenames, stage7_output)
            stage8(stage7_output, stage8_output, previous_years, number_of_previous_year_data)
            stage9(stage8_output, stage9_output, previous_years)
            stage10(stage9_output, stage10_output, previous_years)
            combine_sheets(stage4_output, stage10_output, final_output, sheet1_name, sheet2_name, app=app)

        app.status_label.config(text="Processing complete!")
        messagebox.showinfo("Success", f"Final output saved as {final_output}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        app.process_button.config(state="normal")
        if app.cleanup_var:
            # Clean up temporary files
            for file in [stage1_output, stage2_output, stage3_output, stage4_output, stage5_output, stage7_output,
                         stage8_output, stage9_output, stage10_output]:
                if os.path.exists(file):
                    os.remove(file)
            for file in stage6_output_filenames:
                if os.path.exists(file):
                    os.remove(file)


def combine_sheets(stage4_output, stage5_output, final_output, sheet1_name, sheet2_name, app):
    """Combine sheets from stage4 and stage5 outputs into a single Excel file."""
    # Load workbooks
    wb_stage4 = load_workbook(stage4_output)
    wb_stage5 = load_workbook(stage5_output) if app.availability_per_nationality_path else None

    # Create a new workbook for the final output
    wb_final = load_workbook(stage4_output)  # Start with a copy of stage4 output

    # Rename the sheet from stage4 to the custom sheet1 name
    sheet_stage4 = wb_final.active
    sheet_stage4.title = sheet1_name

    # Add sheet from stage5 if available
    if wb_stage5:
        sheet_stage5 = wb_stage5.active
        # Create a new sheet in the final workbook for Stage5 Results
        new_sheet_stage5 = wb_final.create_sheet(sheet2_name)

        # Copy all cells, styles, and formulas from Stage5 sheet to the new sheet
        for row in sheet_stage5.iter_rows():
            for cell in row:
                new_cell = new_sheet_stage5.cell(
                    row=cell.row, column=cell.column, value=cell.value
                )
                if cell.has_style:
                    new_cell.font = copy(cell.font)  # Use copy function
                    new_cell.border = copy(cell.border)  # Use copy function
                    new_cell.fill = copy(cell.fill)  # Use copy function
                    new_cell.number_format = cell.number_format
                    new_cell.protection = copy(cell.protection)  # Use copy function
                    new_cell.alignment = copy(cell.alignment)  # Use copy function

    # Save the final workbook
    wb_final.save(final_output)
    apply_conditional_formatting(final_output)


def apply_conditional_formatting(file_path):
    # Load the workbook and select the specified worksheet
    wb = openpyxl.load_workbook(file_path)
    ws = wb["εθνικότητες"]

    ws.freeze_panes = "B2"

    # Find columns that start with "Percent difference"
    header_row = ws[1]  # Assuming headers are in the first row
    percent_diff_cols = [cell.column for cell in header_row if
                         cell.value and cell.value.startswith("Percent difference")]

    # Apply conditional formatting to each identified column
    for percent_diff_col in percent_diff_cols:
        percent_diff_range = f"{get_column_letter(percent_diff_col)}2:{get_column_letter(percent_diff_col)}{ws.max_row}"
        color_scale_rule = ColorScaleRule(
            start_type="num", start_value=-1, start_color="FFCCCC",  # Red for negative
            mid_type="num", mid_value=0, mid_color="FFFFFF",  # White for neutral
            end_type="num", end_value=1, end_color="CCFFCC"  # Green for positive
        )
        ws.conditional_formatting.add(percent_diff_range, color_scale_rule)

    # Save the workbook
    wb.save(file_path)
    print("Conditional formatting applied successfully!")

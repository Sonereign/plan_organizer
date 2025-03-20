import os
from copy import copy
from datetime import datetime
from tkinter import messagebox

import openpyxl
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import get_column_letter

from per_zone_stage1 import per_zone_stage1
from per_zone_stage2 import per_zone_stage2
from per_zone_stage3 import per_zone_stage3
from per_zone_stage4 import per_zone_stage4
from per_nat_stage1 import per_nat_stage1
from per_nat_stage2 import per_nat_stage2
from per_nat_stage3 import per_nat_stage3
from per_nat_stage4 import per_nat_stage4
from per_nat_stage5 import per_nat_stage5
from per_nat_stage6 import per_nat_stage6
from logger import logger


def process_files(app):
    try:
        # Generate final output file name and sheet names
        today = datetime.today().strftime("%d-%m-%y")
        final_output = f"{today}_availabilityPerZone&Nationality.xlsx"
        sheet1_name = f"{today}-πληρότητα-units"
        sheet2_name = "εθνικότητες"

        per_zone_stage1_output = "per_zone_stage1_output.xlsx"
        per_zone_stage2_output = "per_zone_stage2_output.xlsx"
        per_zone_stage3_output = "per_zone_stage3_output.xlsx"
        per_zone_stage4_output = "per_zone_stage4_output.xlsx"
        per_nat_stage1_output = "per_nat_stage1_output.xlsx"
        per_nat_stage2_output_filenames = []
        per_nat_stage3_output = "per_nat_stage3_output.xlsx"
        per_nat_stage4_output = "per_nat_stage4_output.xlsx"
        per_nat_stage5_output = "per_nat_stage5_output.xlsx"
        per_nat_stage6_output = "per_nat_stage6_output.xlsx"

        per_zone_stage1(app.availability_per_zone_path, per_zone_stage1_output)
        per_zone_stage2(per_zone_stage1_output, app.availability_per_type_path, per_zone_stage2_output)
        per_zone_stage3(per_zone_stage2_output, per_zone_stage3_output)
        per_zone_stage4(per_zone_stage3_output, per_zone_stage4_output)

        # Run stage5 if nationality file is provided
        if app.availability_per_nationality_path:
            per_nat_stage1(app.availability_per_nationality_path, per_nat_stage1_output)

        # Run Stage 6 for previous years
        for year, file_path in app.previous_years_paths.items():
            output_file = f"per_nat_stage2_output_{year}.xlsx"
            per_nat_stage2_output_filenames.append(output_file)  # Append file name to list
            per_nat_stage2(input_file=file_path, output_file=output_file, year=year)

        previous_years = list(app.previous_years_paths.keys())  # Extract years from dictionary keys
        number_of_previous_year_data = len(per_nat_stage2_output_filenames)

        if not per_nat_stage2_output_filenames:
            # Create final output by combining sheets from stage4 and stage5 outputs
            combine_sheets(per_zone_stage4=per_zone_stage4_output, per_nat_stage1=per_nat_stage1_output, final_output=final_output,
                           sheet1_name=sheet1_name, sheet2_name=sheet2_name, app=app)
        else:
            # Create final output by combining sheets from stage4 and stage10 outputs
            per_nat_stage3(per_nat_stage1_output, per_nat_stage2_output_filenames, per_nat_stage3_output)
            per_nat_stage4(per_nat_stage3_output, per_nat_stage4_output, previous_years, number_of_previous_year_data)
            per_nat_stage5(per_nat_stage4_output, per_nat_stage5_output, previous_years)
            per_nat_stage6(per_nat_stage5_output, per_nat_stage6_output, previous_years)
            combine_sheets(per_zone_stage4_output, per_nat_stage6_output, final_output, sheet1_name, sheet2_name, app=app)

        app.status_label.config(text="Processing complete!")
        messagebox.showinfo("Success", f"Final output saved as {final_output}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        app.process_button.config(state="normal")
        if app.cleanup_var:
            # Clean up temporary files
            for file in [per_zone_stage1_output, per_zone_stage2_output, per_zone_stage3_output, per_zone_stage4_output, per_nat_stage1_output, per_nat_stage3_output,
                         per_nat_stage4_output, per_nat_stage5_output, per_nat_stage6_output]:
                if os.path.exists(file):
                    os.remove(file)
            for file in per_nat_stage2_output_filenames:
                if os.path.exists(file):
                    os.remove(file)


def combine_sheets(per_zone_stage4, per_nat_stage1, final_output, sheet1_name, sheet2_name, app):
    """Combine sheets from stage4 and stage5 outputs into a single Excel file."""
    # Load workbooks
    wb_stage4 = load_workbook(per_zone_stage4)
    wb_stage5 = load_workbook(per_nat_stage1) if app.availability_per_nationality_path else None

    # Create a new workbook for the final output
    wb_final = load_workbook(per_zone_stage4)  # Start with a copy of stage4 output

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
    logger.info("Conditional formatting applied successfully!")

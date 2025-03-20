import os
from copy import copy
from datetime import datetime
from tkinter import messagebox

import openpyxl
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import get_column_letter

from per_zone_per_type_stage1 import per_zone_per_type_stage1
from per_zone_per_type_stage2 import per_zone_per_type_stage2
from per_zone_per_type_stage3 import per_zone_per_type_stage3
from per_zone_per_type_stage4 import per_zone_per_type_stage4
from per_nationality_stage1 import per_nationality_stage1
from per_nationality_stage2_previous_years import per_nationality_stage2_previous_years
from per_nationality_stage3_appending_previous_years import per_nationality_stage3
from per_nationality_stage4 import per_nationality_stage4
from per_nationality_stage5 import per_nationality_stage5
from per_nationality_stage6 import per_nationality_stage6

from logger import logger
from per_zone_per_type_stage5_previous_years import per_zone_per_type_stage5_previous_years


def process_files(app):
    try:
        # Generate final output file name and sheet names
        today = datetime.today().strftime("%d-%m-%y")
        final_output = f"{today}_availabilityPerZone&Nationality.xlsx"
        sheet1_name = f"{today}-πληρότητα-units"
        sheet2_name = "εθνικότητες"

        zone_stage1_output = "zone_stage1_output.xlsx"
        zone_stage2_output = "zone_stage2_output.xlsx"
        zone_stage3_output = "zone_stage3_output.xlsx"
        zone_stage4_output = "zone_stage4_output.xlsx"
        zone_stage5_output_filenames = []

        nat_stage1_output = "nat_stage1_output.xlsx"
        nat_stage2_output_filenames = []
        nat_stage3_output = "nat_stage3_output.xlsx"
        nat_stage4_output = "nat_stage4_output.xlsx"
        nat_stage5_output = "nat_stage5_output.xlsx"
        nat_stage6_output = "nat_stage6_output.xlsx"

        # logger.debug(f'THE AVAILABILITY PATH FOR ZONE IS: !!!!!!!!!!! {app.availability_per_zone_path}')
        per_zone_per_type_stage1(app.availability_per_zone_path, zone_stage1_output)
        per_zone_per_type_stage2(zone_stage1_output, app.availability_per_type_path, zone_stage2_output)
        per_zone_per_type_stage3(zone_stage2_output, zone_stage3_output)
        per_zone_per_type_stage4(zone_stage3_output, zone_stage4_output)

        # Run per_zone_per_type_stage5 for previous years
        for year, file_path in app.previous_years_zone_paths.items():
            if not file_path or "<tkinter" in file_path:  # Skip empty or invalid paths
                continue
            output_file = f"zone_stage5_output_{year}.xlsx"
            zone_stage5_output_filenames.append(output_file)  # Append file name to list
            per_zone_per_type_stage5_previous_years(input_file=file_path, output_file=output_file, year=year)

        # Run per_nationality_stage1 if nationality file is provided
        if app.availability_per_nationality_path:
            per_nationality_stage1(app.availability_per_nationality_path, nat_stage1_output)

        # Run per_nationality_stage2_previous_years for previous years
        for year, file_path in app.previous_years_nationality_paths.items():
            if not file_path or "<tkinter" in file_path:  # Skip empty or invalid paths
                continue
            output_file = f"nat_stage2_output_{year}.xlsx"
            nat_stage2_output_filenames.append(output_file)  # Append file name to list
            per_nationality_stage2_previous_years(input_file=file_path, output_file=output_file, year=year)

        # Extract years from dictionary keys (excluding skipped ones)
        nat_previous_years = [year for year in app.previous_years_nationality_paths.keys() if
                          year in nat_stage2_output_filenames]
        number_of_previous_year_data = len(nat_stage2_output_filenames)

        #Zone
        if not zone_stage5_output_filenames:
            # Do something
            pass
        else:
            # mix zone previous years with current
            pass
         # Zone end

        if not nat_stage2_output_filenames:
            # Create final output by combining sheets from stage4 and stage5 outputs
            combine_sheets(stage4_output=zone_stage4_output, stage5_output=nat_stage1_output, final_output=final_output,
                           sheet1_name=sheet1_name, sheet2_name=sheet2_name, app=app)
        else:
            # Create final output by combining sheets from stage4 and stage10 outputs
            per_nationality_stage3(nat_stage1_output, nat_stage2_output_filenames, nat_stage3_output)
            per_nationality_stage4(nat_stage3_output, nat_stage4_output, nat_previous_years, number_of_previous_year_data)
            per_nationality_stage5(nat_stage4_output, nat_stage5_output, nat_previous_years)
            per_nationality_stage6(nat_stage5_output, nat_stage6_output, nat_previous_years)
            combine_sheets(zone_stage4_output, nat_stage6_output, final_output, sheet1_name, sheet2_name, app=app)

        app.status_label.config(text="Processing complete!")
        messagebox.showinfo("Success", f"Final output saved as {final_output}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        app.process_button.config(state="normal")
        if app.cleanup_var:
            # Clean up temporary files
            for file in [zone_stage1_output, zone_stage2_output, zone_stage3_output, zone_stage4_output,
                         nat_stage1_output, nat_stage3_output,
                         nat_stage4_output, nat_stage5_output, nat_stage6_output]:
                if os.path.exists(file):
                    os.remove(file)
            for file in nat_stage2_output_filenames:
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
    logger.info("Conditional formatting applied successfully!")

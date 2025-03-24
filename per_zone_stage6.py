from datetime import datetime
import pandas as pd
from openpyxl import load_workbook


def load_stage4(input_file):
    return pd.read_excel(input_file)


def load_stage5_files(stage5_files):
    stage5_data = []
    for file in stage5_files:
        df = pd.read_excel(file)
        stage5_data.append((file, df))
    return stage5_data


def parse_date(date_str):
    try:
        # Remove the day of the week (e.g., "Tue ", "Mon ")
        if isinstance(date_str, str):
            date_str = date_str.split(" ")[-1]  # Keep only the date part

        # Add the year if missing (assuming the year is the current year)
        if len(date_str.split("/")) == 2:  # If the date is in "29/04" format
            current_year = datetime.now().year
            date_str += f"/{current_year}"  # Append the year

        # Parse the date with dayfirst=True
        return pd.to_datetime(date_str, dayfirst=True, errors='coerce')
    except:
        return None


def detect_date_range(file_name, df):
    date_columns = df.columns[2:]
    if len(date_columns) > 0:
        start_date = date_columns[0]
        end_date = date_columns[-1]

        # Ensure the end date is a valid date and not a total column
        if isinstance(end_date, str) and end_date.lower() == "total":
            end_date = date_columns[-2] if len(date_columns) > 1 else "Unknown"

        parsed_start = parse_date(start_date)
        parsed_end = parse_date(end_date)

        print(f"File: {file_name} | Date range: {start_date} to {end_date}")
        return parsed_start, parsed_end
    return None, None


def add_empty_columns(input_file, output_file, start_diff):
    """
    Adds empty columns after the "Capacity" column in the Stage 4 file.
    """
    # Load the workbook using openpyxl
    workbook = load_workbook(input_file)
    sheet = workbook.active

    # Find the index of the "Capacity" column
    capacity_col_index = None
    for col in sheet.iter_cols(1, sheet.max_column):
        if col[0].value == "Capacity":
            capacity_col_index = col[0].column
            break

    if capacity_col_index is None:
        print("Error: 'Capacity' column not found in the Excel file.")
        return

    # Add empty columns after the "Capacity" column
    if start_diff > 0:
        sheet.insert_cols(capacity_col_index + 1, start_diff)
        print(f"Added {start_diff} empty columns after the Capacity column.")

    # Save the modified workbook
    workbook.save(output_file)
    print(f"Saved modified Stage 4 file with empty columns to {output_file}.")

def add_empty_cells(sheet, row_index, num_empty_cells):
    """
    Adds a specified number of empty cells at index 2 in a given row.
    """
    if num_empty_cells > 0:
        # Insert empty cells at index 2
        for _ in range(num_empty_cells):
            sheet.insert_cols(2)  # Insert a new column at index 2
            # Set the value of the new cell to None (empty)
            for row in sheet.iter_rows(min_row=row_index, max_row=row_index):
                row[1].value = None

def copy_header(stage5_file, stage4_file, output_file):
    """
    Copies the header from the Stage 5 file and inserts it under the header in the Stage 4 file.
    """
    # Load the Stage 5 file to extract the header
    stage5_df = pd.read_excel(stage5_file)
    stage5_header = stage5_df.columns.tolist()  # Extract header as a list

    # Load the Stage 4 file using openpyxl
    workbook = load_workbook(stage4_file)
    sheet = workbook.active

    # Insert the Stage 5 header as a new row under the Stage 4 header
    sheet.insert_rows(2)  # Insert a new row at position 2 (under the header)
    for col_idx, value in enumerate(stage5_header, start=1):
        sheet.cell(row=2, column=col_idx, value=value)

    # Save the modified workbook
    workbook.save(output_file)
    print(f"Copied header from {stage5_file} to Stage 4 file. Saved to {output_file}.")


def copy_total_rows_from_stage5(stage5_file, stage4_file, output_file):
    """
    Copies rows starting with "Total Accommodation", "Total Youth Hostel", or "Total Camping"
    from the Stage 5 file and inserts them under their respective counterparts in the Stage 4 file.
    """
    # Load the Stage 5 file to extract the total rows
    stage5_df = pd.read_excel(stage5_file)

    # Identify the rows to copy (check if the first column starts with the keywords)
    total_rows = stage5_df[
        stage5_df.iloc[:, 0].str.startswith("Total Accommodation") |
        stage5_df.iloc[:, 0].str.startswith("Total Youth Hostel") |
        stage5_df.iloc[:, 0].str.startswith("Total Camping")
    ]

    if total_rows.empty:
        print(f"Warning: No total rows found in the Stage 5 file {stage5_file}.")
        return

    # Load the Stage 4 file using openpyxl
    workbook = load_workbook(stage4_file)
    sheet = workbook.active

    # Iterate through the total rows from Stage 5
    for idx, row in total_rows.iterrows():
        # Extract the keyword from the row (e.g., "Total Accommodation")
        keyword = next(
            (k for k in ["Total Accommodation", "Total Youth Hostel", "Total Camping"]
            if str(row.iloc[0]).startswith(k)),
            None
        )

        if not keyword:
            continue  # Skip if no matching keyword is found

        # Find the corresponding row in Stage 4
        target_row_index = None
        for row_idx, sheet_row in enumerate(sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=1), start=1):
            if sheet_row[0].value and str(sheet_row[0].value).startswith(keyword):
                target_row_index = row_idx
                break

        if target_row_index is None:
            print(f"Warning: No matching row found in Stage 4 for '{keyword}'.")
            continue

        # Insert the row from Stage 5 below the target row in Stage 4
        sheet.insert_rows(target_row_index + 1)  # Insert a new row below the target row
        for col_idx, value in enumerate(row, start=1):
            sheet.cell(row=target_row_index + 1, column=col_idx, value=value)

    # Save the modified workbook
    workbook.save(output_file)
    print(f"Copied total rows from {stage5_file} to Stage 4 file. Saved to {output_file}.")

def calculate_days_difference(date1, date2):
    """
    Calculates the absolute difference in days between two dates, ignoring the year.
    """
    # Extract month and day from the dates
    date1_md = (date1.month, date1.day)
    date2_md = (date2.month, date2.day)

    # Create datetime objects for the same year (year is irrelevant)
    current_year = datetime.now().year
    date1_fixed = datetime(current_year, date1_md[0], date1_md[1])
    date2_fixed = datetime(current_year, date2_md[0], date2_md[1])

    # Calculate the absolute difference in days
    return abs((date1_fixed - date2_fixed).days)

def extract_month_day(date):
    """
    Extracts the month and day from a date and returns a tuple (month, day).
    """
    return (date.month, date.day)

def process_stage6(input_file, stage5_files, output_file):
    # Load the Stage 4 file
    stage4_df = load_stage4(input_file)
    stage4_start, stage4_end = detect_date_range(input_file, stage4_df)

    if stage4_start is None or stage4_end is None:
        print(f"Warning: Could not determine date range for Stage 4 file {input_file}")

    # Load the Stage 5 files
    stage5_data = load_stage5_files(stage5_files)
    all_start_dates = []
    all_end_dates = []

    for file_name, df in stage5_data:
        start_date, end_date = detect_date_range(file_name, df)
        if start_date:
            all_start_dates.append(start_date)
        if end_date:
            all_end_dates.append(end_date)

    if all_start_dates and all_end_dates and stage4_start and stage4_end:
        # Extract month and day from all starting and ending dates
        all_start_md = [extract_month_day(date) for date in all_start_dates]
        all_end_md = [extract_month_day(date) for date in all_end_dates]

        # Find the earliest starting date and latest ending date (ignoring the year)
        earliest_start_md = min(all_start_md)
        latest_end_md = max(all_end_md)

        # Convert back to datetime objects for the current year (year is irrelevant)
        current_year = datetime.now().year
        earliest_start = datetime(current_year, earliest_start_md[0], earliest_start_md[1])
        latest_end = datetime(current_year, latest_end_md[0], latest_end_md[1])

        # Calculate the days difference (ignoring the year)
        start_diff = calculate_days_difference(stage4_start, earliest_start)
        end_diff = calculate_days_difference(stage4_end, latest_end)

        # Print the results
        print(
            f"Stage4 Starting Date: {stage4_start.strftime('%m-%d')} - Earliest Starting Date found in Stage5 files: {earliest_start.strftime('%m-%d')} - Days difference: {start_diff}")
        print(
            f"Stage4 Ending Date: {stage4_end.strftime('%m-%d')} - Latest Ending Date found in Stage5 files: {latest_end.strftime('%m-%d')} - Days difference: {end_diff}")

        # Add empty columns to Stage 4
        add_empty_columns(input_file, output_file, start_diff)

        # Iterate through all Stage 5 files
        for stage5_file in stage5_files:
            # Copy header from Stage 5 to Stage 4
            copy_header(stage5_file, output_file, output_file)

            # Copy total rows from Stage 5 to Stage 4
            copy_total_rows_from_stage5(stage5_file, output_file, output_file)
    else:
        print("Error: Could not determine valid date ranges for comparison.")

    # Placeholder for further processing
    print("Loaded Stage 4 and Stage 5 files successfully")


if __name__ == "__main__":
    INPUT_FILE = "per_zone_stage4_output.xlsx"
    STAGE5_FILES = ["per_zone_stage5_output_2024.xlsx", "per_zone_stage5_output_2023.xlsx"]
    OUTPUT_FILE = "per_zone_stage6_output.xlsx"
    process_stage6(INPUT_FILE, STAGE5_FILES, OUTPUT_FILE)
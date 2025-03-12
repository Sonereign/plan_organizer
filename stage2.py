import pandas as pd

# Constants
TARGET_CATEGORY = ".Beach Apt / for 5.2.6"
FILTER_KEYWORDS = ["APT", "Beach", "for2", "for5", "for6"]


def load_filtered_data(file, keywords):
    """Load Excel file and filter rows based on category keywords."""
    df = pd.read_excel(file, sheet_name=None)
    sheet_name = list(df.keys())[0]  # Get first sheet
    df = df[sheet_name].copy()
    df.iloc[:, 0] = df.iloc[:, 0].astype(str)  # Ensure Category column is a string
    df_filtered = df[df.iloc[:, 0].apply(lambda x: any(kw in x for kw in keywords))].copy()

    # Format date columns
    date_cols = [col for col in df_filtered.columns if is_date(col)]
    if date_cols:
        df_filtered.rename(columns={col: pd.to_datetime(col, dayfirst=True).strftime("%a %d/%m") for col in date_cols},
                           inplace=True)

    return df_filtered


def is_date(column_name):
    """Check if a column name is a date."""
    try:
        pd.to_datetime(column_name, dayfirst=True)
        return True
    except (ValueError, TypeError):
        return False


def replace_category_row(df_zone, df_type, target_category):
    """Replace the row containing the target category with the new filtered rows, maintaining alignment."""
    mask = df_zone.iloc[:, 0] == target_category
    if mask.any():
        index = mask.idxmax()  # Get first occurrence index
        df_type = df_type.reindex(columns=df_zone.columns, fill_value="")  # Align columns
        df_zone.drop(index, inplace=True)  # Remove the target category row
        df_zone = pd.concat([df_zone.iloc[:index], df_type, df_zone.iloc[index:]], ignore_index=True)
    return df_zone


def stage2(zone_file, type_file, output_file):
    """
    Process the input files (output of stage1 and availabilityPerType) and save the result to the output file.
    """
    df_zone = pd.read_excel(zone_file)
    df_type = load_filtered_data(type_file, FILTER_KEYWORDS)
    df_updated = replace_category_row(df_zone, df_type, TARGET_CATEGORY)
    df_updated.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Stage 2 completed. File saved as {output_file}")


if __name__ == "__main__":
    # Default file paths (for standalone execution)
    ZONE_FILE = "stage1_output.xlsx"
    TYPE_FILE = "availabilityPerType.xls"
    OUTPUT_FILE = "stage2_output.xlsx"

    # Run stage2
    stage2(ZONE_FILE, TYPE_FILE, OUTPUT_FILE)
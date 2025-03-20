import pandas as pd
from logger import logger

# Constants
TARGET_CATEGORY = ".Beach Apt / for 5.2.6"
FILTER_KEYWORDS = ["APT", "Beach", "for2", "for5", "for6"]


def load_filtered_data(file, keywords):
    """Load Excel file and filter rows based on category keywords."""
    try:
        logger.info(f"Loading Excel file: {file}")
        df = pd.read_excel(file, sheet_name=None)
        sheet_name = list(df.keys())[0]  # Get first sheet
        logger.debug(f"Loaded sheet: {sheet_name}")

        df = df[sheet_name].copy()
        df.iloc[:, 0] = df.iloc[:, 0].astype(str)  # Ensure Category column is a string

        df_filtered = df[df.iloc[:, 0].apply(lambda x: any(kw in x for kw in keywords))].copy()
        logger.info(f"Filtered rows based on keywords: {keywords}")
        logger.debug(f"Filtered DataFrame shape: {df_filtered.shape}")

        # Format date columns
        date_cols = [col for col in df_filtered.columns if is_date(col)]
        if date_cols:
            logger.info(f"Formatting date columns: {date_cols}")
            df_filtered.rename(
                columns={col: pd.to_datetime(col, dayfirst=True).strftime("%a %d/%m/%Y") for col in date_cols},
                inplace=True)

        return df_filtered
    except Exception as e:
        logger.error(f"Error loading and filtering data from {file}: {e}", exc_info=True)
        return pd.DataFrame()  # Return an empty DataFrame in case of failure


def is_date(column_name):
    """Check if a column name is a date."""
    try:
        pd.to_datetime(column_name, dayfirst=True)
        return True
    except (ValueError, TypeError):
        return False


def replace_category_row(df_zone, df_type, target_category):
    """Replace the row containing the target category with the new filtered rows, maintaining alignment."""
    try:
        logger.info(f"Replacing category row: {target_category}")
        mask = df_zone.iloc[:, 0] == target_category
        if mask.any():
            index = mask.idxmax()  # Get first occurrence index
            logger.debug(f"Found target category at index: {index}")

            df_type = df_type.reindex(columns=df_zone.columns, fill_value="")  # Align columns
            df_zone.drop(index, inplace=True)  # Remove the target category row
            df_zone = pd.concat([df_zone.iloc[:index], df_type, df_zone.iloc[index:]], ignore_index=True)

        return df_zone
    except Exception as e:
        logger.error(f"Error replacing category row {target_category}: {e}", exc_info=True)
        return df_zone  # Return the original DataFrame in case of failure


def per_zone_per_type_stage2(zone_file, type_file, output_file):
    """
    Process the input files (output of stage1 and availabilityPerType) and save the result to the output file.
    """
    logger.info("#######################################################")
    logger.info(f"Running Stage 2 with {zone_file=} - {type_file=} ....")

    try:
        logger.info(f'Loading stage1_output (perZone): {zone_file}')
        df_zone = pd.read_excel(zone_file)
        logger.debug(f"Loaded DataFrame shape (perZone): {df_zone.shape}")

        logger.info(f'Loading perType file: {type_file}')
        df_type = load_filtered_data(type_file, FILTER_KEYWORDS)
        logger.debug(f"Loaded DataFrame shape (perType): {df_type.shape}")

        logger.info(f'Finding {TARGET_CATEGORY} in perType and replacing with {FILTER_KEYWORDS} in perZone.')
        df_updated = replace_category_row(df_zone, df_type, TARGET_CATEGORY)

        logger.info(f'Saving output file: {output_file}')
        df_updated.to_excel(output_file, index=False, engine='openpyxl')
        logger.info(f"Stage 2 completed successfully. File saved as {output_file}")
    except Exception as e:
        logger.error(f"Error during Stage 2 processing: {e}", exc_info=True)


if __name__ == "__main__":
    # Default file paths (for standalone execution)
    ZONE_FILE = "zone_stage1_output.xlsx"
    TYPE_FILE = "sources/availabilityPerType2025.xls"
    OUTPUT_FILE = "zone_stage2_output.xlsx"

    # Run stage2
    per_zone_per_type_stage2(ZONE_FILE, TYPE_FILE, OUTPUT_FILE)
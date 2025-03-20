import pandas as pd
from logger import logger

# Hardcoded capacities for accommodations
ACCOMMODATION_CAPACITIES = {
    "APT": 2,
    "Beach": 3,
    "for2": 7,
    "for5": 14,
    "for6": 5,
    ".LUX for 4": 51,
    ".Safari Tent 5pax": 12,
    ".Sea Safari 4pax": 31,
    ".Skyline 3pax": 11,
    ".Standard Mobile Home": 31,
    ".ΤΡΟΧΟΣΠΙΤΑ DELUXE": 8,
    ".ΤΡΟΧΟΣΠΙΤΑ SEA VIEW": 24,
    ".ΤΡΟΧΟΣΠΙΤΑ standard": 8
}

# Hardcoded capacities for camping areas
CAMPING_CAPACITIES = {
    "area 1": 0,
    "area 2": 20,
    "area 3": 81,
    "area 4": 23,
    "area 5": 44,
    "area 6": 11,
    "area 7": 32,
    "area Z": 27,  # English Z
    "area K": 80,  # English K
    "area Δ": 12,
    "area Ε": 14,
    "area Ι": 4
}

# Mapping of Greek letters to English equivalents
GREEK_TO_ENGLISH = {
    "Ζ": "Z",  # Greek Zeta to English Z
    "Κ": "K",  # Greek Kappa to English K
}


def normalize_letters(name):
    """Normalize Greek letters to English equivalents."""
    try:
        if pd.isna(name):  # Handle NaN values
            logger.debug("Skipping NaN value in normalize_letters")
            return None

        name = str(name).strip()  # Convert to string and remove leading/trailing spaces
        logger.debug(f"Normalizing letters for: {name}")

        for greek, english in GREEK_TO_ENGLISH.items():
            name = name.replace(greek, english)

        return name
    except Exception as e:
        logger.error(f"Error in normalize_letters with input {name}: {e}", exc_info=True)
        return name  # Return the original name in case of failure


def normalize_camping_area_name(name):
    """Normalize camping area names by adding 'area' prefix if missing."""
    try:
        if pd.isna(name):  # Handle NaN values
            logger.debug("Skipping NaN value in normalize_camping_area_name")
            return None

        name = str(name).strip()  # Convert to string and remove leading/trailing spaces
        logger.debug(f"Normalizing camping area name for: {name}")

        if not name.startswith("area "):
            return f"area {name}"

        return name
    except Exception as e:
        logger.error(f"Error in normalize_camping_area_name with input {name}: {e}", exc_info=True)
        return name  # Return the original name in case of failure


def update_capacity_column(df, capacities, category_type):
    """Update the capacity column with hardcoded values."""
    try:
        logger.info(f"Updating capacity column for {category_type}")
        updated_categories = set()  # Track which categories/areas were updated

        for index, row in df.iterrows():
            category = row[0]  # First column contains the category/area names

            if category_type == "camping areas":
                category = normalize_letters(category)
                category = normalize_camping_area_name(category)
                if category is None:  # Skip NaN values
                    continue

            if category in capacities:
                df.at[index, 1] = capacities[category]  # Update the capacity column (column index 1)
                updated_categories.add(category)
                logger.debug(f"Updated capacity for {category} to {capacities[category]}")

        # Check for skipped categories/areas
        skipped_categories = set(capacities.keys()) - updated_categories
        if skipped_categories:
            logger.info(
                f"The following {category_type} were not found in the Excel file and were skipped: {', '.join(skipped_categories)}"
            )

        return df
    except Exception as e:
        logger.error(f"Error updating capacity column for {category_type}: {e}", exc_info=True)
        return df  # Return original DataFrame in case of failure


def per_zone_per_type_stage3(input_file, output_file):
    """
    Process the input file (output of stage2) and save the result to the output file.
    """
    logger.info("#######################################################")
    logger.info(f"Running Stage 3 with {input_file=} ....")

    try:
        logger.info(f'Loading {input_file} ..')
        df = pd.read_excel(input_file, sheet_name='Sheet1', header=None)
        logger.debug(f"Loaded DataFrame shape: {df.shape}")

        logger.info(f'Updating accommodation capacities based on: {ACCOMMODATION_CAPACITIES}')
        df = update_capacity_column(df, ACCOMMODATION_CAPACITIES, "accommodations")

        logger.info(f'Updating camping capacities based on: {CAMPING_CAPACITIES}')
        df = update_capacity_column(df, CAMPING_CAPACITIES, "camping areas")

        df.to_excel(output_file, index=False, header=False)
        logger.info(f"Stage 3 completed. File saved as {output_file}")
    except Exception as e:
        logger.error(f"Error during Stage 3 processing: {e}", exc_info=True)


if __name__ == "__main__":
    # Default file paths (for standalone execution)
    INPUT_FILE = "zone_stage2_output.xlsx"
    OUTPUT_FILE = "zone_stage3_output.xlsx"

    # Run stage3
    per_zone_per_type_stage3(INPUT_FILE, OUTPUT_FILE)

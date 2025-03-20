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
    if pd.isna(name):  # Handle NaN values
        return None
    name = str(name).lstrip().rstrip()  # Convert to string and remove leading/trailing spaces
    # Replace Greek letters with their English equivalents
    for greek, english in GREEK_TO_ENGLISH.items():
        name = name.replace(greek, english)
    return name


def normalize_camping_area_name(name):
    """Normalize camping area names by adding 'area' prefix if missing."""
    if pd.isna(name):  # Handle NaN values
        return None
    name = str(name).lstrip().rstrip()  # Convert to string and remove leading/trailing spaces
    if not name.startswith("area "):
        return f"area {name}"
    return name


def update_capacity_column(df, capacities, category_type):
    """Update the capacity column with hardcoded values."""
    updated_categories = set()  # Track which categories/areas were updated
    for index, row in df.iterrows():
        category = row[0]  # First column contains the category/area names
        # Normalize camping area names if the category type is "camping areas"
        if category_type == "camping areas":
            # Normalize letters first (e.g., Greek Ζ to English Z)
            category = normalize_letters(category)
            # Normalize camping area names (add "area" prefix if missing)
            category = normalize_camping_area_name(category)
            if category is None:  # Skip NaN values
                continue

        # Check if the category matches any key in the capacities dictionary
        if category in capacities:
            df.at[index, 1] = capacities[category]  # Update the capacity column (column index 1)
            updated_categories.add(category)

    # Check for skipped categories/areas
    skipped_categories = set(capacities.keys()) - updated_categories
    if skipped_categories:
        logger.info(
            f"The following {category_type} were not found in the Excel file and were skipped: {', '.join(skipped_categories)}")

    return df


def per_zone_stage3(input_file, output_file):
    """
    Process the input file (output of stage2) and save the result to the output file.
    """
    logger.info("#######################################################")
    logger.info(f"Running Per Zone Stage 3 with {input_file=} ....")
    # Load the Excel file
    df = pd.read_excel(input_file, sheet_name='Sheet1', header=None)

    # Update the capacity column for accommodations
    df = update_capacity_column(df, ACCOMMODATION_CAPACITIES, "accommodations")

    # Update the capacity column for camping areas
    df = update_capacity_column(df, CAMPING_CAPACITIES, "camping areas")

    # Save the updated DataFrame to a new Excel file
    df.to_excel(output_file, index=False, header=False)
    logger.info(f"Per_zone Stage 3 completed. File saved as {output_file}")


if __name__ == "__main__":
    # Default file paths (for standalone execution)
    INPUT_FILE = "per_zone_stage2_output.xlsx"
    OUTPUT_FILE = "per_zone_stage3_output.xlsx"

    # Run stage3
    per_zone_stage3(INPUT_FILE, OUTPUT_FILE)

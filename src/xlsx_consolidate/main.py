import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytz
from rich.pretty import pprint


def consolidate_files(*, spreadsheet_dir: str):
    """Consolidate all XLSX files in the dir and write new XLSX file."""
    dir_path = Path(spreadsheet_dir)
    xlsx_files = dir_path.glob(pattern="*.xlsx")
    consolidated_df = pd.DataFrame()

    for one_file in xlsx_files:
        in_df = pd.read_excel(one_file)
        consolidated_df = pd.concat([consolidated_df, in_df], ignore_index=True)

    pprint(consolidated_df)

    # Save the consolidated DataFrame to a new XLSX file
    home_dir = Path.home()
    # Assuming us central
    timezone_name = "US/Central"
    # Create a tzinfo object from the timezone name
    tz = pytz.timezone(timezone_name)
    dt_str = datetime.now(tz=tz).strftime("%Y%m%d_%H%M%S")
    consolidated_file_path = home_dir / f"Documents/xlsx_consolidated_{dt_str}.xlsx"
    with pd.ExcelWriter(
        consolidated_file_path,
        date_format="YYYY-MM-DD",
        datetime_format="YYYY-MM-DD HH:MM:SS",
    ) as writer:
        consolidated_df.to_excel(writer, index=False)
    pprint(f"Consolidated XLSX file written to {consolidated_file_path}")


def main():
    """Main function to run the consolidation."""
    parser = argparse.ArgumentParser(
        description="Consolidate all XLSX files in the given directory. Does not search subdirectories or deduplicate data.",
    )

    # Add an argument for the input directory
    parser.add_argument(
        "-d",
        "--directory",
        type=str,  # The type of the argument (string for file path)
        required=True,  # Make this argument mandatory
        help="Path to the XLSX directory.",
    )

    args = parser.parse_args()

    input_dir_path = args.directory

    if not Path(input_dir_path).is_dir():
        pprint(f"Error: The path '{input_dir_path}' is not a valid directory.")
        sys.exit(1)

    consolidate_files(spreadsheet_dir=input_dir_path)


if __name__ == "__main__":
    main()

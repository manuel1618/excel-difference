import os

import pandas as pd

from excel_difference.excel_diff import excel_diff


def test_real_excel_files():
    """Test the excel diff with the actual data files."""
    # Test with the actual data files
    excel_diff("data/excel1.xlsx", "data/excel2.xlsx", "test_real_output.xlsx", 1, 1)

    # Verify the output file was created
    assert os.path.exists("test_real_output.xlsx"), "Output file was not created"

    # Read the output file to verify it has content
    try:
        # Try to read all sheets
        all_sheets = pd.read_excel("test_real_output.xlsx", sheet_name=None)
        assert len(all_sheets) > 0, "Output file has no sheets"

        # Print some info about the output
        print(f"Output file created with {len(all_sheets)} sheets:")
        for sheet_name, df in all_sheets.items():
            print(f"  Sheet '{sheet_name}': {df.shape[0]} rows, {df.shape[1]} columns")

    except Exception as e:
        print(f"Error reading output file: {e}")
        raise

    # Clean up
    os.remove("test_real_output.xlsx")
    print("Test completed successfully!")


if __name__ == "__main__":
    test_real_excel_files()

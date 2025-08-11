import os

import pandas as pd

from excel_difference.excel_diff import excel_diff


def test_auto_detect():
    """Test auto-detection of key row and column."""
    # Create test data with clear header row and ID column
    data1 = {
        "Sheet1": pd.DataFrame(
            {
                "ID": ["A", "B", "C"],
                "Name": ["Alice", "Bob", "Charlie"],
                "Age": [25, 30, 35],
                "City": ["NYC", "LA", "Chicago"],
            }
        )
    }
    data2 = {
        "Sheet1": pd.DataFrame(
            {
                "ID": ["A", "B", "C"],
                "Name": ["Alice", "Bob", "Charlie"],
                "Age": [25, 31, 35],  # Changed Bob's age
                "City": ["NYC", "LA", "Chicago"],
            }
        )
    }

    with pd.ExcelWriter("test_auto1.xlsx") as writer:
        for sheet_name, df in data1.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    with pd.ExcelWriter("test_auto2.xlsx") as writer:
        for sheet_name, df in data2.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Test with auto-detection (key_row=0, key_column=0)
    excel_diff("test_auto1.xlsx", "test_auto2.xlsx", "test_auto_output.xlsx", 0, 0)

    # Verify the output file was created
    assert os.path.exists("test_auto_output.xlsx"), "Output file was not created"

    # Read the output file to verify it has content
    try:
        all_sheets = pd.read_excel("test_auto_output.xlsx", sheet_name=None)
        assert len(all_sheets) > 0, "Output file has no sheets"

        # Check that the data was processed correctly
        df_output = pd.read_excel(
            "test_auto_output.xlsx", sheet_name="Sheet1", skiprows=13
        )
        assert len(df_output) > 0, "No data rows found"

        print("Auto-detection test completed successfully!")
        print(f"Output file created with {len(all_sheets)} sheets")

    except Exception as e:
        print(f"Error reading output file: {e}")
        raise

    # Clean up
    os.remove("test_auto1.xlsx")
    os.remove("test_auto2.xlsx")
    os.remove("test_auto_output.xlsx")


if __name__ == "__main__":
    test_auto_detect()

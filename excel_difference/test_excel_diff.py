import os

import pandas as pd

from excel_difference.excel_diff import excel_diff


def test_excel_diff():
    # Create dummy excel files for testing
    data1 = {
        "Sheet1": pd.DataFrame(
            {
                "A": [1, 2, 3],
                "B": ["apple", "banana", "cherry"],
                "C": [10.1, 20.2, 30.3],
            }
        )
    }
    data2 = {
        "Sheet1": pd.DataFrame(
            {
                "A": [1, 5, 3],
                "B": ["apple", "orange", "cherry"],
                "C": [10.1, 25.2, 30.3],
            }
        )
    }

    with pd.ExcelWriter("test_excel1.xlsx") as writer:
        for sheet_name, df in data1.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    with pd.ExcelWriter("test_excel2.xlsx") as writer:
        for sheet_name, df in data2.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    excel_diff("test_excel1.xlsx", "test_excel2.xlsx", "test_output.xlsx", 1, 1)

    # Read the output file, skipping the mapping information
    # The data starts at row 13 (index 12), with header at row 13 and data at row 14+
    df_output = pd.read_excel("test_output.xlsx", sheet_name="Sheet1", skiprows=13)

    # Now test the actual data (only 2 rows matched due to missing row 3)
    assert df_output.iloc[0, 0] == 0  # First column, first data row (1-1=0)
    assert df_output.iloc[1, 0] == 0  # First column, second data row (3-3=0)

    assert df_output.iloc[0, 1] == "apple"  # Second column, first data row (same)
    assert df_output.iloc[1, 1] == "cherry"  # Second column, second data row (same)

    assert df_output.iloc[0, 2] == 0.0  # Third column, first data row (10.1-10.1=0)
    assert df_output.iloc[1, 2] == 0.0  # Third column, second data row (30.3-30.3=0)

    os.remove("test_excel1.xlsx")
    os.remove("test_excel2.xlsx")
    os.remove("test_output.xlsx")

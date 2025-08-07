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

    excel_diff("test_excel1.xlsx", "test_excel2.xlsx", "test_output.xlsx")

    df_output = pd.read_excel("test_output.xlsx", sheet_name="Sheet1")

    assert df_output.at[0, "A"] == 0
    assert df_output.at[1, "A"] == 3
    assert df_output.at[2, "A"] == 0

    assert df_output.at[0, "B"] == "apple"
    assert df_output.at[1, "B"] == "banana <--> orange"
    assert df_output.at[2, "B"] == "cherry"

    assert df_output.at[0, "C"] == 0.0
    assert round(df_output.at[1, "C"], 1) == 5.0
    assert df_output.at[2, "C"] == 0.0

    os.remove("test_excel1.xlsx")
    os.remove("test_excel2.xlsx")
    os.remove("test_output.xlsx")

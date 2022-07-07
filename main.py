# This app will be a reference for Excel sheet manipulation
# Generally the user wants to repair an Excel file that often has the same errors.
# Step 0 The Excel sheet is opened and converted to a pandas dataframe.
# Step 1 Each row is evaluated. Users can add logic here
# Step 2 Convert the dataframe object to an Excel file and save with timestamp in name.

# "os" library is for manipulating files on your computer
import os

# "pandas" is for manipulating data, in this case Excel files
import pandas

# this makes it so the saved excel files don't have weird formatting
pandas.io.formats.excel.ExcelFormatter.header_style = None


def main(name):
    print(f"Hi, {name}")
    file1_src = "C:/dev/excel_delta/venv/ExampleData/StudentFood.xlsx"
    file1_dst = "C:/dev/excel_delta/venv/ExampleData/StudentFood_changed.xlsx"
    file1_sheet1 = "Sheet1"

    # creates a pandas "dataframe" object that's basically the Excel file sheet
    df = pandas.read_excel(io=file1_src, sheet_name=file1_sheet1)

    # prints to console the Excel file path
    print(f"Excel from {file1_dst}")
    print(df)

    df.reset_index()

    # goes through the Excel file one row at time
    for row in df.index:
        # corrects student name
        if df.at[row, "Student Name"] == "Li Ling":
            df.at[row, "Student Name"] = "Li Ping"

        # corrects food column misspellings
        if df.at[row, "Food"] == "Frys":
            df.at[row, "Food"] = "Fries"
        elif df.at[row, "Food"] == "Cola":
            temp_x = df.at[row, "Food"]
            df.at[row, "Food"] = df.at[row, "Drink"]
            df.at[row, "Drink"] = temp_x

        # corrects drink errors
        if df.at[row, "Drink"] == "H20":
            df.at[row, "Drink"] = "Water"

        # corrects age error
        if df.at[row, "Age"] > 20:
            df.at[row, "Age"] = df.at[row, "Age"] / 10

    print(f"\n")
    print(f"Saved to {file1_dst}")
    print(df)

    df.to_excel(file1_dst, index=False, sheet_name="sch_data_1")


# if the script is run this runs
if __name__ == "__main__":
    main("Portland Public Schools")

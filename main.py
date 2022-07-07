# This app will be a reference for Excel sheet manipulation
# Generally the user wants to repair an Excel file that often has the same errors.
# Step 0 The Excel sheet is opened and converted to a pandas dataframe.
# Step 1 Each row is evaluated. Users can add logic here
# Step 2 Convert the dataframe object to an Excel file and save with timestamp in name.
import os
import pandas as pd

pd.io.formats.excel.ExcelFormatter.header_style = None


def main(name):
    print(f"Hi, {name}")
    dir_src = os.curdir
    file1_src = "C:/dev/excel_delta/venv/ExampleData/StudentFood.xlsx"
    file1_dst = "C:/dev/excel_delta/venv/ExampleData/StudentFood_changed.xlsx"
    file1_sheet1 = "Sheet1"
    df = pd.read_excel(io=file1_src, sheet_name=file1_sheet1)
    # print(f"Type: {type(df)}")
    print(f"Excel from {file1_dst}")
    print(df)

    df.reset_index()

    for index in df.index:
        if df.at[index, "Student Name"] == "Li Ling":
            df.at[index, "Student Name"] = "Li Ping"

        if df.at[index, "Food"] == "Frys":
            df.at[index, "Food"] = "Fries"
        elif df.at[index, "Food"] == "Cola":
            temp_x = df.at[index, "Food"]
            df.at[index, "Food"] = df.at[index, "Drink"]
            df.at[index, "Drink"] = temp_x

        if df.at[index, "Drink"] == "H20":
            df.at[index, "Drink"] = "Water"

        if df.at[index, "Age"] > 20:
            df.at[index, "Age"] = df.at[index, "Age"] / 10

    print(f"\n")
    print(f"Saved to {file1_dst}")
    print(df)

    df.to_excel(file1_dst, index=False, sheet_name="sch_data_1")


# Press the green button in the gutter to run the script.
if __name__ == "__main__":
    main("Portland Public Schools")

# See PyCharm help at https://www.jetbrains.com/help/pycharm/

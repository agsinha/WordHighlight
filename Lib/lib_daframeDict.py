import pandas as pd

def excel_to_df(file_path, sheet_name):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    for index, row in df.iterrows():
        for column in df.columns:
            cell_value = row[column]
            print(f"Row {index}, Column {column}: {cell_value}")

    return df
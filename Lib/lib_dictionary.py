import pandas as pd

def excel_to_dict(file_path, sheet_name):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Convert the DataFrame to a dictionary
    data_dict = df.to_dict(orient='dict')
    for word, color in data_dict.items():
        print(f"{word}, {color}")
    return data_dict
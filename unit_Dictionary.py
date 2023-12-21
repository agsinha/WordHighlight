from Lib.lib_dictionary import excel_to_dict
from Lib.lib_daframeDict import  excel_to_df
# Example usage
file_path = 'data/dict.xlsx'
sheet_name = 'Sheet1'  # Replace with the actual sheet name in your Excel file

# Import data into a dictionary
# excel_data_dict = excel_to_dict(file_path, sheet_name)
excel_data_dict = excel_to_df(file_path, sheet_name)
# Display the dictionary
# print(excel_data_dict)
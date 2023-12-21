# begin
import pandas as pd
from docx import Document


def sow_review():
    file_path = 'data/dict.xlsx'
    sheet_name = 'Sheet1'  # Replace with the actual sheet name in your Excel file
    # Import data into a dictionary
    xl_data = excel_to_df(file_path, sheet_name)

    # Load the Word document
    doc_path = 'data/sow.docx'
    doc = Document(doc_path)

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for idx, word, color, comment in xl_data.itertuples():
                if word in run.text:
                    # Highlight the word with the specified color
                    run.font.highlight_color = color
                    # Add a comment to the paragraph
                    comment_text = f"Comment: {comment}"
                    paragraph.add_comment(comment_text, author='Deals Review',initials= 'ag')
    # Save the modified document
    doc.save('data/sow_h.docx')

def excel_to_df(file_path, sheet_name):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    for idx, word, color, policy in df.itertuples():
        print(f"{word}, Column {color}, {policy}")

    return df

def excel_to_dict(file_path, sheet_name):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Convert the DataFrame to a dictionary
    data_dict = df.to_dict(orient='dict')

    return data_dict


# begin
import pandas as pd
from docx import Document
from docx.shared import RGBColor


def highlight_and_comment():
    file_path = 'data/dict.xlsx'
    sheet_name = 'Sheet1'  # Replace with the actual sheet name in your Excel file
    # Import data into a dictionary
    word_dict = excel_to_dict(file_path, sheet_name)

    # Load the Word document
    doc_path = 'data/sow.docx'
    doc = Document(doc_path)

    for paragraph in doc.paragraphs:
        for word, color in word_dict.items():
            if word in paragraph.text:
                # Highlight the word with the specified color
                for run in paragraph.runs:
                    if word in run.text:
                        run.font.highlight_color = color
                # Add a comment to the paragraph
                comment_text = f"Highlighted word: {word_dict.values}"
                paragraph.add_comment(comment_text, author='Deals Review',initials= 'ag')
# Save the modified document
    doc.save('data/sow_h.docx')

def excel_to_dict(file_path, sheet_name):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Convert the DataFrame to a dictionary
    data_dict = df.to_dict(orient='dict')

    return data_dict


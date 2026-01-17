import pandas as pd
from docx import Document

def word_table_to_excel(docx_path, excel_output_path, table_index=0):
    # Load the Word document
    doc = Document(docx_path)
    
    # Access the specific table
    table = doc.tables[table_index]
    
    # Extract data from rows and cells
    data = []
    for row in table.rows:
        data.append([cell.text.strip() for cell in row.cells])
    
    # Convert to DataFrame (using the first row as headers)
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # Output to Excel
    df.to_excel(excel_output_path, index=False)
    print(f"Table {table_index} successfully exported to {excel_output_path}")

# Usage
word_table_to_excel("WordTest.docx", "output_file.xlsx")

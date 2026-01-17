import io
import pandas as pd
from flask import Flask, render_template_string, request, send_file
from docx import Document

app = Flask(__name__)

# Basic HTML template for the upload form
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head><title>Word Table to Excel</title></head>
<body>
    <h2>Upload Word Document (.docx)</h2>
    <form method="post" action="/convert" enctype="multipart/form-data">
        <input type="file" name="file" accept=".docx" required>
        <button type="submit">Convert to Excel</button>
    </form>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file uploaded", 400
    
    file = request.files['file']
    
    # 1. Read the Word file from the uploaded stream
    doc = Document(file)
    
    if not doc.tables:
        return "No tables found in this document.", 400

    # 2. Extract data (using your logic)
    table = doc.tables[0] # Processes the first table
    data = []
    for row in table.rows:
        data.append([cell.text.strip() for cell in row.cells])
    
    df = pd.DataFrame(data[1:], columns=data[0])

    # 3. Save to an in-memory buffer instead of a disk file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0) # Move to the beginning of the buffer

    # 4. Return the buffer as a downloadable file
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='converted_table.xlsx'
    )

if __name__ == '__main__':
    app.run(debug=True)

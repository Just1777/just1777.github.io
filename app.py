from flask import Flask, request, send_file
import os
import pandas as pd

app = Flask(__name__)

UPLOAD_FOLDER = './uploads'
PROCESSED_FOLDER = './processed'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    
    if file.filename == '':
        return 'No selected file', 400
    
    if file and file.filename.endswith(('.xls', '.xlsx')):
        # Save the uploaded file
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        
        # Process the Excel file
        processed_file_path = process_excel(file_path, file.filename)
        
        # Return the processed file to the user
        return send_file(processed_file_path, as_attachment=True)

    return 'Invalid file type', 400

def process_excel(file_path, filename):
    # Use pandas to read the Excel file
    df = pd.read_excel(file_path)

    # Example modification: Add a new column
    df['Processed'] = 'Yes'

    # Save the modified file
    processed_file_path = os.path.join(PROCESSED_FOLDER, 'modified_' + filename)
    df.to_excel(processed_file_path, index=False)

    return processed_file_path

if __name__ == '__main__':
    app.run(debug=True)
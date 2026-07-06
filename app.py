from flask import Flask, render_template, request, send_file
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font  # <-- NEW: Added 'Font' for bolding
import zipfile
import os
import tempfile

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/merge', methods=['POST'])
def merge():
    # Get the uploaded files
    csv_file = request.files.get('csv_file')
    zip_file = request.files.get('zip_file')

    if not csv_file or not zip_file:
        return "Please upload both files.", 400

    # Create a temporary directory to handle the processing safely
    temp_dir = tempfile.mkdtemp()
    
    # Save and extract the zip file
    zip_path = os.path.join(temp_dir, "photos.zip")
    zip_file.save(zip_path)
    
    extract_folder = os.path.join(temp_dir, "extracted_photos")
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_folder)

    # Read the CSV (with fallback encoding for Excel-generated CSVs)
    try:
        df = pd.read_csv(csv_file, encoding='utf-8')
    except UnicodeDecodeError:
        csv_file.seek(0) # Reset the file reader back to the beginning
        df = pd.read_csv(csv_file, encoding='latin1')
    
    # Create new Excel workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Merged Data"

    # <-- NEW: Set up centering and bold fonts
    center_align = Alignment(horizontal='center', vertical='center')
    bold_font = Font(bold=True)

    # <-- NEW: Clean up the headers to ensure they have proper spaces
    clean_headers = []
    for col in df.columns:
        col_name = str(col).strip()
        
        # Force specific names to have spaces and correct capitalization
        if col_name.lower() in ["fathername", "father name"]:
            col_name = "Father Name"
        elif col_name.lower() in ["fathermobile", "father mobile"]:
            col_name = "Father Mobile"
        elif col_name.lower() in ["rollno", "roll no", "roll"]:
            col_name = "Roll No"
        elif col_name.lower() == "name":
            col_name = "Name"
        elif col_name.lower() == "class":
            col_name = "Class"
        elif col_name.lower() == "section":
            col_name = "Section"
            
        clean_headers.append(col_name)

    # Add "Photo" to the end of the cleaned headers
    headers = clean_headers + ["Photo"]
    ws.append(headers)

    # <-- NEW: Apply centering AND bold text to the header row (Row 1)
    for cell in ws[1]:
        cell.alignment = center_align
        cell.font = bold_font

    # Process each student
    for index, row in df.iterrows():
        ws.append(list(row))
        current_row = ws.max_row 
        
        # Apply centering to every cell in this specific student's row
        for cell in ws[current_row]:
            cell.alignment = center_align
        
        # 1. Get the raw roll number safely using .iloc[0]
        raw_roll = str(row.iloc[0]).strip() 
        
        # 2. Clean it! (This turns "1.0" or "01" into a perfect "1")
        try:
            search_roll = str(int(float(raw_roll)))
        except ValueError:
            search_roll = raw_roll 
        
        # Look for the image in the extracted folder
        img_path = None
        for root, dirs, files in os.walk(extract_folder):
            for file in files:
                file_name_only = os.path.splitext(file)[0].strip()
                
                try:
                    clean_file_name = str(int(float(file_name_only)))
                except ValueError:
                    clean_file_name = file_name_only
                
                if clean_file_name == search_roll and file.lower().endswith(('.png', '.jpg', '.jpeg')):
                    img_path = os.path.join(root, file)
                    break
            
            if img_path:
                break
        
        # If a match was found, insert it into Excel
        if img_path:
            img = Image(img_path)
            img.width, img.height = 80, 80
            cell_letter = openpyxl.utils.get_column_letter(len(headers))
            cell_name = f"{cell_letter}{current_row}"
            ws.add_image(img, cell_name)
            ws.row_dimensions[current_row].height = 65

    # Adjust the width of the final Photo column
    ws.column_dimensions[openpyxl.utils.get_column_letter(len(headers))].width = 15

    # Save the final file
    output_path = os.path.join(temp_dir, "Final_Student_Data.xlsx")
    wb.save(output_path)

    # Send the file to the user to download
    return send_file(output_path, as_attachment=True, download_name="Final_Student_Data.xlsx")

if __name__ == '__main__':
    app.run(debug=True)
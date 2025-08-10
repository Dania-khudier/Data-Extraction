import pandas as pd
from openpyxl import load_workbook
import os

def process_excel_file(filename='contacts.xlsx'):
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    df = pd.read_excel(filename, engine='openpyxl')
    print(df)
    
    all_sheets = pd.read_excel(filename, sheet_name=None, engine='openpyxl')
    all_texts = []
    all_data = []
    
    for sheet_name, df in all_sheets.items():
        print(f"محتوى الورقة: {sheet_name}")
        df_str = df.astype(str)
        sheet_data = df_str.values.flatten().tolist()
        all_data.extend(sheet_data)
        
        for item in sheet_data:
            print(item)
    
    output_dir = "extracted_images"
    os.makedirs(output_dir, exist_ok=True)
    wb = load_workbook(filename)
    
    img_count = 0
    for sheet in wb.worksheets:
        for image in sheet._images:
            img_count += 1
            ext = image.path.split('.')[-1] if image.path else "png"
            img_path = os.path.join(output_dir, f"image_{img_count}.{ext}")
            with open(img_path, "wb") as f:
                f.write(image._data())
            print(f"Extraction images from file {img_path}")
    
    print(f"Extraction images from file {img_count}")

process_excel_file()



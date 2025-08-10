import fitz 
import io
from PIL import Image
import os

def process_pdf_file(pdf_path="معلومات شركة COXIR الكورية.pdf"):
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    doc = fitz.open(pdf_path)
    full_text = ""
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        full_text += text + "\n"
    
    print(full_text)
    
    output_folder = "extracted_images/"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        image_list = page.get_images(full=True)
        
        for img_index, img in enumerate(image_list, start=1):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            
            image = Image.open(io.BytesIO(image_bytes))
            image_filename = f"{output_folder}page{page_num+1}_img{img_index}.{image_ext}"
            image.save(image_filename)
            print(f"تم حفظ الصورة: {image_filename}")

    with open("output.txt", "w", encoding="utf-8") as f:
        f.write(full_text)

process_pdf_file()



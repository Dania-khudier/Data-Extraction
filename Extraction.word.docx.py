from docx import Document
import zipfile
import os

def process_docx_file(docx_path="Product Requirements Document.docx"):
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    # Extract Texts
    doc = Document(docx_path)
    print("Extract Texts From The File")
    for para in doc.paragraphs:
        print(para.text)
    
    # Extract Images
    output_dir = "extracted_images"
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        media_files = [f for f in docx_zip.namelist() if f.startswith('word/media/')]
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        for file in media_files:
            docx_zip.extract(file, output_dir)
            print(f"images extracted to file {output_dir}")

# استدعاء الدالة
process_docx_file()



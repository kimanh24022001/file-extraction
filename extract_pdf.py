import fitz  # PyMuPDF
import os
from fpdf import FPDF
 

def extract_from_pdf(input_path, output_path):
    doc = fitz.open(input_path)
    pdf = fitz.open()  
    text_output = []

    if not os.path.exists(output_path):
        os.makedirs(output_path)
    images_dir =  os.path.join(output_path, 'images')
    if not os.path.exists(images_dir):
        os.makedirs(images_dir)
    
    text_output = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        text_output.append(text)
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            image_name = f"page{page_num+1}_img{img_index+1}.{image_ext}"
            with open(os.path.join(images_dir, image_name), "wb") as img_file:
                img_file.write(image_bytes)
    
    with open(os.path.join(output_path, 'text.txt'), 'w', encoding='utf-8') as f:
        f.writelines(text_output)
    pdf.close()
    doc.close()

def extract_pdf_text_details(pdf_path, output_file):
    output_file =  os.path.join(output_file, "detail.txt")
    doc = fitz.open(pdf_path)
    paragraphs = []

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        blocks = page.get_text("dict")["blocks"]
        
        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text_content = span["text"]
                        font_name = span["font"]
                        font_size = span["size"]
                        font_color = span.get("color", None)
                        
                        flags = str(span.get("flags", "")).lower() 
                        
                        is_bold = "bold" in flags
                        is_italic = "italic" in flags
                        
                        paragraphs.append({
                            "text": text_content,
                            "font": font_name,
                            "size": font_size,
                            "color": font_color,
                            "bold": is_bold,
                            "italic": is_italic
                        })
    
    doc.close()

    # Save to text file
    with open(output_file, 'w', encoding='utf-8') as f:
        for paragraph in paragraphs:
            f.write(f"Text: {paragraph['text']}\n")
            f.write(f"Font: {paragraph['font']}\n")
            f.write(f"Size: {paragraph['size']}\n")
            f.write(f"Color: {paragraph['color']}\n")
            f.write(f"Bold: {paragraph['bold']}\n")
            f.write(f"Italic: {paragraph['italic']}\n\n")

    print(f"text details saved to {output_file}")

def extract_and_uppercase_pdf(input_pdf_path, output_pdf_path):
    doc = fitz.open(input_pdf_path)
    new_doc = fitz.open()

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
        
        blocks = page.get_text("dict")["blocks"]
        
        for block in blocks:
            if block['type'] == 0:  # this block contains text
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"].upper()
                        bbox = fitz.Rect(span['bbox'])

                        # Insert text into the new page with the same formatting
                        new_page.insert_text(bbox.tl, text, fontsize=span['size'], fontname=span['font'], color=span['color'])

    new_doc.save(os.path.join(output_pdf_path, 'uppercase'))
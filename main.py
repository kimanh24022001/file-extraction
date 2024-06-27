import argparse
import os
from extract_pdf import extract_from_pdf, extract_pdf_text_details, extract_and_uppercase_pdf
from extract_docx import extract_from_docx,extract_and_uppercase_docx,extract_docx_text_details
from extract_pptx import translate_pptx

def main():
    parser = argparse.ArgumentParser(description='Process some files.')
    parser.add_argument('file', type=str, help='The path to the file to process')
    parser.add_argument('output', type=str, help='The output directory for the processed files')

    args = parser.parse_args()
    
    file_path = args.file
    output_dir = args.output

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    filename = os.path.basename(file_path)
    
    ext = filename.split('.')[-1].lower()
    if ext == 'pdf':
        output_path = os.path.join(output_dir, filename.split('.')[0].lower()) 
        extract_from_pdf(file_path, output_path)
        extract_pdf_text_details(file_path, output_path)
        extract_and_uppercase_pdf(file_path, output_path)
    elif ext == 'docx':
         output_path = os.path.join(output_dir, filename.split('.')[0].lower()) 
         extract_from_docx(file_path, output_path)
         extract_docx_text_details(file_path, output_path)
         extract_and_uppercase_docx(file_path, output_path)
    elif ext == 'pptx':
        output_path = os.path.join(output_dir, filename.split('.')[0].lower()) 
        translate_pptx(file_path, output_path)
    else:
        print('Unsupported file type')
        return

    print(f'Processed file saved to {output_path}')

if __name__ == "__main__":
    main()

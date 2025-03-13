import os
import re
import argparse
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter

def extract_email(pdf_reader, start_page, end_page):
    """
    Extract email address from a range of pages in a PDF file using PyPDF2.
    
    Args:
        pdf_reader (PdfReader): PyPDF2 reader object
        start_page (int): First page to analyze (0-based)
        end_page (int): Last page to analyze (0-based)
    
    Returns:
        str: Found email address or None
    """
    try:
        # Extract text from the specified range of pages
        text = ""
        for page_num in range(start_page, end_page + 1):
            if page_num < len(pdf_reader.pages):
                text += pdf_reader.pages[page_num].extract_text()
        
        # Search for email pattern - improved version handling special characters
        email_pattern = r'[A-Za-z0-9._%+\-_]+@[A-Za-z0-9.\-_]+\.[A-Z|a-z]{2,}'
        match = re.search(email_pattern, text)
        
        if match:
            return match.group(0)
        return None
    except Exception as e:
        print(f"Error extracting email: {e}")
        return None

def split_doc_to_pdfs(input_doc, output_folder, pages_per_file=2, extract_emails=False, filename_prefix="part"):
    """
    Converts a Word document to PDF and splits it into smaller PDFs.
    
    Args:
        input_doc (str): Path to the Word document
        output_folder (str): Folder where PDFs will be saved
        pages_per_file (int): Number of pages in each output file (default 2)
        extract_emails (bool): Whether to extract emails from pages for filenames
        filename_prefix (str): Custom prefix for output filenames (default "part")
    """
    # Make sure output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Temporary PDF file
    temp_pdf = os.path.join(output_folder, "temp_full.pdf")
    
    # Convert DOC/DOCX to PDF
    convert(input_doc, temp_pdf)
    
    # Open the full PDF
    reader = PdfReader(temp_pdf)
    total_pages = len(reader.pages)
    
    # Split into parts with specified number of pages
    for i in range(0, total_pages, pages_per_file):
        writer = PdfWriter()
        start_page = i
        end_page = min(i + pages_per_file - 1, total_pages - 1)
        
        # Add pages to the new PDF
        for page_num in range(start_page, end_page + 1):
            writer.add_page(reader.pages[page_num])
        
        # Default filename
        if extract_emails:
            file_prefix = filename_prefix
        else:
            file_prefix = f"{filename_prefix}_{i//pages_per_file + 1}"
        
        # Extract email if requested
        if extract_emails:
            email = extract_email(reader, start_page, end_page)
            if email:
                # Use sanitized email as part of filename
                sanitized_email = re.sub(r'[\\/*?:"<>|]', "_", email)
                file_prefix = f"{filename_prefix}_{sanitized_email}"
        
        # Save the new PDF
        output_pdf = os.path.join(output_folder, f"{file_prefix}.pdf")
        with open(output_pdf, "wb") as output_file:
            writer.write(output_file)
    
    # Remove temporary file
    os.remove(temp_pdf)
    print(f"Document split into {(total_pages + pages_per_file - 1) // pages_per_file} PDF files")

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Split Word document into multiple PDFs')
    parser.add_argument('input_file', help='Path to the input Word document')
    parser.add_argument('output_folder', help='Folder where output PDFs will be saved')
    parser.add_argument('-p', '--pages', type=int, default=2, help='Number of pages per output file (default: 2)')
    parser.add_argument('-e', '--extract-emails', action='store_true', 
                        help='Extract email addresses from pages and use them in filenames')
    parser.add_argument('-n', '--name', type=str, default='part',
                        help='Custom prefix for output filenames (default: "part")')
    
    args = parser.parse_args()
    
    # Run the conversion
    split_doc_to_pdfs(args.input_file, args.output_folder, args.pages, args.extract_emails, args.name)

if __name__ == "__main__":
    main()
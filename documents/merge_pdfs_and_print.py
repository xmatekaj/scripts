import os
from PyPDF2 import PdfMerger

def merge_pdfs(directory_path, output_path="merged_output.pdf", print_output=False):
    """
    Merges all PDF files in the specified directory into a single PDF.
    Optionally prints the merged file if requested.
    
    Args:
        directory_path (str): Path to the directory containing PDF files
        output_path (str): Path for the output merged PDF file
        print_output (bool): Whether to print the output file after merging
        
    Returns:
        bool: True if merge was successful, False otherwise
    """
    # Check if the directory exists
    if not os.path.isdir(directory_path):
        print(f"Error: Directory '{directory_path}' does not exist.")
        return False
    
    # Find all PDF files
    pdf_files = []
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.lower().endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                pdf_files.append(pdf_path)
    
    if not pdf_files:
        print("No PDF files found in the directory.")
        return False
    
    # Sort PDF files (for consistent ordering)
    pdf_files.sort()
    
    # Create a PdfMerger object
    merger = PdfMerger()
    
    # Loop through all PDF files and append them to the merger
    for pdf in pdf_files:
        try:
            print(f"Adding: {pdf}")
            merger.append(pdf)
        except Exception as e:
            print(f"Error adding {pdf}: {str(e)}")
    
    # Write the merged PDF to the output file
    try:
        merger.write(output_path)
        merger.close()
        print(f"\nSuccessfully merged {len(pdf_files)} PDFs into {output_path}")
        
        # Print the merged file if requested
        if print_output:
            import subprocess
            import platform
            
            system = platform.system()
            try:
                print(f"Sending merged PDF to printer...")
                
                if system == "Windows":
                    # Windows
                    subprocess.run(['start', '', '/p', output_path], shell=True, check=True)
                elif system == "Darwin":
                    # macOS
                    subprocess.run(['lpr', output_path], check=True)
                elif system == "Linux":
                    # Linux
                    subprocess.run(['lpr', output_path], check=True)
                else:
                    print(f"Unsupported system for printing: {system}")
                    return False
                    
                print("Print job sent successfully!")
                
            except subprocess.SubprocessError as e:
                print(f"Error printing merged PDF: {str(e)}")
                return False
        
        return True
            
    except Exception as e:
        print(f"Error writing merged PDF: {str(e)}")
        return False

if __name__ == "__main__":
    # Get the directory path from user input (default to current directory)
    directory = input("Enter the directory path containing PDFs to merge (press Enter for current directory): ")
    if not directory:
        directory = os.getcwd()
        print(f"Using current directory: {directory}")
    
    # Get the output file path (optional)
    output = input("Enter the output file path (press Enter for default 'merged_output.pdf'): ")
    if not output:
        output = "merged_output.pdf"
    
    # Ask if user wants to print the output
    print_question = input("Do you want to print the merged PDF after creation? (y/n): ")
    should_print = print_question.lower() == 'y'
    
    # Merge all PDFs in the directory
    merge_pdfs(directory, output, print_output=should_print)
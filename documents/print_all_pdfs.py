import os
import subprocess
import platform

def print_pdfs(directory_path):
    """
    Prints all PDF files in the specified directory using the default system printer.
    
    Args:
        directory_path (str): Path to the directory containing PDFs to print
        
    Returns:
        list: List of PDF files that were sent to the printer
    """
    pdf_files = []
    
    # Check if the directory exists
    if not os.path.isdir(directory_path):
        print(f"Error: Directory '{directory_path}' does not exist.")
        return pdf_files
    
    # Determine the operating system to use the correct print command
    system = platform.system()
    
    # Walk through the directory
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            # Check if the file has a .pdf extension (case insensitive)
            if file.lower().endswith('.pdf'):
                # Get the full path to the PDF file
                pdf_path = os.path.join(root, file)
                pdf_files.append(pdf_path)
                
                # Print the PDF file based on operating system
                try:
                    if system == "Windows":
                        # Windows - use the default application to print
                        print(f"Printing: {pdf_path}")
                        # Use ShellExecute in Windows to print
                        os.startfile(pdf_path, "print")
                    elif system == "Darwin":
                        # macOS
                        print(f"Printing: {pdf_path}")
                        subprocess.run(['lpr', pdf_path], check=True)
                    elif system == "Linux":
                        # Linux
                        print(f"Printing: {pdf_path}")
                        subprocess.run(['lpr', pdf_path], check=True)
                    else:
                        print(f"Unsupported system: {system}")
                        return pdf_files
                    
                except subprocess.SubprocessError as e:
                    print(f"Error printing {pdf_path}: {str(e)}")
    
    print(f"\nTotal PDFs sent to printer: {len(pdf_files)}")
    return pdf_files

if __name__ == "__main__":
    # Get directory path (default to current directory if empty)
    directory = input("Enter the directory path to print PDFs from (press Enter for current directory): ")
    if not directory:
        directory = os.getcwd()
        print(f"Using current directory: {directory}")
    
    # Confirm before printing potentially many files
    pdf_count = sum(1 for root, dirs, files in os.walk(directory) 
               for file in files if file.lower().endswith('.pdf'))
    
    if pdf_count > 0:
        confirm = input(f"Found {pdf_count} PDF files. Print them all? (y/n): ")
        if confirm.lower() == 'y':
            # Print all PDFs in the directory
            print_pdfs(directory)
        else:
            print("Printing cancelled.")
    else:
        print("No PDF files found in the directory.")
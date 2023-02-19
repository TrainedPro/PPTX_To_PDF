import os
import subprocess

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Change the current working directory to the script directory
os.chdir(script_dir)

# Get all the files in the current directory with the .ppt or .pptx extension
files = [f for f in os.listdir('.') if f.endswith('.pptx') or f.endswith('.ppt')]

# Sort the files by modified time, oldest first
files.sort(key=lambda x: os.path.getmtime(x))

# Check if the user wants to combine all files into one PDF or have separate PDFs
combine_pdfs = input("Do You Want To Combine All Files Into One PDF? (Y/N): ").lower() == "y"


# Initialize the PDFMerger
if combine_pdfs:
    from PyPDF2 import PdfWriter
    merger = PdfWriter()

pdf_files = []
# Loop through the files and convert each one to a PDF
for file in files:
    # Get the file extension and base filename
    filename = os.path.splitext(file)[0]
    extension = os.path.splitext(file)[1]

    # Get the output PDF filename
    pdf_file = filename + ".pdf"

    print(f"    Starting {pdf_file}")

    # Convert the file to a PDF
    try:
        subprocess.call(['libreoffice', '--headless', '--convert-to', 'pdf', file])
    except subprocess.CalledProcessError:
        print(f"Unable to convert {pdf_file}. Skipping...")    

    os.rename(pdf_file, os.path.join(os.getcwd(), pdf_file))
    pdf_files.append(pdf_file)
    print(f"    Complete {pdf_file}!")
    
    
if combine_pdfs:

    # Append files
    for file in pdf_files:
        merger.append(file)

    # Name the merged document
    merger.write(f"{os.getcwd().split('/')[-1]}_Merged_PDF")
    merger.close()

    # Delete files
    for file in pdf_files:
        os.remove(file)
    
    print("Combination Complete!")





# This script renames all the PDF files in the directory to the title of the PDF file
# The title is extracted from the metadata of the PDF file
# If the title metadata is empty then it extracts the title from the first page of the PDF file
# The title is extracted from the first line of the first page of the PDF file

# Author: @fahim9778
# Date: 25-Mar-2023

# Special thanks to chatGPT, BingAi and GitHub Copilot for their immense help with this script
# And also to the StackOverflow community for their help with the regex

# To run this script, you need to install PyPDF2 using pip or pip3 by running the following command in the terminal
# pip install PyPDF2

import os
import re
from PyPDF2 import PdfReader

def rename_pdf_files():
    # Get the directory of the script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f'Checking in directory: {script_dir} for PDF files... ')
    # Change the working directory to the script directory
    os.chdir(script_dir)

    # Get a list of all the PDF files in the directory
    filenames = [filename for filename in os.listdir('.') if filename.endswith('.pdf')]

    # Check if there are any PDF files in the directory
    if not filenames:
        print('No PDF files found.')
        return
    else:
        print(f"Found {len(filenames)} PDF file(s) with name(s): ")
        # enumerate the filenames with a counter starting at 1
        for counter, filename in enumerate(filenames, start=1):
            print(f'{counter}. {filename}')
    
    for filename in filenames:
            try:
                # with open(filename, 'rb') as f: # for python 3.6 and below use this line
                # for python 3.7 and above use this line. It's the same as the line above but with utf-8 encoding
                with open(filename.encode('utf-8'), 'rb') as f:
                    pdf = PdfReader(f)
                    title = pdf.metadata.title
                    if not title: # if the title metadata is empty then extract the title from the first page 
                            pageObj = pdf.pages[0]
                            text = pageObj.extract_text()
                            lines = text.splitlines()
                            title = lines[0] # assuming the first line of the first page is the title
                    
                    # remove special characters from the title and replace them with an underscore
                    title = re.sub(r'[^\w\s-]', ' _', title)
                    # remove multiple spaces
                    title = re.sub(r'\s+', ' ', title)
                    new_filename = f'{title}.pdf'
                    # close the file before renaming it
                    f.close()
                    # rename the file
                    os.rename(filename, new_filename)
                    print(f'Renamed {filename} to {new_filename}')
                    continue
            except:
                print(f'Error renaming {filename}')
                continue

# Run the script
if __name__ == '__main__':
    rename_pdf_files()
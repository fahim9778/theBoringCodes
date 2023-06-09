# This script renames all the PDF files in the directory to the title of the PDF file
# The title is extracted from the metadata of the PDF file
# If the title metadata is empty then it extracts the title from the first page of the PDF file
# The title is extracted from the first line of the first page of the PDF file

# Author: @fahim9778
# Date: 25-Mar-2023

# Special thanks to chatGPT, BingAi and GitHub Copilot for their immense help with this script
# And also to the StackOverflow community for their help with the regex

# [ATTENTION!]
# To run this script, you MUST NEED to install PyPDF2 using pip or pip3 by running the following command in the terminal
# pip install PyPDF2

# [ATTENTION!]
# To successfully run this script, KEEP THE PDF FILES IN THE SAME DIRECTORY AS THE SCRIPT
# The script will NOT WORK if the PDF files are in a different directory

# [NOTE!]
# if the program escapes the loop before renaming all the PDF files, then kindly run the script again


import os
import re
from PyPDF2 import PdfReader

def rename_pdf_files():
    # Get the directory of the script
    script_dir = os.getcwd()
    print(f'Checking in directory: {script_dir} for PDF files... ')

    # Get a list of all the PDF files in the directory
    filenames = [filename for filename in os.listdir('.') if filename.endswith('.pdf')]

    # Create empty lists to store the filenames of the PDF files that were renamed successfully and the ones that failed to rename
    succeded = []
    failed = []

    # Check if there are any PDF files in the directory
    if not filenames:
        print('No PDF files found.')
        input('Press any key to continue...')
        return
    else:
        print(f"Found {len(filenames)} PDF file(s) with name(s): ")
        # enumerate the filenames with a counter starting at 1
        for counter, filename in enumerate(filenames, start=1):
            print(f'{counter}. {filename}')
        print('Renaming PDF files...\n')

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
                            if not lines: # if the first page is empty then skip the file
                                print(f'No title found in {filename}')
                                print('----------------------------------------\n')
                                continue
                            # if the first line contains a hyperlink then the title could be the second line
                            if contains_hyperlink(lines[0]):
                                title = lines[1]
                            else:
                                title = lines[0] # assuming the first line of the first page is the title
                            print(f'Extracted title: {title} from {filename}')
                                            
                    # remove special characters from the title and replace them with an underscore
                    title = re.sub(r'[^\w\s-]', ' _', title)
                    # remove multiple spaces
                    title = re.sub(r'\s+', ' ', title)
                    # create the new filename
                    new_filename = f'{title}.pdf'
                    # close the file before renaming it
                    f.close()
                    # rename the file
                    os.rename(filename, new_filename)
                    print(f'Renamed {filename} to {new_filename}')
                    succeded.append(filename)
                    print('----------------------------------------\n')
                    continue
            except:
                print(f'Error renaming {filename}')
                failed.append(filename)
                print('!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n')
                continue
    
    # Print the filenames of the PDF files that were renamed successfully and the ones that failed to rename
    if succeded:
        print(f'Renamed {len(succeded)} PDF file(s) successfully.')
    if failed:
        print(f'Failed to rename {len(failed)} PDF file(s): ')
        for counter, failedfilename in enumerate(failed, start=1):
            print(f'{counter}. {failedfilename}')

    input('Press any key to continue...')

# Check if a string contains a hyperlink
def contains_hyperlink(string):
    regex = r"(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'\".,<>?«»“”‘’]))"
    return bool(re.search(regex, string))

# Run the script
if __name__ == '__main__':
    rename_pdf_files()
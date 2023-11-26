# Renewed approach using pdfminer (pattern matching instead of xml name matching)
# R&D: https://www.perplexity.ai/search/6ff640de-385a-47e2-8f85-68b0412b9868 (original thread) and https://www.perplexity.ai/search/1c227236-2441-419e-8ddb-0ced28299f96?s=u
# It worked!!! :) (let's try it now on bigger samples to find potencial new problems)
# v1.2 - Works well with minor issue (cannot convert lowecase to upper due to Windows restrictions)
import os, re, time, datetime
import pandas as pd
from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser


# Define the folder path (input) and *report* output path (ouput)
folder_path = r'C:\Users\Sanchezrl\Downloads\WORK DOWNLOADS\1. Facturas\autoRenameAux'
output_path = r'C:\BECARI@\PENDIENTES 2023\1.2 FACTURAS\pwrQry_facturas\xProcesar'

# Get the list of files in the folder and sort them by creation time (oldest to latest)
files = sorted(
    os.listdir(folder_path),
    key=lambda f: os.path.getctime(os.path.join(folder_path, f)))
print('Sorting files by latest creation time')

# Iterate through the list of all the files in the folder and filter the list to only include PDF files
pdf_files = [f for f in files if f.casefold().endswith('.pdf')]
print('Filtering only PDF files')

# Create empty lists to store the data
data = []
not_found = []
errors = []

# Iterate through all the PDF files in the folder
for file in pdf_files:
    try:
        file_path = os.path.join(folder_path, file)

        with open(file_path, 'rb') as f:
            parser = PDFParser(f)
            doc = PDFDocument(parser)
            rsrcmgr = PDFResourceManager()
            output_string = StringIO()
            device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            for page in PDFPage.create_pages(doc):
                interpreter.process_page(page)

        text = output_string.getvalue()
        print(f'Extracting text from {file}')

        pattern_match = re.search(r'\b\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\b', text)
        if pattern_match:
            pattern = pattern_match.group(0)
            old_filename = file

            # Check if the pattern match is already in the filename in uppercase
            if pattern.upper() in old_filename.upper():
                print(f'Pattern match already in filename: {old_filename}')
                new_filename = old_filename
            else:
                # RENAME the file with the pattern match in uppercase (only variables)
                new_filename = f'{pattern.upper()}.pdf'
                # Check if the file with the new name already exists in the folder
                while os.path.exists(os.path.join(folder_path, new_filename)):
                    # If the new filename already exists, add a counter to the end of the filename
                    counter = 2
                    while os.path.exists(
                        os.path.join(folder_path, f'{pattern.upper()}_{counter}.pdf')
                    ):
                        counter += 1
                    new_filename = f'{pattern.upper()}_{counter}.pdf'

                f.close()  # Close the file before renaming
                time.sleep(1)  # Add a delay before renaming
                os.rename(file_path, os.path.join(folder_path, new_filename))
                print(f'Renaming {old_filename} to {new_filename}')

                # Append the data to the list
                data.append([old_filename, new_filename])
        else:
            print(f'No pattern found in {file}')
            not_found.append(file)
    except Exception as e:
        print(f'Error processing {file}: {str(e)}')
        errors.append([file, str(e)])

print('Finished processing all files')

# Create DataFrames from the data
df_data = pd.DataFrame(data, columns=['Old Filename', 'New Filename'])
df_not_found = pd.DataFrame(not_found, columns=['Filename'])
df_errors = pd.DataFrame(errors, columns=['Filename', 'Error'])
print('Creating DataFrames')

# Create the report file name with the current date
date_str = datetime.datetime.now().strftime('%Y%m%d')
report_filename = f'{date_str}_PDF renaming report.xlsx'

# Check if the report file already exists
counter = 2
while os.path.exists(os.path.join(output_path, report_filename)):
    report_filename = f'{date_str}_PDF renaming report_{counter}.xlsx'
    counter += 1

# Export the DataFrames to Excel
report_path = os.path.join(output_path, report_filename)
with pd.ExcelWriter(report_path) as writer:
    df_data.to_excel(writer, sheet_name='Data', index=False)
    df_not_found.to_excel(writer, sheet_name='Not Found', index=False)
    df_errors.to_excel(writer, sheet_name='Errors', index=False)
    print(f'Exporting to Excel workbook: {report_filename}')

print('Report finished')

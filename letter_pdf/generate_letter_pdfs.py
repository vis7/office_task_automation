"""
This file take excel containing list of students and generate pdf letter for them 
from given template.
"""

import os
from subprocess import  Popen
from docx import Document
from openpyxl import load_workbook

# getting name from excel sheet
workbook = load_workbook('data/joining_student_list.xlsx')

sheet = workbook.active

# generate letter in doc format
for i in range(2, 4):
    name = sheet.cell(row=i, column=1).value
    subject = sheet.cell(row=i, column=2).value
    date = sheet.cell(row=i, column=3).value

    # reformate date to proper format
    if date:
        date = date.strftime('%d/%m/%Y')

        print(f"Generating certificate for: {name} {subject} {date}")

        #######################################
        # code for replacing name in doc
        #######################################
        #open the document
        doc=Document('data/joining_letter_template.docx')

        for p in doc.paragraphs:
            inline = p.runs
            for i in range(len(inline)):
                text = inline[i].text
                text=text.replace('Name',name)
                text=text.replace('Subject', subject)
                text=text.replace('Date', date)
                inline[i].text = text
        
        # save changed document
        doc.save(f'./docs/Internship Confirmation {name}.docx')

# convert doc to pdf
LIBRE_OFFICE = r"/usr/bin/libreoffice"

def convert_to_pdf(input_docx, out_folder):
    p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
               out_folder, input_docx])
    print([LIBRE_OFFICE, '--convert-to', 'pdf', input_docx])
    p.communicate()


sample_doc_dir = 'docs/'
out_folder = 'pdfs/'

for i in os.listdir(sample_doc_dir):
    convert_to_pdf(os.path.join(sample_doc_dir, i), out_folder)


# remove docs
for file_name in os.listdir(sample_doc_dir):
    os.remove(os.path.join(sample_doc_dir, file_name))

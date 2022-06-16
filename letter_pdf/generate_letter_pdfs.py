"""
This file take excel containing list of students and generate pdf letter for them 
from given template.
"""
###############################################################
# Adding absolute path to the root directory for using absolute path
###############################################################
import os
import sys

abs_path_of_directory = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.sys.path.append(abs_path_of_directory)
################################################################
from docx import Document
from openpyxl import load_workbook
from letter_pdf.utils import send_mail, create_session, destroy_session, convert_to_pdf

# getting name from excel sheet
workbook = load_workbook('data/student list.xlsx')
sheet = workbook.active
sample_doc_dir = 'docs/'
out_folder = 'pdfs/'

session = create_session()

# generate letter in doc format
for i in range(47, 492):
    name = sheet.cell(row=i, column=1).value
    subject = sheet.cell(row=i, column=2).value
    date = sheet.cell(row=i, column=3).value
    email = sheet.cell(row=i, column=4).value

    # refine email
    if email:
        email = email.strip()

    # reformate date to proper format
    if date:
        # date = date.strftime('%d/%m/%Y')

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
        letter_doc_path = f'./docs/Internship Confirmation {name}.docx'
        doc.save(letter_doc_path)

        convert_to_pdf(letter_doc_path, out_folder)

        if email:
            # sending mail
            letter_pdf_path = f'./pdfs/Internship Confirmation {name}.pdf'
            send_mail(email, session, letter_pdf_path)

destroy_session(session)

# remove docs
for file_name in os.listdir(sample_doc_dir):
    os.remove(os.path.join(sample_doc_dir, file_name))

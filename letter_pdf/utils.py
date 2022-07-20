# libraries to be imported
import smtplib
import platform
import logging
import re
from docx2pdf import convert
from smtplib import SMTPResponseException
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from subprocess import  Popen

fromaddr = "help.vnurture@gmail.com"
password = "brdeetidhbilhrah"
toaddr = "vishvajeetramanuj95@gmail.com"

os_type = platform.system()

logger = logging.getLogger(__name__)
f_handler = logging.FileHandler('sending_email.log')
f_handler.setLevel(logging.ERROR)

LIBRE_OFFICE = r"/usr/bin/libreoffice"

def send_mail(receiver_email, session, filename, task='send_joining_letter'):
    # task
    # task = 'send_certificate'
    if task == 'send_certificate':
        body_part = "Certificate"
    elif task == 'send_completion_letter':
        body_part = "Completion Letter"
    else: # assuming task is joining letter
        pass

    # string to store the body of the mail
    body = f"""
Greetings from Vnurture Technologies,

Congratulations! You have successfully completed 15 days Internship Programme. 
kindly find your {body_part} which attached below.

Stay tuned with Vnurture Services.

Best wishes for your career and will see you in the session.

Thanks,
Vnurture Technologies
(https://www.vnurture.in/) 
"""

    # instance of MIMEMultipart
    msg = MIMEMultipart()
    
    # storing the senders email address  
    msg['From'] = fromaddr

    # storing the subject 
    msg['Subject'] = f"Vnurture Internship {body_part}"


    # storing the receivers email address 
    msg['To'] = receiver_email

    # attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))

    # open the file to be sent 
    # filename = "File_name_with_extension"
    attachment = open(filename, "rb")

    # instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')
    
    # To change the payload into encoded form
    p.set_payload((attachment).read())
    
    # encode into base64
    encoders.encode_base64(p)
    
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename.split('/')[-1])
    
    # attach the instance 'p' to instance 'msg'
    msg.attach(p)

    # Converts the Multipart msg into a string
    text = msg.as_string()
    
    # sending the mail
    try:
        print('sending mail...')
        session.sendmail(fromaddr, receiver_email, text)
    except SMTPResponseException as e:
        print(f"{receiver_email} {e.smtp_code} {e.smtp_error}")
        logger.error(f"{receiver_email} {e.smtp_code} {e.smtp_error}")

def create_session():
    # creates SMTP session
    session = smtplib.SMTP('smtp.gmail.com', 587)
    
    # start TLS for security
    session.starttls()
    
    # Authentication
    session.login(fromaddr, password)
    return session
  
def destroy_session(session):
    # terminating the session
    session.quit()

# convert doc to pdf
def convert_to_pdf(input_docx, out_folder):
    if os_type == 'Linux':
        out_folder = "letter_pdf/pdfs/" # file will be export in this folder

        p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
                out_folder, input_docx])
        print([LIBRE_OFFICE, '--convert-to', 'pdf', input_docx])
        p.communicate()
    elif os_type == 'Windows':
        convert(input_docx, out_folder)
    else:
        print('Invalid Operation System.')
        print('Works only on Linux and Windows.')

def check_email(email):
    regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    
    if(re.fullmatch(regex, email)):
        # print("Valid Email") 
        return True
    else:
        # print("Invalid Email")
        return False

if __name__ == "__main__":
    session = create_session()
    send_mail(toaddr, session, 'test.pdf')
    destroy_session(session)  

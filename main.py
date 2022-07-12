import tkinter
from tkinter.filedialog import askopenfilename
from letter_pdf.generate_letter_pdfs import send_confirmation_letter
from certificate_pdf.edit_image import send_certificate, send_completion_letter

confirmation_letter_message = """
Greetings from Vnurture Technologies,

Congratulations! You are successfully enrolled in 15 days Internship Programme. 
kindly find your Confirmation letter which attached below.

Stay tuned with Vnurture Services.

Best wishes for your career and will see you in the session.

Thanks,
Vnurture Technologies
(https://www.vnurture.in/) 
"""

def certificate_cb():
    filename = askopenfilename(filetypes=(("Excel File", "*.xlsx"),))
    print(f'filename is {filename}')
    send_certificate(filename)
    print("All letter sent successfully.")
    print("Kindly close the application")

def completion_letter_cb():
    filename = askopenfilename(filetypes=(("Excel File", "*.xlsx"),))
    print(f'filename is {filename}')
    send_completion_letter(filename)
    print("All letter sent successfully.")
    print("Kindly close the application")

# def confirmation_letter_cb():
#     filename = askopenfilename(filetypes=(("Excel File", "*.xlsx"),))
#     print(f'filename is {filename}')
#     send_confirmation_letter(filename, completion_letter_message)
#     print("All letter sent successfully.")
#     print("Kindly close the application")

if __name__ == "__main__":
    main_window = tkinter.Tk(className='vnurture_training_utility')
    main_window.geometry("400x200")

    certificate_btn = tkinter.Button(main_window, text="Send Certificate", command=certificate_cb)
    certificate_btn.config(width=20, height=2)

    completion_letter_btn = tkinter.Button(main_window, text="Send Completion Letter", command=completion_letter_cb)
    completion_letter_btn.config(width=20, height=2)

    certificate_btn.pack()
    completion_letter_btn.pack()
    main_window.mainloop()

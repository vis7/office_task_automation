import tkinter
from tkinter.filedialog import askopenfilename
from letter_pdf.generate_letter_pdfs import send_confirmation_letter

def submitFunction():
    filename = askopenfilename(filetypes=(("Excel File", "*.xlsx"),))
    print(f'filename is {filename}')
    send_confirmation_letter(filename)
    print("All letter sent successfully.")
    print("Kindly close the application")

if __name__ == "__main__":
    main_window = tkinter.Tk(className='demo_window')
    main_window.geometry("400x200")

    button_submit = tkinter.Button(main_window, text="Select the File", command=submitFunction)
    button_submit.config(width=20, height=2)

    button_submit.pack()
    main_window.mainloop()

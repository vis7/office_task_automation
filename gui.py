import tkinter
from tkinter.filedialog import askopenfilename
from turtle import onclick

main_window = tkinter.Tk(className='demo_window')
main_window.geometry("400x200")

def submitFunction():
    filename = askopenfilename(filetypes=(("Excel File", "*.xlsx"),))
    print(f'filename is {filename}')

button_submit = tkinter.Button(main_window, text="Submit", command=submitFunction)
button_submit.config(width=20, height=2)

button_submit.pack()
main_window.mainloop()

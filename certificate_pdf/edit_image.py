"""
This program take certificate image and student details in xlsx file and create
certificates for them based on student details.
"""
from PIL import Image, ImageDraw, ImageFont
from openpyxl import load_workbook

workbook = load_workbook('data/student_details.xlsx')
sheet = workbook.active

size = 50 
color = (0,0,0)
font = ImageFont.truetype('data/fonts/lucida calligraphy italic.ttf',size)

name_position = (300, 480) # 1150
subject_position = (290, 590)
date_position = (260,700)

for i in range(2,7):
    name = sheet.cell(row=i, column=1).value
    subject = sheet.cell(row=i, column=2).value
    date = sheet.cell(row=i, column=3).value

    if date:
        date = date.strftime('%d/%m/%Y')
        if name:
            name = name.title()
            print(f"Generating Certificate for {name} {subject} {date}")
            img = Image.open("certificate.jpeg")
            image_ed = ImageDraw.Draw(img)
            # image_ed.text((0,0), 'vis', fill=(255,0,0))
            image_ed.text(name_position, name, fill=color, font=font)
            image_ed.text(subject_position, subject, color, font=font)
            image_ed.text(date_position, date, color, font=font)
            img.save(f'certificates/{name}_certificate.jpeg')

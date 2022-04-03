# This program is used to put a list of names onto a diploma template.
# It should be run from within the current working directory

# ----- IMPORTS ----- #
import shutil

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PyPDF4 import PdfFileWriter, PdfFileReader
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import os
from sys import platform

try:
    os.mkdir('Intermediate_Files')
except FileExistsError:
    print("There was a FileExistsError")
else:
    pass


date_font = "Times-Roman"
name_font = "Times-Bold"
available_fonts = [
    'Courier',
    'Courier-Bold',
    'Courier-BoldOblique',
    'Courier-Oblique',
    'Helvetica',
    'Helvetica-Bold',
    'Helvetica-BoldOblique',
    'Helvetica-Oblique',
    'Symbol',
    'Times-Bold',
    'Times-BoldItalic',
    'Times-Italic',
    'Times-Roman',
    'ZapfDingbats'
]
available_font_sizes = [n for n in range(10, 100)]

date_font_size = 14
name_font_size = 25
start_date = "May 13, 1999"
end_date = "April 3, 2022"
names_path = "No File Selected"
diploma_path = "No File Selected"
width, height = A4
translate_name = {
    'x': 0,
    'y': 0
}
translate_date = {
    'x': 0,
    'y': 0
}
files_to_merge = []
names = []


# Function used to create separate pdf with only a name on it
def create_name_pdf(c_, name):
    # print(c_.getAvailableFonts())
    c_.translate(height / 2 + translate_name['x'], width * 0.70 + translate_name['y'])
    # c_.setFillColorRGB(255, 255, 255)
    # c_.rect(-width * 0.2, -10, height * 0.4, 25, stroke=0, fill=1)
    c_.setFont(name_font, name_font_size)
    c_.setFillColorRGB(0, 0, 0)
    c_.drawCentredString(0, 0, name)
    c_.translate(-translate_name['x'], -translate_name['y'])
    c_.translate(translate_date['x'], width * 0.07 + translate_date['y'])
    c_.setFont(date_font, date_font_size)
    c_.drawCentredString(0, 0, f"{start_date} - {end_date}")


# Merge to create each unique diploma
def write_name_on_diploma(temp, output, name_pdf):
    name_obj = PdfFileReader(name_pdf)
    name_page = name_obj.getPage(0)

    pdf_reader = PdfFileReader(temp)
    pdf_writer = PdfFileWriter()

    for page in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page)
        page.mergePage(name_page)
        pdf_writer.addPage(page)

    with open(output, "wb") as out:
        pdf_writer.write(out)


# Combine all into a final file
def merge_diplomas(paths, output):
    pdf_writer = PdfFileWriter()

    for path in paths:
        pdf_reader = PdfFileReader(path)
        pdf_writer.addPage(pdf_reader.getPage(0))

    with open(output, "wb") as out:
        pdf_writer.write(out)


def generate_diplomas():
    global selected_name_font, selected_date_font,  selected_name_size, selected_date_size, start_date_input, \
        end_date_input
    global name_font, date_font, start_date, end_date, name_font_size, date_font_size
    name_font = selected_name_font.get()
    date_font = selected_date_font.get()
    start_date = start_date_input.get() if start_date_input.get() != '' else 'January 00, 0000'
    end_date = end_date_input.get() if end_date_input.get() != '' else 'December 99, 9999'
    name_font_size = int(selected_name_size.get())
    date_font_size = int(selected_date_size.get())

    global translate_name, translate_date, translate_name_input, translate_date_input
    if translate_name_input.get() != "":
        translate_name['x'] = int(translate_name_input.get().split(',')[0].strip())
        translate_name['y'] = int(translate_name_input.get().split(',')[1].strip())
    if translate_date_input.get() != "":
        translate_date['x'] = int(translate_date_input.get().split(',')[0].strip())
        translate_date['y'] = int(translate_date_input.get().split(',')[1].strip())

    for curr_name in names:
        c = canvas.Canvas(f"Intermediate_Files/{curr_name}.pdf", pagesize=A4)
        files_to_merge.append(f"Intermediate_Files/output-{curr_name}.pdf")
        create_name_pdf(c, curr_name)
        c.showPage()
        c.save()

    for each in names:
        write_name_on_diploma(diploma_path, f"Intermediate_Files/output-{each}.pdf",
                              f"Intermediate_Files/{each}.pdf")

    # Export final file to desktop
    if platform == "linux" or platform == "darwin":
        merge_diplomas(files_to_merge, os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop/final_output.pdf'))
    if platform == "win32":
        merge_diplomas(files_to_merge, os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop/final_output.pdf'))


def select_names_file():
    filetypes = (
        ('text files', '*.txt'),
        ('All files', '*.*'),
    )

    filename = fd.askopenfile(
        title='Open a file',
        initialdir='.',
        filetypes=filetypes)

    global names_path
    names_path = filename.name
    names_file_label["text"] = names_path
    names_file_label["foreground"] = "blue"

    global names
    with open(names_path, "r") as f:
        names = f.readlines()
        names = [name.strip() for name in names]


def select_diploma_file():
    filetypes = (
        ('pdf files', '*.pdf'),
    )

    filename = fd.askopenfile(
        title='Open a file',
        initialdir='.',
        filetypes=filetypes)

    global diploma_path
    diploma_path = filename.name
    print(diploma_path)
    diploma_file_label["text"] = diploma_path
    diploma_file_label["foreground"] = "blue"


# ----- UI SETUP ----- #
base_col_width = 15

window = tk.Tk()
window.title("Diploma Generator")
window.resizable(True, True)
window.config(padx=100, pady=50)


select_names_file_prompt = tk.Label(text="Choose Names File (.txt only):")
select_names_file_prompt.grid(column=0, row=0, sticky='e')
names_file_label = ttk.Label(
    text=names_path,
)
names_file_label.grid(column=1, row=0, pady=10)
select_names_file_btn = ttk.Button(
    text="Choose name file",
    command=select_names_file,
    width=base_col_width,
)
select_names_file_btn.grid(column=2, row=0, columnspan=1, pady=10, padx=10, sticky='w')

select_diploma_file_prompt = tk.Label(text="Choose Diploma File (.pdf only):")
select_diploma_file_prompt.grid(column=0, row=1, sticky='e')
diploma_file_label = ttk.Label(
    text=diploma_path
)
diploma_file_label.grid(column=1, row=1, pady=10)
select_diploma_file_btn = ttk.Button(
    text="Choose diploma file",
    command=select_diploma_file,
    width=base_col_width,
)
select_diploma_file_btn.grid(column=2, row=1, columnspan=1, pady=10, padx=10, sticky='w')

name_font_prompt = tk.Label(text="Name Font:")
name_font_prompt.grid(column=0, row=2, pady=10, sticky='e')
selected_name_font = tk.StringVar()
selected_name_font.set(name_font)
drop = tk.OptionMenu(window, selected_name_font, *available_fonts)
drop.grid(column=1, row=2)

name_size_prompt = tk.Label(text="Name Font Size:")
name_size_prompt.grid(column=0, row=3, pady=10, sticky='e')
selected_name_size = tk.StringVar()
selected_name_size.set(name_font_size)
drop = tk.OptionMenu(window, selected_name_size, *available_font_sizes)
drop.grid(column=1, row=3)

date_font_prompt = tk.Label(text="Date Font:")
date_font_prompt.grid(column=0, row=4, pady=10, sticky='e')
selected_date_font = tk.StringVar()
selected_date_font.set(date_font)
drop = tk.OptionMenu(window, selected_date_font, *available_fonts)
drop.grid(column=1, row=4)

date_size_prompt = tk.Label(text="Date Font Size:")
date_size_prompt.grid(column=0, row=5, pady=10, sticky='e')
selected_date_size = tk.StringVar()
selected_date_size.set(date_font_size)
drop = tk.OptionMenu(window, selected_date_size, *available_font_sizes)
drop.grid(column=1, row=5)

start_date_prompt = tk.Label(text='Start Date:')
start_date_prompt.grid(column=0, row=6, pady=10, sticky='e')
start_date_input = tk.Entry(window)
start_date_input.grid(column=1, row=6)

end_date_prompt = tk.Label(text='End Date:')
end_date_prompt.grid(column=0, row=7, pady=10, sticky='e')
end_date_input = tk.Entry(window)
end_date_input.grid(column=1, row=7)

translate_name_label = tk.Label(text='Change Name Position:')
translate_name_label.grid(column=0, row=8, pady=10, sticky='e')
translate_name_input = tk.Entry(window)
translate_name_input.grid(column=1, row=8)

translate_date_label = tk.Label(text='Change Date Position:')
translate_date_label.grid(column=0, row=9, pady=10, sticky='e')
translate_date_input = tk.Entry(window)
translate_date_input.grid(column=1, row=9)

generate_btn = ttk.Button(
    text="Generate Diplomas",
    command=generate_diplomas)
generate_btn.grid(column=1, row=10, columnspan=1, pady=20)

warning_label = tk.Label(text="READ INSTRUCTIONS IF YOU ARE CONFUSED", font=('Arial', 30), fg='red')
warning_label.grid(column=0, row=11, columnspan=3)


window.mainloop()


# Drop down for name/date font/size
# Input for positioning text within the document
# Possible input for changing document size


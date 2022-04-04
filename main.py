# This program is used to put a list of names onto a diploma template.
# It should be run from within the current working directory

# ----- IMPORTS ----- #
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PyPDF4 import PdfFileWriter, PdfFileReader
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import os
from sys import platform
import pandas as pd


# Make sure that necessary directories exist
try:
    os.mkdir('Intermediate_Files')
except FileExistsError:
    print("ERROR: There was a FileExistsError.")
else:
    pass


# ----- VARIABLES ----- #
width, height = A4
name_font = "Times-Bold"
name_font_size = 25
date_font = "Times-Roman"
date_font_size = 14
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
start_date = "May 13, 1999"  # Stores beginning date of training
end_date = "April 3, 2022"  # Stores end date of training
names_path = "No File Selected"  # Stores full path to local list of names .txt file
diploma_path = "No File Selected"  # Stores full path to local diploma pdf file
translate_name = {  # Stores the amount to offset the name from default position in pixels
    'x': 0,
    'y': 0
}
translate_date = {  # Stores the amount to offset the date from default position in pixels
    'x': 0,
    'y': 0
}
files_to_merge = []  # Intermediate list to store individual diplomas before combining at the end
names = []  # Stores list of names from file


# ----- PDF METHODS ----- #
def create_name_pdf(c_, name):
    """Takes blank pdf canvas and a name and returns a pdf with only the name on it"""
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


def write_name_on_diploma(temp, name_pdf, output):
    """Takes diploma template and name pdf, returns a merged version to output"""
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


def merge_diplomas(paths, output):
    """Takes a list of single page pdf paths, returns a combined version to output"""
    pdf_writer = PdfFileWriter()

    for path in paths:
        pdf_reader = PdfFileReader(path)
        pdf_writer.addPage(pdf_reader.getPage(0))

    with open(output, "wb") as out:
        pdf_writer.write(out)


def generate_diplomas():
    """Uses all of the above methods to generate a single, multi-page, diploma document"""

    # Make sure that files have been selected
    if names_path == "No File Selected" or diploma_path == "No File Selected":
        showinfo(
            title='No File Selected',
            message='To generate the diplomas, you need to select a diploma template file (PDF) and a file containing '
                    'a list of names (.txt or .xlsx). One or both of these is missing.'
        )
        return

    # Update all variables, these probably should all be global but I was in a rush lol
    global selected_name_font, selected_date_font,  selected_name_size, selected_date_size, start_date_input, \
        end_date_input, name_font, date_font, start_date, end_date, name_font_size, date_font_size, \
        translate_name, translate_date, translate_name_input, translate_date_input
    name_font = selected_name_font.get()
    date_font = selected_date_font.get()
    start_date = start_date_input.get() if start_date_input.get() != '' else 'January 00, 0000'
    end_date = end_date_input.get() if end_date_input.get() != '' else 'December 99, 9999'
    name_font_size = int(selected_name_size.get())
    date_font_size = int(selected_date_size.get())

    if translate_name_input.get() != "":
        try:
            translate_name['x'] = int(translate_name_input.get().split(',')[0].strip())
            translate_name['y'] = int(translate_name_input.get().split(',')[1].strip())
        except ValueError:
            showinfo(
                title='Value Error',
                message='One of the values you entered in the "Change Name Position" field is incorrect. '
                        'Make sure to input 2 integers separated by a comma. '
                        'For example, to shift left 30 pixels and up 60 pixels, you should enter: -30, 60'
            )
            return
        except IndexError:
            showinfo(
                title='Value Error',
                message='One of the values you entered in the "Change Name Position" field is incorrect. '
                        'Make sure to input 2 integers separated by a comma. '
                        'For example, to shift left 30 pixels and up 60 pixels, you should enter: -30, 60'
            )
            return
    if translate_date_input.get() != "":
        try:
            translate_date['x'] = int(translate_date_input.get().split(',')[0].strip())
            translate_date['y'] = int(translate_date_input.get().split(',')[1].strip())
        except ValueError:
            showinfo(
                title='Value Error',
                message='One of the values you entered in the "Change Date Position" field is incorrect. '
                        'Make sure to input 2 integers separated by a comma. '
                        'For example, to shift left 30 pixels and up 60 pixels, you should enter: -30, 60'
            )
            return
        except IndexError:
            showinfo(
                title='Value Error',
                message='One of the values you entered in the "Change Name Position" field is incorrect. '
                        'Make sure to input 2 integers separated by a comma. '
                        'For example, to shift left 30 pixels and up 60 pixels, you should enter: -30, 60'
            )
            return

    # If list of names is empty, return and show error
    if not names:
        showinfo(
            title='Names File Empty',
            message='It looks like the list of names you selected is empty. '
                    'Try selecting another file or adding names to the .txt/.xlxs file you selected.'
        )
        return

    # Otherwise, create name pdf files
    for curr_name in names:
        c = canvas.Canvas(f"Intermediate_Files/{curr_name}.pdf", pagesize=A4)
        files_to_merge.append(f"Intermediate_Files/output-{curr_name}.pdf")
        create_name_pdf(c, curr_name)
        c.showPage()
        c.save()

    # Write the names on the diplomas
    for each in names:
        write_name_on_diploma(diploma_path, f"Intermediate_Files/{each}.pdf", f"Intermediate_Files/output-{each}.pdf")

    # Export final file to desktop
    if platform == "linux" or platform == "darwin":
        merge_diplomas(files_to_merge, os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop/final_output.pdf'))
    if platform == "win32":
        merge_diplomas(files_to_merge, os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop/final_output.pdf'))

    showinfo(
        title='Finished Operation',
        message='The file has been created and is located on your desktop. '
                'Rename/move the file before generating again or the previous version will be overwritten.'
    )


# ----- GUI METHODS ----- #
def select_names_file():
    """Used to select the names.txt file in the GUI"""
    filetypes = (
        ('text files', '*.txt'),
        ('excel files', '*.xlsx'),
        ('excel files', '*.xls'),
    )

    filename = fd.askopenfile(
        title='Open a file',
        initialdir='.',
        filetypes=filetypes)

    if filename:
        global names_path
        names_path = filename.name
        names_file_label["text"] = names_path
        names_file_label["foreground"] = "blue"

        global names
        if names_path[len(names_path) - 4:] == ".txt":
            with open(names_path, "r") as f:
                names = f.readlines()
                names = [name.strip() for name in names]
        if names_path[len(names_path) - 4:] == ".xls" or names_path[len(names_path) - 4:] == "xlsx":
            df = pd.read_excel(rf"{names_path}", header=None)
            df = df.to_dict()[0]
            names = list(df.values())
            names = [n.strip() for n in names]


def select_diploma_file():
    """Used to select the diploma.pdf file in the GUI"""
    filetypes = (
        ('pdf files', '*.pdf'),
    )

    filename = fd.askopenfile(
        title='Open a file',
        initialdir='.',
        filetypes=filetypes)

    if filename:
        global diploma_path
        diploma_path = filename.name
        diploma_file_label["text"] = diploma_path
        diploma_file_label["foreground"] = "blue"
    # if filename not in names:
    #     return True


# ----- GUI SETUP ----- #
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
names_file_label.config(foreground='red')
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
diploma_file_label.config(foreground='red')
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

# This program is used to put a list of names onto a diploma template.
# It should be run from within the current working directory

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, A5
from PyPDF4 import PdfFileWriter, PdfFileReader

date_font = "Times-Roman"
name_font = "Times-Bold"
date_font_size = 14
name_font_size = 25
start_date = "October 18, 2021"
end_date = "January 3, 2022"

files_to_merge = []
names_file = "names.txt"
beginner_names = []
advanced_names = []
beginner_template = "basic_template.pdf"
advanced_template = "advanced_template.pdf"
# template_file = "template.pdf"

# Create a list of names from names.txt
with open(names_file, "r") as f:
    name_list = f.readlines()
    name_list = [name.strip() for name in name_list]
    for n in name_list:
        if n[0] == '.':
            advanced_names.append(n[1:])
        else:
            beginner_names.append(n)


# Function used to create separate pdf with only a name on it
def create_name_pdf(c_, name):
    # print(c_.getAvailableFonts())
    c_.translate(height / 2, width * 0.70)
    # c_.setFillColorRGB(255, 255, 255)
    # c_.rect(-width * 0.2, -10, height * 0.4, 25, stroke=0, fill=1)
    c_.setFont(name_font, name_font_size)
    c_.setFillColorRGB(0, 0, 0)
    c_.drawCentredString(0, 0, name)
    c_.translate(0, width * 0.07)
    c_.setFont(date_font, date_font_size)
    c_.drawCentredString(0, 0, f"{start_date} - {end_date}")



# Use the above function and save all path names for merging at the end
for curr_name in beginner_names:
    c = canvas.Canvas(f"Intermediate_Files/{curr_name}.pdf", pagesize=A4)
    files_to_merge.append(f"Intermediate_Files/output-{curr_name}.pdf")
    width, height = A4
    create_name_pdf(c, curr_name)
    c.showPage()
    c.save()
for curr_name in advanced_names:
    c = canvas.Canvas(f"Intermediate_Files/{curr_name}.pdf", pagesize=A4)
    files_to_merge.append(f"Intermediate_Files/output-{curr_name}.pdf")
    width, height = A4
    create_name_pdf(c, curr_name)
    c.showPage()
    c.save()


# Merge to create each unique diploma
def write_name_on_diploma(template, output, name_pdf):
    name_obj = PdfFileReader(name_pdf)
    name_page = name_obj.getPage(0)

    pdf_reader = PdfFileReader(template)
    pdf_writer = PdfFileWriter()

    for page in range(pdf_reader.getNumPages()):
        page = pdf_reader.getPage(page)
        page.mergePage(name_page)
        pdf_writer.addPage(page)

    with open(output, "wb") as out:
        pdf_writer.write(out)


for each in beginner_names:
    write_name_on_diploma(beginner_template, f"Intermediate_Files/output-{each}.pdf",
                          f"Intermediate_Files/{each}.pdf")
for each in advanced_names:
    write_name_on_diploma(advanced_template, f"Intermediate_Files/output-{each}.pdf",
                          f"Intermediate_Files/{each}.pdf")


# Combine all into a final file
def merge_diplomas(paths, output):
    pdf_writer = PdfFileWriter()

    for path in paths:
        pdf_reader = PdfFileReader(path)
        pdf_writer.addPage(pdf_reader.getPage(0))

    with open(output, "wb") as out:
        pdf_writer.write(out)


merge_diplomas(files_to_merge, "final_output.pdf")


# Single screen interface
# File chooser button for names.txt and diploma.pdf
# Drop down for name/date font/size
# Input for positioning text within the document
# Possible input for changing document size


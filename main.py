# This program is used to put a list of names onto a diploma template.
# It should be run from within the current working directory

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PyPDF4 import PdfFileWriter, PdfFileReader

files_to_merge = []
names_file = "names.txt"
template_file = "template.pdf"

# Create a python list of names from names.txt
with open(names_file, "r") as f:
    name_list = f.readlines()
    name_list = [name.strip() for name in name_list]


# Function used to create separate pdf with only a name on it
def create_name_pdf(c_, name):
    c_.translate(height / 2, width * 0.73)
    c_.setFillColorRGB(255, 255, 255)
    c_.rect(-width * 0.2, -10, height * 0.4, 25, stroke=0, fill=1)
    c_.setFont("Helvetica-Bold", 15)
    c_.setFillColorRGB(0, 0, 0)
    c_.drawCentredString(0, 0, name)


# USe the above function and save all path names for merging at the end
for curr_name in name_list:
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


for each in name_list:
    write_name_on_diploma(template_file, f"Intermediate_Files/output-{each}.pdf",
                          f"Intermediate_Files/{each}.pdf")


# Combine all into a final file
def merge_diplomas(paths, output):
    pdf_writer = PdfFileWriter()

    for path in paths:
        pdf_reader = PdfFileReader(path)
        pdf_writer.addPage(pdf_reader.getPage(0))

    with open(output, "wb") as out:
        pdf_writer.write(out)


merge_diplomas(files_to_merge, "final_output_with_names.pdf")



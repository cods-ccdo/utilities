# CODS Badge generation from google spreadsheet
# pip install pandas PyPDF2 reportlab


import pandas as pd
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
import qrcode
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# make sure to only use the google spreadsheet link before the /edit
SPREADSHEET_URL = 'GOOGLE_SPREADSHEET_LINK' + '/gviz/tq?tqx=out:csv'  # Append to get CSV export

TEMPLATE_PDF_PATH = 'template.pdf'
OUTPUT_PDF_PATH = 'output.pdf'

# Read data from the Google Sheet. If there is no header in the first row use , header=None
data = pd.read_csv(SPREADSHEET_URL)
#print(data)


font_name_bold = "BC Sans Bold"
font_path_bold = "/Users/mschwanzer/Library/Fonts/2023_01_01_BCSans-Bold_2f.ttf"  # Replace with the actual path you found
pdfmetrics.registerFont(TTFont(font_name_bold, font_path_bold))

font_name_regular = "BC Sans Regular"
font_path_regular = "/Users/mschwanzer/Library/Fonts/2023_01_01_BCSans-Regular_2f.ttf"  # Replace with the actual path you found
pdfmetrics.registerFont(TTFont(font_name_regular, font_path_regular))


with open(TEMPLATE_PDF_PATH, 'rb') as file:
    reader = PdfReader(file)
    page = reader.pages[0]
    page_width, page_height =  float(page.mediabox.width), float(page.mediabox.height)
print(f"Template Width: {page_width}, Height: {page_height}")

overlay_file_path = 'overlay.pdf'
c = canvas.Canvas(overlay_file_path, pagesize=(page_width, page_height))
def draw_centered_string(c, text, x, y):
    text_width = c.stringWidth(text, 'Helvetica', 18)  # Assuming font Helvetica size 14; adjust if different
    new_x = (page_width - text_width) / 2
    c.drawString(new_x, y, text)
    return y - 20  # Adjust this value based on the space you want between lines

def generate_qr(link):
    img = qrcode.make(link)
    img_file = "temp_qr.png"
    img.save(img_file)
    return img_file


# Open the template PDF
with open(TEMPLATE_PDF_PATH, 'rb') as template_file:
    
    output = PdfWriter()
    print("ready to iterate over lines")
    # Iterate over rows in spreadsheet and overlay information
    for _, row in data.iterrows():
        template = PdfReader(template_file)
        firstname, lastname, company, link = row['firstname'], row['lastname'], row['company'], str(row['link'])
        print(f"f: {firstname}, l: {lastname}, c: {company}, l: {link}")
        # Create overlay with data
        overlay_file_path = 'overlay.pdf'
        c = canvas.Canvas(overlay_file_path)
        
        c.setFillColorRGB(0, 0, 0)  # Set color to black

        # Drawing the name
        text = f"{firstname} {lastname}"
        c.setFont("BC Sans Bold", 17)  # name BC Sans (Bold), size 17, organisation BC Sans (Regular), size 14
        text_width = c.stringWidth(text, 'BC Sans Bold', 17) 
        new_x = (page_width - text_width) / 2
        c.drawString(new_x, page_height / 2, text)

        # Drawing the company
        c.setFont("BC Sans Regular", 14)  # name BC Sans (Bold), size 17, organisation BC Sans (Regular), size 14
        text = f"{company}"
        c.setFont("BC Sans Regular", 14)  # name BC Sans (Bold), size 17, organisation BC Sans (Regular), size 14
        text_width = c.stringWidth(text, 'BC Sans Regular', 14) 
        new_x = (page_width - text_width) / 2
        c.drawString(new_x, page_height / 2 - 20, text)
        
        if "http" in link: 
            qr_file_path = generate_qr(link)
            #print(f"width = {page_width}, height = {page_height}")
            c.drawImage(qr_file_path, page_width - 61.5, 58, width=27, height=27)  # Adjust coordinates and size as needed

        c.save()

        # Merge template with overlay
        with open(overlay_file_path, 'rb') as overlay_file:
            overlay = PdfReader(overlay_file)
            page = template.pages[0]
            page.merge_page(overlay.pages[0])
            output.add_page(page)

    # Save the final combined output
    with open(OUTPUT_PDF_PATH, 'wb') as output_file:
        output.write(output_file)


# CODS Badge generation from google spreadsheet
# pip install pandas PyPDF2 reportlab


import pandas as pd
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
import qrcode


# Constants
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1of7DdSDiBxhlRAuop_zM-_6YFYCw53dIO4lf5x4XSME' + '/gviz/tq?tqx=out:csv'  # Append to get CSV export
TEMPLATE_PDF_PATH = 'template.pdf'
OUTPUT_PDF_PATH = 'output.pdf'

# Read data from the Google Sheet. If there is no header in the first row use , header=None
data = pd.read_csv(SPREADSHEET_URL)
print(data)


# Create overlay with data
overlay_file_path = 'overlay.pdf'
c = canvas.Canvas(overlay_file_path, pagesize=letter)

page_width, page_height = letter  # Assuming you're using letter-sized paper. Adjust as needed.


def draw_centered_string(c, text, x, y):
    text_width = c.stringWidth(text, 'Helvetica', 14)  # Assuming font Helvetica size 14; adjust if different
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
        y = page_height / 2  # Starting y-coordinate
        x = 300
        template = PdfReader(template_file)
        firstname, lastname, pronouns, role, company, link = row['firstname'], row['lastname'], row['pronouns'], row['role'], row['company'], row['link']
        print(f"f: {firstname}, l: {lastname}")
        # Create overlay with data
        overlay_file_path = 'overlay.pdf'
        c = canvas.Canvas(overlay_file_path)
        c.setFont("Helvetica", 20)  # Set font to Helvetica and size to 14
        c.setFillColorRGB(0, 0, 0)  # Set color to black

        # Drawing the name and pronouns
        y = draw_centered_string(c, f"{firstname} {lastname} ({pronouns})", x, y)

        # Drawing the role
        y = draw_centered_string(c, role, x, y)

        # Drawing the company
        y = draw_centered_string(c, company, x, y)
        
        qr_file_path = generate_qr(link)
        c.drawImage(qr_file_path, (page_width - 100) / 2, y - 100, width=100, height=100)  # Adjust coordinates and size as needed


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


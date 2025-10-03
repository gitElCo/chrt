from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO

def add_watermark(pdf_path, text, position="center", opacity=0.3, angle=45):
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)
    c.setFont("Helvetica", 30)
    c.setFillAlpha(opacity)
    c.rotate(angle)

    if position == "center":
        x, y = 300, 400
    elif position == "top-left":
        x, y = 100, 700
    elif position == "top-right":
        x, y = 500, 700
    elif position == "bottom-left":
        x, y = 100, 100
    elif position == "bottom-right":
        x, y = 500, 100

    c.drawString(x, y, text)
    c.save()

    packet.seek(0)
    watermark = PdfReader(packet)
    watermark_page = watermark.pages[0]

    existing_pdf = PdfReader(open(pdf_path, "rb"))
    output = PdfWriter()

    for page in existing_pdf.pages:
        page.merge_page(watermark_page)
        output.add_page(page)

    with open(pdf_path, "wb") as output_stream:
        output.write(output_stream)

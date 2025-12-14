from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO

def add_footer_to_pdf(input_path, output_path, footer_text):
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for page in reader.pages:
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        # Create footer overlay
        packet = BytesIO()
        c = canvas.Canvas(packet, pagesize=(page_width, page_height))
        c.setFont("Helvetica", 8)
        text_width = c.stringWidth(footer_text, "Helvetica", 8)
        x = (page_width - text_width) / 2
        c.drawString(x, 40, footer_text)
        c.save()

        # Merge footer with page
        packet.seek(0)
        overlay = PdfReader(packet)
        page.merge_page(overlay.pages[0])
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)

if __name__ == "__main__":
    footer = "Corresponding Author: Thomas W Buford, Email: thomas_buford@baylor.edu"
    add_footer_to_pdf("Sample.pdf", "Sample_with_footer.pdf", footer)
    print("Footer added successfully!")

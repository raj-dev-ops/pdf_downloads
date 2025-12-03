import sys
from pathlib import Path
import pandas as pd
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
import re

def normalize_title(title):
    """Normalize title for matching"""
    if pd.isna(title):
        return ""
    title = str(title).lower()
    title = re.sub(r'[^\w\s]', '', title)
    title = re.sub(r'\s+', ' ', title)
    return title.strip()

def match_pdf_to_author(pdf_file, df):
    """Match PDF filename to Excel entry"""
    pdf_title = normalize_title(pdf_file.stem)
    
    # Try exact match first
    for idx, row in df.iterrows():
        excel_title = normalize_title(row['title'])
        if excel_title == pdf_title:
            return row['Corresponding_Author']
    
    # Try fuzzy match (20+ character overlap)
    for idx, row in df.iterrows():
        excel_title = normalize_title(row['title'])
        if len(excel_title) > 20 and len(pdf_title) > 20:
            if excel_title[:50] in pdf_title or pdf_title[:50] in excel_title:
                return row['Corresponding_Author']
    
    return None

def add_footer_to_pdf(pdf_path, footer_text, output_path=None):
    """Add footer to first page of PDF"""
    if output_path is None:
        output_path = pdf_path
    
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    
    # Get first page
    first_page = reader.pages[0]
    page_width = float(first_page.mediabox.width)
    page_height = float(first_page.mediabox.height)
    
    # Create footer overlay
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    # Footer positioning
    footer_x = 72   # 1 inch margin from left
    footer_y = 55   # Changed from 30 to prevent cutoff

    # Add white background rectangle behind footer text
    can.setFillColorRGB(1, 1, 1)
    can.rect(footer_x - 5, footer_y - 5, 500, 20, fill=1, stroke=0)

    # Add footer text
    can.setFillColorRGB(0, 0, 0)
    can.setFont("Times-Roman", 9)
    can.drawString(footer_x, footer_y, footer_text)
    
    can.save()
    packet.seek(0)
    
    # Merge overlay with first page
    overlay = PdfReader(packet)
    first_page.merge_page(overlay.pages[0])
    writer.add_page(first_page)
    
    # Add remaining pages unchanged
    for i in range(1, len(reader.pages)):
        writer.add_page(reader.pages[i])
    
    # Save
    tmp_path = pdf_path.with_suffix('.tmp.pdf')
    with open(tmp_path, 'wb') as f:
        writer.write(f)
    
    tmp_path.replace(output_path)

def process_folder(folder_path, excel_path):
    """Process all PDFs in folder"""
    folder = Path(folder_path)
    
    # Find all PDFs
    pdf_files = sorted(folder.glob('*.pdf'))
    print(f"Found {len(pdf_files)} PDF files")
    
    if not pdf_files:
        print("No PDF files found!")
        return
    
    # Load Excel - handle the correct column
    df = pd.read_excel(excel_path)
    
    # Verify required columns exist
    if 'title' not in df.columns:
        print(f"❌ Error: Excel must have 'title' column")
        print(f"   Available columns: {list(df.columns)}")
        return
    
    if 'Corresponding_Author' not in df.columns:
        print(f"❌ Error: Excel must have 'Corresponding_Author' column")
        print(f"   Available columns: {list(df.columns)}")
        return
    
    print(f"✅ Loaded {len(df)} records from Excel")
    print(f"   Using columns: 'title' and 'Corresponding_Author'")
    
    # Remove rows with empty Corresponding_Author
    df = df[df['Corresponding_Author'].notna()]
    print(f"   Records with author data: {len(df)}")
    
    # Process each PDF
    processed = 0
    skipped = 0
    
    for pdf_file in pdf_files:
        print(f"\nProcessing: {pdf_file.name}")
        
        # Match to Excel
        author_info = match_pdf_to_author(pdf_file, df)
        
        if author_info and pd.notna(author_info):
            print(f"   ✓ Matched → {author_info[:60]}...")
            try:
                add_footer_to_pdf(pdf_file, str(author_info))
                processed += 1
                print(f"   ✓ Footer added successfully")
            except Exception as e:
                print(f"   ❌ Error: {e}")
                skipped += 1
        else:
            print(f"   ⚠️  No matching author data found - skipped")
            skipped += 1
    
    print(f"\n{'='*60}")
    print(f"SUMMARY:")
    print(f"  Total PDFs: {len(pdf_files)}")
    print(f"  Processed: {processed}")
    print(f"  Skipped: {skipped}")
    print(f"{'='*60}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python add_author_footer.py <pdf_folder> <excel_file>")
        print("Example: python add_author_footer.py ./pdfs/ authors.xlsx")
        sys.exit(1)
    
    folder_path = sys.argv[1]
    excel_path = sys.argv[2]
    
    if not Path(folder_path).exists():
        print(f"Error: Folder not found: {folder_path}")
        sys.exit(1)
    
    if not Path(excel_path).exists():
        print(f"Error: Excel file not found: {excel_path}")
        sys.exit(1)
    
    process_folder(folder_path, excel_path)


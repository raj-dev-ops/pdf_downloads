import sys
from pathlib import Path
import pandas as pd
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import re
import warnings

# Register Palatino Regular font (Book Antiqua Linotype equivalent)
try:
    # On macOS, Palatino is the equivalent of Book Antiqua
    # subfontIndex=0 is Regular variant
    pdfmetrics.registerFont(TTFont('Palatino-Regular', '/System/Library/Fonts/Palatino.ttc', subfontIndex=0))
    FOOTER_FONT = 'Palatino-Regular'
except:
    # Fallback to Times-Roman if Palatino not available
    FOOTER_FONT = 'Times-Roman'
    print("⚠️  Warning: Palatino font not found, using Times-Roman instead")

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
    
    # Try fuzzy match (50+ character overlap)
    for idx, row in df.iterrows():
        excel_title = normalize_title(row['title'])
        if len(excel_title) > 50 and len(pdf_title) > 50:
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
    # Footer placed at 20mm from bottom
    footer_x = 53.86  # 19mm margin from left (19mm * 72/25.4)
    footer_y = 56.69  # 20mm from bottom (20mm * 72/25.4)

    # Calculate text width to fit background rectangle (add space for superscript *)
    text_width = can.stringWidth(footer_text, FOOTER_FONT, 10) + 8  # Extra space for superscript

    # Add white background rectangle behind footer text (sized to fit text)
    can.setFillColorRGB(1, 1, 1)
    can.rect(footer_x - 5, footer_y - 3, text_width + 10, 16, fill=1, stroke=0)

    # Add footer text (using Palatino Regular - Book Antiqua Linotype equivalent)
    # Color: #943634 (RGB: 148/255, 54/255, 52/255)
    can.setFillColorRGB(148/255, 54/255, 52/255)
    can.setFont(FOOTER_FONT, 10)

    # Draw asterisk as superscript before "Corresponding"
    if "Corresponding" in footer_text:
        # Draw superscript asterisk
        can.setFont(FOOTER_FONT, 7)  # Smaller font for superscript
        can.drawString(footer_x, footer_y + 3, "*")  # Raised position

        # Draw main text starting after the asterisk
        can.setFont(FOOTER_FONT, 10)  # Back to normal size
        asterisk_width = can.stringWidth("*", FOOTER_FONT, 7)
        can.drawString(footer_x + asterisk_width, footer_y, footer_text)
    else:
        # No "Corresponding" word, draw text normally
        can.drawString(footer_x, footer_y, footer_text)
    
    can.save()
    packet.seek(0)

    # Merge overlay with first page
    # Suppress font warnings that cause merge failures
    overlay = PdfReader(packet)

    try:
        # Try standard merge with warning suppression
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore")
            first_page.merge_page(overlay.pages[0])
    except Exception as e:
        # If merge fails due to font errors, try alternative approach
        # Create a fresh page copy and merge on that
        print(f"      Standard merge failed ({str(e)[:50]}...), trying alternative method")
        try:
            from pypdf.generic import Transformation
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore")
                first_page.merge_transformed_page(
                    overlay.pages[0],
                    Transformation(),
                    expand=False
                )
        except Exception as e2:
            print(f"      Alternative merge also failed: {str(e2)[:50]}...")
            raise Exception(f"All merge methods failed. Font conflict cannot be resolved.")

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
    if len(sys.argv) < 3:
        print("Usage: python add_author_footer.py <pdf_folder> <excel_file>")
        print('Example: python add_author_footer.py ./pdfs/ authors.xlsx')
        print('Note: Use quotes for paths with spaces on Windows:')
        print('  python add_author_footer.py "C:\\My Folder" "data.xlsx"')
        sys.exit(1)

    # Handle paths with spaces - find the Excel file (ends with .xlsx or .xls)
    all_args = sys.argv[1:]

    # Join all arguments and try to find xlsx/xls file
    full_path = ' '.join(all_args)

    # Try to split into folder and Excel file
    folder_path = None
    excel_path = None

    # Method 1: If only 2 args, use them directly
    if len(sys.argv) == 3:
        folder_path = sys.argv[1]
        excel_path = sys.argv[2]
    else:
        # Method 2: Find .xlsx or .xls in the joined path
        import re
        # Look for pattern: path ending with .xlsx or .xls
        match = re.search(r'(.+?)\s+(.+\.xlsx?)$', full_path, re.IGNORECASE)
        if match:
            folder_path = match.group(1)
            excel_path = match.group(2)
        else:
            # Method 3: Try to find existing files/folders
            for i in range(1, len(all_args)):
                potential_folder = ' '.join(all_args[:i])
                potential_excel = ' '.join(all_args[i:])

                if Path(potential_folder).exists() and potential_excel.lower().endswith(('.xlsx', '.xls')):
                    folder_path = potential_folder
                    excel_path = potential_excel
                    break

    if not folder_path or not excel_path:
        print("❌ Error: Could not parse folder and Excel file paths")
        print(f"   Received arguments: {sys.argv[1:]}")
        print("   Please use quotes for paths with spaces:")
        print('   python add_author_footer.py "C:\\Path With Spaces" "file.xlsx"')
        sys.exit(1)

    if not Path(folder_path).exists():
        print(f"❌ Error: Folder not found: {folder_path}")
        sys.exit(1)

    if not Path(excel_path).exists():
        print(f"❌ Error: Excel file not found: {excel_path}")
        sys.exit(1)

    print(f"📁 PDF Folder: {folder_path}")
    print(f"📊 Excel File: {excel_path}")
    print()

    process_folder(folder_path, excel_path)


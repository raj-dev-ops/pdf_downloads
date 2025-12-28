#!/usr/bin/env python3
"""
Check w:pict elements in detail
"""

import sys
from docx import Document
from lxml import etree

def check_pict(docx_path):
    """Check w:pict elements"""

    doc = Document(docx_path)

    print(f"Analyzing w:pict elements in: {docx_path}\n")

    # Find all w:pict elements
    for para_idx, para in enumerate(doc.paragraphs):
        picts = para._element.xpath('.//w:pict')
        if picts:
            print(f"Paragraph {para_idx}: {len(picts)} w:pict element(s)")
            for pict_idx, pict in enumerate(picts):
                print(f"\n  w:pict #{pict_idx + 1}:")
                # Print the raw XML of this pict element
                xml_str = etree.tostring(pict, pretty_print=True, encoding='unicode')
                print(xml_str[:1000])  # First 1000 chars


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python check_pict.py <path_to_docx>")
        sys.exit(1)

    check_pict(sys.argv[1])

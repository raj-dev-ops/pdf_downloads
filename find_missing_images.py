#!/usr/bin/env python3
"""
Find which images are being extracted and which are missing
"""

import sys
from docx import Document
from io import BytesIO
from PIL import Image

def analyze_extraction(docx_path):
    """Analyze which images are being extracted"""

    doc = Document(docx_path)

    # Get all image relationships
    print("All image relationships in document:")
    print("=" * 70)
    for rel_id, rel in doc.part.rels.items():
        if "image" in rel.target_ref:
            print(f"  {rel_id} -> {rel.target_ref}")

    print("\n" + "=" * 70)
    print("Attempting to extract each image:")
    print("=" * 70)

    extracted_count = 0
    failed_count = 0

    for rel_id, rel in doc.part.rels.items():
        if "image" in rel.target_ref:
            try:
                image_part = doc.part.related_parts[rel_id]
                image_bytes = image_part.blob

                # Try to open with PIL
                try:
                    img = Image.open(BytesIO(image_bytes))
                    print(f"✓ {rel_id} ({rel.target_ref}): {img.format} {img.size} - OK")
                    extracted_count += 1
                except Exception as e:
                    print(f"✗ {rel_id} ({rel.target_ref}): Cannot open with PIL - {e}")
                    failed_count += 1

            except Exception as e:
                print(f"✗ {rel_id} ({rel.target_ref}): Cannot extract - {e}")
                failed_count += 1

    print("\n" + "=" * 70)
    print(f"Summary: {extracted_count} images can be opened, {failed_count} failed")
    print("=" * 70)

    # Now check which rel_ids are actually referenced in the document
    print("\nChecking which images are referenced in document paragraphs:")
    print("=" * 70)

    referenced_ids = set()

    for para_idx, paragraph in enumerate(doc.paragraphs):
        # Check modern format
        for run in paragraph.runs:
            if run._element.xpath('.//pic:pic'):
                inline_shapes = run._element.xpath('.//a:blip/@r:embed')
                for rel_id in inline_shapes:
                    referenced_ids.add(rel_id)
                    print(f"  Para {para_idx}: Found modern format image {rel_id}")

        # Check VML format
        pict_elements = paragraph._element.xpath('.//w:pict')
        if pict_elements:
            for pict in pict_elements:
                all_elements = pict.xpath('.//*')
                for elem in all_elements:
                    for attr_name, attr_value in elem.attrib.items():
                        if attr_value and isinstance(attr_value, str) and attr_value.startswith('rId'):
                            referenced_ids.add(attr_value)
                            print(f"  Para {para_idx}: Found VML format image {attr_value}")

    print("\n" + "=" * 70)
    print(f"Total unique referenced images: {len(referenced_ids)}")
    print("=" * 70)

    # Find unreferenced images
    all_image_ids = set()
    for rel_id, rel in doc.part.rels.items():
        if "image" in rel.target_ref:
            all_image_ids.add(rel_id)

    unreferenced = all_image_ids - referenced_ids
    if unreferenced:
        print("\nUnreferenced images (not found in paragraphs):")
        for rel_id in unreferenced:
            rel = doc.part.rels[rel_id]
            print(f"  {rel_id} -> {rel.target_ref}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python find_missing_images.py <path_to_docx>")
        sys.exit(1)

    analyze_extraction(sys.argv[1])

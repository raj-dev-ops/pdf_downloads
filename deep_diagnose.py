#!/usr/bin/env python3
"""
Deep diagnostic to find exactly where images are referenced
"""

import sys
from docx import Document
from docx.oxml.ns import qn
import zipfile
from lxml import etree

def deep_diagnose(docx_path):
    """Deep dive into document structure"""

    doc = Document(docx_path)

    print(f"Deep analysis of: {docx_path}\n")
    print("=" * 70)

    # Get all image relationships
    image_rels = {}
    for rel_id, rel in doc.part.rels.items():
        if "image" in rel.target_ref:
            image_rels[rel_id] = rel.target_ref

    print(f"Found {len(image_rels)} image relationships:")
    for rel_id, target in image_rels.items():
        print(f"  {rel_id} -> {target}")

    # Now search the entire document XML for references to these rel IDs
    print("\n" + "=" * 70)
    print("Searching for image references in document body:")
    print("=" * 70)

    # Get the raw XML of the document
    doc_xml = doc.element
    xml_string = etree.tostring(doc_xml, pretty_print=True, encoding='unicode')

    # Search for each relationship ID
    for rel_id in image_rels.keys():
        if rel_id in xml_string:
            print(f"\n{rel_id} IS referenced in document XML")

            # Find the context
            # Search for r:embed or r:link attributes using qualified names
            r_embed = qn('r:embed')
            r_link = qn('r:link')
            embeds = doc_xml.xpath(f'.//*[@{r_embed}="{rel_id}"]')
            links = doc_xml.xpath(f'.//*[@{r_link}="{rel_id}"]')

            if embeds:
                print(f"  Found in {len(embeds)} embed element(s)")
                for elem in embeds:
                    print(f"    Tag: {elem.tag}")
                    print(f"    Parent: {elem.getparent().tag if elem.getparent() is not None else 'None'}")
                    # Print the element path
                    ancestors = []
                    parent = elem.getparent()
                    while parent is not None:
                        tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
                        ancestors.append(tag)
                        parent = parent.getparent()
                    print(f"    Path: {' -> '.join(reversed(ancestors[-5:]))}")

            if links:
                print(f"  Found in {len(links)} link element(s)")

        else:
            print(f"\n{rel_id} NOT referenced in document XML (orphaned image)")

    # Also check if there are any "object" or "pict" elements
    print("\n" + "=" * 70)
    print("Checking for alternative image containers:")
    print("=" * 70)

    # Check for v:imagedata (VML images - old format)
    vml_images = doc_xml.xpath('.//v:imagedata')
    if vml_images:
        print(f"\nFound {len(vml_images)} VML imagedata elements:")
        for img in vml_images:
            rel_id = img.get(qn('r:id')) or img.get(qn('r:embed')) or img.get(qn('o:relid'))
            print(f"  VML image with rel_id: {rel_id}")

    # Check for w:pict elements (Picture elements)
    pict_elements = doc_xml.xpath('.//w:pict')
    if pict_elements:
        print(f"\nFound {len(pict_elements)} w:pict (picture) elements")

    # Check for w:object elements (embedded objects)
    object_elements = doc_xml.xpath('.//w:object')
    if object_elements:
        print(f"\nFound {len(object_elements)} w:object (embedded object) elements")

    # Check for mc:AlternateContent (compatibility mode content)
    alt_content = doc_xml.xpath('.//mc:AlternateContent')
    if alt_content:
        print(f"\nFound {len(alt_content)} mc:AlternateContent elements (compatibility mode)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python deep_diagnose.py <path_to_docx>")
        sys.exit(1)

    deep_diagnose(sys.argv[1])

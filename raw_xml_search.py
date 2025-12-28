#!/usr/bin/env python3
"""
Search raw XML for image references
"""

import sys
from docx import Document
from lxml import etree

def search_xml_for_rel_id(docx_path, target_rel_ids):
    """Search the entire document XML for relationship IDs"""

    doc = Document(docx_path)

    print("Searching raw XML for relationship IDs:")
    print("=" * 70)

    # Get the raw XML
    doc_xml = doc.element
    xml_string = etree.tostring(doc_xml, pretty_print=True, encoding='unicode')

    for rel_id in target_rel_ids:
        print(f"\nSearching for {rel_id}:")

        if rel_id in xml_string:
            print(f"  ✓ Found in document XML")

            # Find the lines containing this rel_id
            lines = xml_string.split('\n')
            matching_lines = [i for i, line in enumerate(lines) if rel_id in line]

            print(f"  Found on {len(matching_lines)} line(s):")
            for line_num in matching_lines[:5]:  # Show first 5 matches
                # Show context (5 lines before and after)
                start = max(0, line_num - 2)
                end = min(len(lines), line_num + 3)
                print(f"\n  Context around line {line_num}:")
                for i in range(start, end):
                    marker = ">>>" if i == line_num else "   "
                    print(f"  {marker} {lines[i][:120]}")

        else:
            print(f"  ✗ NOT found in document XML")

    # Also try to find all elements with r:id, r:embed, or o:relid attributes
    print("\n" + "=" * 70)
    print("All image-related attributes in document:")
    print("=" * 70)

    # Search for all elements that have relationship attributes
    from docx.oxml.ns import qn

    r_id = qn('r:id')
    r_embed = qn('r:embed')
    o_relid = qn('o:relid')

    for attr in [r_id, r_embed, o_relid]:
        try:
            elements = doc_xml.xpath(f'.//*[@{attr}]')
            if elements:
                print(f"\nFound {len(elements)} elements with {attr} attribute:")
                for elem in elements:
                    rel_id_val = elem.get(attr)
                    if rel_id_val in target_rel_ids:
                        print(f"  >>> {elem.tag}: {rel_id_val} <<<< TARGET FOUND")
                        # Print parent info
                        parent = elem.getparent()
                        if parent is not None:
                            print(f"      Parent: {parent.tag}")
                            grandparent = parent.getparent()
                            if grandparent is not None:
                                print(f"      Grandparent: {grandparent.tag}")
        except Exception as e:
            print(f"Error searching for {attr}: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python raw_xml_search.py <path_to_docx>")
        sys.exit(1)

    search_xml_for_rel_id(sys.argv[1], ['rId20', 'rId7'])

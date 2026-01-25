#!/usr/bin/env python3
"""
Script to change volume numbers in PDF files.
Example: "Vol. 45" -> "Vol. 47"
"""

import os
import re
import sys
from pathlib import Path
import fitz  # PyMuPDF
from typing import Optional, Tuple


def change_volume_number(
    pdf_path: str,
    old_volume: int,
    new_volume: int,
    output_path: Optional[str] = None,
    backup: bool = True
) -> bool:
    """
    Change volume number in a PDF file while preserving exact font formatting.

    Args:
        pdf_path: Path to the PDF file
        old_volume: Current volume number to replace
        new_volume: New volume number
        output_path: Path for the modified PDF (if None, overwrites original)
        backup: Whether to create a backup of the original file

    Returns:
        True if successful, False otherwise
    """
    try:
        import shutil

        # Open the PDF
        doc = fitz.open(pdf_path)
        replacements_made = 0

        # Iterate through all pages
        for page_num in range(len(doc)):
            page = doc[page_num]

            # Get all text with detailed formatting information
            text_dict = page.get_text("dict")

            # Store areas to redact and their replacement text
            redactions = []

            # Iterate through all text blocks
            for block in text_dict.get("blocks", []):
                if block.get("type") != 0:  # Skip non-text blocks
                    continue

                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        span_text = span.get("text", "")

                        # Check if this span contains the volume number
                        if f"Vol. {old_volume}" in span_text or f"Vol.{old_volume}" in span_text or \
                           f"Volume {old_volume}" in span_text or f"vol. {old_volume}" in span_text:

                            # Get exact position and formatting
                            bbox = fitz.Rect(span["bbox"])
                            font_name = span.get("font", "Times-Italic")
                            font_size = span.get("size", 10.0)
                            color = span.get("color", 0)  # 0 = black
                            flags = span.get("flags", 0)

                            # Replace volume number in text
                            new_text = span_text.replace(
                                f"Vol. {old_volume}", f"Vol. {new_volume}"
                            ).replace(
                                f"Vol.{old_volume}", f"Vol.{new_volume}"
                            ).replace(
                                f"Volume {old_volume}", f"Volume {new_volume}"
                            ).replace(
                                f"vol. {old_volume}", f"vol. {new_volume}"
                            ).replace(
                                f"vol.{old_volume}", f"vol.{new_volume}"
                            )

                            # Only process if text actually changed
                            if new_text != span_text:
                                redactions.append({
                                    'bbox': bbox,
                                    'text': new_text,
                                    'font': font_name,
                                    'fontsize': font_size,
                                    'color': color,
                                    'flags': flags
                                })
                                replacements_made += 1

            # Apply redactions
            for redaction in redactions:
                # First, add a white rectangle to cover the old text
                page.add_redact_annot(
                    redaction['bbox'],
                    fill=(1, 1, 1)  # White fill
                )

            # Apply the redactions (this removes the old text)
            page.apply_redactions()

            # Now add the new text with preserved formatting
            for redaction in redactions:
                # Use Times-Italic as the font (standard PDF font name for Times New Roman Italic)
                font_to_use = "Times-Italic"

                # Insert new text at the same position
                rc = page.insert_text(
                    redaction['bbox'].tl + fitz.Point(0, redaction['fontsize'] * 0.8),  # Adjust baseline
                    redaction['text'],
                    fontname=font_to_use,
                    fontsize=redaction['fontsize'],
                    color=redaction['color'] if isinstance(redaction['color'], tuple) else (0, 0, 0),
                    render_mode=0  # Fill text
                )

        # Save if changes were made
        if replacements_made > 0:
            # Determine output path
            if output_path is None:
                # Save to temp file first, then replace original
                temp_path = pdf_path.replace('.pdf', '_temp.pdf')

                if backup:
                    backup_path = pdf_path.replace('.pdf', '_backup.pdf')
                    shutil.copy2(pdf_path, backup_path)
                    print(f"  Backup created: {os.path.basename(backup_path)}")

                # Save to temp file
                doc.save(temp_path, garbage=4, deflate=True)
                doc.close()

                # Replace original with temp
                shutil.move(temp_path, pdf_path)
            else:
                # Save to different location
                doc.save(output_path, garbage=4, deflate=True)
                doc.close()

            print(f"  ✓ Changed {replacements_made} instance(s) of 'Vol. {old_volume}' to 'Vol. {new_volume}'")
            return True
        else:
            print(f"  ℹ No instances of 'Vol. {old_volume}' found")
            doc.close()
            return False

    except Exception as e:
        print(f"  ✗ Error processing {pdf_path}: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def process_folder(
    folder_path: str,
    old_volume: int,
    new_volume: int,
    output_folder: Optional[str] = None,
    backup: bool = True
) -> Tuple[int, int]:
    """
    Process all PDF files in a folder.

    Args:
        folder_path: Path to folder containing PDFs
        old_volume: Current volume number to replace
        new_volume: New volume number
        output_folder: Folder for modified PDFs (if None, overwrites originals)
        backup: Whether to create backups

    Returns:
        Tuple of (successful_count, total_count)
    """
    folder = Path(folder_path)

    if not folder.exists() or not folder.is_dir():
        print(f"Error: '{folder_path}' is not a valid directory")
        return 0, 0

    # Find all PDF files
    pdf_files = list(folder.glob("*.pdf"))

    if not pdf_files:
        print(f"No PDF files found in '{folder_path}'")
        return 0, 0

    print(f"Found {len(pdf_files)} PDF file(s)")
    print(f"Changing 'Vol. {old_volume}' to 'Vol. {new_volume}'")
    print("-" * 60)

    successful = 0

    for pdf_file in pdf_files:
        print(f"\nProcessing: {pdf_file.name}")

        output_path = None
        if output_folder:
            output_folder_path = Path(output_folder)
            output_folder_path.mkdir(parents=True, exist_ok=True)
            output_path = str(output_folder_path / pdf_file.name)

        if change_volume_number(
            str(pdf_file),
            old_volume,
            new_volume,
            output_path,
            backup
        ):
            successful += 1

    print("\n" + "=" * 60)
    print(f"Completed: {successful}/{len(pdf_files)} files successfully processed")

    return successful, len(pdf_files)


def main():
    """Main entry point for the script."""
    print("=" * 60)
    print("PDF Volume Number Changer (Vol. 45 -> Vol. 47)")
    print("=" * 60)

    # Get folder path
    if len(sys.argv) > 1:
        folder_path = sys.argv[1]
    else:
        folder_path = input("\nEnter folder path containing PDFs: ").strip()

    # Default: change Vol. 45 to Vol. 47
    old_volume = 45
    new_volume = 47

    print()  # Empty line for readability

    # Process the folder (no automatic backup)
    process_folder(folder_path, old_volume, new_volume, output_folder=None, backup=False)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Word Document Image Extractor to GIF

Extracts images from Word documents (.doc/.docx) and converts them to GIF format
with specific naming conventions based on image type (Figure or Scheme).
"""

import os
import re
import sys
import argparse
from pathlib import Path
from io import BytesIO
from typing import Tuple, List, Dict

from docx import Document
from PIL import Image
import doc2docx


def extract_document_basename(filepath: str) -> str:
    """
    Extract base name from document filename.

    Example:
        IJRIT-11-139-Figures.doc → IJRIT-11-139

    Args:
        filepath: Path to the document file

    Returns:
        Base name without suffixes like -Figures, -Images, -Schemes
    """
    filename = Path(filepath).stem  # Get filename without extension

    # Remove common suffixes
    suffixes_to_remove = ['-Figures', '-figures', '-Images', '-images',
                          '-Schemes', '-schemes', '-Tables', '-tables']

    for suffix in suffixes_to_remove:
        if filename.endswith(suffix):
            filename = filename[:-len(suffix)]
            break

    return filename


def classify_image(caption_text: str) -> str:
    """
    Classify image based on caption text.

    Args:
        caption_text: The caption or alt text from the image

    Returns:
        'g' for figures, 's' for schemes
    """
    if not caption_text:
        return 'g'  # Default to figure

    caption_lower = caption_text.lower()

    # Check for scheme indicators
    if 'scheme' in caption_lower:
        return 's'

    # Check for figure indicators (including fig, figure)
    if 'figure' in caption_lower or 'fig' in caption_lower:
        return 'g'

    # Default to figure
    return 'g'


def resize_to_width(image: Image.Image, target_width: int = 1500) -> Image.Image:
    """
    Resize image to target width while maintaining aspect ratio.

    Args:
        image: PIL Image object
        target_width: Desired width in pixels

    Returns:
        Resized PIL Image object
    """
    original_width, original_height = image.size

    # Calculate proportional height
    aspect_ratio = original_height / original_width
    target_height = int(target_width * aspect_ratio)

    # Resize image using high-quality resampling
    resized_image = image.resize((target_width, target_height), Image.Resampling.LANCZOS)

    return resized_image


def convert_doc_to_docx(doc_path: str) -> str:
    """
    Convert .doc file to .docx format.

    Args:
        doc_path: Path to .doc file

    Returns:
        Path to converted .docx file
    """
    output_path = doc_path.rsplit('.', 1)[0] + '_converted.docx'

    print(f"Converting {doc_path} to .docx format...")
    doc2docx.convert(doc_path, output_path)

    return output_path


def extract_images_from_docx(docx_path: str) -> List[Dict]:
    """
    Extract all images from a .docx file with their captions.

    Args:
        docx_path: Path to .docx file

    Returns:
        List of dictionaries containing image data and captions
    """
    doc = Document(docx_path)
    images = []

    # Track paragraph index for caption association
    for para_idx, paragraph in enumerate(doc.paragraphs):
        # Check if paragraph contains an image
        for run in paragraph.runs:
            if run._element.xpath('.//pic:pic'):
                # Found an image
                inline_shapes = run._element.xpath('.//a:blip/@r:embed')

                for rel_id in inline_shapes:
                    try:
                        image_part = doc.part.related_parts[rel_id]
                        image_bytes = image_part.blob

                        # Try to get caption from current paragraph or next paragraph
                        caption = paragraph.text.strip()

                        # If current paragraph is empty, check next paragraph
                        if not caption and para_idx + 1 < len(doc.paragraphs):
                            caption = doc.paragraphs[para_idx + 1].text.strip()

                        # Also check previous paragraph
                        if not caption and para_idx > 0:
                            caption = doc.paragraphs[para_idx - 1].text.strip()

                        images.append({
                            'bytes': image_bytes,
                            'caption': caption,
                            'format': image_part.content_type
                        })
                    except Exception as e:
                        print(f"Warning: Could not extract image: {e}")
                        continue

    # Also check for inline shapes in the document
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            try:
                image_bytes = rel.target_part.blob
                # For inline images, we might not have a direct caption
                # Try to find associated text
                caption = ""

                images.append({
                    'bytes': image_bytes,
                    'caption': caption,
                    'format': rel.target_part.content_type
                })
            except Exception as e:
                continue

    # Remove duplicates based on image bytes
    unique_images = []
    seen_bytes = set()

    for img in images:
        img_hash = hash(img['bytes'])
        if img_hash not in seen_bytes:
            seen_bytes.add(img_hash)
            unique_images.append(img)

    return unique_images


def main(input_file: str, output_dir: str = None):
    """
    Main orchestration function.

    Args:
        input_file: Path to input Word document
        output_dir: Optional output directory (defaults to same as input file)
    """
    input_path = Path(input_file)

    if not input_path.exists():
        print(f"Error: File not found: {input_file}")
        sys.exit(1)

    # Determine output directory
    if output_dir:
        output_path = Path(output_dir) / input_path.stem
        output_path.mkdir(parents=True, exist_ok=True)
    else:
        # Create extracted_gifs subfolder with filename-based subfolder
        output_path = input_path.parent / "extracted_gifs" / input_path.stem
        output_path.mkdir(parents=True, exist_ok=True)

    # Extract base name for output files
    base_name = extract_document_basename(str(input_path))

    # Handle .doc vs .docx
    docx_file = str(input_path)
    temp_file = None

    if input_path.suffix.lower() == '.doc':
        temp_file = convert_doc_to_docx(str(input_path))
        docx_file = temp_file
    elif input_path.suffix.lower() != '.docx':
        print(f"Error: Unsupported file format: {input_path.suffix}")
        sys.exit(1)

    # Extract images
    print(f"Extracting images from {input_path.name}...")
    extracted_images = extract_images_from_docx(docx_file)

    if not extracted_images:
        print("No images found in document.")
        if temp_file and os.path.exists(temp_file):
            os.remove(temp_file)
        return

    print(f"Found {len(extracted_images)} image(s)")

    # Process and save images
    figure_counter = 1
    scheme_counter = 1
    figures_created = 0
    schemes_created = 0

    for img_data in extracted_images:
        try:
            # Load image from bytes
            image = Image.open(BytesIO(img_data['bytes']))

            # Convert to RGB if necessary (for GIF compatibility)
            if image.mode not in ('RGB', 'P'):
                image = image.convert('RGB')

            # Resize to target width
            resized_image = resize_to_width(image, target_width=1500)

            # Classify image type
            img_type = classify_image(img_data['caption'])

            # Generate filename
            if img_type == 's':
                filename = f"{base_name}-s{scheme_counter:03d}.gif"
                scheme_counter += 1
                schemes_created += 1
            else:
                filename = f"{base_name}-g{figure_counter:03d}.gif"
                figure_counter += 1
                figures_created += 1

            # Save as GIF
            output_file = output_path / filename
            resized_image.save(output_file, 'GIF', optimize=True)

            width, height = resized_image.size
            print(f"Created: {filename} ({width}x{height})")

        except Exception as e:
            print(f"Error processing image: {e}")
            continue

    # Clean up temp file
    if temp_file and os.path.exists(temp_file):
        os.remove(temp_file)

    # Summary
    print(f"\nSummary: {figures_created} figure(s), {schemes_created} scheme(s) extracted")
    print(f"Output directory: {output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extract images from Word documents and convert to GIF format"
    )
    parser.add_argument(
        "input_file",
        help="Path to input Word document (.doc or .docx)"
    )
    parser.add_argument(
        "-o", "--output",
        help="Output directory (default: extracted_gifs/ in same directory as input)",
        default=None
    )

    args = parser.parse_args()
    main(args.input_file, args.output)

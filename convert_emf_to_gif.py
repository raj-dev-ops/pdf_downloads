#!/usr/bin/env python3
"""
Convert EMF file to GIF using various methods
"""

import sys
from pathlib import Path
from PIL import Image
from io import BytesIO

def convert_emf_to_gif(emf_path, output_path, target_width=1500):
    """Try various methods to convert EMF to GIF"""

    emf_file = Path(emf_path)
    output_file = Path(output_path)

    print(f"Attempting to convert: {emf_file.name}")
    print("=" * 70)

    # Method 1: Try PIL directly with different modes
    print("\nMethod 1: Direct PIL conversion...")
    try:
        with open(emf_file, 'rb') as f:
            img = Image.open(f)
            print(f"  Successfully opened: {img.format} {img.size} {img.mode}")

            # Convert to RGB
            if img.mode not in ('RGB', 'P'):
                img = img.convert('RGB')

            # Resize to target width
            original_width, original_height = img.size
            aspect_ratio = original_height / original_width
            target_height = int(target_width * aspect_ratio)
            img_resized = img.resize((target_width, target_height), Image.Resampling.LANCZOS)

            # Save as GIF
            img_resized.save(output_file, 'GIF', optimize=True)
            print(f"  ✓ Converted successfully: {output_file.name} ({target_width}x{target_height})")
            return True

    except Exception as e:
        print(f"  ✗ Failed: {e}")

    # Method 2: Try loading as WMF explicitly
    print("\nMethod 2: Force WMF/EMF handler...")
    try:
        # Register WMF handler if not already registered
        from PIL import WmfImagePlugin
        Image.register_open(WmfImagePlugin.WmfImageFile.format, WmfImagePlugin.WmfImageFile)

        with open(emf_file, 'rb') as f:
            img = WmfImagePlugin.WmfImageFile(f)
            print(f"  Successfully opened with WmfImagePlugin")

            # Convert to RGB and save to buffer as PNG first
            png_buffer = BytesIO()
            rgb_image = img.convert('RGB')
            rgb_image.save(png_buffer, 'PNG')
            png_buffer.seek(0)

            # Reload from PNG
            img_png = Image.open(png_buffer)

            # Resize
            original_width, original_height = img_png.size
            aspect_ratio = original_height / original_width
            target_height = int(target_width * aspect_ratio)
            img_resized = img_png.resize((target_width, target_height), Image.Resampling.LANCZOS)

            # Save as GIF
            img_resized.save(output_file, 'GIF', optimize=True)
            print(f"  ✓ Converted successfully: {output_file.name} ({target_width}x{target_height})")
            return True

    except Exception as e:
        print(f"  ✗ Failed: {e}")

    # Method 3: Try with explicit format hint
    print("\nMethod 3: Explicit format specification...")
    try:
        with open(emf_file, 'rb') as f:
            data = f.read()
            img = Image.open(BytesIO(data))
            img.format = 'WMF'  # Force format

            # Try to load the data
            img.load()

            # Convert to RGB
            if img.mode not in ('RGB', 'P'):
                img = img.convert('RGB')

            # Resize
            original_width, original_height = img.size
            aspect_ratio = original_height / original_width
            target_height = int(target_width * aspect_ratio)
            img_resized = img.resize((target_width, target_height), Image.Resampling.LANCZOS)

            # Save as GIF
            img_resized.save(output_file, 'GIF', optimize=True)
            print(f"  ✓ Converted successfully: {output_file.name} ({target_width}x{target_height})")
            return True

    except Exception as e:
        print(f"  ✗ Failed: {e}")

    print("\n" + "=" * 70)
    print("All conversion methods failed.")
    print("The EMF file requires ImageMagick or manual conversion.")
    print("\nTo install ImageMagick on macOS:")
    print("  1. Install Homebrew: https://brew.sh")
    print("  2. Run: brew install imagemagick")
    print("  3. Then convert with: convert ijrit-11-139-g002.emf ijrit-11-139-g002.gif")
    return False

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python convert_emf_to_gif.py <input.emf> <output.gif>")
        sys.exit(1)

    success = convert_emf_to_gif(sys.argv[1], sys.argv[2])
    sys.exit(0 if success else 1)

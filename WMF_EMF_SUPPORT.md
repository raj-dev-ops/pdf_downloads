# WMF/EMF Image Support in Word Document Extractor

## Overview

The `extract_images_to_gif.py` script now supports extracting WMF/EMF images (Enhanced Metafiles) that are commonly used in Word documents for embedded Excel charts and diagrams.

## Platform Support

### ✅ Windows (Recommended for automation)

**Native support - works out of the box!**

On Windows, PIL (Pillow) has built-in support for WMF/EMF files through the Windows GDI (Graphics Device Interface). No additional installation is required.

**Setup:**
```bash
pip install -r requirements.txt
python extract_images_to_gif.py your-document.docx
```

That's it! WMF/EMF files will be automatically converted to GIF.

### 🔧 macOS / Linux

Requires ImageMagick installation for WMF/EMF support.

**Setup:**

1. **Install ImageMagick:**

   **macOS (via Homebrew):**
   ```bash
   brew install imagemagick
   ```

   **Ubuntu/Debian:**
   ```bash
   sudo apt-get install imagemagick libmagickwand-dev
   ```

   **CentOS/RHEL:**
   ```bash
   sudo yum install ImageMagick ImageMagick-devel
   ```

2. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   pip install wand>=0.6.0
   ```

3. **Run the script:**
   ```bash
   python extract_images_to_gif.py your-document.docx
   ```

## How It Works

The script uses a multi-method approach to handle WMF/EMF files:

1. **Method 1 (Windows):** Uses PIL's native Windows GDI support
2. **Method 2 (Windows):** Falls back to pywin32 if available
3. **Method 3 (All platforms):** Uses Wand (ImageMagick Python bindings)

If all methods fail (e.g., on Mac/Linux without ImageMagick), the script will:
- Save the raw EMF file (e.g., `ijrit-11-139-g002.emf`)
- Print a warning with installation instructions
- Continue processing other images

## Automated Windows Deployment

For automation on Windows laptops, simply:

1. Ensure Python 3.7+ is installed
2. Install dependencies: `pip install -r requirements.txt`
3. Run the script - WMF/EMF support works automatically

**No additional configuration needed!**

## Testing WMF/EMF Support

To verify WMF/EMF support is working on your system:

```python
from PIL import Image
import platform

print(f"Platform: {platform.system()}")

try:
    # Try to open a WMF/EMF file
    img = Image.open("your-file.emf")
    img.load()
    print("✓ WMF/EMF support is working!")
except Exception as e:
    print(f"✗ WMF/EMF support not available: {e}")
```

## Troubleshooting

### Windows

If extraction fails on Windows:
- Ensure PIL/Pillow is up to date: `pip install --upgrade Pillow`
- Check Python version (3.7+ recommended)

### Mac/Linux

If you see "No suitable converter found":
1. Verify ImageMagick is installed: `which convert` or `which magick`
2. If not found, install ImageMagick (see setup above)
3. Install wand: `pip install wand`
4. Try running `convert -version` to verify ImageMagick works

If ImageMagick is installed but wand fails:
```bash
# Reinstall wand with proper ImageMagick linkage
pip uninstall wand
pip install --no-cache-dir wand
```

## File Format Notes

- **EMF (Enhanced Metafile):** Modern Windows metafile format
- **WMF (Windows Metafile):** Legacy Windows metafile format
- Both formats are vector-based and commonly used for embedded charts/diagrams
- The script converts them to GIF (rasterized) at 1500px width with maintained aspect ratio

## Known Limitations

- Vector quality is lost during conversion to GIF (rasterization)
- Some complex EMF files may not render perfectly
- Color profiles may vary slightly from original
- Requires ImageMagick on non-Windows platforms

## Questions?

For issues or questions, please check the main script documentation or file an issue in the repository.

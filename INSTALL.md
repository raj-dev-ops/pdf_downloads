# Installation Instructions

This guide covers installation for both Windows and macOS users.

---

## Windows Installation

### Prerequisites
- Python 3.8 or higher
- Microsoft Office installed (if you need .doc file conversion)

### Step 1: Update pip
```cmd
python -m pip install --upgrade pip
```

### Step 2: Install core dependencies
```cmd
pip install -r requirements.txt
```

This will install all required packages. The installation should complete without errors.

### Step 3: Legacy .doc File Support (Optional)

If you need to convert legacy `.doc` files to `.docx`, choose one of these options:

**Option A: Use Microsoft Word directly (Recommended if you have MS Office)**
```python
# Your Python code can use COM automation with win32com
pip install pywin32

# Example usage:
import win32com.client
word = win32com.client.Dispatch("Word.Application")
doc = word.Documents.Open("path/to/file.doc")
doc.SaveAs("path/to/file.docx", FileFormat=16)
doc.Close()
word.Quit()
```

**Option B: Manually convert .doc to .docx first**
- Open .doc files in Microsoft Word
- Save As → .docx format
- Then use this tool with .docx files

**Option C: Install doc2docx (May have dependency issues)**
```cmd
# This may fail due to pywin32 version conflicts
pip install pywin32
pip install doc2docx --no-deps
```

**Note:** Most modern Word documents are already in `.docx` format, so you likely don't need `.doc` support.

---

## macOS Installation

### Prerequisites
- Python 3.8 or higher
- Homebrew (optional, for ImageMagick support)

### Step 1: Update pip
```bash
python3 -m pip install --upgrade pip
```

### Step 2: Install dependencies
```bash
pip3 install -r requirements.txt
```

**Note:** On macOS, `doc2docx` will be automatically skipped since it's Windows-only. You can still process `.docx` files, but legacy `.doc` files are not supported.

### Optional: WMF/EMF Image Support (macOS/Linux)
If you need to handle WMF/EMF images, install ImageMagick:

```bash
# Install ImageMagick
brew install imagemagick

# Install Python wrapper
pip3 install wand
```

---

## Verification

After installation, verify everything is working:

**Windows:**
```cmd
python -c "import requests, pypdf, docx, PIL; print('All core packages installed successfully!')"
```

**macOS:**
```bash
python3 -c "import requests, pypdf, docx, PIL; print('All core packages installed successfully!')"
```

**Note:** Using `docx` (from python-docx) instead of `doc2docx` - works for .docx files on both platforms.

---

## Troubleshooting

### Windows: "pywin32 version conflict" or "doc2docx" errors
**Root cause:** The `doc2docx` package requires an outdated pywin32 version (305) that's no longer available.

**Solutions:**
1. **Skip .doc support** - Most documents are .docx anyway, core installation works fine
2. **Use Microsoft Word COM** - Install `pip install pywin32` and use win32com.client (see Option A above)
3. **Manual conversion** - Convert .doc to .docx in Word before processing

### macOS: All packages install successfully
- No special configuration needed
- .docx files work out of the box with python-docx

### Both platforms: "Command not found: python"
- Windows: Make sure Python is added to PATH during installation
- macOS: Use `python3` instead of `python`

---

## What's Installed

### Core Dependencies (Both Platforms)
- **requests, beautifulsoup4, lxml**: Web scraping and HTML parsing
- **click**: Command-line interface
- **pandas, openpyxl**: Excel file handling
- **tqdm**: Progress bars
- **pypdf**: PDF manipulation
- **reportlab**: PDF generation
- **python-docx**: Word document (.docx) handling
- **Pillow**: Image processing

### Windows-Only (Optional)
- **pywin32**: For COM automation with Microsoft Office
- **doc2docx**: Legacy .doc to .docx conversion (has dependency conflicts, use alternatives above)

### macOS/Linux Optional
- **wand**: ImageMagick wrapper for WMF/EMF support

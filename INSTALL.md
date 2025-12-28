# Installation Instructions

This guide covers installation for both Windows and macOS users.

---

## Windows Installation (Recommended for .doc file support)

### Prerequisites
- Python 3.8 or higher
- Microsoft Office installed (for .doc file conversion)

### Step 1: Update pip
```cmd
python -m pip install --upgrade pip
```

### Step 2: Install dependencies
```cmd
pip install -r requirements.txt
```

**If you encounter dependency conflicts**, try one of these solutions:

**Option A: Use legacy resolver**
```cmd
pip install -r requirements.txt --use-deprecated=legacy-resolver
```

**Option B: Install with no dependencies check first, then fix**
```cmd
pip install --upgrade pip setuptools wheel
pip install -r requirements.txt --use-deprecated=legacy-resolver
```

**Option C: Install packages individually (if above fails)**
```cmd
pip install requests beautifulsoup4 lxml click urllib3 pandas openpyxl tqdm pypdf reportlab python-docx Pillow
pip install doc2docx
```

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
python -c "import requests, pypdf, python_docx, doc2docx; print('All packages installed successfully!')"
```

**macOS:**
```bash
python3 -c "import requests, pypdf, docx; print('All packages installed successfully!')"
```

---

## Troubleshooting

### Windows: "Cannot install doc2docx" error
1. Make sure you have the latest pip: `python -m pip install --upgrade pip`
2. Use the legacy resolver: `pip install -r requirements.txt --use-deprecated=legacy-resolver`
3. Ensure Microsoft Office is installed on your system

### macOS: "doc2docx not found" (This is normal)
- doc2docx is Windows-only and will be skipped on macOS
- You can still process .docx files using python-docx

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

### Windows-Only
- **doc2docx**: Legacy .doc to .docx conversion (requires MS Office)

### macOS/Linux Optional
- **wand**: ImageMagick wrapper for WMF/EMF support

# Docling GUI

A Python GUI application for document conversion using CustomTkinter.

## Key Features
- Convert between PDF, DOCX, PPTX, HTML, Markdown, and JSON
- Batch processing support
- OCR with multi-language support
- Table extraction with fast/accurate modes
- Real-time progress tracking

## Installation

### Basic Installation
1. Install Python 3.8+
2. Install Docling:
   ```bash
   pip install docling
   ```
3. Install GUI dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the application:
   ```bash
   python docling-gui.py
   ```

Works on macOS, Linux, and Windows, with support for both x86_64 and arm64 architectures.

### Alternative PyTorch Distributions
Docling depends on PyTorch. For different architectures or CPU-only installations:
```bash
# Example for Linux CPU-only version
pip install docling --extra-index-url https://download.pytorch.org/whl/cpu
```

### OCR Engine Options
Docling supports multiple OCR engines:

| Engine        | Installation                          | Usage               |
|---------------|---------------------------------------|---------------------|
| EasyOCR       | Default in Docling                   | EasyOcrOptions      |
| Tesseract     | System dependency (see below)        | TesseractOcrOptions |
| Tesseract CLI | System dependency                    | TesseractCliOcrOptions |
| OcrMac        | macOS only: `pip install ocrmac`     | OcrMacOptions       |
| RapidOCR      | `pip install rapidocr_onnxruntime`   | RapidOcrOptions     |

### Tesseract Installation
For Tesseract OCR engine:

#### macOS (via Homebrew)
```bash
brew install tesseract leptonica pkg-config
export TESSDATA_PREFIX=/opt/homebrew/share/tessdata/
```

#### Linux (Debian-based)
```bash
sudo apt-get install tesseract-ocr
export TESSDATA_PREFIX=/usr/share/tesseract-ocr/4.00/tessdata/
```

#### Linking Tesseract
For optimal performance:
```bash
pip uninstall tesserocr
pip install --no-binary :all: tesserocr
```

## Development Setup
To contribute to Docling development:
```bash
poetry install --all-extras
```

## Further Details
For more information and advanced usage, refer to the official Docling repository:
https://github.com/DS4SD/docling

## License
MIT License
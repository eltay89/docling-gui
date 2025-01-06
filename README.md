# Docling GUI

A graphical user interface for Docling document processing, built with Python and CustomTkinter.

## Features

- Convert documents between multiple formats (PDF, DOCX, PPTX, HTML, Markdown, JSON)
- Batch processing of multiple files
- OCR support for scanned documents with multi-language support
- Table extraction with fast and accurate modes
- Real-time progress tracking with terminal-style logging
- Detailed error handling and reporting
- Configuration persistence (remembers last used directories and settings)
- Document structure visualization
- Intuitive user interface with tooltips and help system
- Threaded processing with cancellation support

## Installation

1. Ensure you have Python 3.8+ installed
2. Install required dependencies:
   ```bash
   pip install docling customtkinter pillow
   ```
3. Download or clone this repository
4. Run the application:
   ```bash
   python docling-gui.py
   ```

## Usage

1. Select the input format from the dropdown
2. Choose a file or directory for batch processing
3. Select output format (Markdown, HTML, or JSON)
4. Configure OCR languages if needed (comma-separated)
5. Choose table extraction mode (Fast or Accurate)
6. Click "Convert" to start processing
7. View results in the output panel
8. Use the terminal window to monitor detailed conversion progress
9. Access help documentation through the Help button

## Advanced Features

- **Terminal Logging**: Real-time conversion progress and debugging information
- **Error Handling**: Detailed error messages with traceback information
- **Configuration**: Settings are automatically saved in config.json
- **Thread Management**: Safe thread cancellation and cleanup
- **File Validation**: Automatic output directory creation and permission checking

## License

This project is licensed under the MIT License.
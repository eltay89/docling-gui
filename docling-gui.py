import os
import sys
import logging
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import List, Optional
import threading

# Set up logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='pdf2md.log'
)

# Set environment variable for PyTorch compatibility
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'

try:
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.pipeline_options import PdfPipelineOptions, TableFormerMode, EasyOcrOptions, TableStructureOptions
    from docling.datamodel.base_models import InputFormat
    logging.info("Docling modules imported successfully")
except ImportError as e:
    logging.error(f"Failed to import Docling modules: {e}")
    messagebox.showerror("Import Error", f"Failed to import Docling modules: {e}")
    sys.exit(1)

class PDFConverterApp:
    FILE_EXTENSIONS = {
        'pdf': ('PDF Files', '*.pdf'),
        'docx': ('Word Documents', '*.docx'),
        'pptx': ('PowerPoint Presentations', '*.pptx'),
        'html': ('HTML Files', '*.html;*.htm'),
        'image': ('Image Files', '*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.tiff')
    }
    OUTPUT_FORMAT_MAP = {
        'markdown': lambda doc: doc.export_to_markdown(),
        'html': lambda doc: doc.export_to_html(),
        'json': lambda doc: doc.export_to_json(indent=2)
    }

    def __init__(self, root: tk.Tk):
        """Initialize the main application."""
        self.root = root
        self.root.title("Enhanced PDF Converter")
        # Set initial size and make non-resizable
        self.window_width = 800
        self.window_height = 600
        self.root.geometry(f"{self.window_width}x{self.window_height}")
        self.root.resizable(False, False)  # Prevent resizing

        self.conversion_result = None
        self.setup_ui()
        self.setup_defaults()
        self.last_input_dir = None
        self.last_output_dir = None

        # Center the window after setting up UI (which determines final size)
        self.center_window()

    def setup_defaults(self):
        """Initialize default settings for the converter."""
        self.pipeline_options = PdfPipelineOptions(
            do_ocr=True,
            do_table_structure=True,
            table_structure_options=TableStructureOptions(mode=TableFormerMode.FAST, do_cell_matching=True),
            ocr_options=EasyOcrOptions(lang=['en']) # Initialize with EasyOcrOptions
        )

    def setup_ui(self):
        """Create and arrange all GUI components."""
        self.root.style = ttk.Style(self.root)
        self.root.style.theme_use('clam')  # Use a modern theme

        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill='both', expand=True)

        # File Selection
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=0, column=0, sticky='ew', padx=10, pady=10)

        ttk.Label(file_frame, text="Input Format:").grid(row=0, column=0, sticky='w', pady=3)
        self.input_format = tk.StringVar(value='pdf')
        input_formats = list(self.FILE_EXTENSIONS.keys())
        self.input_format_combobox = ttk.Combobox(file_frame, textvariable=self.input_format, values=input_formats, state='readonly')
        self.input_format_combobox.grid(row=0, column=1, sticky='ew', padx=5, pady=3)

        ttk.Label(file_frame, text="File Path:").grid(row=1, column=0, sticky='w', pady=3)
        self.file_path = tk.StringVar()
        self.file_path_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=60)
        self.file_path_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=3)
        self.browse_file_button = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        self.browse_file_button.grid(row=1, column=2, padx=2)
        self.batch_mode_button = ttk.Button(file_frame, text="Batch Mode", command=self.browse_directory)
        self.batch_mode_button.grid(row=1, column=3, padx=2)
        ttk.Button(file_frame, text="Clear", command=lambda: self.clear_paths('file')).grid(row=1, column=4, padx=2)
        file_frame.columnconfigure(1, weight=1)

        # Output Directory
        output_dir_frame = ttk.LabelFrame(main_frame, text="Output Directory", padding="10")
        output_dir_frame.grid(row=1, column=0, sticky='ew', padx=10, pady=10)

        ttk.Label(output_dir_frame, text="Directory:").grid(row=0, column=0, sticky='w', pady=3)
        self.output_directory = tk.StringVar()
        self.output_directory_entry = ttk.Entry(output_dir_frame, textvariable=self.output_directory, width=60)
        self.output_directory_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=3)
        self.browse_output_dir_button = ttk.Button(output_dir_frame, text="Browse", command=self.browse_output_directory)
        self.browse_output_dir_button.grid(row=0, column=2, padx=2)
        ttk.Button(output_dir_frame, text="Clear", command=lambda: self.clear_paths('output')).grid(row=0, column=3, padx=2)
        output_dir_frame.columnconfigure(1, weight=1)

        # Conversion Options
        options_frame = ttk.LabelFrame(main_frame, text="Conversion Options", padding="10")
        options_frame.grid(row=2, column=0, sticky='ew', padx=10, pady=10)

        ttk.Label(options_frame, text="Output Format:").grid(row=0, column=0, sticky='w', pady=3)
        self.output_format = tk.StringVar(value='markdown')
        self.output_markdown_radio = ttk.Radiobutton(options_frame, text="Markdown", variable=self.output_format, value='markdown')
        self.output_markdown_radio.grid(row=0, column=1, sticky='w', padx=5)
        self.output_html_radio = ttk.Radiobutton(options_frame, text="HTML", variable=self.output_format, value='html')
        self.output_html_radio.grid(row=0, column=2, sticky='w', padx=5)
        self.output_json_radio = ttk.Radiobutton(options_frame, text="JSON", variable=self.output_format, value='json')
        self.output_json_radio.grid(row=0, column=3, sticky='w', padx=5)

        ttk.Label(options_frame, text="OCR Languages:").grid(row=1, column=0, sticky='w', pady=3)
        self.ocr_languages = tk.StringVar(value='en')
        self.ocr_languages_entry = ttk.Entry(options_frame, textvariable=self.ocr_languages, width=30)
        self.ocr_languages_entry.grid(row=1, column=1, columnspan=3, sticky='ew', padx=5)
        ttk.Label(options_frame, text="(comma separated, e.g. 'en,fr,de')", font=('TkDefaultFont', 8)).grid(row=1, column=4, sticky='w', padx=5)

        ttk.Label(options_frame, text="Table Extraction:").grid(row=2, column=0, sticky='w', pady=3)
        self.table_mode = tk.StringVar(value='fast')
        self.table_mode_fast_radio = ttk.Radiobutton(options_frame, text="Fast", variable=self.table_mode, value='fast')
        self.table_mode_fast_radio.grid(row=2, column=1, sticky='w', padx=5)
        self.table_mode_accurate_radio = ttk.Radiobutton(options_frame, text="Accurate", variable=self.table_mode, value='accurate')
        self.table_mode_accurate_radio.grid(row=2, column=2, sticky='w', padx=5)
        options_frame.columnconfigure(1, weight=1)

        # Control Buttons
        control_frame = ttk.Frame(main_frame, padding="10")
        control_frame.grid(row=3, column=0, sticky='ew', padx=10, pady=15)
        control_frame.columnconfigure(0, weight=1)
        control_frame.columnconfigure(3, weight=1)

        self.convert_button = ttk.Button(control_frame, text="Convert", command=self.convert, style='Accent.TButton')
        self.convert_button.grid(row=0, column=1, padx=5, pady=10, sticky='ew')

        self.show_structure_button = ttk.Button(control_frame, text="Show Structure", command=self.show_structure)
        self.show_structure_button.grid(row=0, column=2, padx=5, pady=10, sticky='ew')
        self.clear_output_button = ttk.Button(control_frame, text="Clear Output", command=self.clear_output)
        self.clear_output_button.grid(row=0, column=0, padx=5, pady=10, sticky='ew')
        self.help_button = ttk.Button(control_frame, text="Help", command=self.show_help)
        self.help_button.grid(row=0, column=3, padx=5, pady=10, sticky='ew')

        # Output Section
        output_frame = ttk.LabelFrame(main_frame, text="Output", padding="10")
        output_frame.grid(row=4, column=0, sticky='nsew', padx=10, pady=10)
        main_frame.rowconfigure(4, weight=1)

        self.output_text = scrolledtext.ScrolledText(output_frame, wrap='word', state=tk.DISABLED)
        self.output_text.pack(fill='both', expand=True)

        # Progress Bar (indeterminate mode for single file, determinate for batch)
        self.progress_bar = ttk.Progressbar(output_frame, orient=tk.HORIZONTAL, length=self.window_width - 40, mode='indeterminate')
        self.progress_bar.pack(pady=(0, 5))

        # Status Bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief='sunken', anchor='w')
        status_bar.grid(row=5, column=0, sticky='ew', padx=10, pady=(0, 10))

        # Initialize Tooltips after widget creation
        self.create_tooltips()

    def create_tooltips(self):
        ToolTip(self.input_format_combobox, "Select the format of the input file.")
        ToolTip(self.file_path_entry, "Path to the input file or directory for batch conversion.")
        ToolTip(self.browse_file_button, "Browse for an input file.")
        ToolTip(self.batch_mode_button, "Browse for a directory to convert multiple files.")
        ToolTip(self.output_directory_entry, "Directory where converted files will be saved.")
        ToolTip(self.browse_output_dir_button, "Browse for an output directory.")
        ToolTip(self.output_markdown_radio, "Convert to Markdown format.")
        ToolTip(self.output_html_radio, "Convert to HTML format.")
        ToolTip(self.output_json_radio, "Convert to JSON format.")
        ToolTip(self.ocr_languages_entry, "Specify languages for OCR (comma-separated).")
        ToolTip(self.table_mode_fast_radio, "Faster table extraction, may be less accurate.")
        ToolTip(self.table_mode_accurate_radio, "More accurate table extraction, may take longer.")
        ToolTip(self.convert_button, "Start the conversion process.")
        ToolTip(self.show_structure_button, "Display the structure of the converted document.")
        ToolTip(self.clear_output_button, "Clear the output area.")
        ToolTip(self.help_button, "Show help information.")

    def browse_file(self):
        """Handle file browsing based on the selected input format."""
        selected_format = self.input_format.get().lower()
        if selected_format in self.FILE_EXTENSIONS:
            file_type_name, file_pattern = self.FILE_EXTENSIONS[selected_format]
            title = f"Select {file_type_name}"
            filetypes = [(file_type_name, file_pattern)]
            initialdir = self.last_input_dir if self.last_input_dir else os.path.expanduser("~")
            filename = filedialog.askopenfilename(title=title, filetypes=filetypes, initialdir=initialdir)
            if filename:
                self.file_path.set(filename)
                self.last_input_dir = os.path.dirname(filename)
                self.status_var.set(f"Selected file: {filename}")
        else:
            messagebox.showerror("Error", "Invalid input format selected.")

    def browse_directory(self):
        """Handle directory selection for batch processing."""
        initialdir = self.last_input_dir if self.last_input_dir else os.path.expanduser("~")
        directory = filedialog.askdirectory(initialdir=initialdir)
        if directory:
            self.file_path.set(directory)
            self.last_input_dir = directory
            self.status_var.set(f"Selected directory: {directory}")

    def browse_output_directory(self):
        """Handle browsing for the output directory."""
        initialdir = self.last_output_dir if self.last_output_dir else os.path.expanduser("~")
        directory = filedialog.askdirectory(initialdir=initialdir)
        if directory:
            self.output_directory.set(directory)
            self.last_output_dir = directory
            self.status_var.set(f"Selected output directory: {directory}")

    def clear_paths(self, target):
        """Clears the file path or output directory entry."""
        if target == 'file':
            self.file_path.set("")
        elif target == 'output':
            self.output_directory.set("")
        self.status_var.set("Ready")

    def convert(self):
        """Handle the conversion process in a separate thread."""
        source = self.file_path.get()
        if not source:
            messagebox.showwarning("Input Required", "Please select a file or directory first.")
            return

        self.progress_bar["mode"] = "indeterminate"
        self.progress_bar.start()  # Start indeterminate progress

        threading.Thread(target=self._perform_conversion).start()
        self.status_var.set("Converting...")

    def _perform_conversion(self):
        """Perform the conversion process."""
        source = self.file_path.get()
        try:
            # Update pipeline options based on user selections
            ocr_languages = [
                lang.strip() for lang in self.ocr_languages.get().split(',')
            ]
            self.pipeline_options.ocr_options = EasyOcrOptions(lang=ocr_languages)

            table_mode_str = self.table_mode.get()
            table_mode = TableFormerMode.ACCURATE if table_mode_str == 'accurate' else TableFormerMode.FAST
            self.pipeline_options.table_structure_options = TableStructureOptions(mode=table_mode)

            # Determine the input format and create DocumentConverter
            input_format_str = self.input_format.get()
            input_format = InputFormat[input_format_str.upper()]  # Directly use InputFormat members
            
            allowed_formats = [input_format] # Use the InputFormat member
            format_options = {}
            if input_format == InputFormat.PDF:
                format_options = {InputFormat.PDF: PdfFormatOption(pipeline_options=self.pipeline_options)}

            converter = DocumentConverter(
                allowed_formats=allowed_formats,
                format_options=format_options
            )

            # Handle conversion based on whether a file or directory is selected
            if os.path.isfile(source):
                self._convert_single_file(converter, source)
            elif os.path.isdir(source):
                self._convert_batch(converter, source, input_format_str) # Pass the string format for file extension matching
            else:
                self.root.after(0, messagebox.showerror, "Invalid Input", "Please select a valid file or directory.")
                self.root.after(0, self.status_var.set, "Ready")

        except Exception as e:
            self.handle_error(e)
            self.root.after(0, self.status_var.set, "Error occurred")

    def _convert_single_file(self, converter: DocumentConverter, source: str):
        """Convert a single file using the provided converter."""
        self.root.after(0, self.status_var.set, f"Converting {os.path.basename(source)}...")

        try:
            self.conversion_result = converter.convert(source)  # Store the conversion result
            output_content = self.get_output_content(self.conversion_result)

            # Determine the output directory
            output_dir = self.output_directory.get() or os.path.dirname(source)
            output_filename = Path(output_dir) / Path(os.path.basename(source)).with_suffix(f".{self.output_format.get()}")

            # Save output file
            with open(output_filename, 'w', encoding='utf-8') as f:
                f.write(output_content)

            # Update UI
            self.root.after(0, self.update_output, f"Conversion complete!\nSaved to: {output_filename}\n\n{output_content}")
            self.root.after(0, self.progress_bar.stop)  # Stop indeterminate progress
            self.root.after(0, self.status_var.set, f"Conversion complete: {output_filename}")
        except Exception as e:
            self.handle_error(e)
            self.root.after(0, self.status_var.set, "Error during conversion")

    def _convert_batch(self, converter: DocumentConverter, directory: str, input_format_str: str):
        """Convert all files of the selected format in a directory."""
        file_patterns = self.FILE_EXTENSIONS.get(input_format_str, ("", ""))[1]
        if not file_patterns:
            self.root.after(0, messagebox.showerror, "Error", "Invalid input format for batch conversion.")
            self.root.after(0, self.status_var.set, "Error")
            return

        files = Path(directory).glob(file_patterns)
        file_list = list(files)

        if not file_list:
            self.root.after(0, messagebox.showinfo, "No Files Found", f"No {input_format_str.upper()} files found in the selected directory.")
            self.root.after(0, self.status_var.set, "Ready")
            return

        # Determine the output directory
        output_dir = self.output_directory.get() or directory

        self.root.after(0, self.status_var.set, f"Converting {len(file_list)} files...")
        
        total_files = len(file_list)
        results = []
        for i, input_file in enumerate(file_list):
            try:
                conversion_result = converter.convert(str(input_file))
                output_content = self.get_output_content(conversion_result)
                output_filename = Path(output_dir) / Path(input_file).with_suffix(f".{self.output_format.get()}")
                output_filename.write_text(output_content, encoding="utf-8")
                results.append(f"✔ {input_file.name} → {output_filename.name}")
            except Exception as e:
                results.append(f"✘ {input_file.name} failed: {str(e)}")
            finally:
                progress = int(((i + 1) / total_files) * 100)
                self.root.after(0, self.progress_bar.config, {"mode": "determinate", "value": progress}) # Update progress
                self.root.after(0, self.update_output, f"Processing: {input_file.name} ({progress}%)")

        self.root.after(0, self.update_output, "Batch conversion results:\n\n" + "\n".join(results))
        self.root.after(0, self.progress_bar.stop)  # Stop indeterminate progress
        self.root.after(0, self.status_var.set, f"Batch conversion complete: {len(file_list)} files processed")

    def get_output_content(self, result) -> str:
        """Get content in the selected output format."""
        output_format = self.output_format.get()
        if output_format in self.OUTPUT_FORMAT_MAP:
            return self.OUTPUT_FORMAT_MAP[output_format](result.document)
        return ""

    def update_output(self, text: str):
        """Update the output text area."""
        self.output_text.config(state=tk.NORMAL)
        self.output_text.insert(tk.END, text + "\n")  # Append new text with newline
        self.output_text.config(state=tk.DISABLED)
        self.output_text.see(tk.END)  # Scroll to the end

    def clear_output(self):
        """Clear the output text area."""
        self.update_output("")
        self.status_var.set("Ready")
        self.conversion_result = None

    def handle_error(self, error: Exception):
        """Handle and display errors."""
        error_msg = f"Error: {str(error)}\n{traceback.format_exc()}"
        logging.error(error_msg)
        self.root.after(0, self.update_output, error_msg)
        self.root.after(0, messagebox.showerror, "Conversion Error", str(error))

    def show_structure(self):
        """Display the structure of the converted document in a new window."""
        if self.conversion_result and self.conversion_result.document:
            structure_window = tk.Toplevel(self.root)
            structure_window.title("Document Structure")
            text_area = scrolledtext.ScrolledText(structure_window, wrap='none')
            text_area.pack(expand=True, fill='both')
            element_tree = self.conversion_result.document.print_element_tree(repr_mode=True)
            text_area.insert(tk.END, element_tree)
        else:
            messagebox.showinfo("Info", "No document has been converted yet.")

    def show_help(self):
        """Display help information for the application."""
        help_text = """PDF to Markdown Converter Help

1. Select the input file type and browse for the file or directory for batch processing.
2. Choose the desired output format (Markdown, HTML, or JSON).
3. Configure OCR languages (comma separated, e.g., 'en,fr,de') if processing scanned documents.
4. Select the table extraction mode (Fast or Accurate) for PDF documents.
5. Optionally, select an output directory to save the converted files. If no directory is selected, files will be saved in the same directory as the input files.
6. Click 'Convert' to start the conversion process.
7. Click 'Show Structure' after converting a document to see its internal structure.

The converted files will be saved in the specified output directory or the same directory as the input files with the appropriate extension.

Note: For best OCR results, specify all languages present in your documents.
"""
        messagebox.showinfo("Help", help_text)

    def center_window(self):
        """Centers the application window on the screen."""
        self.root.update_idletasks()  # Required to get accurate dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - self.window_width) // 2
        y = (screen_height - self.window_height) // 2
        self.root.geometry(f"+{x}+{y}")

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)  # Remove border
        tooltip_label = ttk.Label(self.tooltip_window, text=self.text, background="#ffffe0", relief="solid", borderwidth=1, padding=5)
        tooltip_label.pack()
        self.tooltip_window.wm_geometry(f"+{x}+{y}")

    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()
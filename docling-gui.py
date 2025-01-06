import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import json
from pathlib import Path
import logging
import logging.handlers
import subprocess
from typing import Optional
import traceback

import customtkinter as ctk
from PIL import Image, ImageTk

# Set KMP_DUPLICATE_LIB_OK Environment Variable
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
print(f"KMP_DUPLICATE_LIB_OK: {os.environ.get('KMP_DUPLICATE_LIB_OK')}")

try:
    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.pipeline_options import PdfPipelineOptions, TableFormerMode, TableStructureOptions
    from docling.datamodel.base_models import InputFormat

    logging.info("Docling modules imported successfully")

except ImportError as e:
    logging.error(f"Failed to import Docling modules: {e}")
    messagebox.showerror("Import Error", f"Failed to import Docling modules: {e}")
    sys.exit(1)

# Logging Setup
log_file = Path("pdf2md.log")
file_handler = logging.handlers.RotatingFileHandler(
    log_file, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8"
)
logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(message)s",
    handlers=[file_handler, logging.StreamHandler(sys.stdout)],
)

def check_for_multiple_openmp_libs():
    """Attempts to find multiple instances of libiomp5md.dll (Windows)."""
    try:
        if os.name == "nt":
            result = subprocess.run(
                ["tasklist", "/m", "libiomp5md.dll"],
                capture_output=True,
                text=True,
                check=True,
            )
            output = result.stdout
            if output.count("libiomp5md.dll") > 1:
                logging.warning("Multiple instances of libiomp5md.dll detected:")
                logging.warning(output)
            else:
                logging.info(
                    "No multiple instances of libiomp5md.dll found (using tasklist)."
                )
    except Exception as e:
        logging.error(f"Error checking for OpenMP libraries: {e}")

check_for_multiple_openmp_libs()

class PDFConverterApp:
    FILE_EXTENSIONS = {
        "pdf": ("PDF Files", "*.pdf"),
        "docx": ("Word Documents", "*.docx"),
        "pptx": ("PowerPoint Presentations", "*.pptx"),
        "html": ("HTML Files", "*.html;*.htm"),
        "image": ("Image Files", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.tiff"),
    }
    OUTPUT_FORMAT_MAP = {
        "markdown": lambda doc: doc.document.export_to_markdown(),
        "html": lambda doc: doc.document.export_to_html(),
        "json": lambda doc: doc.document.export_to_json(indent=2),
    }
    CONFIG_FILE = "config.json"

    def __init__(self, root: ctk.CTk):
        self.root = root
        self.root.title("Enhanced PDF Converter")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.root.geometry("800x600")

        self.cancelled = False
        self.conversion_thread = None
        self.conversion_result = None
        self.last_input_dir = None
        self.last_output_dir = None
        self.pipeline_options = PdfPipelineOptions()
        self.cancel_event = threading.Event()

        self.setup_ui()
        self.load_config()

    def setup_ui(self):
        self.main_frame = ctk.CTkFrame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tabview = ctk.CTkTabview(self.main_frame)
        self.tabview.add("Input")
        self.tabview.add("Output")
        self.tabview.add("Advanced")
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        self.setup_input_tab(self.tabview.tab("Input"))
        self.setup_output_tab(self.tabview.tab("Output"))
        self.setup_advanced_tab(self.tabview.tab("Advanced"))

        self.control_frame = ctk.CTkFrame(self.main_frame)
        self.setup_control_frame(self.control_frame).pack(fill="x", padx=10, pady=10)

        self.setup_output_text()
        self.setup_status()

    def setup_input_tab(self, parent_frame):
        ctk.CTkLabel(parent_frame, text="Input Format:").grid(
            row=0, column=0, padx=10, pady=(10, 5), sticky="w"
        )
        self.input_format = ctk.StringVar(value="pdf")
        self.input_format_combobox = ctk.CTkComboBox(
            parent_frame,
            values=list(self.FILE_EXTENSIONS.keys()),
            variable=self.input_format,
            state="readonly",
        )
        self.input_format_combobox.grid(
            row=0, column=1, padx=10, pady=(10, 5), sticky="ew"
        )

        ctk.CTkLabel(parent_frame, text="File/Directory:").grid(
            row=1, column=0, padx=10, pady=5, sticky="w"
        )
        self.file_path = ctk.StringVar()
        self.file_path_entry = ctk.CTkEntry(parent_frame, textvariable=self.file_path)
        self.file_path_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        self.browse_file_button = ctk.CTkButton(
            parent_frame, text="Browse File", command=self.browse_file
        )
        self.browse_file_button.grid(row=1, column=2, padx=5, pady=5, sticky="w")

        self.batch_mode_button = ctk.CTkButton(
            parent_frame, text="Batch Mode", command=self.browse_directory
        )
        self.batch_mode_button.grid(row=1, column=3, padx=5, pady=5, sticky="w")

        ctk.CTkButton(
            parent_frame, text="Clear", command=lambda: self.clear_paths("file")
        ).grid(row=1, column=4, padx=5, pady=5, sticky="w")

        parent_frame.columnconfigure(1, weight=1)

    def setup_output_tab(self, parent_frame):
        ctk.CTkLabel(parent_frame, text="Output Format:").grid(
            row=0, column=0, padx=10, pady=(10, 5), sticky="w"
        )
        self.output_format = ctk.StringVar(value="markdown")
        self.output_markdown_radio = ctk.CTkRadioButton(
            parent_frame,
            text="Markdown",
            variable=self.output_format,
            value="markdown",
        )
        self.output_markdown_radio.grid(
            row=0, column=1, padx=10, pady=(10, 5), sticky="w"
        )
        self.output_html_radio = ctk.CTkRadioButton(
            parent_frame, text="HTML", variable=self.output_format, value="html"
        )
        self.output_html_radio.grid(row=0, column=2, padx=10, pady=(10, 5), sticky="w")
        self.output_json_radio = ctk.CTkRadioButton(
            parent_frame, text="JSON", variable=self.output_format, value="json"
        )
        self.output_json_radio.grid(row=0, column=3, padx=10, pady=(10, 5), sticky="w")

        ctk.CTkLabel(parent_frame, text="Output Directory:").grid(
            row=1, column=0, padx=10, pady=5, sticky="w"
        )
        self.output_directory = ctk.StringVar()
        self.output_directory_entry = ctk.CTkEntry(
            parent_frame, textvariable=self.output_directory
        )
        self.output_directory_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.browse_output_dir_button = ctk.CTkButton(
            parent_frame, text="Browse", command=self.browse_output_directory
        )
        self.browse_output_dir_button.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        ctk.CTkButton(
            parent_frame, text="Clear", command=lambda: self.clear_paths("output")
        ).grid(row=1, column=3, padx=5, pady=5, sticky="w")

        parent_frame.columnconfigure(1, weight=1)

    def setup_advanced_tab(self, parent_frame):
        ctk.CTkLabel(parent_frame, text="OCR Languages (comma separated):").grid(
            row=0, column=0, padx=10, pady=(10, 5), sticky="w"
        )
        self.ocr_languages = ctk.StringVar(value="en")
        self.ocr_languages_entry = ctk.CTkEntry(
            parent_frame,
            textvariable=self.ocr_languages,
            placeholder_text="Enter language codes (e.g., en,de,fr)"
        )
        self.ocr_languages_entry.grid(
            row=0, column=1, padx=10, pady=(10, 5), sticky="ew"
        )

        # Add help button with tooltip for supported languages
        self.ocr_help_button = ctk.CTkButton(
            parent_frame,
            text="?",
            width=30,
            command=lambda: messagebox.showinfo(
                "Supported OCR Languages",
                "Supported language codes:\n\n"
                "English: en\n"
                "German: de\n"
                "French: fr\n"
                "Spanish: es\n"
                "Italian: it\n"
                "Portuguese: pt\n"
                "Russian: ru\n"
                "Chinese: zh\n"
                "Japanese: ja\n"
                "Korean: ko\n"
                "Arabic: ar\n"
                "Hindi: hi\n"
                "And many more...\n\n"
                "Separate multiple languages with commas."
            )
        )
        self.ocr_help_button.grid(row=0, column=2, padx=(0, 10), pady=(10, 5), sticky="w")

        ctk.CTkLabel(parent_frame, text="Table Mode:").grid(
            row=1, column=0, padx=10, pady=5, sticky="w"
        )
        self.table_mode = ctk.StringVar(value="fast")
        self.table_mode_fast_radio = ctk.CTkRadioButton(
            parent_frame, text="Fast", variable=self.table_mode, value="fast"
        )
        self.table_mode_fast_radio.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        self.table_mode_accurate_radio = ctk.CTkRadioButton(
            parent_frame, text="Accurate", variable=self.table_mode, value="accurate"
        )
        self.table_mode_accurate_radio.grid(row=1, column=2, padx=10, pady=5, sticky="w")

        parent_frame.columnconfigure(1, weight=1)

    def setup_control_frame(self, parent_frame):
        self.convert_button = ctk.CTkButton(
            parent_frame, text="Convert", command=self.convert
        )
        self.convert_button.grid(row=0, column=0, padx=10, pady=5, sticky="ew")

        self.cancel_button = ctk.CTkButton(
            parent_frame, text="Cancel", command=self.cancel_conversion, state=ctk.DISABLED
        )
        self.cancel_button.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        self.show_structure_button = ctk.CTkButton(
            parent_frame,
            text="Show Structure",
            command=self.show_structure,
            state=ctk.DISABLED,
        )
        self.show_structure_button.grid(row=0, column=2, padx=10, pady=5, sticky="ew")

        self.clear_output_button = ctk.CTkButton(
            parent_frame, text="Clear Output", command=self.clear_output
        )
        self.clear_output_button.grid(row=0, column=3, padx=10, pady=5, sticky="ew")

        self.help_button = ctk.CTkButton(
            parent_frame, text="Help", command=self.show_help
        )
        self.help_button.grid(row=0, column=4, padx=10, pady=5, sticky="ew")

        for i in range(5):
            parent_frame.columnconfigure(i, weight=1)
        return parent_frame

    def setup_output_text(self):
        self.output_scrollable_frame = ctk.CTkScrollableFrame(
            self.main_frame, corner_radius=0, border_width=0
        )
        self.output_scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Use standard tkinter color options for scrolledtext.ScrolledText
        self.output_text = scrolledtext.ScrolledText(
            self.output_scrollable_frame,
            wrap=tk.WORD,
            bg="white",
            fg="black",  # Set colors directly
            state=tk.DISABLED,
            insertbackground="black",  # Color of the cursor
        )
        self.output_text.pack(fill="both", expand=True)

    def setup_status(self):
        # Status label
        self.status_var = ctk.StringVar(value="Ready")
        self.status_label = ctk.CTkLabel(
            self.main_frame, textvariable=self.status_var, anchor="w"
        )
        self.status_label.pack(fill="x", padx=10, pady=(0, 5))

        # Terminal output window
        self.terminal_frame = ctk.CTkFrame(self.main_frame)
        self.terminal_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        self.terminal_text = scrolledtext.ScrolledText(
            self.terminal_frame,
            wrap=tk.WORD,
            bg="black",
            fg="white",
            state=tk.DISABLED,
            insertbackground="white",
            font=("Consolas", 10)
        )
        self.terminal_text.pack(fill="both", expand=True, padx=5, pady=5)

    def log_terminal(self, message: str):
        """Logs messages to the terminal output window."""
        self.terminal_text.configure(state=tk.NORMAL)
        self.terminal_text.insert(tk.END, f"> {message}\n")
        self.terminal_text.configure(state=tk.DISABLED)
        self.terminal_text.see(tk.END)

    def load_config(self):
        try:
            with open(self.CONFIG_FILE, "r") as f:
                config = json.load(f)
            self.last_input_dir = config.get("last_input_dir")
            self.last_output_dir = config.get("last_output_dir")
            self.input_format.set(config.get("input_format", "pdf"))
            self.output_format.set(config.get("output_format", "markdown"))
            self.ocr_languages.set(config.get("ocr_languages", "en"))
            self.table_mode.set(config.get("table_mode", "fast"))
        except FileNotFoundError:
            pass
        except json.JSONDecodeError:
            logging.warning("Invalid config file, using defaults.")

    def save_config(self):
        config = {
            "last_input_dir": self.last_input_dir,
            "last_output_dir": self.last_output_dir,
            "input_format": self.input_format.get(),
            "output_format": self.output_format.get(),
            "ocr_languages": self.ocr_languages.get(),
            "table_mode": self.table_mode.get(),
        }
        try:
            with open(self.CONFIG_FILE, "w") as f:
                json.dump(config, f)
        except Exception as e:
            logging.error(f"Error saving config: {e}")

    def browse_file(self):
        selected_format = self.input_format.get().lower()
        if selected_format in self.FILE_EXTENSIONS:
            file_type_name, file_pattern = self.FILE_EXTENSIONS[selected_format]
            title = f"Select {file_type_name}"
            filetypes = [(file_type_name, file_pattern)]
            initialdir = self.last_input_dir if self.last_input_dir else os.path.expanduser("~")
            filename = filedialog.askopenfilename(
                title=title, filetypes=filetypes, initialdir=initialdir
            )
            if filename:
                self.file_path.set(filename)
                self.last_input_dir = os.path.dirname(filename)
                self.save_config()

    def browse_directory(self):
        initialdir = self.last_input_dir if self.last_input_dir else os.path.expanduser("~")
        directory = filedialog.askdirectory(
            title="Select Directory", initialdir=initialdir
        )
        if directory:
            self.file_path.set(directory)
            self.last_input_dir = directory
            self.save_config()

    def browse_output_directory(self):
        initialdir = self.last_output_dir if self.last_output_dir else os.path.expanduser("~")
        directory = filedialog.askdirectory(
            title="Select Output Directory", initialdir=initialdir
        )
        if directory:
            self.output_directory.set(directory)
            self.last_output_dir = directory
            self.save_config()

    def clear_paths(self, path_type: str):
        if path_type == "file":
            self.file_path.set("")
        elif path_type == "output":
            self.output_directory.set("")

    def _validate_file_paths(self, input_path: str) -> Optional[str]:
        """Validates file paths and sets output directory."""
        output_dir = self.output_directory.get()
        if not output_dir:
            # If no output directory selected, use the last used output directory
            output_dir = self.last_output_dir if self.last_output_dir else os.path.dirname(input_path)
            self.output_directory.set(output_dir)

        # Ensure output directory exists and is writable
        try:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Test write permissions
            test_file = os.path.join(output_dir, ".permission_test")
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)
            
        except OSError as e:
            self.handle_error(f"Error accessing output directory {output_dir}: {e}")
            return None
        except Exception as e:
            self.handle_error(f"Permission denied for output directory {output_dir}: {e}")
            return None

        # Update last output directory
        self.last_output_dir = output_dir
        self.save_config()
        return output_dir

    def convert(self):
        # Reset state before starting new conversion
        if self.conversion_thread and self.conversion_thread.is_alive():
            self.cancel_conversion()
            self.conversion_thread.join(timeout=1.0)
            
        # Clear previous results
        self.conversion_result = None
        self.clear_output()
        
        input_path = self.file_path.get()
        if not input_path:
            messagebox.showwarning(
                "No File/Directory Selected",
                "Please select a file or directory to convert.",
            )
            return

        # Reset conversion state
        self.cancelled = False
        self.cancel_event.clear()
        self.log_progress("Starting conversion...")
        
        # Update UI state
        self.convert_button.configure(state=ctk.DISABLED, text="Converting...")
        self.cancel_button.configure(state=ctk.NORMAL)
        self.show_structure_button.configure(state=ctk.DISABLED)
        
        # Ensure output directory is set
        if not self.output_directory.get() and self.last_output_dir:
            self.output_directory.set(self.last_output_dir)

        if os.path.isfile(input_path):
            self.conversion_thread = threading.Thread(target=self.convert_single_file)
        elif os.path.isdir(input_path):
            self.conversion_thread = threading.Thread(target=self.convert_batch)
        else:
            messagebox.showerror(
                "Invalid Path", "The specified path is neither a file nor a directory."
            )
            self.convert_button.configure(state=ctk.NORMAL)
            return

        self.conversion_thread.start()

    def log_progress(self, message):
        """Logs progress messages to terminal and status bar."""
        logging.info(message)
        self.status_var.set(message)
        self.log_terminal(message)

    def convert_single_file(self):
        input_path = self.file_path.get()
        output_dir = self._validate_file_paths(input_path)
        if output_dir is None:
            self.root.after(0, self.convert_button.configure, {"state": ctk.NORMAL})
            return

        try:
            converter = self._configure_converter()
            self.log_progress(f"Starting conversion of {os.path.basename(input_path)}")
            
            # Clear previous output
            self.clear_output()
            
            # Add detailed logging
            self.log_terminal(f"Input file: {input_path}")
            self.log_terminal(f"Output directory: {output_dir}")
            self.log_terminal(f"Output format: {self.output_format.get()}")
            
            conversion_results = converter.convert_all([input_path], raises_on_error=False)
            
            # Ensure thread cleanup
            self.conversion_thread = None
            
            if not conversion_results:
                self.log_terminal("Error: No conversion results returned")
                self.handle_error("Conversion failed - no results returned")
                return

            for result in conversion_results:
                if result.status == "success":
                    self.conversion_result = result
                    output_content = self.get_output_content(self.conversion_result)
                    output_extension = ".md" if self.output_format.get() == "markdown" else f".{self.output_format.get()}"
                    output_filename = Path(output_dir) / Path(os.path.basename(input_path)).with_suffix(output_extension)

                    try:
                        with open(output_filename, "w", encoding="utf-8") as f:
                            f.write(output_content)
                            
                        self.log_terminal(f"Successfully saved output to: {output_filename}")
                        self.conversion_result = result
                        self.root.after(0, self.update_output, f"Conversion complete: {output_filename}")
                        self.root.after(0, self.status_var.set, "Conversion complete")
                        if self.conversion_result:
                            self.root.after(0, self.show_structure_button.configure, {"state": ctk.NORMAL})
                        
                    except IOError as e:
                        self.log_terminal(f"Error writing output file: {str(e)}")
                        self.handle_error(f"Failed to write output file: {str(e)}")
                        
                else:
                    error_msg = f"Error converting {input_path}: {result.error_message}"
                    self.log_terminal(error_msg)
                    self.root.after(0, self.update_output, error_msg)
                    self.handle_error(result.error_message)

        except Exception as e:
            error_msg = f"Conversion error: {str(e)}"
            self.log_terminal(error_msg)
            self.log_terminal(traceback.format_exc())
            self.handle_error(e)
        finally:
            # Ensure proper cleanup
            self.conversion_thread = None
            self.cancel_event.clear()
            
            # Update UI state
            self.root.after(0, lambda: [
                self.convert_button.configure(state=ctk.NORMAL, text="Convert"),
                self.cancel_button.configure(state=ctk.DISABLED),
                self.show_structure_button.configure(state=ctk.NORMAL if self.conversion_result else ctk.DISABLED)
            ])
            
            # Save last output directory
            if self.output_directory.get():
                self.last_output_dir = self.output_directory.get()
                self.save_config()

    def _configure_converter(self) -> DocumentConverter:
        """Configures the DocumentConverter with the selected options."""
        input_format_str = self.input_format.get().lower()
        
        # Validate OCR languages
        ocr_languages = []
        if self.ocr_languages.get().strip():
            try:
                ocr_languages = [lang.strip() for lang in self.ocr_languages.get().split(",")]
                # Validate each language code
                supported_languages = ["en", "de", "fr", "es", "it", "pt", "ru", "zh", "ja", "ko", "ar", "hi"]
                invalid_languages = [lang for lang in ocr_languages if lang not in supported_languages]
                
                if invalid_languages:
                    raise ValueError(f"Unsupported OCR languages: {', '.join(invalid_languages)}")
                    
            except Exception as e:
                self.handle_error(f"Invalid OCR language configuration: {str(e)}")
                raise

        table_mode = (
            TableFormerMode.FAST
            if self.table_mode.get() == "fast"
            else TableFormerMode.ACCURATE
        )

        if input_format_str == "pdf":
            pdf_format_options = PdfFormatOption(
                pipeline_options=PdfPipelineOptions(
                    do_ocr=True if ocr_languages else False,
                    ocr_lang=",".join(ocr_languages),
                    table_structure_options=TableStructureOptions(mode=table_mode)
                )
            )
            converter = DocumentConverter(
                format_options={InputFormat.PDF: pdf_format_options}
            )
        else:
            converter = DocumentConverter(
                allowed_formats=[InputFormat(input_format_str)]
            )
        return converter

    def convert_batch(self):
        input_directory = self.file_path.get()
        input_format_str = self.input_format.get().lower()
        converter = self._configure_converter()

        try:
            file_list = list(Path(input_directory).glob(f"*.{input_format_str}"))
        except OSError as e:
            self.handle_error(f"Error accessing directory {input_directory}: {e}")
            self.root.after(0, self.convert_button.configure, {"state": ctk.NORMAL})
            return

        total_files = len(file_list)
        if not file_list:
            self.root.after(
                0,
                messagebox.showwarning,
                "No Files Found",
                f"No {input_format_str.upper()} files found in the selected directory.",
            )
            return

        self.root.after(0, self.convert_button.configure, {"state": ctk.DISABLED})
        self.root.after(0, self.cancel_button.configure, {"state": ctk.NORMAL})
        self.log_progress(f"Starting batch conversion of {total_files} files")

        for index, file in enumerate(file_list, start=1):
            if self.cancel_event.is_set():
                self.log_progress("Batch conversion cancelled")
                break

            self.log_progress(f"Converting {file.name} ({index}/{total_files})")

            try:
                output_dir = self._validate_file_paths(str(file))
                conversion_results = converter.convert_all([str(file)], raises_on_error=False)

                for result in conversion_results:
                    if result.status == "success":
                        self.conversion_result = result
                        output_content = self.get_output_content(self.conversion_result)
                        output_extension = ".md" if self.output_format.get() == "markdown" else f".{self.output_format.get()}"
                        output_filename = Path(output_dir) / file.with_suffix(output_extension)

                        with open(output_filename, "w", encoding="utf-8") as f:
                            f.write(output_content)

                        self.root.after(
                            0,
                            self.update_output,
                            f"Converted: {file.name} -> {output_filename.name}",
                        )
                    else:
                        self.root.after(0, self.update_output, f"Error converting {file.name}: {result.error_message}")
                        self.handle_error(result.error_message)

            except Exception as e:
                self.handle_error(e)
                self.root.after(0, self.update_output, f"Error converting {file.name}: {str(e)}")

        if not self.cancel_event.is_set():
            self.root.after(0, self.status_var.set, "Batch conversion complete")
            self.root.after(0, self.update_progress_bar, 1.0)

        self.root.after(0, self.convert_button.configure, {"state": ctk.NORMAL, "text": "Convert"})
        self.root.after(0, self.cancel_button.configure, {"state": ctk.DISABLED})
        self.root.after(
            0, self.show_structure_button.configure, {"state": ctk.NORMAL}
        )

    def get_output_content(self, conversion_result):
        output_format = self.output_format.get()
        if output_format == "markdown":
            return conversion_result.document.export_to_markdown()
        elif output_format == "html":
            return conversion_result.document.export_to_html()
        elif output_format == "json":
            return conversion_result.document.export_to_json(indent=2)
        else:
            raise ValueError(f"Unsupported output format: {output_format}")

    def cancel_conversion(self):
        self.cancel_event.set()
        self.status_var.set("Cancelling...")
        self.convert_button.configure(state=ctk.DISABLED)
        self.cancel_button.configure(state=ctk.DISABLED)

        # Forcefully terminate the conversion thread if it's running
        if self.conversion_thread and self.conversion_thread.is_alive():
            try:
                # Clean up any resources used by the thread
                # You don't have a 'converter' attribute that needs cleanup in the class
                # if hasattr(self, 'converter'):
                #     self.converter.cleanup()

                # Wait for thread to finish with timeout
                self.conversion_thread.join(timeout=1.0)

                if self.conversion_thread.is_alive():
                    # If thread is still alive after timeout, terminate it
                    import ctypes
                    id = self.conversion_thread.ident
                    if id:
                        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(ctypes.c_long(id), ctypes.py_object(SystemExit))
                        if res == 0:
                            raise ValueError("Invalid thread ID")
                        elif res != 1:
                            ctypes.pythonapi.PyThreadState_SetAsyncExc(ctypes.c_long(id), None)
                            raise SystemError("PyThreadState_SetAsyncExc failed")

                self.status_var.set("Conversion cancelled")
                self.update_output("Conversion cancelled by user")
            except Exception as e:
                self.handle_error(f"Error cancelling conversion: {e}")

    def update_output(self, message):
        self.output_text.configure(state=tk.NORMAL)
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.configure(state=tk.DISABLED)
        self.output_text.see(tk.END)

    def clear_output(self):
        self.output_text.configure(state=tk.NORMAL)
        self.output_text.delete("1.0", tk.END)
        self.output_text.configure(state=tk.DISABLED)

    def show_structure(self):
        if self.conversion_result and hasattr(self.conversion_result, 'document'):
            try:
                output_format = self.output_format.get()
                if output_format == "markdown":
                    output_content = self.conversion_result.document.export_to_markdown()
                elif output_format == "html":
                    output_content = self.conversion_result.document.export_to_html()
                elif output_format == "json":
                    output_content = self.conversion_result.document.export_to_json(indent=2)
                
                self.output_text.configure(state=tk.NORMAL)
                self.output_text.delete("1.0", tk.END)
                self.output_text.insert(tk.END, output_content)
                self.output_text.configure(state=tk.DISABLED)
                self.output_text.see(tk.END)
            except Exception as e:
                self.handle_error(f"Error displaying content: {e}")
        else:
            messagebox.showinfo("No Conversion", "Please convert a document first.")

    def show_help(self):
        help_text = """
    Enhanced PDF Converter Help:

    1. Input Tab:
       - Select the input format (PDF, DOCX, PPTX, HTML, or Image)
       - Choose a file or directory for batch conversion

    2. Output Tab:
       - Select the output format (Markdown, HTML, or JSON)
       - Choose an output directory (optional)

    3. Advanced Tab:
       - Set OCR languages (comma-separated language codes, e.g., en, de, fr, ...)
       - Choose table extraction mode (Fast or Accurate)

    4. Convert:
       - Click 'Convert' to start the conversion process
       - For batch conversion, select a directory in the Input tab

    5. Show Structure:
       - After conversion, click to view the document's structure

    6. Clear Output:
       - Clears the output text area

    Note: For best results, ensure input files are readable and not corrupted.
    """
        messagebox.showinfo("Help", help_text)

    def handle_error(self, error):
        error_message = (
            f"An error occurred: {str(error)}\n\nTraceback:\n{traceback.format_exc()}"
        )
        logging.error(error_message)
        self.root.after(0, self.update_output, error_message)
        self.root.after(0, self.status_var.set, "Error occurred")
        self.root.after(
            0, self.show_structure_button.configure, {"state": ctk.DISABLED}
        )

if __name__ == "__main__":
    root = ctk.CTk()
    app = PDFConverterApp(root)
    root.mainloop()
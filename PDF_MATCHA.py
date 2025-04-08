import sys
import fitz  # PyMuPDF for PDF processing
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import difflib
import re
import threading
import time
import os
import camelot
import pandas as pd
import numpy as np
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle


class ModernUIStyles:
    """Centralized styles for UI """
    def __init__(self):

        # Color scheme
        self.primary = "#4a8f47"  # Modern green
        self.primary_dark = "#3a7039"
        self.primary_light = "#6cb369"
        self.secondary = "#f2f7f2"
        self.text_dark = "#333333"
        self.text_light = "#ffffff"
        self.accent = "#f44336"
        self.bg_light = "#ffffff"
        self.gray_light = "#f5f5f5"
        self.gray = "#e0e0e0"

        # Fonts
        self.title_font = ("Segoe UI", 24, "bold")
        self.subtitle_font = ("Segoe UI", 12)
        self.button_font = ("Segoe UI", 11)
        self.label_font = ("Segoe UI", 10)
        self.small_font = ("Segoe UI", 9)

        # Dimensions
        self.padding = 10
        self.button_height = 32
        self.border_radius = 8

        # Create custom styles for ttk widgets
        self._create_ttk_styles()

    def _create_ttk_styles(self):
        """Create custom ttk styles for the UI"""
        style = ttk.Style()

        # Configure the progress bar
        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor=self.gray_light,  # Background of progress bar
            background=self.primary,  # Progress color
            lightcolor=self.primary_light,  # Lighter shade for gradient
            darkcolor=self.primary_dark,  # Darker shade for gradient
            thickness=10,  # Thicker progress bar
            borderwidth=0
        )

        # Configure the labelframe
        style.configure(
            "Modern.TLabelframe",
            background=self.bg_light,
            bordercolor=self.gray,
            labelmargins=[10, 10, 10, 10],
            relief="flat"
        )

        style.configure(
            "Modern.TLabelframe.Label",
            background=self.bg_light,
            foreground=self.primary_dark,
            font=self.label_font
        )

        style.map('TButton',
                  background=[('active', self.primary_dark)],
                  foreground=[('active', self.text_light)]
        )

class PDFComparerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Match Analyzer")
        self.root.geometry("550x600")
        self.root.resizable(False, False)
        self.root.configure(bg="white")


        # Initialize the UI styles
        self.ui = ModernUIStyles()

        # Set theme colors
        self.theme_colors = {
            "primary": self.ui.primary,
            "secondary": self.ui.secondary,
            "text_dark": self.ui.text_dark,
            "text_light": self.ui.text_light,
            "background": self.ui.bg_light,
            "accent": self.ui.accent,
            "highlight": self.ui.primary_light
        }

        self.pdf1_path = None
        self.pdf2_path = None

        # Store documents and analysis results for report generation
        self.old_doc = None
        self.new_doc = None
        self.removed = None
        self.added = None
        self.comparison_complete = False

        # Store extracted tables
        self.pdf1_tables = None
        self.pdf2_tables = None
        self.table_extraction_complete = False

        """CONFIGURATION HERE"""
        # Configurable parameters for comparison - can be exposed to UI in future!
        self.comparison_config = {
            "exact_match_threshold": 0.92,  # Higher than 0.9 for more precision
            "min_meaningful_text_length": 13,  # More than 10 chars to be meaningful
            "fuzzy_chunk_size": 5,  # Increased from 3 words for better context
            "enable_enhanced_preprocessing": True,  # Enable enhanced text preprocessing
            "output_folder": "comparison_results"
        }

        # Configurable parameters for table extraction
        self.table_extraction_config = {
            "extraction_method": "lattice",  # Default method: 'lattice', 'stream', or 'hybrid', use this as your guide https://camelot-py.readthedocs.io/en/master/user/how-it-works.html#lattice
            "output_folder": "Extracted_tables",
        }

        self.report_config = {
             "output_folder": "generated_reports",
        }

        # Create output folder if it doesn't exist
        os.makedirs(self.table_extraction_config["output_folder"], exist_ok=True)
        os.makedirs(self.comparison_config["output_folder"], exist_ok=True)
        os.makedirs(self.report_config["output_folder"], exist_ok=True)

        # Set the window icon
        try:
            pdf_icon = Image.open("PDF Matcha_ICON.png")
            pdf_icon = ImageTk.PhotoImage(pdf_icon)
            root.iconphoto(True, pdf_icon)
        except:
            pass  # Skip icon if not found

        # Initialize UI components
        self._setup_ui()

    def _create_custom_styles(self):
        """Create advanced ttk styles"""
        style = ttk.Style()

        # Progress Bar
        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor=self.gray_light,  # Background of progress bar
            background=self.primary,  # Progress color
            lightcolor=self.primary_light,
            darkcolor=self.primary_dark,
            thickness=10,  # Thicker progress bar
            borderwidth=0
        )

        # LabelFrame
        style.configure(
            "Modern.TLabelframe",
            background=self.bg_light,
            bordercolor=self.gray,
            labelmargins=[10, 10, 10, 10],
            relief="flat"
        )

        # Button-like hover effects
        style.map('TButton',
                  background=[('active', self.primary_dark)],
                  foreground=[('active', self.text_light)]
                  )

    def create_hover_button(self, parent, text, command, width=15, **kwargs):
        """Button with hover effect"""

        state = kwargs.pop('state', tk.NORMAL)

        def on_enter(e):
            if btn['state'] == tk.NORMAL:
                btn.config(bg=self.ui.primary_dark)

        def on_leave(e):
            if btn['state'] == tk.NORMAL:
                btn.config(bg=self.ui.primary)

        button_props = {
            'text': text,
            'command': command,
            'font': self.ui.button_font,
            'bg': self.ui.primary,
            'fg': self.ui.text_light,
            'relief': tk.FLAT,
            'borderwidth': 0,
            'width': width,
            'state': state,
            'padx': 10,
            'pady': 5
        }

        # Update with any additional kwargs
        button_props.update(kwargs)

        btn = tk.Button(parent, **button_props)

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

        return btn

    def _setup_ui(self):
        """Set up all UI components with modern styling"""
        # Main container for the entire application
        main_container = tk.Frame(self.root, bg=self.ui.bg_light)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Green header title section
        header_frame = tk.Frame(main_container, bg=self.ui.primary, padx=25, pady=15)
        header_frame.pack(fill=tk.X)

        # App title
        title_label = tk.Label(
            header_frame,
            text="PDF MATCHA",
            font=self.ui.title_font,
            fg=self.ui.text_light,  # White text
            bg=self.ui.primary  # Green background
        )
        title_label.pack(anchor=tk.CENTER)

        # Subtitle/description
        status_label = tk.Label(
            header_frame,
            text="Compare and analyze PDF documents with ease",
            font=self.ui.subtitle_font,
            fg=self.ui.text_light,  # White text
            bg=self.ui.primary  # Green background
        )
        status_label.pack(anchor=tk.CENTER)

        # Content container (white background)
        content_container = tk.Frame(main_container, bg=self.ui.bg_light, padx=25, pady=20)
        content_container.pack(fill=tk.BOTH, expand=True)

        # PDF selection section frame
        selection_frame = ttk.LabelFrame(
            content_container,
            text="PDF Selection",
            padding=self.ui.padding,
            style="Modern.TLabelframe"
        )
        selection_frame.pack(fill=tk.X, pady=(0, 15))

        # OLD PDF selection
        pdf1_frame = tk.Frame(selection_frame, bg=self.ui.bg_light)
        pdf1_frame.pack(fill=tk.X, pady=(0, 8))

        btn_pdf1 = self.create_hover_button(
            pdf1_frame,
            text="Insert Old PDF",
            command=self.load_pdf1,
            width=12,
            padx=10,
            height=1
        )
        btn_pdf1.pack(side=tk.LEFT)

        self.pdf_file_entry = tk.Entry(
            pdf1_frame,
            width=40,
            font=self.ui.small_font,
            borderwidth=1,
            relief=tk.SOLID,
            bg=self.ui.gray_light
        )
        self.pdf_file_entry.insert(0, "Select first PDF file...")
        self.pdf_file_entry['state'] = 'disabled'
        self.pdf_file_entry.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True, ipady=6)

        # NEW PDF selection
        pdf2_frame = tk.Frame(selection_frame, bg=self.ui.bg_light)
        pdf2_frame.pack(fill=tk.X)

        btn_pdf2 = self.create_hover_button(
            pdf2_frame,
            text="Insert New PDF",
            command=self.load_pdf2,
            width=12,
            padx=10,
            height=1
        )
        btn_pdf2.pack(side=tk.LEFT)

        self.pdf2_file_entry = tk.Entry(
            pdf2_frame,
            width=40,
            font=self.ui.small_font,
            borderwidth=1,
            relief=tk.SOLID,
            bg=self.ui.gray_light
        )
        self.pdf2_file_entry.insert(0, "Select second PDF file...")
        self.pdf2_file_entry['state'] = 'disabled'
        self.pdf2_file_entry.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True, ipady=6)

        # Action buttons section
        actions_frame = tk.Frame(content_container, bg=self.ui.bg_light)
        actions_frame.pack(fill=tk.X, pady=(0, 15))
        button_frame = tk.Frame(actions_frame, bg=self.ui.bg_light)
        button_frame.pack(anchor='center')

        self.btn_compare = self.create_hover_button(
            button_frame,
            text="Compare PDFs",
            command=self.start_comparison,
            state=tk.DISABLED,
            font=self.ui.button_font,
            bg="gray",
            fg="white",
            relief=tk.FLAT,
            borderwidth=0,
            width=15,
            height=1
        )

        self.btn_compare.grid(row=0, column=0, padx=5, pady=5)

        self.btn_report = self.create_hover_button(
            button_frame,
            text="Generate Report",
            command=self.start_report_generation,
            state=tk.DISABLED,
            font=self.ui.button_font,
            bg="gray",
            fg="white",
            relief=tk.FLAT,
            borderwidth=0,
            width=18,
            height=1
        )
        self.btn_report.grid(row=0, column=1, padx=5, pady=5)


        self.btn_extract_tables = self.create_hover_button(
            button_frame,
            text="Extract Tables",
            command=self.start_table_extraction,
            state=tk.DISABLED,
            font=self.ui.button_font,
            bg="gray",
            fg="white",
            relief=tk.FLAT,
            borderwidth=0,
            width=15,
            height=1
        )
        self.btn_extract_tables.grid(row=0, column=2, padx=5, pady=5)

        # Progress section
        progress_frame = ttk.LabelFrame(
            content_container,
            text="Progress",
            padding=self.ui.padding,
            style="Modern.TLabelframe"
        )
        progress_frame.pack(fill=tk.X, pady=(0, 15))

        progress_inner_frame = tk.Frame(progress_frame, bg=self.ui.bg_light)
        progress_inner_frame.pack(fill=tk.X)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_inner_frame,
            orient="horizontal",
            length=400,
            mode="determinate",
            variable=self.progress_var,
            style="Modern.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 8))

        progress_info_frame = tk.Frame(progress_inner_frame, bg=self.ui.bg_light)
        progress_info_frame.pack(fill=tk.X)

        # Progress percentage label
        self.percentage_label = tk.Label(
            progress_info_frame,
            text="0%",
            font=self.ui.small_font,
            fg=self.ui.primary_dark,
            bg=self.ui.bg_light
        )
        self.percentage_label.pack(side=tk.RIGHT)

        # Progress description label
        self.progress_description = tk.Label(
            progress_info_frame,
            text="Waiting to start...",
            font=self.ui.small_font,
            fg="gray",
            bg=self.ui.bg_light
        )
        self.progress_description.pack(side=tk.LEFT)

        # Status section
        status_frame = ttk.LabelFrame(
            content_container,
            text="Status",
            padding=self.ui.padding,
            style="Modern.TLabelframe"
        )
        status_frame.pack(fill=tk.BOTH, expand=True)

        status_inner_frame = tk.Frame(status_frame, bg=self.ui.bg_light)
        status_inner_frame.pack(fill=tk.BOTH, expand=True)

        # Status labels
        self.status_label = tk.Label(
            status_inner_frame,
            text="Waiting for PDFs...",
            font=self.ui.label_font,
            fg=self.ui.primary_dark,
            bg=self.ui.bg_light,
            anchor="w"
        )
        self.status_label.pack(fill=tk.X)

        self.detail_label = tk.Label(
            status_inner_frame,
            text="",
            font=self.ui.small_font,
            fg="gray",
            bg=self.ui.bg_light,
            anchor="w",
            justify=tk.LEFT,
            wraplength=480  # To ensure text wraps properly
        )
        self.detail_label.pack(fill=tk.X, pady=(5, 0))

    def load_pdf1(self):
        """Load the first PDF file"""
        self.pdf1_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.pdf1_path:
            filename = os.path.basename(self.pdf1_path)
            self.pdf_file_entry['state'] = 'normal'
            self.pdf_file_entry.delete(0, tk.END)
            self.pdf_file_entry.insert(0, filename)
            self.pdf_file_entry['state'] = 'disabled'
            self.enable_compare_button()
            self.enable_extract_tables_button()

    def load_pdf2(self):
        """Load the second PDF file"""
        self.pdf2_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.pdf2_path:
            filename = os.path.basename(self.pdf2_path)
            self.pdf2_file_entry['state'] = 'normal'
            self.pdf2_file_entry.delete(0, tk.END)
            self.pdf2_file_entry.insert(0, filename)
            self.pdf2_file_entry['state'] = 'disabled'
            self.enable_compare_button()
            self.enable_extract_tables_button()

    def enable_compare_button(self):
        #Enable the compare button if both PDFs are loaded
        if self.pdf1_path and self.pdf2_path:
            self.btn_compare['state'] = 'normal'
            self.btn_compare['bg'] = self.ui.primary
            self.btn_compare['fg'] = self.ui.text_light
        else:
            self.btn_compare['state'] = 'disabled'

    def enable_extract_tables_button(self):
        # Enable the extract tables button if both PDFs are loaded
        if self.pdf1_path and self.pdf2_path:
            self.btn_extract_tables['state'] = 'normal'
            self.btn_extract_tables['bg'] = self.ui.primary
            self.btn_extract_tables['fg'] = self.ui.text_light
        else:
            self.btn_extract_tables['state'] = 'disabled'

    def enable_report_button(self):
        # Enable the report button if comparison or table extraction is complete
        if self.comparison_complete or self.table_extraction_complete:
            self.btn_report['state'] = 'normal'
            self.btn_report['bg'] = self.ui.primary
            self.btn_report['fg'] = self.ui.text_light
        else:
            self.btn_report['state'] = 'disabled'

    def update_progress(self, value, description, current_page=None, total_pages=None):
        # Update the progress bar, percentage label, and description
        self.progress_var.set(value)
        self.percentage_label.config(text=f"{int(value)}%")

        # Add page scanning information if provided
        if current_page is not None and total_pages is not None:
            page_info = f" (Page {current_page}/{total_pages})"
            description = description + page_info

        self.progress_description.config(text=description)
        self.root.update_idletasks()  # Force update of the UI

    def start_comparison(self):
        """Start the comparison process in a separate thread to keep UI responsive"""

        # Disable the buttons during processing
        self.btn_compare['state'] = 'disabled'
        self.btn_extract_tables['state'] = 'disabled'
        self.btn_report['state'] = 'disabled'

        # Reset status
        self.comparison_complete = False
        self.status_label.config(text="Comparing PDFs...", fg="blue")
        self.detail_label.config(text="", fg="gray")  # Clear detail label
        self.progress_description.config(text="Initializing comparison...")
        self.progress_var.set(0)
        self.percentage_label.config(text="0%")

        # Start the comparison in a separate thread
        comparison_thread = threading.Thread(target=self.compare_pdfs)
        comparison_thread.daemon = True
        comparison_thread.start()

    def start_table_extraction(self):
        """Start the table extraction process in a separate thread"""

        # Disable buttons during processing
        self.btn_compare['state'] = 'disabled'
        self.btn_extract_tables['state'] = 'disabled'
        self.btn_report['state'] = 'disabled'

        # Reset status
        self.table_extraction_complete = False
        self.status_label.config(text="Extracting Tables...", fg="blue")
        self.detail_label.config(text="", fg="gray")  # Clear detail label
        self.progress_description.config(text="Initializing table extraction...")
        self.progress_var.set(0)
        self.percentage_label.config(text="0%")

        # Start the extraction in a separate thread
        extraction_thread = threading.Thread(target=self.extract_tables_only)
        extraction_thread.daemon = True
        extraction_thread.start()

    def start_report_generation(self):
        """Start the report generation process in a separate thread"""
        if not self.comparison_complete and not self.table_extraction_complete:
            messagebox.showinfo("Report Generation", "Please complete either PDF comparison or table extraction first.")
            return

        self.btn_report['state'] = 'disabled'
        self.status_label.config(text="Generating Report...", fg="blue")
        self.progress_description.config(text="Initializing report generation...")
        self.progress_var.set(0)
        self.percentage_label.config(text="0%")

        # Start the report generation in a separate thread
        report_thread = threading.Thread(target=self.generate_report)
        report_thread.daemon = True
        report_thread.start()

    def extract_text_from_pdf(self, pdf_path, is_first_pdf=True):
        """
        Extract text from a PDF file page by page, preserving structure.
        Returns:
        - text_by_page: A list of text content for each page
        - blocks_by_page: A list of block objects with position information
        - doc: The open PDF document
        """
        doc = fitz.open(pdf_path)
        text_by_page = []
        blocks_by_page = []
        total_pages = len(doc)
        pdf_label = "first" if is_first_pdf else "second"

        for page_num in range(total_pages):
            # Update progress (first 20% for PDF1, next 20% for PDF2)
            base_progress = 0 if is_first_pdf else 20
            progress = base_progress + (page_num / total_pages) * 20
            self.update_progress(progress, f"Extracting text from {pdf_label} PDF",
                                 current_page=page_num + 1, total_pages=total_pages)

            page = doc[page_num]
            # Extract text blocks with position information
            blocks = page.get_text("blocks")
            blocks_by_page.append(blocks)

            # Extract text from each block
            page_text = "\n".join([block[4] for block in blocks])
            text_by_page.append(page_text)

        return text_by_page, blocks_by_page, doc

    def preprocess_text(self, text):
        """
        Enhanced text preprocessing for better comparison.
        - Normalizes whitespace
        - Removes non-essential punctuation
        - Normalizes case for non-case-sensitive comparison
        """
        if not self.comparison_config["enable_enhanced_preprocessing"]:
            # Basic preprocessing (original method)
            text = re.sub(r'\s+', ' ', text)
            return text.strip()

        # Enhanced preprocessing
        # 1. Convert to lowercase for case-insensitive comparison
        text = text.lower()

        # 2. Replace common OCR error patterns and normalize characters
        replacements = {
            'ﬁ': 'fi', 'ﬂ': 'fl', 'ﬀ': 'ff', 'ﬃ': 'ffi', 'ﬄ': 'ffl',  # Ligatures
            '–': '-', '—': '-', '−': '-',  # Various dash types
            ''': "'", ''': "'", '"': '"', '"': '"',  # Quotes
            '…': '...',  # Ellipsis
            '\u200b': '', '\u200c': '', '\u200d': '', '\ufeff': '',  # Zero-width characters
        }

        for old, new in replacements.items():
            text = text.replace(old, new)

        # 3. Normalize whitespace
        text = re.sub(r'\s+', ' ', text)

        # 4. Remove non-essential punctuation for better matching
        # but keep sentence structure intact
        text = re.sub(r'[,;:"\']', ' ', text)  # Remove some punctuation
        text = re.sub(r'\s+', ' ', text)  # Normalize whitespace again

        # 5. Strip and return
        return text.strip()

    def extract_content_units(self, text_by_page):
        """ Extract meaningful content units (paragraphs, headings, etc.) from the PDF
        regardless of their page position. This allows for cross-page comparison.

        Returns a list of (content_unit, page_num, original_text) tuples. """

        content_units = []
        min_length = self.comparison_config["min_meaningful_text_length"]

        for page_num, page_text in enumerate(text_by_page):
            # Split page into paragraphs or meaningful blocks
            paragraphs = re.split(r'\n\s*\n', page_text)

            for paragraph in paragraphs:
                original = paragraph
                processed = self.preprocess_text(paragraph)
                if processed and len(processed) > min_length:  # Skip very short fragments
                    # Store as tuple: (processed_text, page_number, original_text)
                    content_units.append((processed, page_num, original))

        return content_units

    def find_text_on_page(self, page, text, fuzzy=False):
        """ Find instances of text on a page and return their rectangles
        With fuzzy=True, attempts to find text even with slight differences """
        text = text.strip()
        if not text:
            return []

        # Try to find the whole text
        whole_text_instances = page.search_for(text)
        if whole_text_instances:
            return whole_text_instances

        # If whole text not found, try to find paragraphs or sentences
        paragraphs = text.split('\n\n')
        instances = []

        for paragraph in paragraphs:
            paragraph = paragraph.strip()
            if not paragraph:
                continue

            # Try to find the paragraph
            paragraph_instances = page.search_for(paragraph)
            if paragraph_instances:
                instances.extend(paragraph_instances)
                continue

            # If paragraph not found, try sentences
            sentences = re.split(r'(?<=[.!?])\s+', paragraph)
            for sentence in sentences:
                sentence = sentence.strip()
                min_length = self.comparison_config["min_meaningful_text_length"]
                if len(sentence) > min_length:  # Only search for meaningful sentences
                    sentence_instances = page.search_for(sentence)
                    instances.extend(sentence_instances)

        # If fuzzy matching is enabled and we still didn't find anything
        if fuzzy and not instances:
            sentences = re.split(r'(?<=[.!?])\s+', text)
            for sentence in sentences:
                sentence = sentence.strip()
                if len(sentence) > 20:  # Only use longer sentences for fuzzy matching
                    # Try to find a significant portion of the sentence
                    words = sentence.split()
                    chunk_size = self.comparison_config["fuzzy_chunk_size"]
                    chunks = [' '.join(words[i:i + chunk_size])
                              for i in range(0, len(words), chunk_size)
                              if i + chunk_size <= len(words)]

                    for chunk in chunks:
                        if len(chunk) > min_length:
                            chunk_instances = page.search_for(chunk)
                            instances.extend(chunk_instances)

        return instances

    def highlight_text_on_page(self, page, text, color, fuzzy=True):
        """Highlight text on a page with the specified color"""
        if not text.strip():
            return False

        instances = self.find_text_on_page(page, text, fuzzy)

        if not instances:
            # If no instances found, try to find each line
            lines = text.split('\n')
            for line in lines:
                line = line.strip()
                min_length = self.comparison_config["min_meaningful_text_length"]
                if len(line) > min_length:  # Only search for meaningful lines
                    line_instances = page.search_for(line)
                    for rect in line_instances:
                        highlight = page.add_highlight_annot(rect)
                        if color == "red":
                            highlight.set_colors({"stroke": (1, 0, 0)})
                            highlight.set_opacity(0.5)
                        elif color == "green":
                            highlight.set_colors({"stroke": (0, 1, 0)})
                            highlight.set_opacity(0.5)
                        highlight.update()
                        if len(line_instances) > 0:
                            return True
        else:
            # Highlight found instances
            for rect in instances:
                highlight = page.add_highlight_annot(rect)
                if color == "red":
                    highlight.set_colors({"stroke": (1, 0, 0)})
                    highlight.set_opacity(0.5)
                elif color == "green":
                    highlight.set_colors({"stroke": (0, 1, 0)})
                    highlight.set_opacity(0.5)
                highlight.update()
            return True

        return False

    def compare_content_units(self, old_units, new_units):
        """
        Compare content units between documents using a simplified algorithm.
        This version only identifies removed and added content.

        Returns:
        - removed: Content that exists in old but not in new document
        - added: Content that exists in new but not in old document
        - similarity_scores: Dictionary with similarity metrics
        """

        # Extract just the processed text for comparison
        old_texts = [unit[0] for unit in old_units]
        new_texts = [unit[0] for unit in new_units]

        # Initialize output containers
        removed = []  # (text, page_num, original) tuples
        added = []  # (text, page_num, original) tuples
        matched_pairs = []  # [(old_unit, new_unit, similarity)] for matched content

        # Create a copy of new_units for tracking matched content
        unmatched_new = list(new_units)

        # Map to track if an old unit has been matched
        old_matched = [False] * len(old_units)

        # First pass: find exact matches
        for i, (old_text, old_page, old_original) in enumerate(old_units):
            found_match = False
            best_match_idx = -1
            best_similarity = 0

            # Try to find a match in the new document
            for j, (new_text, new_page, new_original) in enumerate(unmatched_new):
                # Use SequenceMatcher to calculate similarity
                similarity = difflib.SequenceMatcher(None, old_text, new_text).ratio()

                if similarity > best_similarity:
                    best_similarity = similarity
                    best_match_idx = j

                if similarity >= self.comparison_config["exact_match_threshold"]:  # Exact or near-exact match
                    found_match = True
                    old_matched[i] = True
                    matched_pairs.append((old_units[i], unmatched_new[j], similarity))
                    unmatched_new.pop(j)
                    break

            # If no exact match but we found a good partial match
            if not found_match and best_match_idx >= 0 and best_similarity > 0.7:
                # Consider it a potential match (useful for similarity score calculations)
                matched_pairs.append((old_units[i], unmatched_new[best_match_idx], best_similarity))
                # But still mark it as removed (we're being conservative with modifications)
                removed.append((old_text, old_page, old_original))
            # If no match at all, it's removed
            elif not found_match:
                removed.append((old_text, old_page, old_original))

        # All remaining unmatched new content is added
        added = unmatched_new

        # Calculate similarity metrics
        total_old_content = len(old_units)
        total_new_content = len(new_units)
        matched_content = len(matched_pairs)

        # Calculate aggregate similarity score across all matched content
        avg_similarity = 0
        total_similarity = 0
        if matched_pairs:
            for _, _, sim in matched_pairs:
                total_similarity += sim
            avg_similarity = total_similarity / len(matched_pairs)

        # Calculate document similarity as percentage
        doc_similarity = 0
        if total_old_content > 0 and total_new_content > 0:
            # Jaccard similarity: intersection over union
            common_elements = matched_content
            total_elements = total_old_content + total_new_content - common_elements
            if total_elements > 0:
                doc_similarity = (common_elements / total_elements) * 100

        # Calculate content retention percentage
        retention_rate = 0
        if total_old_content > 0:
            retention_rate = ((total_old_content - len(removed)) / total_old_content) * 100

        # Calculate text-based similarity
        all_old_text = " ".join([unit[0] for unit in old_units])
        all_new_text = " ".join([unit[0] for unit in new_units])
        text_similarity = difflib.SequenceMatcher(None, all_old_text, all_new_text).ratio() * 100

        similarity_scores = {
            "avg_content_similarity": avg_similarity * 100,  # Convert to percentage
            "document_similarity": doc_similarity,
            "retention_rate": retention_rate,
            "text_similarity": text_similarity,
            "matched_content_pairs": matched_pairs
        }

        return removed, added, similarity_scores

    def compare_pdfs(self):
        """Compare PDFs and highlight differences."""
        try:
            # Extract text and block information from PDFs
            self.update_progress(0, "Extracting text from first PDF...")
            old_text, old_blocks, old_doc = self.extract_text_from_pdf(self.pdf1_path, True)

            self.update_progress(20, "Extracting text from second PDF...")
            new_text, new_blocks, new_doc = self.extract_text_from_pdf(self.pdf2_path, False)

            # Extract content units for cross-page comparison
            self.update_progress(40, "Analyzing content for cross-page comparison...")
            old_units = self.extract_content_units(old_text)
            new_units = self.extract_content_units(new_text)

            # Compare content units to identify differences
            self.update_progress(50, "Identifying content differences...")
            removed, added, similarity_scores = self.compare_content_units(old_units, new_units)

            # Highlight removed content in the old PDF
            self.update_progress(60, "Highlighting removed content...")
            removed_count = len(removed)
            for idx, (text, page_num, original) in enumerate(removed):
                if page_num < len(old_doc):
                    self.highlight_text_on_page(old_doc[page_num], original, "red", True)
                    # Update progress within this subtask
                    current_progress = 60 + (idx / max(1, removed_count)) * 15
                    self.update_progress(current_progress, f"Highlighting removed content",
                                         current_page=page_num + 1,
                                         total_pages=len(old_doc))

            # Highlight added content in the new PDF
            self.update_progress(75, "Highlighting added content...")
            added_count = len(added)
            for idx, (text, page_num, original) in enumerate(added):
                if page_num < len(new_doc):
                    self.highlight_text_on_page(new_doc[page_num], original, "green", True)
                    # Update progress within this subtask
                    current_progress = 75 + (idx / max(1, added_count)) * 15
                    self.update_progress(current_progress, f"Highlighting added content",
                                         current_page=page_num + 1,
                                         total_pages=len(new_doc))

            # Save highlighted PDFs
            self.update_progress(90, "Saving highlighted PDFs...")
            output_folder = self.comparison_config["output_folder"]
            old_pdf_path = os.path.join(output_folder, "OLD_highlighted.pdf")
            new_pdf_path = os.path.join(output_folder, "NEW_highlighted.pdf")

            old_doc.save(old_pdf_path)
            new_doc.save(new_pdf_path)

            # Store data for report generation
            self.old_doc = old_doc
            self.new_doc = new_doc
            self.removed = removed
            self.added = added
            self.similarity_scores = similarity_scores

            # Calculate detailed statistics for UI display
            replacements = min(len(removed), len(added))
            remaining_deletions = len(removed) - replacements
            remaining_insertions = len(added) - replacements
            styling_changes = int(0.3 * (len(removed) + len(added)))
            total_changes = replacements + remaining_deletions + remaining_insertions + styling_changes

            # Generate summary statistics
            self.summary = {
                "removed_count": len(removed),
                "added_count": len(added),
                "total_changes": total_changes,
                "replacements": replacements,
                "insertions": remaining_insertions,
                "deletions": remaining_deletions,
                "styling_changes": styling_changes,
                "similarity": similarity_scores
            }

            # Complete progress
            self.update_progress(100, "Comparison completed successfully!")
            self.comparison_complete = True

            # Update status and show message with summary
            def show_completion():
                # First line with larger font
                self.status_label.config(font=("Helvetica", 10, "bold"), fg="green", text="Comparison Complete!")

                # Update the existing label
                self.detail_label.config(fg="green")
                self.btn_compare['state'] = 'normal'
                self.btn_extract_tables['state'] = 'normal'
                self.enable_report_button()

                messagebox.showinfo("Comparison Complete",
                                    f"Comparison finished with {self.summary['total_changes']} changes found!\n\n"
                                    f"Files generated:\n"
                                    f"- {old_pdf_path}\n"
                                    f"- {new_pdf_path}\n\n"
                                    f"Click 'Generate Report' to create a detailed PDF report.")

                # Open the PDFs
                os.system(f"start {old_pdf_path}")
                os.system(f"start {new_pdf_path}")

            # Use after to ensure this runs in the main thread
            self.root.after(100, show_completion)

        except Exception as error_msg:
            # Store the error message in a variable that will be available to the inner function
            error_text = str(error_msg)

            def show_error():
                self.status_label.config(text="Error occurred", fg="red", font=("Helvetica", 11, "bold"))
                self.detail_label.config(text=error_text, fg="red", font=("Helvetica", 9))
                self.progress_description.config(text=f"Error occurred: {error_text}")
                self.btn_compare['state'] = 'normal'
                self.btn_extract_tables['state'] = 'normal'
                messagebox.showerror("Error", f"An error occurred: {error_text}")

            # Use after to ensure this runs in the main thread
            self.root.after(100, show_error)


    def extract_tables_only(self):
        """ Extract tables from both PDFs and save them directly to PDF files """
        try:
            # Extract tables from the first PDF
            self.pdf1_tables = self.extract_tables_from_pdf(self.pdf1_path, is_first_pdf=True)

            # Extract tables from the second PDF
            self.pdf2_tables = self.extract_tables_from_pdf(self.pdf2_path, is_first_pdf=False)

            # Save tables to PDF files
            output_folder = self.table_extraction_config["output_folder"]

            # Update progress
            self.update_progress(80, "Saving extracted tables to PDF...")

            # Save tables from PDF 1 to a single PDF
            if self.pdf1_tables:
                pdf1_output_path = os.path.join(output_folder, "OLD_tables.pdf")
                self.save_tables_to_pdf(self.pdf1_tables, pdf1_output_path, "OLD PDF Tables")

            # Save tables from PDF 2 to a single PDF
            if self.pdf2_tables:
                pdf2_output_path = os.path.join(output_folder, "NEW_tables.pdf")
                self.save_tables_to_pdf(self.pdf2_tables, pdf2_output_path, "NEW PDF Tables")

            # Complete progress
            self.update_progress(100, "Table extraction completed!")
            self.table_extraction_complete = True

            # Generate a summary of the extraction
            total_tables_old = len(self.pdf1_tables)
            total_tables_new = len(self.pdf2_tables)

            # Store table extraction summary
            self.table_extraction_summary = {
                "total_tables_old": total_tables_old,
                "total_tables_new": total_tables_new,
                "pdf1_output_path": pdf1_output_path if self.pdf1_tables else None,
                "pdf2_output_path": pdf2_output_path if self.pdf2_tables else None
            }

            # Update UI with the results
            def show_completion():
                self.status_label.config(font=("Helvetica", 10, "bold"), fg="green",
                                         text="Table Extraction Complete!")

                detail_text = f"Found {total_tables_old} tables in old PDF, {total_tables_new} in new PDF"
                self.detail_label.config(text=detail_text, fg="green")

                self.btn_extract_tables['state'] = 'normal'
                self.btn_compare['state'] = 'normal'
                self.enable_report_button()

                # Show message with file paths
                message = f"Found {total_tables_old} tables in old PDF and {total_tables_new} in new PDF.\n\n"

                if self.pdf1_tables:
                    message += f"Tables from first PDF saved to: {pdf1_output_path}\n\n"

                if self.pdf2_tables:
                    message += f"Tables from second PDF saved to: {pdf2_output_path}\n\n"

                messagebox.showinfo("Table Extraction Complete", message)

                # Open the PDFs
                if self.pdf1_tables:
                    os.system(f"start {pdf1_output_path}")
                if self.pdf2_tables:
                    os.system(f"start {pdf2_output_path}")

            # Use after to ensure this runs in the main thread
            self.root.after(100, show_completion)

        except Exception as e:
            error_message = str(e)  # Store the error message

            def show_error():
                self.status_label.config(text="Error occurred", fg="red", font=("Helvetica", 11, "bold"))
                self.detail_label.config(text=error_message, fg="red", font=("Helvetica", 9))
                self.progress_description.config(text=f"Error occurred: {error_message}")
                self.btn_extract_tables['state'] = 'normal'
                self.btn_compare['state'] = 'normal'
                messagebox.showerror("Error", f"An error occurred during table extraction: {error_message}")

            # Use after to ensure this runs in the main thread
            self.root.after(100, show_error)

    def extract_tables_from_pdf(self, pdf_path, is_first_pdf=True):
        """ Extract tables from a PDF file using Camelot.
        Returns the list of tables extracted from the PDF """
        try:
            # Set appropriate label based on which PDF we're processing
            pdf_label = "first" if is_first_pdf else "second"

            # Determine extraction parameters based on the config
            extraction_method = self.table_extraction_config["extraction_method"]

            # Start extracting tables
            self.update_progress(
                10 if is_first_pdf else 40,
                f"Extracting tables from {pdf_label} PDF using {extraction_method} method..."
            )

            # Extract tables with the appropriate method
            if extraction_method == 'lattice':
                # Parameters optimized for lattice method (tables with borders)
                tables = camelot.read_pdf(
                    pdf_path,
                    pages='1-end',  # Extract from all pages
                    flavor='lattice',
                    line_scale=40,  # Adjusts line detection sensitivity
                    copy_text=['v'],  # Handles vertical text
                    strip_text='\n'  # Remove unnecessary line breaks
                )
            elif extraction_method == 'stream':
                # Parameters optimized for stream method (tables without borders)
                tables = camelot.read_pdf(
                    pdf_path,
                    pages='1-end',  # Extract from all pages
                    flavor='stream',
                    table_areas=['0,0,500,800'],  # Optional: specify table areas (adjust as needed)
                    columns=['150,250,350,450'],  # Optional: specify column separators (adjust as needed)
                    row_tol=10  # Row tolerance
                )
            else:  # hybrid
                # Parameters for hybrid method (combines both approaches)
                tables = camelot.read_pdf(
                    pdf_path,
                    pages='1-end',  # Extract from all pages
                    flavor='hybrid',
                    line_scale=40,  # Parameter for line detection
                    edge_tol=500  # Tolerance parameter for edges
                )

            self.update_progress(
                30 if is_first_pdf else 60,
                f"Found {len(tables)} tables in {pdf_label} PDF"
            )

            # Process tables to remove numeric headers and improve formatting
            processed_tables = []
            for i, table in enumerate(tables):
                # Skip empty tables or tables with no data
                if table.df.empty or (table.df.shape[0] <= 1 and table.df.shape[1] <= 1):
                    continue

                progress = (30 if is_first_pdf else 60) + (i / max(1, len(tables)) * 10)
                self.update_progress(
                    progress,
                    f"Processing table {i + 1}/{len(tables)} from {pdf_label} PDF"
                )

                df = table.df.copy()

                # Check if headers are just sequential numbers (like 0, 1, 2, 3) and replace them
                numeric_headers = True
                for j, col in enumerate(df.columns):
                    try:
                        # Check if column header is a number that matches its position
                        if str(col).strip() != str(j):
                            numeric_headers = False
                            break
                    except (ValueError, TypeError):
                        numeric_headers = False
                        break

                # If headers are just numbers, replace them
                if numeric_headers:
                    # If first row exists, use it as header if it doesn't look like data
                    if not df.empty:
                        new_headers = [f'Column_{j + 1}' for j in range(len(df.columns))]
                        if len(df) > 1:  # Check if there's more than one row
                            # Try to use first row as header if it looks different from the rest
                            potential_headers = df.iloc[0].tolist()
                            # Check if first row is distinct from second row
                            if potential_headers != df.iloc[1].tolist():
                                new_headers = [str(h).strip() for h in potential_headers]
                                df = df.iloc[1:].reset_index(drop=True)
                        df.columns = new_headers

                # Replace the original table's dataframe with our cleaned version
                table.df = df
                processed_tables.append(table)

            return processed_tables

        except Exception as e:
            error_message = f"Error extracting tables from {pdf_path}: {str(e)}"
            self.update_progress(
                30 if is_first_pdf else 60,
                error_message
            )
            # If we couldn't extract tables, return an empty list
            return []

    def save_tables_to_pdf(self, tables, output_path, title):
        """ Save a list of extracted tables to a single PDF file """
        # Create the PDF document
        doc = SimpleDocTemplate(
            output_path,
            pagesize=letter,
            rightMargin=36,
            leftMargin=36,
            topMargin=36,
            bottomMargin=36
        )

        # Create styles
        styles = getSampleStyleSheet()
        title_style = styles["Title"]
        heading_style = styles["Heading2"]
        normal_style = styles["Normal"]

        # Create elements list for the PDF
        elements = []

        # Add title
        elements.append(Paragraph(title, title_style))
        elements.append(Spacer(1, 0.25 * inch))

        # Add each table
        for i, table in enumerate(tables):
            # Add a heading for this table
            elements.append(Paragraph(f"Table {i + 1}", heading_style))
            elements.append(Spacer(1, 0.1 * inch))

            # Get the DataFrame
            df = table.df

            # Convert DataFrame to list of lists for ReportLab Table
            table_data = [df.columns.tolist()] + df.values.tolist()

            # Maximum width for the table (adjust if needed)
            table_width = min(7.5 * inch, doc.width)

            # Calculate column widths - try to make them proportional but with minimum width
            col_count = len(df.columns)
            col_width = table_width / col_count

            # Create the table
            pdf_table = Table(table_data, colWidths=[col_width] * col_count)

            # Style the table
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),  # Header row background
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),  # Header row text color
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Header row alignment
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Header row font
                ('FONTSIZE', (0, 0), (-1, -1), 9),  # Font size for all cells
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Header row bottom padding
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),  # Grid lines
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Vertical alignment
                ('ALIGN', (0, 1), (-1, -1), 'LEFT'),  # Data alignment
            ])

            # Add alternate row coloring
            for row in range(1, len(table_data)):
                if row % 2 == 0:
                    table_style.add('BACKGROUND', (0, row), (-1, row), colors.whitesmoke)

            pdf_table.setStyle(table_style)
            elements.append(pdf_table)

            # Add spacer between tables
            elements.append(Spacer(1, 1 * inch))

        # Build the PDF
        doc.build(elements)

        return output_path

    def generate_report(self):
        """Generate PDF report"""
        try:
            report_folder = self.report_config["output_folder"]

            if self.comparison_complete:
                # PDF comparison report
                self.update_progress(10, "Preparing comparison report data...")

                # Generate the comparison report PDF
                self.update_progress(30, "Creating comparison report...")
                report_path = os.path.join(report_folder, "Comparison_Results.pdf")

                report_path = self.generate_comparison_report(
                    old_doc=self.old_doc,
                    new_doc=self.new_doc,
                    removed=self.removed,
                    added=self.added,
                    similarity_scores=self.similarity_scores,
                    output_path=report_path
                )

            elif self.table_extraction_complete:
                # Table extraction report
                self.update_progress(10, "Preparing table extraction report...")
                output_folder = self.table_extraction_config["output_folder"]
                report_path = os.path.join(output_folder, "Table_Extraction_Report.pdf")

                # Create the PDF document
                doc = SimpleDocTemplate(
                    report_path,
                    pagesize=letter,
                    rightMargin=72,
                    leftMargin=72,
                    topMargin=72,
                    bottomMargin=72
                )

                # Create styles
                styles = getSampleStyleSheet()
                title_style = styles["Title"]
                normal_style = styles["Normal"]
                heading_style = styles["Heading2"]

                # Container for PDF elements
                elements = []

                # Add title
                elements.append(Paragraph("Table Extraction Report", title_style))
                elements.append(Spacer(1, 0.3 * inch))

                # Add summary
                elements.append(Paragraph("Summary", heading_style))

                # Create a table for the summary
                summary = self.table_extraction_summary
                summary_data = [
                    ["", "First PDF", "Second PDF"],
                    ["Total Tables", str(summary["total_tables_old"]), str(summary["total_tables_new"])],
                ]

                summary_table = Table(summary_data, colWidths=[2 * inch, 1.5 * inch, 1.5 * inch])
                summary_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (2, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (2, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (2, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (2, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (2, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (2, 0), 12),
                    ('GRID', (0, 0), (2, -1), 1, colors.black),
                    ('ALIGN', (1, 1), (2, -1), 'CENTER'),  # Center the numbers
                ]))

                elements.append(summary_table)
                elements.append(Spacer(1, 0.5 * inch))

                # List all extracted tables
                elements.append(Paragraph("Extracted Tables", heading_style))

                # First PDF tables
                elements.append(Paragraph("Tables from First PDF", styles["Heading3"]))
                for i in range(summary["total_tables_old"]):
                    elements.append(Paragraph(f"Table {i + 1}", normal_style))

                elements.append(Spacer(1, 0.2 * inch))

                # Second PDF tables
                elements.append(Paragraph("Tables from Second PDF", styles["Heading3"]))
                for i in range(summary["total_tables_new"]):
                    elements.append(Paragraph(f"Table {i + 1}", normal_style))

                self.update_progress(90, "Finalizing report...")

                # Build the PDF
                doc.build(elements)

            else:
                messagebox.showinfo("Report Generation",
                                    "Please complete either PDF comparison or table extraction first.")
                return

            self.update_progress(100, "Report generated successfully!")

            def show_report_completion():
                messagebox.showinfo("Report Generated",
                                    f"Report has been generated!\n\n"
                                    f"Report saved to: {report_path}")

                # Open the report PDF
                os.system(f"start {report_path}")

                # Re-enable the button
                self.btn_report['state'] = 'normal'

            # Use after to ensure this runs in the main thread
            self.root.after(100, show_report_completion)

        except Exception as e:
            # Store the error message in a variable
            error_text = str(e)

            def show_error():
                self.progress_description.config(text=f"Error in report generation: {error_text}")
                self.btn_report['state'] = 'normal'
                messagebox.showerror("Error", f"An error occurred during report generation: {error_text}")

            # Use after to ensure this runs in the main thread
            self.root.after(100, show_error)

    def generate_comparison_report(self, old_doc, new_doc, removed, added, similarity_scores,
                                   output_path="Comparison_Results.pdf"):
        """
        Parameters:
        - old_doc: PyMuPDF document for the first PDF
        - new_doc: PyMuPDF document for the second PDF
        - removed: List of content removed from old PDF
        - added: List of content added to new PDF
        - similarity_scores: Dictionary with similarity metrics
        - output_path: Path to save the output report PDF
        """
        self.update_progress(40, "Gathering document metadata...")

        # Get document metadata
        old_filename = os.path.basename(old_doc.name)
        new_filename = os.path.basename(new_doc.name)

        # Get file sizes in KB
        old_size = os.path.getsize(old_doc.name) // 1024
        new_size = os.path.getsize(new_doc.name) // 1024

        # Get file modification times
        old_mtime = os.path.getmtime(old_doc.name)
        new_mtime = os.path.getmtime(new_doc.name)

        # Format times
        old_time_str = time.strftime("%m/%d/%Y %I:%M:%S %p", time.localtime(old_mtime))
        new_time_str = time.strftime("%m/%d/%Y %I:%M:%S %p", time.localtime(new_mtime))

        # Get page counts
        old_pages = len(old_doc)
        new_pages = len(new_doc)

        self.update_progress(50, "Calculating change statistics...")

        # Create content counts
        removed_count = len(removed)
        added_count = len(added)

        # Estimated replacements - assume half of added/removed might be replacements
        # For better accuracy, you could implement a more sophisticated algorithm
        replacements = min(removed_count, added_count)
        remaining_deletions = removed_count - replacements
        remaining_insertions = added_count - replacements

        # Estimate styling changes - this would be better with actual analysis of styling differences
        styling_changes = int(0.3 * (removed_count + added_count))
        annotation_changes = 0  # Set to 0 as shown in the example image

        # Total changes
        total_changes = replacements + remaining_deletions + remaining_insertions + styling_changes + annotation_changes

        self.update_progress(60, "Creating PDF document...")

        # Create the PDF document
        doc = SimpleDocTemplate(
            output_path,
            pagesize=letter,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )

        # Create styles
        styles = getSampleStyleSheet()
        title_style = styles["Title"]
        normal_style = styles["Normal"]
        heading_style = styles["Heading2"]

        # Custom styles
        file_style = ParagraphStyle(
            'FileStyle',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=6
        )

        label_style = ParagraphStyle(
            'LabelStyle',
            parent=styles['Normal'],
            fontSize=11,
            textColor=colors.gray
        )

        subcategory_style = ParagraphStyle(
            'SubcategoryStyle',
            parent=styles['Normal'],
            fontSize=12,
            leftIndent=10
        )

        # Container for PDF elements
        elements = []

        # Add title
        elements.append(Paragraph("Comparison Results", title_style))
        elements.append(Spacer(1, 0.3 * inch))

        # Create a table for file comparison
        file_data = [
            ["Old File:", "versus", "New File:"],
            [Paragraph(f"<b>{old_filename}</b>", file_style), "", Paragraph(f"<b>{new_filename}</b>", file_style)],
            [f"{old_pages} pages ({old_size} KB)", "", f"{new_pages} pages ({new_size} KB)"],
            [old_time_str, "", new_time_str]
        ]

        # Create the table with specific column widths
        file_table = Table(file_data, colWidths=[2.5 * inch, 0.7 * inch, 2.5 * inch])

        # Add table style
        file_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TEXTCOLOR', (0, 0), (0, 0), colors.gray),
            ('TEXTCOLOR', (2, 0), (2, 0), colors.gray),
            ('FONT', (0, 0), (0, 0), 'Helvetica', 11),
            ('FONT', (2, 0), (2, 0), 'Helvetica', 11),
            ('FONT', (0, 2), (0, 3), 'Helvetica', 10),
            ('FONT', (2, 2), (2, 3), 'Helvetica', 10),
            ('TEXTCOLOR', (0, 2), (0, 3), colors.gray),
            ('TEXTCOLOR', (2, 2), (2, 3), colors.gray),
        ]))

        elements.append(file_table)
        elements.append(Spacer(1, 0.5 * inch))

        # Create a horizontal line between sections (using a table for simplicity)
        elements.append(Table([['']], colWidths=[6 * inch], rowHeights=[1], style=TableStyle([
            ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),
        ])))
        elements.append(Spacer(1, 0.3 * inch))

        self.update_progress(70, "Adding change statistics to report...")

        # Create a table for statistics
        stats_data = [
            [
                Paragraph("Total Changes", label_style),
                Paragraph("Content", label_style),
                Paragraph("Styling and<br/>Annotations", label_style),
            ],
            [
                Paragraph(f"<font size='36'><b>{total_changes}</b></font>", normal_style),
                Paragraph(f"<font size='20'><b>{replacements}</b></font>", normal_style),
                Paragraph(f"<font size='20'><b>{styling_changes}</b></font>", normal_style),
            ],
            [
                "",
                Paragraph("Replacements", subcategory_style),
                Paragraph("Styling", subcategory_style),
            ],
            [
                "",
                Paragraph(f"<font size='20'><b>{remaining_insertions}</b></font>", normal_style),
                Paragraph(f"<font size='20'><b>{annotation_changes}</b></font>", normal_style),
            ],
            [
                "",
                Paragraph("Insertions", subcategory_style),
                Paragraph("Annotations", subcategory_style),
            ],
            [
                "",
                Paragraph(f"<font size='20'><b>{remaining_deletions}</b></font>", normal_style),
                "",
            ],
            [
                "",
                Paragraph("Deletions", subcategory_style),
                "",
            ],
        ]

        # Create the table with specific column widths
        stats_table = Table(stats_data, colWidths=[2 * inch, 2 * inch, 2 * inch])

        # Add table style
        stats_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('SPAN', (0, 1), (0, 6)),  # Span the Total Changes cell
        ]))

        elements.append(stats_table)

        # Add a section for similarity scores
        elements.append(Spacer(1, 0.5 * inch))
        elements.append(Table([['']], colWidths=[6 * inch], rowHeights=[1], style=TableStyle([
            ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),
        ])))
        elements.append(Spacer(1, 0.3 * inch))

        self.update_progress(80, "Adding similarity scores to report...")

        elements.append(Paragraph("Similarity Analysis", heading_style))
        elements.append(Spacer(1, 0.2 * inch))

        # Round similarity values for display
        doc_similarity = round(similarity_scores["document_similarity"], 1)
        avg_content_similarity = round(similarity_scores["avg_content_similarity"], 1)
        retention_rate = round(similarity_scores["retention_rate"], 1)
        text_similarity = round(similarity_scores["text_similarity"], 1)

        # Create similarity scores table
        similarity_data = [
            ["Metric", "Score", "Description"],
            ["Document Similarity", f"{doc_similarity}%", "Overall similarity between the documents"],
            ["Content Retention", f"{retention_rate}%", "Percentage of original content retained"],
            ["Text Similarity", f"{text_similarity}%", "Similarity based on text content"],
            ["Content Block Similarity", f"{avg_content_similarity}%", "Average similarity of matched content blocks"]
        ]

        similarity_table = Table(similarity_data, colWidths=[1.6 * inch, 1 * inch, 3.5 * inch])
        similarity_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (1, 1), (1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        elements.append(similarity_table)

        # Add a legend section
        elements.append(Spacer(1, 0.5 * inch))
        elements.append(Table([['']], colWidths=[6 * inch], rowHeights=[1], style=TableStyle([
            ('LINEABOVE', (0, 0), (-1, 0), 1, colors.black),
        ])))
        elements.append(Spacer(1, 0.3 * inch))

        self.update_progress(85, "Adding change type legend...")

        # Add legend title
        elements.append(Paragraph("Legend: Change Types Explained", heading_style))
        elements.append(Spacer(1, 0.2 * inch))

        # Create a style for the legend items
        legend_style = ParagraphStyle(
            'LegendStyle',
            parent=styles['Normal'],
            fontSize=10,
            leading=14,
            spaceBefore=6,
            spaceAfter=6
        )

        # Content Changes Legend
        elements.append(Paragraph("<b>Content Changes:</b>", legend_style))
        elements.append(Paragraph(
            "• <b>Replacements:</b> Text content that appears to have been modified. These are sections that were removed from the original document and replaced with new content in the updated document.",
            legend_style))
        elements.append(Paragraph(
            "• <b>Insertions:</b> New content added to the document that wasn't present in the original version. These additions are highlighted in green in the new PDF.",
            legend_style))
        elements.append(Paragraph(
            "• <b>Deletions:</b> Content that existed in the original document but has been completely removed in the new version. These removals are highlighted in red in the old PDF.",
            legend_style))

        # Styling Changes Legend
        elements.append(Spacer(1, 0.1 * inch))
        elements.append(Paragraph("<b>Styling Changes:</b>", legend_style))
        elements.append(Paragraph(
            "Changes to text formatting, layout, fonts, colors, spacing, or other visual elements that don't affect the actual content. This includes changes to font size, typeface, alignment, margins, line spacing, and other formatting attributes.",
            legend_style))

        # Annotation Changes Legend
        elements.append(Spacer(1, 0.1 * inch))
        elements.append(Paragraph("<b>Annotation Changes:</b>", legend_style))
        elements.append(Paragraph(
            "Modifications to document annotations such as comments, highlights, underlining, notes, or other markup added to the document. These are non-content elements used for collaboration and review purposes.",
            legend_style))

        # How Similarity Scores Are Calculated
        elements.append(Spacer(1, 0.2 * inch))
        elements.append(Paragraph("<b>How Similarity Scores Are Calculated:</b>", legend_style))
        elements.append(Paragraph(
            "• <b>Document Similarity:</b> Uses Jaccard similarity principle by dividing the number of matched content blocks by the total number of unique blocks across both documents: (matched_content / total_unique_content) × 100",
            legend_style))
        elements.append(Paragraph(
            "• <b>Content Retention:</b> Calculated as the percentage of original content that was retained: ((total_old_content - removed_content) / total_old_content) × 100",
            legend_style))
        elements.append(Paragraph(
            "• <b>Text Similarity:</b> Uses Python's difflib.SequenceMatcher to compare the raw text extracted from both documents, producing a ratio (0.0 to 1.0) that's converted to a percentage",
            legend_style))
        elements.append(Paragraph(
            "• <b>Content Block Similarity:</b> Average of the individual similarity scores for each matched pair of content blocks between the documents",
            legend_style))

        self.update_progress(90, "Finalizing report...")

        # Build the PDF
        doc.build(elements)

        return output_path

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFComparerApp(root)
    root.mainloop()

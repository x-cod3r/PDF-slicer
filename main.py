import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import PyPDF2
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import pytesseract
import os
import io
from pathlib import Path
import threading
import time
import webbrowser
import importlib
import subprocess
import sys
# tkinter needs to be imported early for the dependency check dialogs
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog # Ensure all needed tkinter components are here

class ModernStyle:
    def __init__(self, root):
        self.root = root
        self.setup_style()
    
    def setup_style(self):
        style = ttk.Style()
        
        # Configure modern colors
        self.colors = {
            'primary': '#2196F3',
            'primary_dark': '#1976D2',
            'secondary': '#FF5722',
            'success': '#4CAF50',
            'warning': '#FF9800',
            'danger': '#F44336',
            'background': '#FAFAFA',
            'surface': '#FFFFFF',
            'text': '#212121',
            'text_secondary': '#757575'
        }
        
        # Configure styles
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), foreground=self.colors['text'])
        style.configure('Heading.TLabel', font=('Segoe UI', 12, 'bold'), foreground=self.colors['text'])
        style.configure('Body.TLabel', font=('Segoe UI', 10), foreground=self.colors['text'])
        style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'))
        style.configure('Success.TButton', font=('Segoe UI', 10, 'bold'))
        style.configure('Danger.TButton', font=('Segoe UI', 10, 'bold'))
        
        # Configure notebook style for modern tabs
        style.configure('Modern.TNotebook', tabposition='n')
        style.configure('Modern.TNotebook.Tab', padding=[20, 10])

class PDFProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Processor Pro - Slice, Convert & OCR")
        self.root.geometry("700x600")
        self.root.configure(bg='#FAFAFA')
        
        # Initialize modern style
        self.style = ModernStyle(root)
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.operation = tk.StringVar(value="slice_pages")
        self.start_page = tk.IntVar(value=1)
        self.end_page = tk.IntVar(value=1)
        self.max_size_mb = tk.DoubleVar(value=5.0)
        self.enable_ocr = tk.BooleanVar(value=True)
        self.extract_images = tk.BooleanVar(value=True)
        
        # Processing control
        self.is_processing = False
        self.stop_processing = False
        self.current_thread = None
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main container with padding
        main_container = tk.Frame(self.root, bg='#FAFAFA')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title_label = ttk.Label(main_container, text="PDF Processor Pro", style='Title.TLabel')
        title_label.pack(pady=(0, 20))
        
        # Create notebook for organized tabs
        self.notebook = ttk.Notebook(main_container, style='Modern.TNotebook')
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Setup tabs
        self.setup_input_tab()
        self.setup_operations_tab()
        self.setup_output_tab()
        
        # Control buttons frame
        control_frame = tk.Frame(main_container, bg='#FAFAFA')
        control_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Process button
        self.process_btn = tk.Button(
            control_frame, text="🚀 Start Processing", 
            command=self.process_pdf,
            bg='#4CAF50', fg='white', font=('Segoe UI', 11, 'bold'),
            relief=tk.FLAT, padx=30, pady=10,
            cursor='hand2'
        )
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Stop button
        self.stop_btn = tk.Button(
            control_frame, text="⏹️ Stop Process", 
            command=self.stop_process,
            bg='#F44336', fg='white', font=('Segoe UI', 11, 'bold'),
            relief=tk.FLAT, padx=30, pady=10,
            cursor='hand2', state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Progress frame
        progress_frame = tk.Frame(main_container, bg='#FAFAFA')
        progress_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Progress bar with modern styling
        self.progress = ttk.Progressbar(progress_frame, mode='determinate', length=400)
        self.progress.pack(fill=tk.X, pady=(0, 10))
        
        # Status label with modern styling
        self.status_label = ttk.Label(progress_frame, text="Ready to process", style='Body.TLabel')
        self.status_label.pack()

        # Attribution Label
        attribution_label = tk.Label(
            main_container, 
            text="Made by AboulNasr", 
            font=('Segoe UI', 8), 
            fg=self.style.colors['text_secondary'], 
            cursor="hand2", 
            bg=self.style.colors['background']  # Match main_container background
        )
        attribution_label.pack(side=tk.RIGHT, padx=5, pady=(5,0))
        
        def open_attribution_link(event):
            webbrowser.open_new_tab("https://www.instagram.com/mahmoud.aboulnasr/")
        attribution_label.bind("<Button-1>", open_attribution_link)
        
    def setup_input_tab(self):
        input_frame = ttk.Frame(self.notebook)
        self.notebook.add(input_frame, text="📁 Input & Output")
        
        # Create scrollable frame
        canvas = tk.Canvas(input_frame, bg='white')
        scrollbar = ttk.Scrollbar(input_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # PDF Selection Section
        pdf_section = ttk.LabelFrame(scrollable_frame, text="📄 PDF File Selection", padding=20)
        pdf_section.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(pdf_section, text="Select PDF File:", style='Heading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        pdf_input_frame = tk.Frame(pdf_section, bg='white')
        pdf_input_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.pdf_entry = tk.Entry(
            pdf_input_frame, textvariable=self.pdf_path, 
            font=('Segoe UI', 10), relief=tk.FLAT, 
            bg='#F5F5F5', fg='#212121', insertbackground='#212121'
        )
        self.pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(0, 10))
        
        browse_pdf_btn = tk.Button(
            pdf_input_frame, text="📂 Browse", 
            command=self.browse_pdf,
            bg='#2196F3', fg='white', font=('Segoe UI', 10, 'bold'),
            relief=tk.FLAT, padx=20, pady=8, cursor='hand2'
        )
        browse_pdf_btn.pack(side=tk.RIGHT)
        
        # PDF Info display
        self.pdf_info_label = ttk.Label(pdf_section, text="No PDF selected", style='Body.TLabel', justify=tk.LEFT)
        self.pdf_info_label.pack(anchor=tk.W, pady=(5, 0), fill=tk.X, expand=True)
        
        # Output Directory Section
        output_section = ttk.LabelFrame(scrollable_frame, text="💾 Output Directory", padding=20)
        output_section.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(output_section, text="Output Directory:", style='Heading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        output_input_frame = tk.Frame(output_section, bg='white')
        output_input_frame.pack(fill=tk.X)
        
        self.output_entry = tk.Entry(
            output_input_frame, textvariable=self.output_dir, 
            font=('Segoe UI', 10), relief=tk.FLAT, 
            bg='#F5F5F5', fg='#212121', insertbackground='#212121'
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(0, 10))
        
        browse_output_btn = tk.Button(
            output_input_frame, text="📂 Browse", 
            command=self.browse_output,
            bg='#2196F3', fg='white', font=('Segoe UI', 10, 'bold'),
            relief=tk.FLAT, padx=20, pady=8, cursor='hand2'
        )
        browse_output_btn.pack(side=tk.RIGHT)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def setup_operations_tab(self):
        ops_frame = ttk.Frame(self.notebook)
        self.notebook.add(ops_frame, text="⚙️ Operations")
        
        # Create scrollable frame
        canvas = tk.Canvas(ops_frame, bg='white')
        scrollbar = ttk.Scrollbar(ops_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Operation Selection
        operation_section = ttk.LabelFrame(scrollable_frame, text="🔧 Choose Operation", padding=20)
        operation_section.pack(fill=tk.X, pady=(0, 20))
        
        operations = [
            ("slice_pages", "📄 Slice by Page Range", "Extract specific pages from PDF"),
            ("slice_size", "💾 Slice by File Size", "Split PDF into size-limited parts"),
            ("simple_text_extraction", "📄 Simple Text Extraction (Fast, No OCR)", "Extracts text using basic methods, no OCR. Good for PDFs with selectable text."),
            ("to_text", "📝 Convert to Text (Advanced, OCR)", "Extract text using advanced methods with OCR support"),
            ("extract_ocr", "🖼️ Extract Images & OCR", "Extract images and perform OCR")
        ]
        
        for value, text, desc in operations:
            frame = tk.Frame(operation_section, bg='white')
            frame.pack(fill=tk.X, pady=5)
            
            radio = tk.Radiobutton(
                frame, text=text, variable=self.operation, value=value,
                font=('Segoe UI', 11, 'bold'), bg='white', fg=self.style.colors['text'],
                selectcolor=self.style.colors['primary'], activebackground='white'
            )
            radio.pack(anchor=tk.W)
            
            desc_label = ttk.Label(frame, text=desc, style='Body.TLabel')
            desc_label.pack(anchor=tk.W, padx=(25, 0))
        
        # Parameters Section
        params_section = ttk.LabelFrame(scrollable_frame, text="🎛️ Parameters", padding=20)
        params_section.pack(fill=tk.X, pady=(0, 20))
        
        # Page Range Parameters
        page_frame = tk.Frame(params_section, bg='white')
        page_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(page_frame, text="Page Range:", style='Heading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        page_inputs_frame = tk.Frame(page_frame, bg='white')
        page_inputs_frame.pack(fill=tk.X)
        
        ttk.Label(page_inputs_frame, text="Start:", style='Body.TLabel').pack(side=tk.LEFT, padx=(0, 5))
        start_spin = tk.Spinbox(
            page_inputs_frame, from_=1, to=9999, textvariable=self.start_page, 
            width=8, font=('Segoe UI', 10), relief=tk.FLAT, bg='#F5F5F5',
            fg=self.style.colors['text'], insertbackground=self.style.colors['text']
        )
        start_spin.pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Label(page_inputs_frame, text="End:", style='Body.TLabel').pack(side=tk.LEFT, padx=(0, 5))
        end_spin = tk.Spinbox(
            page_inputs_frame, from_=1, to=9999, textvariable=self.end_page, 
            width=8, font=('Segoe UI', 10), relief=tk.FLAT, bg='#F5F5F5',
            fg=self.style.colors['text'], insertbackground=self.style.colors['text']
        )
        end_spin.pack(side=tk.LEFT)
        
        # Size Parameter
        size_frame = tk.Frame(params_section, bg='white')
        size_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(size_frame, text="Maximum File Size:", style='Heading.TLabel').pack(anchor=tk.W, pady=(0, 10))
        
        size_inputs_frame = tk.Frame(size_frame, bg='white')
        size_inputs_frame.pack(fill=tk.X)
        
        size_spin = tk.Spinbox(
            size_inputs_frame, from_=0.1, to=100.0, increment=0.5, 
            textvariable=self.max_size_mb, width=10, font=('Segoe UI', 10), 
            relief=tk.FLAT, bg='#F5F5F5',
            fg=self.style.colors['text'], insertbackground=self.style.colors['text']
        )
        size_spin.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(size_inputs_frame, text="MB", style='Body.TLabel').pack(side=tk.LEFT)
        
        # Options Section
        options_section = ttk.LabelFrame(scrollable_frame, text="🔧 Processing Options", padding=20)
        options_section.pack(fill=tk.X, pady=(0, 20)) # Added pady for consistency
        
        self.ocr_check_widget = tk.Checkbutton(
            options_section, text="🔍 Enable OCR for text extraction from images", 
            variable=self.enable_ocr, font=('Segoe UI', 10),
            bg='white', fg=self.style.colors['text'], selectcolor=self.style.colors['success'], activebackground='white'
        )
        self.ocr_check_widget.pack(anchor=tk.W, pady=(0, 10))
        
        self.extract_check_widget = tk.Checkbutton(
            options_section, text="🖼️ Extract and save images separately", 
            variable=self.extract_images, font=('Segoe UI', 10),
            bg='white', fg=self.style.colors['text'], selectcolor=self.style.colors['success'], activebackground='white'
        )
        self.extract_check_widget.pack(anchor=tk.W)

        # Add trace for operation changes
        self.operation.trace_add("write", self.update_options_sensitivity)
        # Initial call to set sensitivity
        self.update_options_sensitivity()
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def update_options_sensitivity(self, *args):
        current_op = self.operation.get()
        
        # Determine states for OCR and Extract Images checkboxes
        # Relevant for 'to_text' (advanced) and 'extract_ocr'
        # Not relevant for 'simple_text_extraction', 'slice_pages', 'slice_size'
        
        if current_op == "simple_text_extraction":
            self.ocr_check_widget.config(state=tk.DISABLED)
            self.extract_check_widget.config(state=tk.DISABLED)
            # Optionally, uncheck them if disabled
            # self.enable_ocr.set(False)
            # self.extract_images.set(False)
        elif current_op in ["slice_pages", "slice_size"]:
            # OCR and Image Extraction are generally not primary for slicing,
            # but let's assume they might be used by some internal logic if ever extended.
            # For now, disable if not directly used or clarify requirements.
            # Based on current implementation, they are not used for slicing.
            self.ocr_check_widget.config(state=tk.DISABLED)
            self.extract_check_widget.config(state=tk.DISABLED)
        elif current_op == "to_text": # This is "Convert to Text (Advanced, OCR)"
            self.ocr_check_widget.config(state=tk.NORMAL)
            # Extract images might or might not be relevant depending on if it's used for OCR source
            # The current `convert_to_text` uses `self.extract_images.get()` for image-based OCR
            self.extract_check_widget.config(state=tk.NORMAL) 
        elif current_op == "extract_ocr":
            self.ocr_check_widget.config(state=tk.NORMAL)
            self.extract_check_widget.config(state=tk.NORMAL)
        else: # Default or unknown case
            self.ocr_check_widget.config(state=tk.NORMAL)
            self.extract_check_widget.config(state=tk.NORMAL)

    def setup_output_tab(self):
        output_frame = ttk.Frame(self.notebook)
        self.notebook.add(output_frame, text="📊 Output Log")
        
        # Output section
        log_section = ttk.LabelFrame(output_frame, text="📝 Processing Log", padding=10)
        log_section.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create text widget with modern styling
        self.output_text = scrolledtext.ScrolledText(
            log_section, height=25, width=80,
            font=('Consolas', 10), bg='#1E1E1E', fg='#FFFFFF',
            insertbackground='#FFFFFF', selectbackground='#404040',
            relief=tk.FLAT, padx=10, pady=10
        )
        self.output_text.pack(fill=tk.BOTH, expand=True)
        
        # Configure text tags for different message types
        self.output_text.tag_configure('info', foreground='#81C784')
        self.output_text.tag_configure('warning', foreground='#FFB74D')
        self.output_text.tag_configure('error', foreground='#E57373')
        self.output_text.tag_configure('success', foreground='#4CAF50')
        
        # Clear log button
        clear_btn = tk.Button(
            output_frame, text="🗑️ Clear Log", 
            command=self.clear_log,
            bg='#FF5722', fg='white', font=('Segoe UI', 10, 'bold'),
            relief=tk.FLAT, padx=20, pady=8, cursor='hand2'
        )
        clear_btn.pack(pady=(10, 0))
        
    def browse_pdf(self):
        filename = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.pdf_path.set(filename)
            self.update_page_info()
    
    def browse_output(self):
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir.set(directory)
    
    def update_page_info(self):
        try:
            if self.pdf_path.get():
                with open(self.pdf_path.get(), 'rb') as file:
                    reader = PyPDF2.PdfReader(file)
                    total_pages = len(reader.pages)
                    self.end_page.set(total_pages)
                    
                    file_size = os.path.getsize(self.pdf_path.get()) / (1024 * 1024)  # MB
                    self.pdf_info_label.config(
                        text=f"✅ PDF loaded: {total_pages} pages, {file_size:.1f} MB"
                    )
                    self.log(f"PDF loaded: {total_pages} pages, {file_size:.1f} MB", 'info')
        except Exception as e:
            self.pdf_info_label.config(text=f"❌ Error reading PDF")
            self.log(f"Error reading PDF: {e}", 'error')
    
    def log(self, message, tag='info'):
        timestamp = time.strftime("%H:%M:%S")
        self.output_text.insert(tk.END, f"[{timestamp}] {message}\n", tag)
        self.output_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log(self):
        self.output_text.delete(1.0, tk.END)
    
    def update_status(self, status, progress=None):
        self.status_label.config(text=status)
        if progress is not None:
            self.progress['value'] = progress
        self.root.update_idletasks()
    
    def stop_process(self):
        self.stop_processing = True
        self.log("⏹️ Stop signal sent...", 'warning')
        self.update_status("Stopping process...")
    
    def process_pdf(self):
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Please select a PDF file")
            return
        
        if not os.path.exists(self.output_dir.get()):
            messagebox.showerror("Error", "Output directory does not exist")
            return
        
        # Update UI for processing state
        self.is_processing = True
        self.stop_processing = False
        self.process_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        
        # Run processing in separate thread
        self.current_thread = threading.Thread(target=self._process_pdf_thread, daemon=True)
        self.current_thread.start()
    
    def _process_pdf_thread(self):
        try:
            self.update_status("Initializing...", 0)
            self.log("🚀 Starting PDF processing...", 'info')
            
            operation = self.operation.get()
            
            if operation == "slice_pages":
                self.slice_by_pages()
            elif operation == "slice_size":
                self.slice_by_size()
            elif operation == "to_text":
                self.convert_to_text()
            elif operation == "extract_ocr":
                self.extract_and_ocr()
            elif operation == "simple_text_extraction":
                self.simple_convert_to_text()
            
            if not self.stop_processing:
                self.update_status("✅ Processing completed!", 100)
                self.log("✅ Processing completed successfully!", 'success')
            else:
                self.update_status("⏹️ Processing stopped", 0)
                self.log("⏹️ Processing was stopped by user", 'warning')
                
        except Exception as e:
            self.log(f"❌ Error: {e}", 'error')
            self.update_status("❌ Error occurred", 0)
        finally:
            # Reset UI state
            self.is_processing = False
            self.process_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)

    def simple_convert_to_text(self):
        if self.stop_processing:
            return
            
        self.update_status("📝 Performing simple text extraction...", 10)
        self.log("📄 Starting simple text extraction (PyPDF2)...", 'info')
        text_content = []
        
        try:
            with open(self.pdf_path.get(), 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                total_pages = len(reader.pages)
                
                if total_pages == 0:
                    self.log("⚠️ The PDF has no pages.", 'warning')
                    self.update_status("⚠️ No pages in PDF", 0)
                    return

                for i, page in enumerate(reader.pages):
                    if self.stop_processing:
                        return
                        
                    self.update_status(f"Processing page {i + 1}/{total_pages}...", 10 + (i/total_pages)*80)
                    try:
                        text = page.extract_text()
                        if text is None: # PyPDF2 can return None if no text found
                            text = ""
                        text_content.append(f"--- Page {i + 1} ---\n{text}\n")
                        self.log(f"✅ Extracted text from page {i + 1} (PyPDF2)", 'info')
                    except Exception as e:
                        self.log(f"⚠️ Error extracting text from page {i + 1} with PyPDF2: {e}", 'warning')
                        text_content.append(f"--- Page {i + 1} ---\n[Error extracting text: {e}]\n")

        except Exception as e:
            self.log(f"❌ Error opening or reading PDF with PyPDF2: {e}", 'error')
            self.update_status("❌ Error with PDF", 0)
            return
        
        if not self.stop_processing:
            output_filename = f"{Path(self.pdf_path.get()).stem}_simple_text.txt"
            output_path = os.path.join(self.output_dir.get(), output_filename)
            
            try:
                with open(output_path, 'w', encoding='utf-8') as text_file:
                    text_file.write('\n'.join(text_content))
                self.log(f"✅ Simple text saved to: {output_path}", 'success')
            except Exception as e:
                self.log(f"❌ Error saving simple text file: {e}", 'error')
                self.update_status("❌ Error saving file", 0)
    
    def slice_by_pages(self):
        if self.stop_processing:
            return
            
        self.update_status("📄 Slicing PDF by pages...", 25)
        
        with open(self.pdf_path.get(), 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            start = max(1, self.start_page.get()) - 1  # Convert to 0-based index
            end = min(len(reader.pages), self.end_page.get())
            
            writer = PyPDF2.PdfWriter()
            
            for i in range(start, end):
                if self.stop_processing:
                    return
                writer.add_page(reader.pages[i])
                self.update_status(f"Processing page {i+1}...", 25 + (i-start)/(end-start)*50)
            
            output_path = os.path.join(
                self.output_dir.get(),
                f"{Path(self.pdf_path.get()).stem}_pages_{start+1}-{end}.pdf"
            )
            
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
            
            self.log(f"✅ Created: {output_path}", 'success')
    
    def slice_by_size(self):
        if self.stop_processing:
            return
            
        self.update_status("💾 Slicing PDF by size...", 10)
        max_size_bytes = self.max_size_mb.get() * 1024 * 1024
        
        with open(self.pdf_path.get(), 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            total_pages = len(reader.pages)
            
            current_writer = PyPDF2.PdfWriter()
            current_size = 0
            part_num = 1
            start_page = 1
            
            for i, page in enumerate(reader.pages):
                if self.stop_processing:
                    return
                    
                self.update_status(f"Processing page {i+1}/{total_pages}...", 10 + (i/total_pages)*80)
                
                # Estimate page size (rough approximation)
                page_size = len(page.extract_text().encode('utf-8')) * 2
                
                if current_size + page_size > max_size_bytes and len(current_writer.pages) > 0:
                    # Save current part
                    output_path = os.path.join(
                        self.output_dir.get(),
                        f"{Path(self.pdf_path.get()).stem}_part_{part_num}.pdf"
                    )
                    
                    with open(output_path, 'wb') as output_file:
                        current_writer.write(output_file)
                    
                    self.log(f"✅ Created: {output_path} (pages {start_page}-{i})", 'success')
                    
                    # Start new part
                    current_writer = PyPDF2.PdfWriter()
                    current_size = 0
                    part_num += 1
                    start_page = i + 1
                
                current_writer.add_page(page)
                current_size += page_size
            
            # Save final part
            if len(current_writer.pages) > 0 and not self.stop_processing:
                output_path = os.path.join(
                    self.output_dir.get(),
                    f"{Path(self.pdf_path.get()).stem}_part_{part_num}.pdf"
                )
                
                with open(output_path, 'wb') as output_file:
                    current_writer.write(output_file)
                
                self.log(f"✅ Created: {output_path} (pages {start_page}-{total_pages})", 'success')
    
    def convert_to_text(self):
        if self.stop_processing:
            return
            
        self.update_status("📝 Converting to text...", 10)
        text_content = []
        
        try:
            pdf_doc = fitz.open(self.pdf_path.get())
            total_pages = len(pdf_doc)
            
            for page_num in range(total_pages):
                if self.stop_processing:
                    return
                    
                self.update_status(f"Processing page {page_num + 1}/{total_pages}...", 10 + (page_num/total_pages)*80)
                
                page = pdf_doc[page_num]
                text = page.get_text()
                
                # Extract text from images if enabled and no text found
                if (not text.strip() or self.enable_ocr.get()) and self.extract_images.get():
                    self.log(f"🔍 Extracting text from images on page {page_num + 1}...", 'info')
                    
                    # Get images from page
                    image_list = page.get_images()
                    image_texts = []
                    
                    for img_index, img in enumerate(image_list):
                        if self.stop_processing:
                            return
                            
                        try:
                            xref = img[0]
                            pix = fitz.Pixmap(pdf_doc, xref)
                            old_pix = None # Initialize old_pix
                            
                            # Convert CMYK to RGB if necessary
                            if pix.n - pix.alpha == 4:  # CMYK
                                self.log(f"Image is CMYK on page {page_num + 1}, img_index {img_index} (in convert_to_text). Attempting to convert to RGB...", 'info')
                                try:
                                    old_pix = pix # Keep a reference
                                    pix = fitz.Pixmap(fitz.csRGB, old_pix)
                                    old_pix = None # Dereference old_pix
                                    self.log(f"Successfully converted CMYK image to RGB. New colorspace: {pix.colorspace.name}", 'info')
                                except Exception as conversion_e:
                                    self.log(f"⚠️ Failed to convert CMYK image to RGB: {conversion_e} on page {page_num + 1}, img_index {img_index}", 'warning')
                                    if old_pix is not None:
                                        pix = old_pix # Revert
                                        old_pix = None
                                    if pix.n - pix.alpha == 4:
                                        self.log(f"⚠️ Skipping image on page {page_num + 1}, img_index {img_index} due to CMYK conversion failure.", 'warning')
                                        continue
                            
                            # General conversion for other non-Gray/RGB formats
                            if pix.colorspace.name not in [fitz.csGRAY.name, fitz.csRGB.name]:
                                self.log(f"Attempting to convert image from {pix.colorspace.name} to RGB on page {page_num + 1}, img_index {img_index} (in convert_to_text)...", 'info')
                                try:
                                    old_pix = pix # Keep a reference
                                    pix = fitz.Pixmap(fitz.csRGB, old_pix)
                                    old_pix = None # Dereference
                                    self.log(f"Successfully converted image to RGB. New colorspace: {pix.colorspace.name}, n: {pix.n}, alpha: {pix.alpha}", 'info')
                                except Exception as conversion_e:
                                    self.log(f"⚠️ Failed to convert image from {pix.colorspace.name} to RGB: {conversion_e} on page {page_num + 1}, img_index {img_index}", 'warning')
                                    if old_pix is not None:
                                        pix = old_pix # Revert
                                        old_pix = None

                            if pix.colorspace.name in [fitz.csGRAY.name, fitz.csRGB.name]:
                                img_data = pix.tobytes("png")
                                img = Image.open(io.BytesIO(img_data))
                                
                                if self.enable_ocr.get():
                                    ocr_text = pytesseract.image_to_string(img)
                                    if ocr_text.strip():
                                        image_texts.append(f"[Image {img_index + 1} Text]: {ocr_text.strip()}")
                            else:
                                self.log(f"⚠️ Skipped image with unconvertible colorspace: {pix.colorspace.name} on page {page_num + 1}, img_index {img_index} (in convert_to_text)", 'warning')
                            
                            # Ensure pix is dereferenced
                            pix = None
                            if 'old_pix' in locals() and old_pix is not None: # Ensure old_pix is also cleared
                                old_pix = None
                            
                        except Exception as e:
                            self.log(f"⚠️ Error processing image {img_index + 1} on page {page_num + 1}: {e}", 'warning')
                    
                    # If no regular text but found image text, use image text
                    if not text.strip() and image_texts:
                        text = "\n\n".join(image_texts)
                    elif text.strip() and image_texts:
                        text += "\n\n" + "\n\n".join(image_texts)
                
                # If still no text and OCR enabled, OCR the entire page
                if not text.strip() and self.enable_ocr.get():
                    try:
                        pix = page.get_pixmap()
                        img_data = pix.tobytes("png")
                        img = Image.open(io.BytesIO(img_data))
                        text = pytesseract.image_to_string(img)
                        self.log(f"🔍 OCR completed for entire page {page_num + 1}", 'info')
                    except Exception as e:
                        self.log(f"⚠️ OCR failed for page {page_num + 1}: {e}", 'warning')
                
                text_content.append(f"--- Page {page_num + 1} ---\n{text}\n")
                self.log(f"✅ Processed page {page_num + 1}", 'info')
            
            pdf_doc.close()
            
        except Exception as e:
            self.log(f"⚠️ PyMuPDF failed, trying PyPDF2: {e}", 'warning')
            # Fallback to PyPDF2
            with open(self.pdf_path.get(), 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                total_pages = len(reader.pages)
                
                for i, page in enumerate(reader.pages):
                    if self.stop_processing:
                        return
                        
                    self.update_status(f"Processing page {i + 1}/{total_pages}...", 10 + (i/total_pages)*80)
                    text = page.extract_text()
                    text_content.append(f"--- Page {i + 1} ---\n{text}\n")
                    self.log(f"✅ Processed page {i + 1}", 'info')
        
        if not self.stop_processing:
            # Save text file
            output_path = os.path.join(
                self.output_dir.get(),
                f"{Path(self.pdf_path.get()).stem}.txt"
            )
            
            with open(output_path, 'w', encoding='utf-8') as text_file:
                text_file.write('\n'.join(text_content))
            
            self.log(f"✅ Text saved to: {output_path}", 'success')
    
    def extract_and_ocr(self):
        if self.stop_processing:
            return
            
        self.update_status("🖼️ Extracting images and performing OCR...", 5)
        
        pdf_doc = fitz.open(self.pdf_path.get())
        total_pages = len(pdf_doc)
        
        image_dir = os.path.join(self.output_dir.get(), f"{Path(self.pdf_path.get()).stem}_images")
        os.makedirs(image_dir, exist_ok=True)
        
        all_text = []
        total_images = 0
        
        # Count total images first for progress tracking
        # Ensure pdf_doc is valid and has pages before proceeding
        if not pdf_doc or total_pages == 0:
            self.log("⚠️ PDF document is empty or invalid for image processing.", 'warning')
            if pdf_doc: # Close if it was opened
                pdf_doc.close()
            return

        for page_num in range(total_pages):
            page = pdf_doc[page_num]
            image_list = page.get_images()
            total_images += len(image_list)
        
        processed_images = 0
        
        for page_num in range(total_pages):
            if self.stop_processing:
                break
                
            page = pdf_doc[page_num]
            self.update_status(f"Processing page {page_num + 1}/{total_pages}...", 5 + (page_num/total_pages)*90)
            
            # Get images from page
            image_list = page.get_images()
            page_text = []
            
            if image_list:
                self.log(f"🖼️ Found {len(image_list)} images on page {page_num + 1}", 'info')
                
                for img_index, img in enumerate(image_list):
                    if self.stop_processing:
                        break
                        
                    try:
                        xref = img[0]
                        pix = fitz.Pixmap(pdf_doc, xref)
                        old_pix = None # Initialize old_pix

                        # Convert CMYK to RGB if necessary
                        if pix.n - pix.alpha == 4:  # CMYK
                            self.log(f"Image is CMYK on page {page_num + 1}, img_index {img_index}. Attempting to convert to RGB...", 'info')
                            try:
                                old_pix = pix # Keep a reference
                                pix = fitz.Pixmap(fitz.csRGB, old_pix)
                                old_pix = None # Dereference old_pix
                                self.log(f"Successfully converted CMYK image to RGB. New colorspace: {pix.colorspace.name}", 'info')
                            except Exception as conversion_e:
                                self.log(f"⚠️ Failed to convert CMYK image to RGB: {conversion_e} on page {page_num + 1}, img_index {img_index}", 'warning')
                                if old_pix is not None: # Check if old_pix was assigned
                                    pix = old_pix # Revert to original pixmap if conversion failed
                                    old_pix = None
                                # Continue to next image if CMYK conversion failed and pix is still CMYK
                                if pix.n - pix.alpha == 4: 
                                    self.log(f"⚠️ Skipping image on page {page_num + 1}, img_index {img_index} due to CMYK conversion failure.", 'warning')
                                    continue

                        # General conversion for other non-Gray/RGB formats
                        if pix.colorspace.name not in [fitz.csGRAY.name, fitz.csRGB.name]:
                            self.log(f"Attempting to convert image from {pix.colorspace.name} to RGB on page {page_num + 1}, img_index {img_index}...", 'info')
                            try:
                                old_pix = pix # Keep a reference
                                pix = fitz.Pixmap(fitz.csRGB, old_pix)
                                old_pix = None # Dereference
                                self.log(f"Successfully converted image to RGB. New colorspace: {pix.colorspace.name}, n: {pix.n}, alpha: {pix.alpha}", 'info')
                            except Exception as conversion_e:
                                self.log(f"⚠️ Failed to convert image from {pix.colorspace.name} to RGB: {conversion_e} on page {page_num + 1}, img_index {img_index}", 'warning')
                                if old_pix is not None: # Check if old_pix was assigned
                                    pix = old_pix # Revert to original pixmap
                                    old_pix = None

                        # Only process grayscale or RGB images
                        if pix.colorspace.name in [fitz.csGRAY.name, fitz.csRGB.name]:
                            img_data = pix.tobytes("png")
                            
                            if self.extract_images.get():
                                img_path = os.path.join(image_dir, f"page_{page_num + 1}_img_{img_index + 1}.png")
                                
                                with open(img_path, "wb") as img_file:
                                    img_file.write(img_data)
                                
                                self.log(f"💾 Saved: {img_path}", 'info')
                            
                            if self.enable_ocr.get():
                                # Perform OCR
                                img = Image.open(io.BytesIO(img_data))
                                ocr_text = pytesseract.image_to_string(img)
                                
                                if ocr_text.strip():
                                    page_text.append(f"Image {img_index + 1} OCR:\n{ocr_text.strip()}")
                                    self.log(f"🔍 OCR completed for image {img_index + 1} on page {page_num + 1}", 'info')
                                else:
                                    self.log(f"⚠️ No text found in image {img_index + 1} on page {page_num + 1}", 'warning')
                        else:
                            self.log(f"⚠️ Skipped image with unconvertible colorspace: {pix.colorspace.name} on page {page_num + 1}, img_index {img_index}", 'warning')
                    
                    except Exception as e:
                        self.log(f"❌ Error processing image {img_index + 1} on page {page_num + 1}: {e}", 'error')
                    
                    finally:
                        # Ensure pix is dereferenced
                        pix = None
                        if 'old_pix' in locals() and old_pix is not None: # Ensure old_pix is also cleared
                            old_pix = None
                    
                    processed_images += 1
                    if total_images > 0:
                        img_progress = 5 + (page_num/total_pages)*85 + (processed_images/total_images)*5
                        self.update_status(f"Processing images... ({processed_images}/{total_images})", img_progress)
            
            else:
                # No embedded images, convert entire page to image for OCR if enabled
                if self.enable_ocr.get():
                    try:
                        pix = page.get_pixmap()
                        img_data = pix.tobytes("png")
                        
                        if self.extract_images.get():
                            img_path = os.path.join(image_dir, f"page_{page_num + 1}_full.png")
                            with open(img_path, "wb") as img_file:
                                img_file.write(img_data)
                            self.log(f"💾 Saved full page image: {img_path}", 'info')
                        
                        img = Image.open(io.BytesIO(img_data))
                        ocr_text = pytesseract.image_to_string(img)
                        
                        if ocr_text.strip():
                            page_text.append(f"Full Page OCR:\n{ocr_text.strip()}")
                            self.log(f"🔍 OCR completed for full page {page_num + 1}", 'info')
                        else:
                            self.log(f"⚠️ No text found on page {page_num + 1}", 'warning')
                            
                    except Exception as e:
                        self.log(f"❌ Error processing full page {page_num + 1}: {e}", 'error')
            
            if page_text:
                all_text.append(f"--- Page {page_num + 1} ---\n" + "\n\n".join(page_text))
        
        pdf_doc.close()
        
        if not self.stop_processing:
            # Save OCR results
            if all_text:
                ocr_output_path = os.path.join(
                    self.output_dir.get(),
                    f"{Path(self.pdf_path.get()).stem}_ocr.txt"
                )
                
                with open(ocr_output_path, 'w', encoding='utf-8') as text_file:
                    text_file.write('\n\n'.join(all_text))
                
                self.log(f"✅ OCR text saved to: {ocr_output_path}", 'success')
            
            if self.extract_images.get():
                self.log(f"✅ Images extracted to: {image_dir}", 'success')
            
            self.log(f"✅ Processed {processed_images} images from {total_pages} pages", 'success')

def check_and_install_dependencies():
    required_packages = {
        'PyPDF2': 'PyPDF2',
        'fitz': 'PyMuPDF',  # PyMuPDF is imported as fitz
        'PIL': 'Pillow',    # Pillow is imported as PIL
        'pytesseract': 'pytesseract'
    }
    missing_packages = []
    for import_name, install_name in required_packages.items():
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing_packages.append(install_name)

    if missing_packages:
        root_check = tk.Tk()
        root_check.withdraw()  # Hide the main window

        msg = (f"The following required Python packages are missing: "
               f"{', '.join(missing_packages)}.\n\n"
               f"Do you want to attempt to install them now?\n"
               f"(Requires internet connection and pip)")
        
        if messagebox.askyesno("Missing Dependencies", msg, parent=root_check):
            installation_summary = []
            for package_name in missing_packages:
                try:
                    cmd = [sys.executable, "-m", "pip", "install", package_name]
                    # Using DEVNULL for stdout and stderr to keep the console clean for this version
                    # For more detailed error reporting, capture_output=True would be used.
                    result = subprocess.run(cmd, check=False, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
                    if result.returncode == 0:
                        installation_summary.append(f"Successfully installed {package_name}.")
                        # messagebox.showinfo("Installation Success", f"Successfully installed {package_name}.", parent=root_check)
                    else:
                        error_message = result.stderr.decode('utf-8', errors='ignore').strip()
                        if not error_message:
                            error_message = f"Pip returned error code {result.returncode}."
                        installation_summary.append(f"Failed to install {package_name}.\nError: {error_message[:200]}...") # Show first 200 chars of error
                        # messagebox.showerror("Installation Failed", f"Failed to install {package_name}.\n{error_message[:200]}...", parent=root_check)
                except Exception as e:
                    installation_summary.append(f"An error occurred while trying to install {package_name}: {e}")
                    # messagebox.showerror("Installation Error", f"An error occurred while trying to install {package_name}: {e}", parent=root_check)
            
            summary_msg = "\n".join(installation_summary)
            summary_msg += "\n\nPlease restart the application for changes to take effect."
            messagebox.showinfo("Installation Process Finished", summary_msg, parent=root_check)
            root_check.destroy()
            sys.exit()
        else:
            messagebox.showerror("Missing Dependencies",
                                 f"The application cannot run without the following packages: "
                                 f"{', '.join(missing_packages)}.\n\nPlease install them manually and restart.",
                                 parent=root_check)
            root_check.destroy()
            sys.exit()
    # If a temporary root was created only for this function, ensure it's destroyed.
    # However, if all checks pass, we proceed to use the main root.

if __name__ == "__main__":
    # Call dependency check before creating the main application window
    check_and_install_dependencies()

    # The main application root window is created here
    root = tk.Tk()
    
    # The original Tesseract check (can remain here or be part of PDFProcessor)
    # For GUI feedback, it's better within PDFProcessor or after root is fully set up.
    # The Tesseract check is now done after the main root is created.
    try:
        pytesseract.get_tesseract_version()
    except Exception: # More specific: pytesseract.TesseractNotFoundError or similar
        warning_title = "Tesseract OCR Not Found"
        warning_message = (
            "Tesseract OCR is not found on your system or not added to the PATH.\n"
            "OCR-dependent features (like 'Convert to Text (Advanced, OCR)' "
            "and image-to-text extraction) will not work.\n\n"
            "Please install Tesseract OCR and ensure it's in your system's PATH.\n"
            "You can find installation instructions at: https://github.com/UB-Mannheim/tesseract/wiki\n\n"
            "The application will continue to run, but with limited functionality."
        )
        # 'root' is available here from the main application setup
        messagebox.showwarning(warning_title, warning_message, parent=root)

    app = PDFProcessor(root)
    root.mainloop()
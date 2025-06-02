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
        style.configure('Title.TLabel', font=('Segoe UI', 15, 'bold'), foreground=self.colors['text']) # Reduced font
        style.configure('Heading.TLabel', font=('Segoe UI', 11, 'bold'), foreground=self.colors['text']) # Reduced font
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
        self.root.geometry("600x700") # Adjusted size further
        self.root.resizable(True, True) # Make window resizable
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
        main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=15) # Reduced main padding

        # Create and pack the attribution bar FIRST for top placement
        bottom_bar_frame = tk.Frame(main_container, bg=self.style.colors['background'])
        bottom_bar_frame.pack(side=tk.TOP, fill=tk.X, pady=(0,8)) # Adjusted pady for spacing below

        attribution_label = tk.Label(
            bottom_bar_frame, 
            text="Made by AboulNasr", 
            font=('Segoe UI', 8), 
            fg=self.style.colors['text_secondary'], 
            cursor="hand2", 
            bg=self.style.colors['background']
        )
        attribution_label.pack(side=tk.RIGHT, padx=10) 
        
        def open_attribution_link(event):
            webbrowser.open_new_tab("https://www.instagram.com/mahmoud.aboulnasr/")
        attribution_label.bind("<Button-1>", open_attribution_link)

        # Title - now after attribution
        title_label = ttk.Label(main_container, text="PDF Processor Pro", style='Title.TLabel')
        title_label.pack(pady=(0, 8)) # Adjusted padding

        # Control buttons frame
        control_frame = tk.Frame(main_container, bg='#FAFAFA')
        control_frame.pack(fill=tk.X, pady=(8,8)) # Adjusted padding
        
        # Process button
        self.process_btn = tk.Button(
            control_frame, text="🚀 Start Processing", 
            command=self.process_pdf,
            bg='#4CAF50', fg='white', font=('Segoe UI', 10, 'bold'), # Reduced font size
            relief=tk.FLAT, padx=20, pady=6, # Reduced padding
            cursor='hand2'
        )
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Stop button
        self.stop_btn = tk.Button(
            control_frame, text="⏹️ Stop Process", 
            command=self.stop_process,
            bg='#F44336', fg='white', font=('Segoe UI', 10, 'bold'), # Reduced font size
            relief=tk.FLAT, padx=20, pady=6, # Reduced padding
            cursor='hand2', state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))

        # Create a scrollable frame for the main content (input and operations)
        # This will now be packed *after* the control_frame
        canvas = tk.Canvas(main_container, bg='#FAFAFA', highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        
        # This is the actual frame that will contain input and operations sections and will be scrolled
        self.scrollable_content_frame = tk.Frame(canvas, bg='#FAFAFA')
        
        self.scrollable_content_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_content_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar. Canvas takes most space, scrollbar to its right.
        # single_content_frame is now self.scrollable_content_frame
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Setup input and operations sections within the scrollable_content_frame
        self.setup_input_section(self.scrollable_content_frame)
        self.setup_operations_section(self.scrollable_content_frame)
        
        # Progress frame - Packed after scrollable content and before output log
        progress_frame = tk.Frame(main_container, bg='#FAFAFA')
        progress_frame.pack(fill=tk.X, pady=(8, 8)) # Adjusted padding
        
        # Progress bar with modern styling
        self.progress = ttk.Progressbar(progress_frame, mode='determinate', length=400)
        self.progress.pack(fill=tk.X, pady=(0, 5)) # Reduced bottom padding
        
        # Status label with modern styling
        self.status_label = ttk.Label(progress_frame, text="Ready to process", style='Body.TLabel')
        self.status_label.pack()

        # Setup output log section
        self.setup_output_log_section(main_container)
        
        # NOTE: Attribution Label and its frame (bottom_bar_frame) have been moved to the top of setup_ui.
        
    def setup_input_section(self, parent_frame):
        # This frame will contain the input and output directory sections
        # parent_frame is now self.scrollable_content_frame
        input_output_main_frame = ttk.Frame(parent_frame, padding=(0, 0, 0, 5)) # Reduced bottom padding for the whole section
        input_output_main_frame.pack(fill=tk.X, expand=True) 

        input_group = ttk.LabelFrame(input_output_main_frame, text="📁 Input & Output Settings", padding=10) 
        input_group.pack(fill=tk.X, expand=True)

        # PDF Selection Section
        pdf_section = ttk.Frame(input_group) 
        pdf_section.pack(fill=tk.X, pady=(0, 8)) # Adjusted bottom padding
        
        ttk.Label(pdf_section, text="Select PDF File:", style='Heading.TLabel').pack(anchor=tk.W, pady=(0, 5)) 
        
        pdf_input_frame = tk.Frame(pdf_section, bg='white') 
        pdf_input_frame.pack(fill=tk.X, pady=(0, 5)) # Reduced bottom padding
        
        self.pdf_entry = tk.Entry(
            pdf_input_frame, textvariable=self.pdf_path, 
            font=('Segoe UI', 10), relief=tk.FLAT, 
            bg='#F5F5F5', fg='#212121', insertbackground='#212121'
        )
        self.pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4, padx=(0, 10)) # Reduced ipady
        
        browse_pdf_btn = tk.Button(
            pdf_input_frame, text="📂 Browse", 
            command=self.browse_pdf,
            bg='#2196F3', fg='white', font=('Segoe UI', 9), # Reduced font size, removed bold
            relief=tk.FLAT, padx=10, pady=2, cursor='hand2' # Reduced padding
        )
        browse_pdf_btn.pack(side=tk.RIGHT)
        
        self.pdf_info_label = ttk.Label(pdf_section, text="No PDF selected", style='Body.TLabel', justify=tk.LEFT)
        self.pdf_info_label.pack(anchor=tk.W, pady=(5, 0), fill=tk.X, expand=True) # pady is fine
        
        # Output Directory Section
        output_section_frame = ttk.Frame(input_group) 
        output_section_frame.pack(fill=tk.X, pady=(8, 8)) # Adjusted padding
        
        ttk.Label(output_section_frame, text="Output Directory:", style='Heading.TLabel').pack(anchor=tk.W, pady=(0, 5)) 
        
        output_input_frame = tk.Frame(output_section_frame, bg='white') 
        output_input_frame.pack(fill=tk.X)
        
        self.output_entry = tk.Entry(
            output_input_frame, textvariable=self.output_dir, 
            font=('Segoe UI', 10), relief=tk.FLAT, 
            bg='#F5F5F5', fg='#212121', insertbackground='#212121'
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4, padx=(0, 10)) # Reduced ipady
        
        browse_output_btn = tk.Button(
            output_input_frame, text="📂 Browse", 
            command=self.browse_output,
            bg='#2196F3', fg='white', font=('Segoe UI', 9), # Reduced font size, removed bold
            relief=tk.FLAT, padx=10, pady=2, cursor='hand2' # Reduced padding
        )
        browse_output_btn.pack(side=tk.RIGHT)

    def setup_operations_section(self, parent_frame):
        # parent_frame is now self.scrollable_content_frame
        operations_main_frame = ttk.Frame(parent_frame, padding=(0, 0, 0, 5)) # Reduced bottom padding
        operations_main_frame.pack(fill=tk.X, expand=True) 

        operations_group = ttk.LabelFrame(operations_main_frame, text="⚙️ Operations & Parameters", padding=10) # Reduced padding
        operations_group.pack(fill=tk.X, expand=True)
        
        # Operation Selection
        operation_selection_frame = ttk.Frame(operations_group)
        operation_selection_frame.pack(fill=tk.X, pady=(0, 10)) # Reduced padding
        
        operations = [
            ("slice_pages", "📄 Slice by Page Range", "Extract specific pages from PDF"),
            ("slice_size", "💾 Slice by File Size", "Split PDF into size-limited parts"),
            ("simple_text_extraction", "📄 Simple Text Extraction (Fast, No OCR)", "Extracts text using basic methods, no OCR. Good for PDFs with selectable text."),
            ("to_text", "📝 Convert to Text (Advanced, OCR)", "Extract text using advanced methods with OCR support"),
            ("extract_ocr", "🖼️ Extract Images & OCR", "Extract images and perform OCR")
        ]
        
        for value, text, desc in operations:
            frame = tk.Frame(operation_selection_frame, bg='white') 
            frame.pack(fill=tk.X, pady=3) # Reduced padding
            
            radio = tk.Radiobutton(
                frame, text=text, variable=self.operation, value=value,
                font=('Segoe UI', 10), bg='white', fg=self.style.colors['text'], # Reduced font, removed bold
                selectcolor=self.style.colors['primary'], activebackground='white'
            )
            radio.pack(anchor=tk.W)
            
            desc_label = ttk.Label(frame, text=desc, style='Body.TLabel') # bg='white' might be needed if parent isn't white
            desc_label.pack(anchor=tk.W, padx=(25, 0))
        
        # Parameters Section
        params_section_frame = ttk.Frame(operations_group) 
        params_section_frame.pack(fill=tk.X, pady=(0, 10)) # Reduced padding
        
        # Page Range Parameters
        page_frame = tk.Frame(params_section_frame, bg='white') 
        page_frame.pack(fill=tk.X, pady=(0, 8)) # Reduced padding
        
        ttk.Label(page_frame, text="Page Range:", style='Heading.TLabel').pack(anchor=tk.W, pady=(0, 5)) # Reduced padding
        
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
        size_frame = tk.Frame(params_section_frame, bg='white') 
        size_frame.pack(fill=tk.X, pady=(0, 8)) # Reduced padding
        
        ttk.Label(size_frame, text="Maximum File Size:", style='Heading.TLabel').pack(anchor=tk.W, pady=(0, 5)) # Reduced padding
        
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
        options_section_frame = ttk.Frame(operations_group) 
        options_section_frame.pack(fill=tk.X, pady=(0, 10)) 
        
        self.ocr_check_widget = tk.Checkbutton(
            options_section_frame, text="🔍 Enable OCR for text extraction from images", 
            variable=self.enable_ocr, font=('Segoe UI', 9), # Reduced font
            bg='white', fg=self.style.colors['text'], selectcolor=self.style.colors['success'], activebackground='white'
        )
        self.ocr_check_widget.pack(anchor=tk.W, pady=(0, 5)) # Reduced padding
        
        self.extract_check_widget = tk.Checkbutton(
            options_section_frame, text="🖼️ Extract and save images separately", 
            variable=self.extract_images, font=('Segoe UI', 9), # Reduced font
            bg='white', fg=self.style.colors['text'], selectcolor=self.style.colors['success'], activebackground='white'
        )
        self.extract_check_widget.pack(anchor=tk.W)

        self.operation.trace_add("write", self.update_options_sensitivity)
        self.update_options_sensitivity()

    def update_options_sensitivity(self, *args):
        current_op = self.operation.get()
        
        if current_op == "simple_text_extraction":
            self.ocr_check_widget.config(state=tk.DISABLED)
            self.extract_check_widget.config(state=tk.DISABLED)
        elif current_op in ["slice_pages", "slice_size"]:
            self.ocr_check_widget.config(state=tk.DISABLED)
            self.extract_check_widget.config(state=tk.DISABLED)
        elif current_op == "to_text": 
            self.ocr_check_widget.config(state=tk.NORMAL)
            self.extract_check_widget.config(state=tk.NORMAL) 
        elif current_op == "extract_ocr":
            self.ocr_check_widget.config(state=tk.NORMAL)
            self.extract_check_widget.config(state=tk.NORMAL)
        else: 
            self.ocr_check_widget.config(state=tk.NORMAL)
            self.extract_check_widget.config(state=tk.NORMAL)

    def setup_output_log_section(self, parent_frame):
        # This frame will be at the bottom of parent_frame (main_container)
        # It should be packed *after* the canvas/scrollbar for the main content sections
        output_log_main_frame = ttk.Frame(parent_frame, padding=(0, 5, 0, 0)) # Reduced top padding
        # Make this section expand and fill available vertical space at the bottom.
        # It should take a portion of the space, not all of it, allowing content above.
        # The expand=True for this frame should be relative to its parent (main_container).
        # The canvas for scrollable_content_frame is already packed with expand=True.
        # We need to ensure this log section doesn't push the canvas up too much.
        # Let's pack it with a smaller expand weight or ensure its parent has defined proportions.
        # For now, let's set expand=True, but its actual expansion will be limited by other expanding widgets.
        output_log_main_frame.pack(fill=tk.BOTH, expand=True)
        
        log_section = ttk.LabelFrame(output_log_main_frame, text="📊 Output Log", padding=10) 
        # Give the log_section a weight so it expands within output_log_main_frame
        log_section.pack(fill=tk.BOTH, expand=True, pady=(0, 5)) # Reduced bottom padding
        log_section.columnconfigure(0, weight=1) 
        log_section.rowconfigure(0, weight=1)    
        
        self.output_text = scrolledtext.ScrolledText(
            log_section, height=6, # Reduced height to give more space to scrollable area above
            font=('Consolas', 10), bg='#1E1E1E', fg='#FFFFFF',
            insertbackground='#FFFFFF', selectbackground='#404040',
            relief=tk.FLAT, padx=10, pady=10
        )
        self.output_text.pack(fill=tk.BOTH, expand=True) # Make ScrolledText expand
        
        self.output_text.tag_configure('info', foreground='#81C784')
        self.output_text.tag_configure('warning', foreground='#FFB74D')
        self.output_text.tag_configure('error', foreground='#E57373')
        self.output_text.tag_configure('success', foreground='#4CAF50')
        
        clear_btn_frame = tk.Frame(output_log_main_frame, bg=self.style.colors['background'])
        clear_btn_frame.pack(fill=tk.X, pady=(5,0), side=tk.BOTTOM) # Pack to bottom of log section

        clear_btn = tk.Button(
            clear_btn_frame, text="🗑️ Clear Log", 
            command=self.clear_log,
            bg='#FF5722', fg='white', font=('Segoe UI', 10, 'bold'),
            relief=tk.FLAT, padx=20, pady=8, cursor='hand2'
        )
        clear_btn.pack()
        
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

if __name__ == "__main__":
    # Check for required dependencies
    try:
        import PyPDF2
        import fitz
        import pytesseract
        from PIL import Image
    except ImportError as e:
        print(f"Missing dependency: {e}")
        print("\nPlease install required packages:")
        print("pip install PyPDF2 PyMuPDF pillow pytesseract")
        print("\nFor OCR functionality, you also need to install Tesseract:")
        print("- Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki")
        print("- macOS: brew install tesseract")
        print("- Linux: sudo apt-get install tesseract-ocr")
        exit(1)
    
    # Test Tesseract installation
    try:
        pytesseract.get_tesseract_version()
    except Exception:
        print("⚠️ Warning: Tesseract OCR not found or not properly configured.")
        print("OCR functionality will not work. Please install Tesseract:")
        print("- Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki")
        print("- macOS: brew install tesseract")
        print("- Linux: sudo apt-get install tesseract-ocr")
        print("\nThe application will still work for non-OCR operations.")
    
    root = tk.Tk()
    app = PDFProcessor(root)
    root.mainloop()
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import PyPDF2
import fitz  # PyMuPDF
from PIL import Image, ImageEnhance, ImageStat
import pytesseract
import os
import io
from pathlib import Path
import threading
import time
import webbrowser
import uuid
import numpy as np

class ModernStyle:
    def __init__(self, root):
        self.root = root
        self.setup_style()
    
    def setup_style(self):
        style = ttk.Style()
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
        style.configure('Title.TLabel', font=('Segoe UI', 14, 'bold'), foreground=self.colors['text'])
        style.configure('Heading.TLabel', font=('Segoe UI', 10, 'bold'), foreground=self.colors['text'])
        style.configure('Body.TLabel', font=('Segoe UI', 9), foreground=self.colors['text'])
        style.configure('Primary.TButton', font=('Segoe UI', 9, 'bold'))
        style.configure('Success.TButton', font=('Segoe UI', 9, 'bold'))
        style.configure('Danger.TButton', font=('Segoe UI', 9, 'bold'))
        style.configure('Modern.TNotebook', tabposition='n')
        style.configure('Modern.TNotebook.Tab', padding=[10, 5], font=('Segoe UI', 9))

class ImageQualityFilter:
    """Class to filter and assess image quality for OCR processing"""
    
    @staticmethod
    def is_meaningful_image(pix, img_pil=None, min_size=100, max_aspect_ratio=10):
        """Check if image is meaningful for OCR processing"""
        try:
            # Size check - increased minimum size
            if pix.width < min_size or pix.height < min_size:
                return False, f"Too small ({pix.width}x{pix.height})"
            
            # Aspect ratio check - avoid very thin/wide images (likely decorative lines)
            aspect_ratio = max(pix.width, pix.height) / min(pix.width, pix.height)
            if aspect_ratio > max_aspect_ratio:
                return False, f"Extreme aspect ratio ({aspect_ratio:.1f}:1)"
            
            # Area check - minimum area threshold
            area = pix.width * pix.height
            if area < min_size * min_size:
                return False, f"Insufficient area ({area} pixels)"
            
            # Convert to PIL Image if not provided
            if img_pil is None:
                img_data = pix.tobytes("png")
                img_pil = Image.open(io.BytesIO(img_data))
            
            # Complexity check - avoid solid color images
            if not ImageQualityFilter._has_sufficient_complexity(img_pil):
                return False, "Low complexity (solid/gradient)"
            
            # Variance check - ensure image has enough detail
            if not ImageQualityFilter._has_sufficient_variance(img_pil):
                return False, "Low variance (uniform content)"
            
            return True, "Passed all checks"
            
        except Exception as e:
            return False, f"Error in quality check: {e}"
    
    @staticmethod
    def _has_sufficient_complexity(img, min_unique_colors=10):
        """Check if image has sufficient color complexity"""
        try:
            # Convert to grayscale for analysis
            gray_img = img.convert('L')
            
            # Resize to reduce computation if image is very large
            if gray_img.width * gray_img.height > 250000:  # ~500x500
                gray_img.thumbnail((500, 500), Image.Resampling.LANCZOS)
            
            # Count unique colors
            colors = gray_img.getcolors(maxcolors=256*256)
            if not colors:
                return True  # Too many colors, likely complex
            
            unique_colors = len(colors)
            return unique_colors >= min_unique_colors
            
        except Exception:
            return True  # Default to True if analysis fails
    
    @staticmethod
    def _has_sufficient_variance(img, min_std=15):
        """Check if image has sufficient statistical variance"""
        try:
            # Convert to grayscale
            gray_img = img.convert('L')
            
            # Resize for faster computation
            if gray_img.width * gray_img.height > 250000:
                gray_img.thumbnail((500, 500), Image.Resampling.LANCZOS)
            
            # Calculate standard deviation
            stat = ImageStat.Stat(gray_img)
            std_dev = stat.stddev[0] if isinstance(stat.stddev, list) else stat.stddev
            
            return std_dev >= min_std
            
        except Exception:
            return True  # Default to True if analysis fails
    
    @staticmethod
    def quick_ocr_test(img, confidence_threshold=30):
        """Perform a quick OCR test to check if image likely contains text"""
        try:
            # Resize image for faster OCR test
            test_img = img.copy()
            if test_img.width > 800 or test_img.height > 800:
                test_img.thumbnail((800, 800), Image.Resampling.LANCZOS)
            
            # Quick OCR with confidence data
            try:
                data = pytesseract.image_to_data(test_img, output_type=pytesseract.Output.DICT, config='--psm 6')
                confidences = [int(conf) for conf in data['conf'] if int(conf) > 0]
                
                if not confidences:
                    return False, "No text detected"
                
                avg_confidence = sum(confidences) / len(confidences)
                detected_text = ' '.join([data['text'][i] for i in range(len(data['text'])) 
                                        if int(data['conf'][i]) > confidence_threshold])
                
                if avg_confidence >= confidence_threshold and len(detected_text.strip()) > 3:
                    return True, f"Text detected (confidence: {avg_confidence:.1f}%)"
                else:
                    return False, f"Low confidence text ({avg_confidence:.1f}%)"
                    
            except Exception as e:
                # Fallback to simple text extraction
                text = pytesseract.image_to_string(test_img, config='--psm 6')
                if len(text.strip()) > 3:
                    return True, "Text detected (fallback method)"
                else:
                    return False, "No meaningful text detected"
                    
        except Exception as e:
            return False, f"OCR test failed: {e}"

class PDFProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Processor Pro")
        self.root.geometry("600x600")
        self.root.resizable(True, True)
        self.root.configure(bg='#FAFAFA')
        
        self.style = ModernStyle(root)
        self.image_filter = ImageQualityFilter()
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.operation = tk.StringVar(value="slice_pages")
        self.start_page = tk.IntVar(value=1)
        self.end_page = tk.IntVar(value=1)
        self.max_size_mb = tk.DoubleVar(value=5.0)
        self.enable_ocr = tk.BooleanVar(value=True)
        self.extract_images = tk.BooleanVar(value=True)
        self.smart_filtering = tk.BooleanVar(value=True)
        self.min_image_size = tk.IntVar(value=150)
        
        self.is_processing = False
        self.stop_processing = False
        self.current_thread = None
        
        self.setup_ui()
        
    def setup_ui(self):
        main_container = tk.Frame(self.root, bg='#FAFAFA')
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(2, weight=1)
        
        # Top frame for always-visible elements
        top_frame = tk.Frame(main_container, bg='#FAFAFA')
        top_frame.grid(row=0, column=0, sticky='ew')
        top_frame.columnconfigure(0, weight=1)
        
        # Title and attribution
        title_label = ttk.Label(top_frame, text="PDF Processor Pro", style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
        attribution_label = tk.Label(
            top_frame, text="Made by AboulNasr", 
            font=('Segoe UI', 8), fg=self.style.colors['text_secondary'], 
            cursor="hand2", bg='#FAFAFA'
        )
        attribution_label.pack(side=tk.RIGHT)
        attribution_label.bind("<Button-1>", lambda e: webbrowser.open_new_tab("https://www.instagram.com/mahmoud.aboulnasr/"))
        
        # Control buttons
        control_frame = tk.Frame(main_container, bg='#FAFAFA')
        control_frame.grid(row=1, column=0, sticky='ew', pady=(5, 5))
        control_frame.columnconfigure(0, weight=1)
        control_frame.columnconfigure(1, weight=1)
        
        self.process_btn = tk.Button(
            control_frame, text="🚀 Start", 
            command=self.process_pdf,
            bg='#4CAF50', fg='white', font=('Segoe UI', 9, 'bold'),
            relief=tk.FLAT, padx=12, pady=3, cursor='hand2'
        )
        self.process_btn.grid(row=0, column=0, sticky='e', padx=(0, 5))
        
        self.stop_btn = tk.Button(
            control_frame, text="⏹️ Stop", 
            command=self.stop_process,
            bg='#F44336', fg='white', font=('Segoe UI', 9, 'bold'),
            relief=tk.FLAT, padx=12, pady=3, cursor='hand2', state=tk.DISABLED
        )
        self.stop_btn.grid(row=0, column=1, sticky='w')
        
        # Notebook for tabs
        notebook = ttk.Notebook(main_container, style='Modern.TNotebook')
        notebook.grid(row=2, column=0, sticky='nsew')
        
        # File Loading Tab
        file_tab = tk.Frame(notebook, bg='#FAFAFA')
        notebook.add(file_tab, text="📁 Files")
        file_tab.columnconfigure(1, weight=1)
        
        ttk.Label(file_tab, text="PDF File:", style='Heading.TLabel').grid(row=0, column=0, sticky='w', pady=(5, 2))
        pdf_frame = tk.Frame(file_tab, bg='white')
        pdf_frame.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(0, 5))
        self.pdf_entry = tk.Entry(
            pdf_frame, textvariable=self.pdf_path, 
            font=('Segoe UI', 9), relief=tk.FLAT, 
            bg='#F5F5F5', fg='#212121', insertbackground='#212121'
        )
        self.pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=2)
        tk.Button(
            pdf_frame, text="📂", command=self.browse_pdf,
            bg='#2196F3', fg='white', font=('Segoe UI', 8),
            relief=tk.FLAT, padx=6, pady=1, cursor='hand2'
        ).pack(side=tk.RIGHT)
        self.pdf_info_label = ttk.Label(file_tab, text="No PDF selected", style='Body.TLabel')
        self.pdf_info_label.grid(row=2, column=0, columnspan=2, sticky='w', pady=(2, 5))
        
        ttk.Label(file_tab, text="Output Dir:", style='Heading.TLabel').grid(row=3, column=0, sticky='w', pady=(5, 2))
        output_frame = tk.Frame(file_tab, bg='white')
        output_frame.grid(row=4, column=0, columnspan=2, sticky='ew')
        self.output_entry = tk.Entry(
            output_frame, textvariable=self.output_dir, 
            font=('Segoe UI', 9), relief=tk.FLAT, 
            bg='#F5F5F5', fg='#212121', insertbackground='#212121'
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=2)
        tk.Button(
            output_frame, text="📂", command=self.browse_output,
            bg='#2196F3', fg='white', font=('Segoe UI', 8),
            relief=tk.FLAT, padx=6, pady=1, cursor='hand2'
        ).pack(side=tk.RIGHT)
        
        # Operation Selection Tab
        op_tab = tk.Frame(notebook, bg='#FAFAFA')
        notebook.add(op_tab, text="⚙️ Operations")
        op_tab.columnconfigure(1, weight=1)
        
        operations = [
            ("slice_pages", "📄 Slice Pages", "Extract specific pages"),
            ("slice_size", "💾 Slice by Size", "Split by file size"),
            ("simple_text_extraction", "📄 Simple Text", "Fast text extraction"),
            ("to_text", "📝 Text + OCR", "Advanced text with OCR"),
            ("extract_ocr", "🖼️ Images + OCR", "Extract images and OCR")
        ]
        
        for i, (value, text, desc) in enumerate(operations):
            frame = tk.Frame(op_tab, bg='white')
            frame.grid(row=i, column=0, columnspan=2, sticky='ew', pady=1)
            tk.Radiobutton(
                frame, text=text, variable=self.operation, value=value,
                font=('Segoe UI', 9), bg='white', fg=self.style.colors['text'],
                selectcolor=self.style.colors['primary']
            ).pack(anchor=tk.W)
            ttk.Label(frame, text=desc, style='Body.TLabel').pack(anchor=tk.W, padx=(15, 0))
        
        params_frame = tk.Frame(op_tab, bg='white')
        params_frame.grid(row=len(operations), column=0, columnspan=2, sticky='ew', pady=(5, 0))
        
        ttk.Label(params_frame, text="Page Range:", style='Heading.TLabel').pack(anchor=tk.W)
        page_inputs = tk.Frame(params_frame, bg='white')
        page_inputs.pack(fill=tk.X)
        ttk.Label(page_inputs, text="Start:", style='Body.TLabel').pack(side=tk.LEFT)
        tk.Spinbox(
            page_inputs, from_=1, to=9999, textvariable=self.start_page, 
            width=5, font=('Segoe UI', 9), relief=tk.FLAT, bg='#F5F5F5'
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(page_inputs, text="End:", style='Body.TLabel').pack(side=tk.LEFT)
        tk.Spinbox(
            page_inputs, from_=1, to=9999, textvariable=self.end_page, 
            width=5, font=('Segoe UI', 9), relief=tk.FLAT, bg='#F5F5F5'
        ).pack(side=tk.LEFT)
        
        ttk.Label(params_frame, text="Max Size:", style='Heading.TLabel').pack(anchor=tk.W, pady=(3, 0))
        size_inputs = tk.Frame(params_frame, bg='white')
        size_inputs.pack(fill=tk.X)
        tk.Spinbox(
            size_inputs, from_=0.1, to=100.0, increment=0.5, 
            textvariable=self.max_size_mb, width=5, font=('Segoe UI', 9), 
            relief=tk.FLAT, bg='#F5F5F5'
        ).pack(side=tk.LEFT)
        ttk.Label(size_inputs, text="MB", style='Body.TLabel').pack(side=tk.LEFT)
        
        self.ocr_check_widget = tk.Checkbutton(
            params_frame, text="🔍 Enable OCR", 
            variable=self.enable_ocr, font=('Segoe UI', 8),
            bg='white', fg=self.style.colors['text'], selectcolor=self.style.colors['success']
        )
        self.ocr_check_widget.pack(anchor=tk.W)
        
        self.extract_check_widget = tk.Checkbutton(
            params_frame, text="🖼️ Save Images", 
            variable=self.extract_images, font=('Segoe UI', 8),
            bg='white', fg=self.style.colors['text'], selectcolor=self.style.colors['success']
        )
        self.extract_check_widget.pack(anchor=tk.W)
        
        # New smart filtering options
        self.smart_check_widget = tk.Checkbutton(
            params_frame, text="🧠 Smart Image Filtering", 
            variable=self.smart_filtering, font=('Segoe UI', 8),
            bg='white', fg=self.style.colors['text'], selectcolor=self.style.colors['success']
        )
        self.smart_check_widget.pack(anchor=tk.W)
        
        # Minimum image size setting
        size_filter_frame = tk.Frame(params_frame, bg='white')
        size_filter_frame.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(size_filter_frame, text="Min Image Size:", style='Body.TLabel').pack(side=tk.LEFT)
        tk.Spinbox(
            size_filter_frame, from_=50, to=500, increment=25, 
            textvariable=self.min_image_size, width=5, font=('Segoe UI', 9), 
            relief=tk.FLAT, bg='#F5F5F5'
        ).pack(side=tk.LEFT, padx=(5, 2))
        ttk.Label(size_filter_frame, text="px", style='Body.TLabel').pack(side=tk.LEFT)
        
        # Log Tab
        log_tab = tk.Frame(notebook, bg='#FAFAFA')
        notebook.add(log_tab, text="📊 Logs")
        log_tab.columnconfigure(0, weight=1)
        log_tab.rowconfigure(0, weight=1)
        
        self.output_text = scrolledtext.ScrolledText(
            log_tab, height=6, font=('Consolas', 9), 
            bg='#1E1E1E', fg='#FFFFFF', insertbackground='#FFFFFF', 
            selectbackground='#404040', relief=tk.FLAT
        )
        self.output_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.output_text.tag_configure('info', foreground='#81C784')
        self.output_text.tag_configure('warning', foreground='#FFB74D')
        self.output_text.tag_configure('error', foreground='#E57373')
        self.output_text.tag_configure('success', foreground='#4CAF50')
        self.output_text.tag_configure('filter', foreground='#FF9E80')
        
        clear_btn = tk.Button(
            log_tab, text="🗑️ Clear", 
            command=self.clear_log,
            bg='#FF5722', fg='white', font=('Segoe UI', 9, 'bold'),
            relief=tk.FLAT, padx=12, pady=3, cursor='hand2'
        )
        clear_btn.pack(pady=(5, 0))
        
        # Progress bar
        progress_frame = tk.Frame(main_container, bg='#FAFAFA')
        progress_frame.grid(row=3, column=0, sticky='ew', pady=(5, 0))
        self.progress = ttk.Progressbar(progress_frame, mode='determinate', length=250)
        self.progress.pack(fill=tk.X)
        self.status_label = ttk.Label(progress_frame, text="Ready to process", style='Body.TLabel')
        self.status_label.pack()
        
        self.operation.trace_add("write", self.update_options_sensitivity)
        self.update_options_sensitivity()

    def update_options_sensitivity(self, *args):
        current_op = self.operation.get()
        state = tk.DISABLED if current_op in ["slice_pages", "slice_size", "simple_text_extraction"] else tk.NORMAL
        self.ocr_check_widget.config(state=state)
        self.extract_check_widget.config(state=state)
        self.smart_check_widget.config(state=state)

    def browse_pdf(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if filename:
            self.pdf_path.set(filename)
            self.update_page_info()
    
    def browse_output(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir.set(directory)
    
    def update_page_info(self):
        try:
            if self.pdf_path.get():
                with open(self.pdf_path.get(), 'rb') as file:
                    reader = PyPDF2.PdfReader(file)
                    total_pages = len(reader.pages)
                    self.end_page.set(total_pages)
                    file_size = os.path.getsize(self.pdf_path.get()) / (1024 * 1024)
                    self.pdf_info_label.config(text=f"✅ {total_pages} pages, {file_size:.1f} MB")
                    self.log(f"PDF loaded: {total_pages} pages, {file_size:.1f} MB", 'info')
        except Exception as e:
            self.pdf_info_label.config(text="❌ Error reading PDF")
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
        self.log("⏹️ Stopping...", 'warning')
        self.update_status("Stopping process...")
    
    def process_pdf(self):
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Please select a PDF file")
            return
        if not os.path.exists(self.output_dir.get()):
            messagebox.showerror("Error", "Output directory does not exist")
            return
        self.is_processing = True
        self.stop_processing = False
        self.process_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.current_thread = threading.Thread(target=self._process_pdf_thread, daemon=True)
        self.current_thread.start()
    
    def _process_pdf_thread(self):
        try:
            self.update_status("Initializing...", 0)
            self.log("🚀 Starting...", 'info')
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
                self.update_status("✅ Completed!", 100)
                self.log("✅ Completed successfully!", 'success')
            else:
                self.update_status("⏹️ Stopped", 0)
                self.log("⏹️ Stopped by user", 'warning')
        except Exception as e:
            self.log(f"❌ Error: {e}", 'error')
            self.update_status("❌ Error occurred", 0)
        finally:
            self.is_processing = False
            self.process_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)

    def is_image_worth_processing(self, pix, page_num, img_index):
        """Enhanced image filtering with detailed logging"""
        if not self.smart_filtering.get():
            # Basic size check only
            min_size = self.min_image_size.get()
            if pix.width < min_size or pix.height < min_size:
                self.log(f"⏭️ Page {page_num + 1}, Image {img_index + 1}: Too small ({pix.width}x{pix.height})", 'filter')
                return False
            return True
        
        # Advanced filtering
        try:
            img_data = pix.tobytes("png")
            img_pil = Image.open(io.BytesIO(img_data))
            
            # Quality assessment
            is_worthy, reason = self.image_filter.is_meaningful_image(
                pix, img_pil, min_size=self.min_image_size.get()
            )
            
            if not is_worthy:
                self.log(f"⏭️ Page {page_num + 1}, Image {img_index + 1}: {reason}", 'filter')
                return False
            
            # Quick OCR test if OCR is enabled
            if self.enable_ocr.get():
                has_text, ocr_reason = self.image_filter.quick_ocr_test(img_pil)
                if not has_text:
                    self.log(f"⏭️ Page {page_num + 1}, Image {img_index + 1}: {ocr_reason}", 'filter')
                    return False
                else:
                    self.log(f"✅ Page {page_num + 1}, Image {img_index + 1}: {ocr_reason}", 'info')
            
            return True
            
        except Exception as e:
            self.log(f"⚠️ Error filtering image {img_index + 1}: {e}", 'warning')
            return True  # Default to processing if filtering fails

    def simple_convert_to_text(self):
        if self.stop_processing:
            return
        self.update_status("📝 Extracting text...", 10)
        self.log("📄 Starting simple text extraction...", 'info')
        text_content = []
        try:
            with open(self.pdf_path.get(), 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                total_pages = len(reader.pages)
                for i, page in enumerate(reader.pages):
                    if self.stop_processing:
                        return
                    self.update_status(f"Page {i + 1}/{total_pages}...", 10 + (i/total_pages)*80)
                    text = page.extract_text() or ""
                    text_content.append(f"--- Page {i + 1} ---\n{text}\n")
                    self.log(f"✅ Page {i + 1} extracted", 'info')
        except Exception as e:
            self.log(f"❌ Error reading PDF: {e}", 'error')
            return
        if not self.stop_processing:
            output_path = os.path.join(self.output_dir.get(), f"{Path(self.pdf_path.get()).stem}_simple_text.txt")
            try:
                with open(output_path, 'w', encoding='utf-8') as text_file:
                    text_file.write('\n'.join(text_content))
                self.log(f"✅ Saved: {output_path}", 'success')
            except Exception as e:
                self.log(f"❌ Error saving: {e}", 'error')
    
    def slice_by_pages(self):
        if self.stop_processing:
            return
        self.update_status("📄 Slicing pages...", 25)
        try:
            with open(self.pdf_path.get(), 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                start = max(1, self.start_page.get()) - 1
                end = min(len(reader.pages), self.end_page.get())
                writer = PyPDF2.PdfWriter()
                
                for i in range(start, end):
                    if self.stop_processing:
                        return
                    writer.add_page(reader.pages[i])
                    self.update_status(f"Page {i+1}...", 25 + (i-start)/(end-start)*50)
                    self.log(f"✅ Added page {i+1}", 'info')
                
                if not self.stop_processing:
                    output_path = os.path.join(self.output_dir.get(), 
                                             f"{Path(self.pdf_path.get()).stem}_pages_{start+1}-{end}.pdf")
                    with open(output_path, 'wb') as output_file:
                        writer.write(output_file)
                    self.log(f"✅ Saved: {output_path}", 'success')
                    self.update_status("✅ Pages sliced successfully!", 100)
                    
        except Exception as e:
            self.log(f"❌ Error slicing pages: {e}", 'error')
            self.update_status("❌ Error occurred", 0)
    
    def slice_by_size(self):
        if self.stop_processing:
            return
        self.update_status("💾 Slicing by size...", 10)
        self.log("💾 Starting size-based slicing...", 'info')
        
        try:
            max_size_bytes = self.max_size_mb.get() * 1024 * 1024
            
            with open(self.pdf_path.get(), 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                total_pages = len(reader.pages)
                
                current_writer = PyPDF2.PdfWriter()
                current_size = 0
                part_number = 1
                pages_in_current_part = 0
                
                for i, page in enumerate(reader.pages):
                    if self.stop_processing:
                        return
                    
                    self.update_status(f"Processing page {i+1}/{total_pages}...", 
                                     10 + (i/total_pages)*80)
                    
                    # Add page to current writer
                    current_writer.add_page(page)
                    pages_in_current_part += 1
                    
                    # Estimate current size
                    temp_output = io.BytesIO()
                    current_writer.write(temp_output)
                    current_size = temp_output.tell()
                    temp_output.close()
                    
                    # Check if we need to save current part
                    if current_size >= max_size_bytes or i == total_pages - 1:
                        if not self.stop_processing:
                            output_path = os.path.join(
                                self.output_dir.get(),
                                f"{Path(self.pdf_path.get()).stem}_part_{part_number}.pdf"
                            )
                            with open(output_path, 'wb') as output_file:
                                current_writer.write(output_file)
                            
                            self.log(f"✅ Saved part {part_number}: {pages_in_current_part} pages, "
                                   f"{current_size/1024/1024:.1f} MB", 'success')
                            
                            # Reset for next part
                            current_writer = PyPDF2.PdfWriter()
                            current_size = 0
                            part_number += 1
                            pages_in_current_part = 0
                
        except Exception as e:
            self.log(f"❌ Error slicing by size: {e}", 'error')
            self.update_status("❌ Error occurred", 0)
    
    def convert_to_text(self):
        if self.stop_processing:
            return
        self.update_status("📝 Converting to text with OCR...", 10)
        self.log("📝 Starting advanced text extraction...", 'info')
        
        text_content = []
        
        try:
            doc = fitz.open(self.pdf_path.get())
            total_pages = len(doc)
            
            for page_num in range(total_pages):
                if self.stop_processing:
                    return
                
                self.update_status(f"Processing page {page_num + 1}/{total_pages}...", 
                                 10 + (page_num/total_pages)*80)
                
                page = doc[page_num]
                
                # Extract text using PyMuPDF
                text = page.get_text()
                
                # If OCR is enabled and text is minimal, try OCR
                if self.enable_ocr.get() and len(text.strip()) < 50:
                    try:
                        # Convert page to image
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Higher resolution
                        img_data = pix.tobytes("png")
                        img = Image.open(io.BytesIO(img_data))
                        
                        # Perform OCR
                        ocr_text = pytesseract.image_to_string(img, config='--psm 1')
                        if len(ocr_text.strip()) > len(text.strip()):
                            text = ocr_text
                            self.log(f"📖 OCR applied to page {page_num + 1}", 'info')
                        
                    except Exception as ocr_error:
                        self.log(f"⚠️ OCR failed for page {page_num + 1}: {ocr_error}", 'warning')
                
                text_content.append(f"--- Page {page_num + 1} ---\n{text}\n")
                self.log(f"✅ Page {page_num + 1} processed", 'info')
            
            doc.close()
            
            # Save text file
            if not self.stop_processing:
                output_path = os.path.join(self.output_dir.get(), 
                                         f"{Path(self.pdf_path.get()).stem}_text_ocr.txt")
                with open(output_path, 'w', encoding='utf-8') as text_file:
                    text_file.write('\n'.join(text_content))
                
                self.log(f"✅ Text saved: {output_path}", 'success')
                self.update_status("✅ Text extraction completed!", 100)
                
        except Exception as e:
            self.log(f"❌ Error in text conversion: {e}", 'error')
            self.update_status("❌ Error occurred", 0)
    
    def extract_and_ocr(self):
        if self.stop_processing:
            return
        self.update_status("🖼️ Extracting images and performing OCR...", 10)
        self.log("🖼️ Starting image extraction and OCR...", 'info')
        
        try:
            doc = fitz.open(self.pdf_path.get())
            total_pages = len(doc)
            
            # Create output directories
            base_name = Path(self.pdf_path.get()).stem
            images_dir = os.path.join(self.output_dir.get(), f"{base_name}_images")
            os.makedirs(images_dir, exist_ok=True)
            
            all_ocr_text = []
            total_images_processed = 0
            total_images_saved = 0
            
            for page_num in range(total_pages):
                if self.stop_processing:
                    return
                
                page = doc[page_num]
                self.update_status(f"Processing page {page_num + 1}/{total_pages}...", 
                                 10 + (page_num/total_pages)*80)
                
                # Get images from page
                image_list = page.get_images()
                page_ocr_text = []
                
                self.log(f"📄 Page {page_num + 1}: Found {len(image_list)} images", 'info')
                
                for img_index, img in enumerate(image_list):
                    if self.stop_processing:
                        return
                    
                    total_images_processed += 1
                    
                    try:
                        # Get image data
                        xref = img[0]
                        pix = fitz.Pixmap(doc, xref)
                        
                        # Log image details for debugging
                        self.log(f"🔍 Page {page_num + 1}, Image {img_index + 1}: "
                               f"{pix.width}x{pix.height}, {pix.n} channels", 'info')
                        
                        # Handle CMYK images by converting them
                        if pix.n - pix.alpha >= 4:  # CMYK
                            self.log(f"🔄 Converting CMYK image {img_index + 1} to RGB", 'info')
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                        
                        # Check if image is worth processing
                        if not self.is_image_worth_processing(pix, page_num, img_index):
                            pix = None
                            continue
                        
                        # Convert to PIL Image
                        img_data = pix.tobytes("png")
                        img_pil = Image.open(io.BytesIO(img_data))
                        
                        # Save image if enabled
                        if self.extract_images.get():
                            img_filename = f"page_{page_num + 1}_img_{img_index + 1}.png"
                            img_path = os.path.join(images_dir, img_filename)
                            img_pil.save(img_path)
                            total_images_saved += 1
                            self.log(f"💾 Saved: {img_filename} ({pix.width}x{pix.height})", 'success')
                        
                        # Perform OCR if enabled
                        if self.enable_ocr.get():
                            try:
                                # Enhance image for better OCR
                                enhanced_img = self.enhance_image_for_ocr(img_pil)
                                
                                # Perform OCR with less restrictive character set
                                ocr_text = pytesseract.image_to_string(enhanced_img, config='--psm 6')
                                
                                if ocr_text.strip():
                                    page_ocr_text.append(f"Image {img_index + 1}: {ocr_text.strip()}")
                                    self.log(f"📖 OCR completed for image {img_index + 1}: "
                                           f"{len(ocr_text.strip())} characters", 'success')
                                else:
                                    self.log(f"📖 OCR found no text in image {img_index + 1}", 'info')
                                
                            except Exception as ocr_error:
                                self.log(f"⚠️ OCR failed for image {img_index + 1}: {ocr_error}", 'warning')
                        
                        pix = None
                        
                    except Exception as img_error:
                        self.log(f"⚠️ Error processing image {img_index + 1}: {img_error}", 'warning')
                
                # Add page OCR results
                if page_ocr_text:
                    all_ocr_text.append(f"--- Page {page_num + 1} ---\n" + "\n".join(page_ocr_text) + "\n")
            
            doc.close()
            
            # Save OCR results
            if not self.stop_processing and all_ocr_text:
                ocr_output_path = os.path.join(self.output_dir.get(), f"{base_name}_ocr_results.txt")
                with open(ocr_output_path, 'w', encoding='utf-8') as ocr_file:
                    ocr_file.write('\n'.join(all_ocr_text))
                self.log(f"✅ OCR results saved: {ocr_output_path}", 'success')
            
            # Summary
            self.log(f"📊 Summary: {total_images_processed} images processed, "
                   f"{total_images_saved} images saved", 'success')
            self.update_status("✅ Image extraction and OCR completed!", 100)
            
        except Exception as e:
            self.log(f"❌ Error in image extraction: {e}", 'error')
            self.update_status("❌ Error occurred", 0)
    
    def enhance_image_for_ocr(self, img):
        """Enhance image quality for better OCR results"""
        try:
            # Convert to grayscale if not already
            if img.mode != 'L':
                img = img.convert('L')
            
            # Resize if too small (OCR works better on larger images)
            if img.width < 300 or img.height < 300:
                scale_factor = max(300 / img.width, 300 / img.height)
                new_size = (int(img.width * scale_factor), int(img.height * scale_factor))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
            
            # Enhance contrast
            enhancer = ImageEnhance.Contrast(img)
            img = enhancer.enhance(1.2)
            
            # Enhance sharpness
            enhancer = ImageEnhance.Sharpness(img)
            img = enhancer.enhance(1.1)
            
            return img
            
        except Exception:
            return img  # Return original if enhancement fails

# Main application entry point
def main():
    try:
        # Check if Tesseract is available
        pytesseract.get_tesseract_version()
    except:
        messagebox.showerror(
            "Tesseract Not Found", 
            "Tesseract OCR is not installed or not in PATH.\n"
            "OCR features will be disabled.\n\n"
            "To install Tesseract:\n"
            "1. Download from: https://github.com/UB-Mannheim/tesseract/wiki\n"
            "2. Add to system PATH\n"
            "3. Restart the application"
        )
    
    root = tk.Tk()
    app = PDFProcessor(root)
    
    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()
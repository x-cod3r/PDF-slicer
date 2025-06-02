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

class PDFProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Processor - Slice, Convert & OCR")
        self.root.geometry("800x700")
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.operation = tk.StringVar(value="slice_pages")
        self.start_page = tk.IntVar(value=1)
        self.end_page = tk.IntVar(value=1)
        self.max_size_mb = tk.DoubleVar(value=5.0)
        self.enable_ocr = tk.BoolVar(value=True)
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # PDF Selection
        ttk.Label(main_frame, text="Select PDF File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        pdf_frame = ttk.Frame(main_frame)
        pdf_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Entry(pdf_frame, textvariable=self.pdf_path, width=60).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(pdf_frame, text="Browse", command=self.browse_pdf).grid(row=0, column=1, padx=(5, 0))
        
        pdf_frame.columnconfigure(0, weight=1)
        
        # Output Directory
        ttk.Label(main_frame, text="Output Directory:").grid(row=2, column=0, sticky=tk.W, pady=(20, 5))
        
        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Entry(output_frame, textvariable=self.output_dir, width=60).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(output_frame, text="Browse", command=self.browse_output).grid(row=0, column=1, padx=(5, 0))
        
        output_frame.columnconfigure(0, weight=1)
        
        # Operation Selection
        ttk.Label(main_frame, text="Operation:").grid(row=4, column=0, sticky=tk.W, pady=(20, 5))
        
        operations_frame = ttk.LabelFrame(main_frame, text="Choose Operation", padding="10")
        operations_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Radiobutton(operations_frame, text="Slice by Page Range", 
                       variable=self.operation, value="slice_pages").grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(operations_frame, text="Slice by File Size", 
                       variable=self.operation, value="slice_size").grid(row=1, column=0, sticky=tk.W)
        ttk.Radiobutton(operations_frame, text="Convert to Text", 
                       variable=self.operation, value="to_text").grid(row=2, column=0, sticky=tk.W)
        ttk.Radiobutton(operations_frame, text="Extract Images & OCR", 
                       variable=self.operation, value="extract_ocr").grid(row=3, column=0, sticky=tk.W)
        
        # Parameters Frame
        params_frame = ttk.LabelFrame(main_frame, text="Parameters", padding="10")
        params_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Page Range Parameters
        ttk.Label(params_frame, text="Start Page:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        ttk.Spinbox(params_frame, from_=1, to=9999, textvariable=self.start_page, width=10).grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        
        ttk.Label(params_frame, text="End Page:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        ttk.Spinbox(params_frame, from_=1, to=9999, textvariable=self.end_page, width=10).grid(row=0, column=3, sticky=tk.W)
        
        # Size Parameter
        ttk.Label(params_frame, text="Max Size (MB):").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(10, 0))
        ttk.Spinbox(params_frame, from_=0.1, to=100.0, increment=0.5, textvariable=self.max_size_mb, width=10).grid(row=1, column=1, sticky=tk.W, pady=(10, 0))
        
        # OCR Option
        ttk.Checkbutton(params_frame, text="Enable OCR for images", variable=self.enable_ocr).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        
        # Process Button
        ttk.Button(main_frame, text="Process PDF", command=self.process_pdf, 
                  style="Accent.TButton").grid(row=7, column=0, columnspan=3, pady=20)
        
        # Progress Bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Status Label
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.grid(row=9, column=0, columnspan=3, pady=5)
        
        # Output Text Area
        ttk.Label(main_frame, text="Output Log:").grid(row=10, column=0, sticky=tk.W, pady=(20, 5))
        self.output_text = scrolledtext.ScrolledText(main_frame, height=10, width=80)
        self.output_text.grid(row=11, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.rowconfigure(11, weight=1)
        
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
                    self.log(f"PDF loaded: {total_pages} pages")
        except Exception as e:
            self.log(f"Error reading PDF: {e}")
    
    def log(self, message):
        self.output_text.insert(tk.END, f"{message}\n")
        self.output_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self, status):
        self.status_label.config(text=status)
        self.root.update_idletasks()
    
    def process_pdf(self):
        if not self.pdf_path.get():
            messagebox.showerror("Error", "Please select a PDF file")
            return
        
        if not os.path.exists(self.output_dir.get()):
            messagebox.showerror("Error", "Output directory does not exist")
            return
        
        # Run processing in separate thread to prevent UI freezing
        threading.Thread(target=self._process_pdf_thread, daemon=True).start()
    
    def _process_pdf_thread(self):
        try:
            self.progress.start()
            self.update_status("Processing...")
            
            operation = self.operation.get()
            
            if operation == "slice_pages":
                self.slice_by_pages()
            elif operation == "slice_size":
                self.slice_by_size()
            elif operation == "to_text":
                self.convert_to_text()
            elif operation == "extract_ocr":
                self.extract_and_ocr()
                
            self.update_status("Complete!")
            
        except Exception as e:
            self.log(f"Error: {e}")
            self.update_status("Error occurred")
        finally:
            self.progress.stop()
    
    def slice_by_pages(self):
        with open(self.pdf_path.get(), 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            start = max(1, self.start_page.get()) - 1  # Convert to 0-based index
            end = min(len(reader.pages), self.end_page.get())
            
            writer = PyPDF2.PdfWriter()
            
            for i in range(start, end):
                writer.add_page(reader.pages[i])
            
            output_path = os.path.join(
                self.output_dir.get(),
                f"{Path(self.pdf_path.get()).stem}_pages_{start+1}-{end}.pdf"
            )
            
            with open(output_path, 'wb') as output_file:
                writer.write(output_file)
            
            self.log(f"Created: {output_path}")
    
    def slice_by_size(self):
        max_size_bytes = self.max_size_mb.get() * 1024 * 1024
        
        with open(self.pdf_path.get(), 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            current_writer = PyPDF2.PdfWriter()
            current_size = 0
            part_num = 1
            start_page = 1
            
            for i, page in enumerate(reader.pages):
                # Estimate page size (rough approximation)
                page_size = len(page.extract_text().encode('utf-8')) * 2  # Rough estimate
                
                if current_size + page_size > max_size_bytes and len(current_writer.pages) > 0:
                    # Save current part
                    output_path = os.path.join(
                        self.output_dir.get(),
                        f"{Path(self.pdf_path.get()).stem}_part_{part_num}.pdf"
                    )
                    
                    with open(output_path, 'wb') as output_file:
                        current_writer.write(output_file)
                    
                    self.log(f"Created: {output_path} (pages {start_page}-{i})")
                    
                    # Start new part
                    current_writer = PyPDF2.PdfWriter()
                    current_size = 0
                    part_num += 1
                    start_page = i + 1
                
                current_writer.add_page(page)
                current_size += page_size
            
            # Save final part
            if len(current_writer.pages) > 0:
                output_path = os.path.join(
                    self.output_dir.get(),
                    f"{Path(self.pdf_path.get()).stem}_part_{part_num}.pdf"
                )
                
                with open(output_path, 'wb') as output_file:
                    current_writer.write(output_file)
                
                self.log(f"Created: {output_path} (pages {start_page}-{len(reader.pages)})")
    
    def convert_to_text(self):
        text_content = []
        
        # Try PyMuPDF first for better text extraction
        try:
            pdf_doc = fitz.open(self.pdf_path.get())
            
            for page_num in range(len(pdf_doc)):
                page = pdf_doc[page_num]
                text = page.get_text()
                
                if not text.strip() and self.enable_ocr.get():
                    # If no text found, try OCR on page image
                    self.log(f"No text on page {page_num + 1}, trying OCR...")
                    pix = page.get_pixmap()
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data))
                    text = pytesseract.image_to_string(img)
                
                text_content.append(f"--- Page {page_num + 1} ---\n{text}\n")
                self.log(f"Processed page {page_num + 1}")
            
            pdf_doc.close()
            
        except Exception as e:
            self.log(f"PyMuPDF failed, trying PyPDF2: {e}")
            # Fallback to PyPDF2
            with open(self.pdf_path.get(), 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                
                for i, page in enumerate(reader.pages):
                    text = page.extract_text()
                    text_content.append(f"--- Page {i + 1} ---\n{text}\n")
                    self.log(f"Processed page {i + 1}")
        
        # Save text file
        output_path = os.path.join(
            self.output_dir.get(),
            f"{Path(self.pdf_path.get()).stem}.txt"
        )
        
        with open(output_path, 'w', encoding='utf-8') as text_file:
            text_file.write('\n'.join(text_content))
        
        self.log(f"Text saved to: {output_path}")
    
    def extract_and_ocr(self):
        pdf_doc = fitz.open(self.pdf_path.get())
        
        image_dir = os.path.join(self.output_dir.get(), f"{Path(self.pdf_path.get()).stem}_images")
        os.makedirs(image_dir, exist_ok=True)
        
        all_text = []
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc[page_num]
            
            # Get images from page
            image_list = page.get_images()
            
            page_text = []
            
            if image_list:
                self.log(f"Found {len(image_list)} images on page {page_num + 1}")
                
                for img_index, img in enumerate(image_list):
                    xref = img[0]
                    pix = fitz.Pixmap(pdf_doc, xref)
                    
                    if pix.n - pix.alpha < 4:  # GRAY or RGB
                        img_data = pix.tobytes("png")
                        img_path = os.path.join(image_dir, f"page_{page_num + 1}_img_{img_index + 1}.png")
                        
                        with open(img_path, "wb") as img_file:
                            img_file.write(img_data)
                        
                        if self.enable_ocr.get():
                            # Perform OCR
                            img = Image.open(io.BytesIO(img_data))
                            ocr_text = pytesseract.image_to_string(img)
                            
                            if ocr_text.strip():
                                page_text.append(f"Image {img_index + 1} OCR:\n{ocr_text}")
                                self.log(f"OCR completed for image {img_index + 1} on page {page_num + 1}")
                        
                        self.log(f"Saved: {img_path}")
                    
                    pix = None
            else:
                # No embedded images, convert entire page to image for OCR
                if self.enable_ocr.get():
                    pix = page.get_pixmap()
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data))
                    ocr_text = pytesseract.image_to_string(img)
                    
                    if ocr_text.strip():
                        page_text.append(f"Page OCR:\n{ocr_text}")
                        self.log(f"OCR completed for page {page_num + 1}")
            
            if page_text:
                all_text.append(f"--- Page {page_num + 1} ---\n" + "\n\n".join(page_text))
        
        pdf_doc.close()
        
        # Save OCR results
        if all_text:
            ocr_output_path = os.path.join(
                self.output_dir.get(),
                f"{Path(self.pdf_path.get()).stem}_ocr.txt"
            )
            
            with open(ocr_output_path, 'w', encoding='utf-8') as text_file:
                text_file.write('\n\n'.join(all_text))
            
            self.log(f"OCR text saved to: {ocr_output_path}")
        
        self.log(f"Images extracted to: {image_dir}")

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
    
    root = tk.Tk()
    app = PDFProcessor(root)
    root.mainloop()

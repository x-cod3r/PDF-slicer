# 🚀 PDF Processor Pro 🚀

**PDF Processor Pro** is a versatile desktop application designed to help you manage and process your PDF files with ease. Slice, convert, and extract information from your PDFs using a modern and user-friendly interface.

---

## ✨ Features ✨

*   📄 **Slice by Page Range:** Extract a specific range of pages from your PDF.
*   💾 **Slice by File Size:** Split your PDF into smaller parts, each under a specified file size limit.
*   📝 **Convert to Text (Advanced, with OCR):** Extract all text content from your PDF. Includes powerful OCR capabilities to pull text from images within the PDF.
*   📑 **Simple Text Extraction (Fast, No OCR):** Quickly extract text from PDFs that have selectable text. This mode is faster as it doesn't perform OCR.
*   🖼️ **Extract Images & Perform OCR:** Pull out all images from your PDF and optionally perform OCR on them to extract text. Extracted images can be saved separately.
*   ✨ **Modern & Adaptive GUI:** A clean, intuitive, and responsive graphical user interface built with Tkinter.
*   📊 **Detailed Processing Log:** Keep track of all operations and see detailed logs in real-time.
*   ⏹️ **Stop Current Process:** Easily stop any ongoing PDF processing task.
*   🔗 **Attribution:** [Made by AboulNasr](https://www.instagram.com/mahmoud.aboulnasr/)

---

## 🛠️ Technologies Used 🛠️

*   **Python 3**
*   **Tkinter** (for the GUI)
*   **PyMuPDF (Fitz)** (for advanced PDF manipulation and image extraction)
*   **Pillow (PIL)** (for image processing)
*   **Pytesseract** (for Optical Character Recognition - OCR)
*   **PyPDF2** (for simple text extraction and basic PDF operations)

---

## ⚙️ Dependencies & Installation ⚙️

**1. Python Packages:**
You can install the necessary Python packages using pip:
```bash
pip install PyPDF2 PyMuPDF Pillow pytesseract
```

**2. Tesseract OCR Engine:**
For OCR functionalities (used in "Convert to Text (Advanced, OCR)" and "Extract Images & OCR"), you need to have Google's Tesseract OCR engine installed on your system.

*   **Windows:** Download and run the installer from [Tesseract at UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki). Ensure you add Tesseract to your system's PATH during installation or note the installation path for `pytesseract`.
*   **macOS:** You can install Tesseract using Homebrew:
    ```bash
    brew install tesseract
    ```
*   **Linux (Debian/Ubuntu):**
    ```bash
    sudo apt-get install tesseract-ocr
    ```
The application will show a warning if Tesseract is not found but will still allow you to use non-OCR features.

---

## ▶️ How to Run ▶️

1.  Ensure all dependencies listed above are installed.
2.  Clone this repository or download the `main.py` file.
3.  Navigate to the directory containing `main.py` in your terminal.
4.  Run the application using:
    ```bash
    python main.py
    ```

---

#PDFprocessor #Python #Tkinter #OCR #PDFtools #Productivity #DesktopApp #Utility

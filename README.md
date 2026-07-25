# 🚀 ForgePDF - Smart Offline PDF Toolkit

A powerful, high-performance, 100% offline desktop application for Windows built with Python, PyMuPDF, `pdf2docx`, and `CustomTkinter`.

---

## ✨ Key Features

1. 📄 **PDF to Word Converter (.docx)**
   - High-fidelity conversion preserving original layout, margins, font sizes, and line spacing.
   - **Spatial & Layout-Aware OCR Engine**: Preserves 1:1 visual match for scanned image-based PDFs.

2. 🗜️ **PDF Compressor**
   - Targeted compression presets:
     - ⚡ **Ultra Small (~1 - 2 MB)**
     - ⚖️ **Medium (~3 - 5 MB)**
     - 🎨 **High Quality (~6 - 10 MB)**
   - Smart Size Safety Guard guarantees files will strictly decrease in size.
   - 100% true-color compositing for transparent PNGs and complex vector graphics.

3. 🔢 **Smart Page Numbering**
   - Auto-detects `BAB I` / `PENDAHULUAN`.
   - Automatically formats preliminary pages as Roman numerals (`i, ii, iii...`) and switches to Latin (`1, 2, 3...`) from Chapter 1.
   - Automatically skips Cover (Page 1).

4. 🛡️ **100% Offline & Private**
   - Zero internet connection required.
   - All processing is done locally on your CPU/RAM.

---

## 🚀 Quick Setup & Desktop Shortcut (1-Click Installation)

1. **Clone or Download Repository**:
   ```bash
   git clone https://github.com/YOUR_USERNAME/ForgePDF.git
   cd ForgePDF
   ```
2. **Run 1-Click Installer**:
   - Double-click **`install.bat`** inside the folder.
   - It will automatically install all dependencies and **create a clean silent `ForgePDF` shortcut on your Desktop**!

---

## 🛠️ Tech Stack
- **GUI**: CustomTkinter / Tkinter
- **PDF & Image Engine**: PyMuPDF (`fitz`), Pillow
- **DOCX Engine**: `pdf2docx`, `python-docx`
- **OCR Engine**: EasyOCR, PyTesseract

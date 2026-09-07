# 🚀 ForgePDF - Smart Offline PDF Toolkit

A powerful, high-performance, 100% offline desktop application for Windows & Linux built with Python, PyMuPDF, `pdf2docx`, and `CustomTkinter`.

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

4. 🔤 **OCR Gambar (Ekstrak Teks)**
   - Extract text from photos, scans, or screenshots — 100% local (RapidOCR/onnx).
   - Results sorted by position (left-to-right, top-to-bottom), copyable or savable as `.txt`.

4. 🛡️ **100% Offline & Private**
   - Zero internet connection required.
   - All processing is done locally on your CPU/RAM.

---

## 💻 How to Install & Run

### 🪟 On Windows OS (1-Click Setup)

1. **Clone or Download Repository**:
   ```bash
   git clone https://github.com/jinsakai-j/ForgePDF.git
   cd ForgePDF
   ```
2. **Run 1-Click Installer**:
   - Double-click **`install.bat`** inside the folder.
   - It will automatically install all dependencies and **create a clean silent `ForgePDF` shortcut on your Desktop**!

3. **Or Run Manually via Terminal**:
   ```bash
   pip install -r requirements.txt
   python main.py
   ```

---

### 🐧 On Linux OS (Kali Linux / Ubuntu / Debian)

1. **Clone Repository**:
   ```bash
   git clone https://github.com/jinsakai-j/ForgePDF.git
   cd ForgePDF
   ```
2. **Install Required Packages**:
   ```bash
   pip3 install -r requirements.txt --break-system-packages
   ```
3. **Launch Application**:
   ```bash
   python3 main.py
   ```

---

## 🛠️ Tech Stack
- **GUI**: CustomTkinter / Tkinter
- **PDF & Image Engine**: PyMuPDF (`fitz`), Pillow
- **DOCX Engine**: `pdf2docx`, `python-docx`
- **OCR Engine**: EasyOCR (PDF → Word), RapidOCR (OCR Gambar)

import os
import time
import io
import re
import fitz  # PyMuPDF
from pdf2docx import Converter
import docx
from docx.shared import Pt, Inches
from PIL import Image

class CancellationException(Exception):
    pass

# Global EasyOCR Reader instance
EASYOCR_READER = None

def get_easyocr_reader(status_callback=None):
    global EASYOCR_READER
    if EASYOCR_READER is None:
        try:
            import easyocr
            if status_callback:
                status_callback("Menyiapkan model OCR (Bahasa Indonesia & Inggris)...", 0.15)
            EASYOCR_READER = easyocr.Reader(['id', 'en'], gpu=False, verbose=False)
        except Exception:
            EASYOCR_READER = False
    return EASYOCR_READER if EASYOCR_READER is not False else None


def check_cancel(status_callback, msg, prog):
    if status_callback:
        res = status_callback(msg, prog)
        if res is False:
            raise CancellationException("Proses dibatalkan oleh pengguna.")


def compress_pdf(pdf_path, output_pdf_path=None, quality="medium", status_callback=None):
    """
    Compresses PDF using targeted size presets:
    - 'ultra_small' / 'low': Target ~1-2 MB (96 DPI, JPEG Quality 55)
    - 'medium': Target ~3-5 MB (130 DPI, JPEG Quality 72)
    - 'high': Target ~6-10 MB (180 DPI, JPEG Quality 85)
    """
    if not os.path.exists(pdf_path):
        return False, f"File PDF tidak ditemukan: {pdf_path}", 0, 0

    if not output_pdf_path:
        output_pdf_path = os.path.splitext(pdf_path)[0] + "_compressed.pdf"

    try:
        orig_size = os.path.getsize(pdf_path)
        check_cancel(status_callback, "Membuka PDF untuk kompresi...", 0.1)

        doc = fitz.open(pdf_path)
        total_pages = len(doc)

        preset_cfg = {
            "ultra_small": {"dpi": 96, "quality": 55},
            "low": {"dpi": 96, "quality": 55},
            "medium": {"dpi": 130, "quality": 72},
            "high": {"dpi": 180, "quality": 85}
        }
        cfg = preset_cfg.get(quality, preset_cfg["medium"])

        target_dpi = cfg["dpi"]
        jpeg_q = cfg["quality"]

        new_doc = fitz.open()

        for pno in range(total_pages):
            prog = 0.1 + (0.80 * (pno / total_pages))
            check_cancel(status_callback, f"Mengompresi Halaman {pno + 1}/{total_pages} ({target_dpi} DPI)...", prog)

            page = doc[pno]
            rect = page.rect

            # Render page with alpha enabled
            pix = page.get_pixmap(dpi=target_dpi, alpha=True)
            pil_img = Image.open(io.BytesIO(pix.tobytes("png")))

            # Composite onto solid white background canvas
            bg = Image.new("RGB", pil_img.size, (255, 255, 255))
            if pil_img.mode == "RGBA":
                bg.paste(pil_img, mask=pil_img.split()[3])
            else:
                bg.paste(pil_img)

            out_buffer = io.BytesIO()
            bg.save(out_buffer, format="JPEG", quality=jpeg_q, optimize=True)
            img_bytes = out_buffer.getvalue()

            new_page = new_doc.new_page(width=rect.width, height=rect.height)
            new_page.insert_image(rect, stream=img_bytes)

        check_cancel(status_callback, "Menyimpan PDF hasil kompresi...", 0.92)

        new_doc.save(
            output_pdf_path,
            garbage=4,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True
        )
        new_doc.close()
        doc.close()

        new_size = os.path.getsize(output_pdf_path)

        if new_size >= orig_size:
            doc_raw = fitz.open(pdf_path)
            doc_raw.save(
                output_pdf_path,
                garbage=4,
                deflate=True,
                deflate_images=True,
                deflate_fonts=True
            )
            doc_raw.close()
            new_size = os.path.getsize(output_pdf_path)

        return True, output_pdf_path, orig_size, new_size

    except CancellationException as ce:
        return False, str(ce), 0, 0
    except Exception as e:
        return False, f"Terjadi kesalahan saat kompresi: {str(e)}", 0, 0


def convert_pdf_ocr_1to1(pdf_path, output_docx_path, pages_to_convert=None, status_callback=None):
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    
    if pages_to_convert is None:
        pages_to_convert = list(range(total_pages))

    word_doc = docx.Document()
    
    check_cancel(status_callback, "Memuat engine OCR 1:1 Visual Match...", 0.1)
    
    reader = get_easyocr_reader(status_callback)
    total_sel = len(pages_to_convert)
    start_time = time.time()

    for idx, pno in enumerate(pages_to_convert):
        prog = 0.2 + (0.75 * (idx / total_sel))
        check_cancel(status_callback, f"Memproses presisi 1:1 Halaman {pno + 1}/{total_pages}...", prog)

        page = doc[pno]
        page_width_in = page.rect.width / 72.0
        page_height_in = page.rect.height / 72.0

        pix = page.get_pixmap(dpi=300)
        img_bytes = pix.tobytes("png")
        pil_img = Image.open(io.BytesIO(img_bytes))
        img_w, img_h = pil_img.size

        if idx == 0:
            section = word_doc.sections[0]
        else:
            section = word_doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)

        section.page_width = Inches(page_width_in)
        section.page_height = Inches(page_height_in)
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

        ocr_results = []
        if reader is not None:
            try:
                raw_res = reader.readtext(img_bytes, detail=1)
                for item in raw_res:
                    try:
                        box = item[0]
                        txt = str(item[1]).strip()
                        c = float(item[2])
                        
                        x1 = float(box[0][0])
                        y1 = float(box[0][1])
                        x2 = float(box[1][0])
                        y2 = float(box[2][1])
                        
                        if txt and c > 0.15:
                            ocr_results.append((y1, x1, x2, y2, txt))
                    except Exception:
                        continue
            except Exception:
                ocr_results = []

        if ocr_results:
            sorted_blocks = sorted(ocr_results, key=lambda b: (b[0], b[1]))

            lines_group = []
            curr_line = []
            curr_y = -1

            for y1, x1, x2, y2, txt in sorted_blocks:
                if curr_y == -1 or abs(y1 - curr_y) < (img_h * 0.015):
                    curr_line.append((y1, x1, x2, y2, txt))
                    curr_y = y1
                else:
                    lines_group.append(curr_line)
                    curr_line = [(y1, x1, x2, y2, txt)]
                    curr_y = y1
            if curr_line:
                lines_group.append(curr_line)

            for line_blocks in lines_group:
                p = word_doc.add_paragraph()
                
                line_left = min(b[1] for b in line_blocks)
                line_right = max(b[2] for b in line_blocks)
                line_height = max(b[3] - b[0] for b in line_blocks)

                rel_center = (line_left + line_right) / 2 / img_w
                if rel_center < 0.35 and (line_right - line_left) < (img_w * 0.5):
                    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT
                elif rel_center > 0.65 and (line_right - line_left) < (img_w * 0.5):
                    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT
                elif 0.35 <= rel_center <= 0.65 and (line_right - line_left) < (img_w * 0.7):
                    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT

                line_text = "  ".join(b[4] for b in line_blocks)
                run = p.add_run(line_text)
                
                est_pt = max(8, min(24, int(line_height * 72 / img_h * 1.4)))
                run.font.size = Pt(est_pt)
                
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.15

        else:
            img_stream = io.BytesIO()
            pil_img.save(img_stream, format='PNG')
            img_stream.seek(0)
            p = word_doc.add_paragraph()
            p.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
            word_doc.add_picture(img_stream, width=Inches(page_width_in - 0.8))

    word_doc.save(output_docx_path)
    doc.close()
    
    elapsed = round(time.time() - start_time, 2)
    return True, elapsed


def int_to_roman(num):
    val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
    syb = ["m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i"]
    roman_num = ''
    i = 0
    while num > 0:
        for _ in range(num // val[i]):
            roman_num += syb[i]
            num -= val[i]
        i += 1
    return roman_num


def parse_pages(pages_str, total_pages):
    if not pages_str or not pages_str.strip():
        return None
    
    pages = set()
    parts = pages_str.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            subparts = part.split('-')
            if len(subparts) == 2 and subparts[0].isdigit() and subparts[1].isdigit():
                start = int(subparts[0]) - 1
                end = int(subparts[1]) - 1
                for p in range(max(0, start), min(total_pages, end + 1)):
                    pages.add(p)
        elif part.isdigit():
            p = int(part) - 1
            if 0 <= p < total_pages:
                pages.add(p)
                
    if not pages:
        return None
    return sorted(list(pages))


def detect_bab1_page(doc):
    for pno in range(len(doc)):
        text = doc[pno].get_text("text").upper()
        if re.search(r'\bBAB\s+(I|1|01)\b', text) or re.search(r'\bPENDAHULUAN\b', text):
            return pno
    return None


def add_smart_page_numbers(
    pdf_path, 
    output_pdf_path=None, 
    mode="auto", 
    pos_vert="bawah", 
    pos_horiz="tengah", 
    skip_cover=True,
    roman_end_page=None,
    status_callback=None
):
    if not os.path.exists(pdf_path):
        return False, f"File PDF tidak ditemukan: {pdf_path}"

    if not output_pdf_path:
        output_pdf_path = os.path.splitext(pdf_path)[0] + "_numbered.pdf"

    try:
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        
        check_cancel(status_callback, "Memulai penambahan nomor halaman...", 0.1)

        bab1_index = None
        if mode == "auto":
            bab1_index = detect_bab1_page(doc)
            if status_callback:
                if bab1_index is not None:
                    check_cancel(status_callback, f"Deteksi otomatis: BAB I ditemukan di Halaman {bab1_index + 1}!", 0.2)
                else:
                    check_cancel(status_callback, "BAB I tidak terdeteksi otomatis, me-nomorkan standar...", 0.2)

        start_latin_index = None
        if mode == "auto" and bab1_index is not None:
            start_latin_index = bab1_index
        elif mode == "custom" and roman_end_page is not None:
            start_latin_index = max(0, min(total_pages - 1, roman_end_page))

        for i, page in enumerate(doc):
            if skip_cover and i == 0:
                continue

            rect = page.rect
            margin = 30

            if start_latin_index is not None:
                if i < start_latin_index:
                    roman_val = i if skip_cover else (i + 1)
                    text = int_to_roman(max(1, roman_val))
                else:
                    latin_val = (i - start_latin_index) + 1
                    text = str(latin_val)
            elif mode == "roman_all":
                roman_val = i if skip_cover else (i + 1)
                text = int_to_roman(max(1, roman_val))
            else:
                text = str(i if skip_cover else (i + 1))

            if pos_vert == "atas":
                y = margin
            else:
                y = rect.height - margin

            if pos_horiz == "kiri":
                x = margin
                align = fitz.TEXT_ALIGN_LEFT
            elif pos_horiz == "kanan":
                x = rect.width - margin
                align = fitz.TEXT_ALIGN_RIGHT
            else:
                x = rect.width / 2
                align = fitz.TEXT_ALIGN_CENTER

            page.insert_text(
                (x, y),
                text,
                fontsize=10,
                fontname="helv",
                color=(0.2, 0.2, 0.2),
                align=align
            )

            check_cancel(status_callback, f"Menambahkan nomor halaman {i + 1}/{total_pages}...", 0.2 + 0.7 * (i / total_pages))

        doc.save(output_pdf_path)
        doc.close()
        return True, output_pdf_path

    except CancellationException as ce:
        return False, str(ce)
    except Exception as e:
        return False, f"Gagal menambahkan nomor halaman: {str(e)}"


def convert_pdf_to_word(pdf_path, output_docx_path=None, pages_str="", use_ocr=False, status_callback=None):
    if not os.path.exists(pdf_path):
        return False, f"File PDF tidak ditemukan: {pdf_path}"
    
    if not output_docx_path:
        output_docx_path = os.path.splitext(pdf_path)[0] + ".docx"

    try:
        if use_ocr:
            check_cancel(status_callback, "Memulai konversi mode OCR 1:1 Visual Presisi...", 0.1)
            
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            pages_to_convert = parse_pages(pages_str, total_pages)
            doc.close()

            success, elapsed = convert_pdf_ocr_1to1(pdf_path, output_docx_path, pages_to_convert, status_callback)
            if status_callback:
                status_callback(f"Selesai OCR 1:1 dalam {elapsed} detik!", 1.0)
            return True, output_docx_path

        else:
            check_cancel(status_callback, "Membuka file PDF...", 0.1)

            cv = Converter(pdf_path)
            total_pages = len(cv.pages)
            pages_to_convert = parse_pages(pages_str, total_pages)
            
            check_cancel(status_callback, f"Memulai konversi PDF standard ({total_pages} halaman)...", 0.3)

            start_time = time.time()
            if pages_to_convert is not None:
                cv.convert(output_docx_path, pages=pages_to_convert)
            else:
                cv.convert(output_docx_path, start=0, end=None)

            cv.close()

            elapsed = round(time.time() - start_time, 2)
            
            if status_callback:
                status_callback(f"Selesai dalam {elapsed} detik!", 1.0)
                
            return True, output_docx_path

    except CancellationException as ce:
        return False, str(ce)
    except Exception as e:
        return False, f"Terjadi kesalahan me-render Word: {str(e)}"

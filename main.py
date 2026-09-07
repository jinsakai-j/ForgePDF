import sys
import os
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Try importing customtkinter for modern UI, fallback to standard tkinter if missing
try:
    import customtkinter as ctk
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    HAS_CTK = True
except ImportError:
    HAS_CTK = False

from converter import convert_pdf_to_word, compress_pdf, add_smart_page_numbers

class PDFToWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ForgePDF - Smart PDF Converter, Compressor & Page Numbering")
        self.root.geometry("780x700")
        self.root.minsize(700, 620)

        self.selected_files = []
        self.output_dir = ""
        self.is_processing = False
        self.cancel_requested = False
        self.ocr_selected_files = []

        if HAS_CTK:
            self._setup_ctk_ui()
        else:
            self._setup_tk_ui()

    def _setup_ctk_ui(self):
        # Header
        self.header_frame = ctk.CTkFrame(self.root, corner_radius=10, fg_color="#1E293B")
        self.header_frame.pack(fill="x", padx=20, pady=(20, 10))

        self.title_label = ctk.CTkLabel(
            self.header_frame, 
            text="📄 ForgePDF", 
            font=("Segoe UI", 24, "bold"),
            text_color="#38BDF8"
        )
        self.title_label.pack(anchor="w", padx=15, pady=(15, 2))

        self.subtitle_label = ctk.CTkLabel(
            self.header_frame, 
            text="Konversi PDF ke Word, Kompres Ukuran, & Penomoran Halaman Pintar (100% Offline & Aman)", 
            font=("Segoe UI", 12),
            text_color="#94A3B8"
        )
        self.subtitle_label.pack(anchor="w", padx=15, pady=(0, 15))

        # Main Body TabView
        self.tabview = ctk.CTkTabview(self.root, corner_radius=10)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=(0, 10))

        self.tab_convert = self.tabview.add("📄 PDF ke Word")
        self.tab_compress = self.tabview.add("🗜️ Kompres PDF")
        self.tab_numbering = self.tabview.add("🔢 Penomoran Halaman")
        self.tab_ocr = self.tabview.add("🔤 OCR Gambar")

        self._setup_convert_tab()
        self._setup_compress_tab()
        self._setup_numbering_tab()
        self._setup_ocr_tab()

        # Shared File Selection & Output Area at bottom
        self.shared_frame = ctk.CTkFrame(self.root, corner_radius=10)
        self.shared_frame.pack(fill="x", padx=20, pady=(0, 15))

        self.file_btn_frame = ctk.CTkFrame(self.shared_frame, fg_color="transparent")
        self.file_btn_frame.pack(fill="x", padx=15, pady=(10, 5))

        self.select_btn = ctk.CTkButton(
            self.file_btn_frame,
            text="📁 Pilih File PDF",
            font=("Segoe UI", 13, "bold"),
            command=self.browse_files,
            height=36,
            fg_color="#2563EB",
            hover_color="#1D4ED8"
        )
        self.select_btn.pack(side="left", padx=(0, 10))

        self.clear_btn = ctk.CTkButton(
            self.file_btn_frame,
            text="🗑️ Hapus Pilihan",
            font=("Segoe UI", 12),
            command=self.clear_files,
            height=36,
            fg_color="#EF4444",
            hover_color="#DC2626",
            width=120
        )
        self.clear_btn.pack(side="left")

        # Output Folder Picker
        self.out_dir_entry = ctk.CTkEntry(self.file_btn_frame, placeholder_text="Folder Output: Sama dengan folder asal PDF", width=280)
        self.out_dir_entry.pack(side="left", padx=10)

        self.browse_out_btn = ctk.CTkButton(
            self.file_btn_frame, text="Browse...", width=80, height=36, command=self.browse_output_dir
        )
        self.browse_out_btn.pack(side="left")

        # Files Display Box
        self.file_display = ctk.CTkTextbox(self.shared_frame, height=55, font=("Consolas", 10))
        self.file_display.pack(fill="x", padx=15, pady=5)
        self.file_display.insert("1.0", "Belum ada file PDF yang dipilih. Klik 'Pilih File PDF' untuk mulai.")
        self.file_display.configure(state="disabled")

        # Progress & Stop Control Section
        self.progress_frame = ctk.CTkFrame(self.shared_frame, fg_color="transparent")
        self.progress_frame.pack(fill="x", padx=15, pady=(5, 10))

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.progress_bar.set(0.0)

        self.cancel_btn = ctk.CTkButton(
            self.progress_frame,
            text="⛔ Batal / Stop",
            font=("Segoe UI", 12, "bold"),
            command=self.request_cancel,
            height=32,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            width=110,
            state="disabled"
        )
        self.cancel_btn.pack(side="right")

        self.status_label = ctk.CTkLabel(self.shared_frame, text="ForgePDF Siap Memproses Dokumen", font=("Segoe UI", 11), text_color="#38BDF8")
        self.status_label.pack(anchor="w", padx=15, pady=(0, 10))

    def _setup_convert_tab(self):
        frame = self.tab_convert
        
        ctk.CTkLabel(frame, text="Pengaturan Konversi PDF ke Word (.docx)", font=("Segoe UI", 13, "bold"), text_color="#F8FAFC").pack(anchor="w", padx=10, pady=(10, 5))
        
        opts = ctk.CTkFrame(frame, fg_color="transparent")
        opts.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(opts, text="Halaman Spesifik:", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.pages_entry = ctk.CTkEntry(opts, placeholder_text="Kosongkan = semua (cth: 1-5, 8)", width=350)
        self.pages_entry.grid(row=0, column=1, sticky="w", padx=10, pady=5)

        self.ocr_switch = ctk.CTkCheckBox(
            frame,
            text="🔍 Mode OCR 1:1 Presisi (Gunakan jika PDF berupa Scan Kamera/Foto)",
            font=("Segoe UI", 12, "bold"),
            text_color="#F59E0B",
            onvalue=True,
            offvalue=False
        )
        self.ocr_switch.pack(anchor="w", padx=10, pady=10)

        self.convert_btn = ctk.CTkButton(
            frame,
            text="🚀 Konversi Sekarang ke Word (.docx)",
            font=("Segoe UI", 14, "bold"),
            command=lambda: self.start_action_thread("convert"),
            height=42,
            fg_color="#10B981",
            hover_color="#059669"
        )
        self.convert_btn.pack(fill="x", padx=10, pady=10)

    def _setup_compress_tab(self):
        frame = self.tab_compress

        ctk.CTkLabel(frame, text="Pengaturan Target Ukuran File Kompresi PDF", font=("Segoe UI", 13, "bold"), text_color="#F8FAFC").pack(anchor="w", padx=10, pady=(10, 5))

        opts = ctk.CTkFrame(frame, fg_color="transparent")
        opts.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(opts, text="Target Ukuran File:", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.compress_quality_menu = ctk.CTkOptionMenu(
            opts,
            values=[
                "⚡ Sangat Kecil (Target ~1 - 2 MB)",
                "⚖️ Sedang (Target ~3 - 5 MB - Rekomendasi)",
                "🎨 Tinggi / Kualitas Maksimal (Target ~6 - 10 MB)"
            ],
            width=420
        )
        self.compress_quality_menu.set("⚖️ Sedang (Target ~3 - 5 MB - Rekomendasi)")
        self.compress_quality_menu.grid(row=0, column=1, sticky="w", padx=10, pady=5)

        self.compress_btn = ctk.CTkButton(
            frame,
            text="🗜️ Kompres Ukuran PDF Sekarang",
            font=("Segoe UI", 14, "bold"),
            command=lambda: self.start_action_thread("compress"),
            height=42,
            fg_color="#8B5CF6",
            hover_color="#7C3AED"
        )
        self.compress_btn.pack(fill="x", padx=10, pady=15)

    def _setup_numbering_tab(self):
        frame = self.tab_numbering

        ctk.CTkLabel(frame, text="Pengaturan Penomoran Halaman Otomatis (Smart Numbering)", font=("Segoe UI", 13, "bold"), text_color="#F8FAFC").pack(anchor="w", padx=10, pady=(10, 5))

        opts = ctk.CTkFrame(frame, fg_color="transparent")
        opts.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(opts, text="Mode Penomoran:", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        self.num_mode_menu = ctk.CTkOptionMenu(
            opts,
            values=[
                "🤖 Otomatis (Romawi sebelum BAB I, Angka 1 mulai BAB I)",
                "🔢 Angka Biasa Semua (1, 2, 3...)",
                "🔤 Angka Romawi Semua (i, ii, iii...)"
            ],
            width=420
        )
        self.num_mode_menu.grid(row=0, column=1, sticky="w", padx=10, pady=5)

        ctk.CTkLabel(opts, text="Posisi Horisontal:", font=("Segoe UI", 12, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        self.num_horiz_menu = ctk.CTkOptionMenu(
            opts,
            values=["Tengah (Center)", "Kanan (Right)", "Kiri (Left)"],
            width=200
        )
        self.num_horiz_menu.grid(row=1, column=1, sticky="w", padx=10, pady=5)

        ctk.CTkLabel(opts, text="Posisi Vertikal:", font=("Segoe UI", 12, "bold")).grid(row=2, column=0, sticky="w", pady=5)
        self.num_vert_menu = ctk.CTkOptionMenu(
            opts,
            values=["Bawah (Footer)", "Atas (Header)"],
            width=200
        )
        self.num_vert_menu.grid(row=2, column=1, sticky="w", padx=10, pady=5)

        self.skip_cover_switch = ctk.CTkCheckBox(
            frame,
            text=" Sembunyikan Nomor Halaman di Cover (Halaman 1)",
            font=("Segoe UI", 12, "bold"),
            text_color="#38BDF8",
            onvalue=True,
            offvalue=False
        )
        self.skip_cover_switch.select()
        self.skip_cover_switch.pack(anchor="w", padx=10, pady=5)

        self.number_btn = ctk.CTkButton(
            frame,
            text="🔢 Tambahkan Nomor Halaman ke PDF",
            font=("Segoe UI", 14, "bold"),
            command=lambda: self.start_action_thread("numbering"),
            height=42,
            fg_color="#0284C7",
            hover_color="#0369A1"
        )
        self.number_btn.pack(fill="x", padx=10, pady=10)

    def _setup_ocr_tab(self):
        frame = self.tab_ocr

        ctk.CTkLabel(frame, text="🔤 Ekstrak Teks dari Gambar (OCR)", font=("Segoe UI", 13, "bold"),
                     text_color="#F8FAFC").pack(anchor="w", padx=10, pady=(10, 2))
        ctk.CTkLabel(frame, text="Buat foto/scan/screenshot yang berisi teks — diproses 100% lokal (RapidOCR/onnx). "
                                "Butuh: pip install rapidocr_onnxruntime",
                     font=("Segoe UI", 10), text_color="#94A3B8").pack(anchor="w", padx=10, pady=(0, 8))

        btn_row = ctk.CTkFrame(frame, fg_color="transparent")
        btn_row.pack(fill="x", padx=10, pady=2)
        self.ocr_pick_btn = ctk.CTkButton(btn_row, text="🖼️ Pilih Gambar", font=("Segoe UI", 12, "bold"),
                                          command=self._pick_ocr_images, height=34,
                                          fg_color="#16A34A", hover_color="#15803D")
        self.ocr_pick_btn.pack(side="left")
        self.ocr_clear_btn = ctk.CTkButton(btn_row, text="🗑️ Hapus", font=("Segoe UI", 11),
                                           command=self._clear_ocr_files, height=34, width=90,
                                           fg_color="#EF4444", hover_color="#DC2626")
        self.ocr_clear_btn.pack(side="left", padx=(8, 0))

        self.ocr_files_display = ctk.CTkTextbox(frame, height=62, font=("Consolas", 10))
        self.ocr_files_display.pack(fill="x", padx=10, pady=6)
        self._set_ocr_display("Belum ada gambar dipilih.")

        self.ocr_btn = ctk.CTkButton(frame, text="🔍 Ekstrak Teks Sekarang", font=("Segoe UI", 14, "bold"),
                                     command=self.start_ocr_extract, height=40,
                                     fg_color="#10B981", hover_color="#059669")
        self.ocr_btn.pack(fill="x", padx=10, pady=6)

        result_bar = ctk.CTkFrame(frame, fg_color="transparent")
        result_bar.pack(fill="x", padx=10, pady=(4, 0))
        ctk.CTkLabel(result_bar, text="Hasil Teks:", font=("Segoe UI", 11, "bold"),
                     text_color="#F8FAFC").pack(side="left")
        ctk.CTkButton(result_bar, text="💾 Simpan .txt", width=100, height=28, font=("Segoe UI", 10),
                      fg_color="#3B82F6", hover_color="#2563EB", command=self._save_ocr_text).pack(side="right")
        ctk.CTkButton(result_bar, text="📋 Salin", width=80, height=28, font=("Segoe UI", 10),
                      command=self._copy_ocr_text).pack(side="right", padx=(6, 0))

        self.ocr_result = ctk.CTkTextbox(frame, font=("Consolas", 10))
        self.ocr_result.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.ocr_result.insert("1.0", "Hasil ekstraksi teks akan muncul di sini...")

    def _set_ocr_display(self, text):
        self.ocr_files_display.configure(state="normal")
        self.ocr_files_display.delete("1.0", tk.END)
        self.ocr_files_display.insert("1.0", text)
        self.ocr_files_display.configure(state="disabled")

    def _pick_ocr_images(self):
        files = filedialog.askopenfilenames(
            title="Pilih Gambar (foto/scan/screenshot)",
            filetypes=[("Gambar", "*.png *.jpg *.jpeg *.webp *.bmp"), ("Semua File", "*.*")]
        )
        if not files:
            return
        self.ocr_selected_files = list(files)
        lines = "\n".join(f"{i+1}. {os.path.basename(f)}" for i, f in enumerate(files))
        self._set_ocr_display(f"Terpilih {len(files)} gambar:\n{lines}")

    def _clear_ocr_files(self):
        self.ocr_selected_files = []
        self._set_ocr_display("Belum ada gambar dipilih.")

    def start_ocr_extract(self):
        if self.is_processing:
            return
        if not self.ocr_selected_files:
            messagebox.showwarning("Peringatan", "Silakan pilih setidaknya 1 gambar terlebih dahulu!")
            return
        self.cancel_requested = False
        self.set_buttons_state(True)
        threading.Thread(target=self._run_ocr_extract, daemon=True).start()

    def _run_ocr_extract(self):
        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError:
            self.root.after(0, lambda: (self.set_buttons_state(False),
                                        messagebox.showerror(
                                            "OCR", "RapidOCR belum terpasang.\n\nJalankan:\npip install rapidocr_onnxruntime\nLalu buka aplikasi lagi.")))
            return
        try:
            engine = RapidOCR()
        except Exception as e:
            self.root.after(0, lambda: (self.set_buttons_state(False),
                                        messagebox.showerror("OCR", f"Gagal memuat engine OCR: {e}")))
            return

        total = len(self.ocr_selected_files)
        chunks = []
        for i, f in enumerate(self.ocr_selected_files):
            if self.cancel_requested:
                break
            name = os.path.basename(f)
            self.update_status(f"[{i+1}/{total}] OCR {name}...", i / total)
            try:
                res, _ = engine(f)
                body = self._ocr_texts_to_readable(res)
                if body:
                    chunks.append(f"[{name}]\n{body}")
            except Exception:
                pass

        result_text = "\n\n".join(chunks).strip()

        def _finish():
            self.set_buttons_state(False)
            if self.cancel_requested:
                self.update_status("⛔ OCR dibatalkan oleh pengguna.", 0.0)
                return
            self.ocr_result.configure(state="normal")
            self.ocr_result.delete("1.0", tk.END)
            self.ocr_result.insert("1.0", result_text or "Tidak ada teks yang terdeteksi.")
            self.ocr_result.configure(state="disabled")
            self.update_status(f"OCR selesai: {total} gambar diproses.", 1.0)

        self.root.after(0, _finish)

    @staticmethod
    def _ocr_texts_to_readable(res):
        if not res:
            return ""
        items = []
        for r in res:
            if isinstance(r, dict):
                box = r.get("box") or r.get("boxes") or []
                text = str(r.get("text") or r.get("txt") or "").strip()
            else:
                text = str(r[1]).strip() if len(r) > 1 else ""
                box = r[0] if r and isinstance(r[0], list) else []
            if text:
                items.append((box, text))
        if not items:
            return ""
        items.sort(key=lambda it: (it[0][0][1], it[0][0][0]) if it[0] else (0, 0))
        return "\n".join(t for _, t in items)

    def _copy_ocr_text(self):
        try:
            text = self.ocr_result.get("1.0", "end").strip()
        except Exception:
            return
        if not text:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.update_status("Teks hasil OCR disalin ke clipboard.", 1.0)

    def _save_ocr_text(self):
        try:
            text = self.ocr_result.get("1.0", "end").strip()
        except Exception:
            return
        if not text:
            messagebox.showwarning("OCR", "Belum ada hasil OCR untuk disimpan.")
            return
        path = filedialog.asksaveasfilename(
            title="Simpan hasil OCR", defaultextension=".txt",
            filetypes=[("Teks", "*.txt"), ("Semua File", "*.*")],
            initialfile="hasil_ocr.txt")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(text)
            self.update_status(f"Hasil OCR disimpan: {path}", 1.0)
        except Exception as e:
            messagebox.showerror("OCR", f"Gagal menyimpan: {e}")

    def _tk_fallback_ocr(self):
        files = filedialog.askopenfilenames(
            title="Pilih Gambar (foto/scan/screenshot)",
            filetypes=[("Gambar", "*.png *.jpg *.jpeg *.webp *.bmp"), ("Semua File", "*.*")]
        )
        if not files:
            return
        try:
            from rapidocr_onnxruntime import RapidOCR
        except ImportError:
            messagebox.showerror("OCR", "RapidOCR belum terpasang.\n\nJalankan:\npip install rapidocr_onnxruntime")
            return
        try:
            engine = RapidOCR()
        except Exception as e:
            messagebox.showerror("OCR", f"Gagal memuat engine OCR: {e}")
            return

        self.is_processing = True
        self.cancel_btn.config(state="normal")
        saved = []
        for i, f in enumerate(files):
            if self.cancel_requested:
                break
            self.status_label.config(text=f"[{i+1}/{len(files)}] OCR {os.path.basename(f)}...")
            self.progress_bar['value'] = i / len(files) * 100
            try:
                res, _ = engine(f)
                body = self._ocr_texts_to_readable(res)
                if body:
                    base = os.path.splitext(os.path.basename(f))[0]
                    out_path = os.path.join(os.path.dirname(f), base + "_ocr.txt")
                    n = 2
                    while os.path.exists(out_path):
                        out_path = os.path.join(os.path.dirname(f), f"{base}_ocr_{n}.txt")
                        n += 1
                    with open(out_path, "w", encoding="utf-8") as fh:
                        fh.write(body)
                    saved.append(f"{os.path.basename(f)} -> {os.path.basename(out_path)}")
            except Exception:
                pass

        self.is_processing = False
        self.cancel_btn.config(state="disabled")
        self.status_label.config(text="OCR selesai." if saved else "Tidak ada teks terdeteksi.")
        if saved:
            messagebox.showinfo("Hasil OCR", "Teks sudah disimpan sebagai .txt:\n\n" + "\n".join(saved))
        else:
            messagebox.showinfo("Hasil OCR", "Tidak ada teks yang terdeteksi dari gambar terpilih.")

    def _setup_tk_ui(self):
        # Fallback Tkinter UI
        self.root.configure(bg="#121826")
        
        header = tk.Frame(self.root, bg="#1E293B", padx=15, pady=15)
        header.pack(fill="x", padx=15, pady=15)
        
        title = tk.Label(header, text="📄 ForgePDF", font=("Segoe UI", 18, "bold"), fg="#38BDF8", bg="#1E293B")
        title.pack(anchor="w")
        sub = tk.Label(header, text="Konversi PDF ke Word, Kompres, & Penomoran Halaman (100% Offline)", font=("Segoe UI", 10), fg="#94A3B8", bg="#1E293B")
        sub.pack(anchor="w")

        body = tk.Frame(self.root, bg="#1E293B", padx=15, pady=15)
        body.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        btn_box = tk.Frame(body, bg="#1E293B")
        btn_box.pack(fill="x", pady=(0, 10))

        self.select_btn = tk.Button(btn_box, text="📁 Pilih File PDF", font=("Segoe UI", 11, "bold"), bg="#2563EB", fg="white", command=self.browse_files, relief="flat", padx=15, pady=5)
        self.select_btn.pack(side="left", padx=(0, 10))

        self.clear_btn = tk.Button(btn_box, text="🗑️ Hapus Pilihan", font=("Segoe UI", 10), bg="#EF4444", fg="white", command=self.clear_files, relief="flat", padx=10, pady=5)
        self.clear_btn.pack(side="left")

        self.file_display = tk.Text(body, height=4, font=("Consolas", 10), bg="#0F172A", fg="#E2E8F0", relief="flat", padx=10, pady=10)
        self.file_display.pack(fill="x", pady=5)
        self.file_display.insert("1.0", "Belum ada file PDF yang dipilih.\nKlik 'Pilih File PDF' untuk mulai.")
        self.file_display.configure(state="disabled")

        self.progress_bar = ttk.Progressbar(body, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill="x", pady=(10, 5))

        self.status_label = tk.Label(body, text="ForgePDF Siap Memproses Dokumen", font=("Segoe UI", 10), fg="#38BDF8", bg="#1E293B")
        self.status_label.pack(anchor="w")

        actions = tk.Frame(self.root, bg="#121826")
        actions.pack(fill="x", padx=15, pady=(0, 15))

        self.convert_btn = tk.Button(actions, text="🚀 Konversi Word", font=("Segoe UI", 10, "bold"), bg="#10B981", fg="white", command=lambda: self.start_action_thread("convert"), relief="flat", pady=8)
        self.convert_btn.pack(side="left", expand=True, fill="x", padx=2)

        self.compress_btn = tk.Button(actions, text="🗜️ Kompres PDF", font=("Segoe UI", 10, "bold"), bg="#8B5CF6", fg="white", command=lambda: self.start_action_thread("compress"), relief="flat", pady=8)
        self.compress_btn.pack(side="left", expand=True, fill="x", padx=2)

        self.number_btn = tk.Button(actions, text="🔢 Nomor Halaman", font=("Segoe UI", 10, "bold"), bg="#0284C7", fg="white", command=lambda: self.start_action_thread("numbering"), relief="flat", pady=8)
        self.number_btn.pack(side="left", expand=True, fill="x", padx=2)

        self.ocr_btn_tk = tk.Button(actions, text="🔤 OCR Gambar", font=("Segoe UI", 10, "bold"), bg="#16A34A", fg="white", command=self._tk_fallback_ocr, relief="flat", pady=8)
        self.ocr_btn_tk.pack(side="left", expand=True, fill="x", padx=2)

        self.cancel_btn = tk.Button(actions, text="⛔ Batal", font=("Segoe UI", 10, "bold"), bg="#DC2626", fg="white", command=self.request_cancel, relief="flat", pady=8, state="disabled")
        self.cancel_btn.pack(side="left", padx=2)

        self.open_folder_btn = tk.Button(actions, text="📂 Buka Folder", font=("Segoe UI", 10), bg="#475569", fg="white", command=self.open_output_folder, relief="flat", pady=8)
        self.open_folder_btn.pack(side="right", padx=2)

    def browse_files(self):
        files = filedialog.askopenfilenames(
            title="Pilih File PDF",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if files:
            self.selected_files = list(files)
            self._update_file_display()

    def browse_output_dir(self):
        folder = filedialog.askdirectory(title="Pilih Folder Tujuan Penyimpanan")
        if folder:
            self.output_dir = folder
            self.out_dir_entry.delete(0, tk.END)
            self.out_dir_entry.insert(0, folder)

    def clear_files(self):
        self.selected_files = []
        self._update_file_display()

    def _update_file_display(self):
        self.file_display.configure(state="normal")
        self.file_display.delete("1.0", tk.END)
        if not self.selected_files:
            self.file_display.insert("1.0", "Belum ada file PDF yang dipilih.\nKlik 'Pilih File PDF' untuk mulai.")
        else:
            text = f"Terpilih {len(self.selected_files)} file:\n"
            for idx, fpath in enumerate(self.selected_files, 1):
                size_mb = os.path.getsize(fpath) / (1024 * 1024)
                text += f"{idx}. {os.path.basename(fpath)} ({size_mb:.2f} MB)\n"
            self.file_display.insert("1.0", text)
        self.file_display.configure(state="disabled")

    def update_status(self, msg, progress_val=None):
        def _update():
            if HAS_CTK:
                self.status_label.configure(text=msg)
                if progress_val is not None:
                    self.progress_bar.set(progress_val)
            else:
                self.status_label.config(text=msg)
                if progress_val is not None:
                    self.progress_bar['value'] = progress_val * 100
        self.root.after(0, _update)

    def set_buttons_state(self, processing):
        self.is_processing = processing
        state = "disabled" if processing else "normal"
        cancel_state = "normal" if processing else "disabled"

        if HAS_CTK:
            self.convert_btn.configure(state=state)
            self.compress_btn.configure(state=state)
            self.number_btn.configure(state=state)
            self.ocr_btn.configure(state=state)
            self.ocr_pick_btn.configure(state=state)
            self.ocr_clear_btn.configure(state=state)
            self.cancel_btn.configure(state=cancel_state)
        else:
            self.convert_btn.config(state=state)
            self.compress_btn.config(state=state)
            self.number_btn.config(state=state)
            self.ocr_btn_tk.config(state=state)
            self.cancel_btn.config(state=cancel_state)

    def request_cancel(self):
        if self.is_processing:
            self.cancel_requested = True
            self.update_status("⚠️ Menghentikan dan membatalkan proses...")

    def check_if_cancelled(self, msg="", prog=None):
        if self.cancel_requested:
            return False
        return True

    def start_action_thread(self, action_type):
        if not self.selected_files:
            messagebox.showwarning("Peringatan", "Silakan pilih setidaknya 1 file PDF terlebih dahulu!")
            return

        if self.is_processing:
            return

        self.cancel_requested = False
        self.set_buttons_state(True)
        
        if action_type == "convert":
            use_ocr = bool(self.ocr_switch.get()) if HAS_CTK else False
            threading.Thread(target=self._run_conversion, args=(use_ocr,), daemon=True).start()
        elif action_type == "compress":
            q_str = "medium"
            if HAS_CTK:
                sel_val = self.compress_quality_menu.get()
                if "Sangat Kecil" in sel_val:
                    q_str = "ultra_small"
                elif "Tinggi" in sel_val:
                    q_str = "high"
                else:
                    q_str = "medium"
            threading.Thread(target=self._run_compression, args=(q_str,), daemon=True).start()
        elif action_type == "numbering":
            threading.Thread(target=self._run_numbering, daemon=True).start()

    def make_status_cb(self, index, total, file_name):
        def _cb(msg, step_val):
            if self.cancel_requested:
                return False
            prog = (index / total) + (step_val / total)
            self.update_status(f"[{index+1}/{total}] {file_name}: {msg}", prog)
            return True
        return _cb

    def _run_conversion(self, use_ocr):
        pages_str = self.pages_entry.get().strip() if HAS_CTK else ""
        custom_out_dir = self.out_dir_entry.get().strip()

        total = len(self.selected_files)
        success_count = 0
        failed_files = []

        for i, pdf_path in enumerate(self.selected_files):
            if self.cancel_requested:
                break

            file_name = os.path.basename(pdf_path)
            if custom_out_dir and os.path.exists(custom_out_dir):
                out_path = os.path.join(custom_out_dir, os.path.splitext(file_name)[0] + ".docx")
            else:
                out_path = os.path.splitext(pdf_path)[0] + ".docx"

            mode_str = " (OCR)" if use_ocr else ""
            self.update_status(f"[{i+1}/{total}] Mengonversi {file_name}{mode_str}...", (i / total))

            status_cb = self.make_status_cb(i, total, file_name)
            success, res = convert_pdf_to_word(pdf_path, out_path, pages_str=pages_str, use_ocr=use_ocr, status_callback=status_cb)

            if self.cancel_requested:
                break

            if success:
                success_count += 1
            else:
                failed_files.append((file_name, res))

        def _finish_ui():
            self.set_buttons_state(False)
            if self.cancel_requested:
                self.update_status("⛔ Proses berhasil dibatalkan oleh pengguna", 0.0)
                messagebox.showinfo("Dibatalkan", "Proses telah dibatalkan oleh pengguna.")
            else:
                self.update_status(f"Konversi selesai! Berhasil: {success_count}/{total}", 1.0)
                if success_count == total:
                    messagebox.showinfo("Sukses!", f"Seluruh {total} file PDF berhasil dikonversi ke Word (.docx)!")
                else:
                    msg = f"{success_count} file berhasil, {len(failed_files)} gagal.\n\nDetail:\n"
                    for fname, err in failed_files:
                        msg += f"- {fname}: {err}\n"
                    messagebox.showwarning("Hasil Konversi", msg)

        self.root.after(0, _finish_ui)

    def _run_compression(self, quality_mode):
        custom_out_dir = self.out_dir_entry.get().strip()
        total = len(self.selected_files)
        success_count = 0
        summary_results = []

        for i, pdf_path in enumerate(self.selected_files):
            if self.cancel_requested:
                break

            file_name = os.path.basename(pdf_path)
            if custom_out_dir and os.path.exists(custom_out_dir):
                out_path = os.path.join(custom_out_dir, os.path.splitext(file_name)[0] + "_compressed.pdf")
            else:
                out_path = os.path.splitext(pdf_path)[0] + "_compressed.pdf"

            self.update_status(f"[{i+1}/{total}] Mengompres {file_name}...", (i / total))

            status_cb = self.make_status_cb(i, total, file_name)
            success, res, orig_sz, new_sz = compress_pdf(pdf_path, out_path, quality=quality_mode, status_callback=status_cb)

            if self.cancel_requested:
                break

            if success:
                success_count += 1
                orig_mb = orig_sz / (1024 * 1024)
                new_mb = new_sz / (1024 * 1024)
                pct = round((1 - (new_sz / orig_sz)) * 100, 1) if orig_sz > 0 else 0
                summary_results.append(f"• {file_name}:\n  {orig_mb:.2f} MB ➡️ {new_mb:.2f} MB (Hemat {pct}%)")
            else:
                summary_results.append(f"• {file_name}: Gagal ({res})")

        def _finish_ui():
            self.set_buttons_state(False)
            if self.cancel_requested:
                self.update_status("⛔ Proses kompresi berhasil dibatalkan", 0.0)
                messagebox.showinfo("Dibatalkan", "Proses kompresi telah dibatalkan.")
            else:
                self.update_status(f"Kompresi selesai! Berhasil: {success_count}/{total}", 1.0)
                msg = f"Kompresi Selesai ({success_count}/{total} file berhasil):\n\n" + "\n".join(summary_results)
                messagebox.showinfo("Hasil Kompresi PDF", msg)

        self.root.after(0, _finish_ui)

    def _run_numbering(self):
        custom_out_dir = self.out_dir_entry.get().strip()
        total = len(self.selected_files)
        success_count = 0
        failed_files = []

        mode_val = "auto"
        horiz_val = "tengah"
        vert_val = "bawah"
        skip_cover = True

        if HAS_CTK:
            mode_str = self.num_mode_menu.get()
            if "Angka Biasa" in mode_str:
                mode_val = "latin_all"
            elif "Angka Romawi" in mode_str:
                mode_val = "roman_all"

            horiz_str = self.num_horiz_menu.get()
            if "Kiri" in horiz_str:
                horiz_val = "kiri"
            elif "Kanan" in horiz_str:
                horiz_val = "kanan"

            vert_str = self.num_vert_menu.get()
            if "Atas" in vert_str:
                vert_val = "atas"

            skip_cover = bool(self.skip_cover_switch.get())

        for i, pdf_path in enumerate(self.selected_files):
            if self.cancel_requested:
                break

            file_name = os.path.basename(pdf_path)
            if custom_out_dir and os.path.exists(custom_out_dir):
                out_path = os.path.join(custom_out_dir, os.path.splitext(file_name)[0] + "_numbered.pdf")
            else:
                out_path = os.path.splitext(pdf_path)[0] + "_numbered.pdf"

            self.update_status(f"[{i+1}/{total}] Penomoran {file_name}...", (i / total))

            status_cb = self.make_status_cb(i, total, file_name)
            success, res = add_smart_page_numbers(
                pdf_path, 
                out_path, 
                mode=mode_val, 
                pos_vert=vert_val, 
                pos_horiz=horiz_val, 
                skip_cover=skip_cover, 
                status_callback=status_cb
            )

            if self.cancel_requested:
                break

            if success:
                success_count += 1
            else:
                failed_files.append((file_name, res))

        def _finish_ui():
            self.set_buttons_state(False)
            if self.cancel_requested:
                self.update_status("⛔ Proses penomoran berhasil dibatalkan", 0.0)
                messagebox.showinfo("Dibatalkan", "Proses penomoran telah dibatalkan.")
            else:
                self.update_status(f"Penomoran selesai! Berhasil: {success_count}/{total}", 1.0)
                if success_count == total:
                    messagebox.showinfo("Sukses!", f"Penomoran halaman berhasil diterapkan ke {total} file PDF!")
                else:
                    msg = f"{success_count} file berhasil, {len(failed_files)} gagal.\n\nDetail:\n"
                    for fname, err in failed_files:
                        msg += f"- {fname}: {err}\n"
                    messagebox.showwarning("Hasil Penomoran Halaman", msg)

        self.root.after(0, _finish_ui)

    def open_output_folder(self):
        folder = self.out_dir_entry.get().strip()
        if not folder or not os.path.exists(folder):
            if self.selected_files:
                folder = os.path.dirname(self.selected_files[0])
            else:
                folder = os.path.expanduser("~")
        
        if os.path.exists(folder):
            if sys.platform == "win32":
                os.startfile(folder)
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", folder])

if __name__ == "__main__":
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("JinsakaiCorp.ForgePDF.1")
    except Exception:
        pass

    if HAS_CTK:
        root = ctk.CTk()
    else:
        root = tk.Tk()

    _icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "forgepdf_icon.ico")
    try:
        root.iconbitmap(_icon_path)
    except Exception:
        pass

    app = PDFToWordApp(root)
    root.mainloop()

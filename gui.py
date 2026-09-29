import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from converter import word_to_pdf, pdf_to_word, batch_convert

# Set CustomTkinter default appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class FileConverterApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Document Converter - Word ⮂ PDF")
        self.geometry("820x620")
        self.minsize(750, 550)

        # Variables
        self.single_input_file = tk.StringVar()
        self.single_output_file = tk.StringVar()
        self.conversion_mode = tk.StringVar(value="auto")
        
        self.batch_input_dir = tk.StringVar()
        self.batch_output_dir = tk.StringVar()
        self.batch_mode = tk.StringVar(value="auto")

        self.is_processing = False

        self._build_ui()

    def _build_ui(self):
        # Header Frame
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=("gray85", "gray17"))
        header_frame.pack(fill="x", padx=0, pady=0)

        title_label = ctk.CTkLabel(
            header_frame, 
            text="📄 Word ⮂ PDF Converter", 
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=("gray10", "gray90")
        )
        title_label.pack(side="left", padx=20, pady=15)

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="High-Fidelity Document Type Converter",
            font=ctk.CTkFont(size=12, slant="italic"),
            text_color="gray60"
        )
        subtitle_label.pack(side="left", padx=0, pady=15)

        mode_switch = ctk.CTkOptionMenu(
            header_frame,
            values=["Dark Mode", "Light Mode", "System"],
            command=self._change_theme,
            width=120
        )
        mode_switch.pack(side="right", padx=20, pady=15)

        # Tabview for Single vs Batch mode
        self.tabview = ctk.CTkTabview(self, width=780, height=480)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=15)

        self.tab_single = self.tabview.add("Single File Conversion")
        self.tab_batch = self.tabview.add("Batch Folder Conversion")
        self.tab_info = self.tabview.add("Diagnostics & Info")

        self._build_single_tab()
        self._build_batch_tab()
        self._build_info_tab()

        # Footer Status Bar
        self.status_frame = ctk.CTkFrame(self, height=40, corner_radius=0)
        self.status_frame.pack(fill="x", side="bottom")

        self.status_label = ctk.CTkLabel(self.status_frame, text="Ready", anchor="w", font=ctk.CTkFont(size=12))
        self.status_label.pack(side="left", padx=15, pady=5)

        self.progress_bar = ctk.CTkProgressBar(self.status_frame, width=200)
        self.progress_bar.pack(side="right", padx=15, pady=5)
        self.progress_bar.set(0)

    def _build_single_tab(self):
        # Container
        container = ctk.CTkFrame(self.tab_single, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=15, pady=15)

        # Input File Selection
        lbl_in = ctk.CTkLabel(container, text="Input File (.docx, .doc, .pdf):", font=ctk.CTkFont(weight="bold"))
        lbl_in.pack(anchor="w", pady=(5, 2))

        in_frame = ctk.CTkFrame(container, fg_color="transparent")
        in_frame.pack(fill="x", pady=(0, 15))

        entry_in = ctk.CTkEntry(in_frame, textvariable=self.single_input_file, placeholder_text="Select a Word or PDF document...")
        entry_in.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_browse_in = ctk.CTkButton(in_frame, text="Browse...", width=100, command=self._browse_single_input)
        btn_browse_in.pack(side="right")

        # Conversion Direction Selection
        lbl_mode = ctk.CTkLabel(container, text="Conversion Direction:", font=ctk.CTkFont(weight="bold"))
        lbl_mode.pack(anchor="w", pady=(5, 2))

        mode_frame = ctk.CTkFrame(container, fg_color="transparent")
        mode_frame.pack(fill="x", pady=(0, 15))

        r1 = ctk.CTkRadioButton(mode_frame, text="Auto-Detect Extension", variable=self.conversion_mode, value="auto")
        r1.pack(side="left", padx=(0, 20))

        r2 = ctk.CTkRadioButton(mode_frame, text="Word ➔ PDF", variable=self.conversion_mode, value="word2pdf")
        r2.pack(side="left", padx=(0, 20))

        r3 = ctk.CTkRadioButton(mode_frame, text="PDF ➔ Word (.docx)", variable=self.conversion_mode, value="pdf2word")
        r3.pack(side="left")

        # Output File Selection
        lbl_out = ctk.CTkLabel(container, text="Output Path (Optional, defaults to same folder):", font=ctk.CTkFont(weight="bold"))
        lbl_out.pack(anchor="w", pady=(5, 2))

        out_frame = ctk.CTkFrame(container, fg_color="transparent")
        out_frame.pack(fill="x", pady=(0, 20))

        entry_out = ctk.CTkEntry(out_frame, textvariable=self.single_output_file, placeholder_text="Default output path...")
        entry_out.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_browse_out = ctk.CTkButton(out_frame, text="Save As...", width=100, command=self._browse_single_output)
        btn_browse_out.pack(side="right")

        # Big Convert Button
        self.btn_convert_single = ctk.CTkButton(
            container,
            text="🚀 Convert Document",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=45,
            command=self._start_single_conversion
        )
        self.btn_convert_single.pack(fill="x", pady=10)

        # Log Text Box
        self.log_textbox = ctk.CTkTextbox(container, height=140, font=ctk.CTkFont(family="Consolas", size=11))
        self.log_textbox.pack(fill="both", expand=True, pady=(10, 0))
        self.log_textbox.insert("end", "Welcome to Document Converter.\nSelect a file above and click 'Convert Document'.\n")

    def _build_batch_tab(self):
        container = ctk.CTkFrame(self.tab_batch, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=15, pady=15)

        # Input Directory
        lbl_in = ctk.CTkLabel(container, text="Input Directory:", font=ctk.CTkFont(weight="bold"))
        lbl_in.pack(anchor="w", pady=(5, 2))

        in_frame = ctk.CTkFrame(container, fg_color="transparent")
        in_frame.pack(fill="x", pady=(0, 10))

        entry_in = ctk.CTkEntry(in_frame, textvariable=self.batch_input_dir, placeholder_text="Select input directory containing documents...")
        entry_in.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_browse_in = ctk.CTkButton(in_frame, text="Choose Folder", width=110, command=self._browse_batch_input)
        btn_browse_in.pack(side="right")

        # Output Directory
        lbl_out = ctk.CTkLabel(container, text="Output Directory (Optional):", font=ctk.CTkFont(weight="bold"))
        lbl_out.pack(anchor="w", pady=(5, 2))

        out_frame = ctk.CTkFrame(container, fg_color="transparent")
        out_frame.pack(fill="x", pady=(0, 15))

        entry_out = ctk.CTkEntry(out_frame, textvariable=self.batch_output_dir, placeholder_text="Default: same as input directory")
        entry_out.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_browse_out = ctk.CTkButton(out_frame, text="Choose Folder", width=110, command=self._browse_batch_output)
        btn_browse_out.pack(side="right")

        # Batch Mode
        lbl_bmode = ctk.CTkLabel(container, text="Batch Conversion Mode:", font=ctk.CTkFont(weight="bold"))
        lbl_bmode.pack(anchor="w", pady=(5, 2))

        bmode_frame = ctk.CTkFrame(container, fg_color="transparent")
        bmode_frame.pack(fill="x", pady=(0, 15))

        rb1 = ctk.CTkRadioButton(bmode_frame, text="Convert All (.docx ➔ PDF and PDF ➔ .docx)", variable=self.batch_mode, value="auto")
        rb1.pack(side="left", padx=(0, 20))

        rb2 = ctk.CTkRadioButton(bmode_frame, text="Only Word ➔ PDF", variable=self.batch_mode, value="word2pdf")
        rb2.pack(side="left", padx=(0, 20))

        rb3 = ctk.CTkRadioButton(bmode_frame, text="Only PDF ➔ Word", variable=self.batch_mode, value="pdf2word")
        rb3.pack(side="left")

        # Start Batch Button
        self.btn_start_batch = ctk.CTkButton(
            container,
            text="⚡ Start Batch Conversion",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=40,
            command=self._start_batch_conversion
        )
        self.btn_start_batch.pack(fill="x", pady=10)

        # Batch Log Box
        self.batch_log_textbox = ctk.CTkTextbox(container, height=140, font=ctk.CTkFont(family="Consolas", size=11))
        self.batch_log_textbox.pack(fill="both", expand=True, pady=(5, 0))
        self.batch_log_textbox.insert("end", "Batch processing log will appear here...\n")

    def _build_info_tab(self):
        container = ctk.CTkFrame(self.tab_info, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=20, pady=20)

        info_title = ctk.CTkLabel(container, text="System & Backend Engine Status", font=ctk.CTkFont(size=16, weight="bold"))
        info_title.pack(anchor="w", pady=(0, 15))

        # Check installed engines
        backend_status = []
        
        # Word COM check
        try:
            import win32com.client
            backend_status.append("✅ Microsoft Word COM Automation: Available (High Quality DOCX ➔ PDF)")
        except Exception:
            backend_status.append("⚠️ Microsoft Word COM Automation: Not detected")

        # pdf2docx check
        try:
            import pdf2docx
            backend_status.append("✅ pdf2docx Engine: Installed (High Quality PDF ➔ DOCX)")
        except Exception:
            backend_status.append("❌ pdf2docx Engine: Missing")

        # reportlab check
        try:
            import reportlab
            backend_status.append("✅ ReportLab Engine: Installed (Fallback DOCX ➔ PDF)")
        except Exception:
            backend_status.append("❌ ReportLab Engine: Missing")

        status_box = ctk.CTkTextbox(container, height=220, font=ctk.CTkFont(size=12))
        status_box.pack(fill="x", pady=(0, 15))

        text_content = "\n".join(backend_status) + "\n\n"
        text_content += f"Python Version: {sys.version.split()[0]}\n"
        text_content += f"OS Platform: {sys.platform}\n\n"
        text_content += "Usage Tips:\n"
        text_content += "• Word to PDF uses MS Word COM automation on Windows for exact document layout reproduction.\n"
        text_content += "• If MS Word is not installed, it seamlessly uses the ReportLab fallback converter.\n"
        text_content += "• PDF to Word converts complex PDF layouts, tables, headers, and images into editable .docx format.\n"

        status_box.insert("end", text_content)
        status_box.configure(state="disabled")

    def _change_theme(self, new_mode):
        if new_mode == "Dark Mode":
            ctk.set_appearance_mode("Dark")
        elif new_mode == "Light Mode":
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("System")

    def _browse_single_input(self):
        filename = filedialog.askopenfilename(
            title="Select Document",
            filetypes=[("All Supported Files", "*.docx *.doc *.pdf"), ("Word Files", "*.docx *.doc"), ("PDF Files", "*.pdf")]
        )
        if filename:
            self.single_input_file.set(filename)
            self._log(f"Selected input file: {filename}")

    def _browse_single_output(self):
        filename = filedialog.asksaveasfilename(
            title="Save Output File",
            filetypes=[("PDF Document", "*.pdf"), ("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if filename:
            self.single_output_file.set(filename)

    def _browse_batch_input(self):
        folder = filedialog.askdirectory(title="Select Folder to Convert")
        if folder:
            self.batch_input_dir.set(folder)
            self._batch_log(f"Selected batch folder: {folder}")

    def _browse_batch_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.batch_output_dir.set(folder)

    def _log(self, message: str):
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")

    def _batch_log(self, message: str):
        self.batch_log_textbox.insert("end", message + "\n")
        self.batch_log_textbox.see("end")

    def _update_status(self, text: str, progress: float = 0.0):
        self.status_label.configure(text=text)
        self.progress_bar.set(progress)

    def _start_single_conversion(self):
        input_path = self.single_input_file.get().strip()
        if not input_path:
            messagebox.showwarning("Missing Input", "Please select an input file first.")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("File Not Found", f"The input file does not exist:\n{input_path}")
            return

        output_path = self.single_output_file.get().strip() or None
        mode = self.conversion_mode.get()

        self.btn_convert_single.configure(state="disabled")
        self._update_status("Converting document...", 0.3)
        self.progress_bar.start()

        def worker():
            try:
                ext = os.path.splitext(input_path)[1].lower()
                if mode == "word2pdf" or (mode == "auto" and ext in ['.docx', '.doc']):
                    self._log(f"⏳ Converting Word to PDF: {os.path.basename(input_path)}...")
                    res = word_to_pdf(input_path, output_path)
                    self._log(f"✅ Success! Created: {res}")
                    self.after(0, lambda: messagebox.showinfo("Conversion Complete", f"Successfully converted to PDF:\n{res}"))

                elif mode == "pdf2word" or (mode == "auto" and ext == '.pdf'):
                    self._log(f"⏳ Converting PDF to Word: {os.path.basename(input_path)}...")
                    res = pdf_to_word(input_path, output_path)
                    self._log(f"✅ Success! Created: {res}")
                    self.after(0, lambda: messagebox.showinfo("Conversion Complete", f"Successfully converted to Word:\n{res}"))

                else:
                    raise ValueError(f"Cannot auto-detect conversion type for extension: {ext}")

                self.after(0, lambda: self._update_status("Conversion Complete!", 1.0))
            except Exception as e:
                self._log(f"❌ Error: {e}")
                self.after(0, lambda: messagebox.showerror("Conversion Failed", str(e)))
                self.after(0, lambda: self._update_status("Conversion Failed", 0.0))
            finally:
                self.after(0, lambda: self.progress_bar.stop())
                self.after(0, lambda: self.btn_convert_single.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def _start_batch_conversion(self):
        input_dir = self.batch_input_dir.get().strip()
        if not input_dir:
            messagebox.showwarning("Missing Folder", "Please select an input directory.")
            return

        if not os.path.isdir(input_dir):
            messagebox.showerror("Invalid Folder", f"The directory does not exist:\n{input_dir}")
            return

        output_dir = self.batch_output_dir.get().strip() or None
        mode = self.batch_mode.get()

        self.btn_start_batch.configure(state="disabled")
        self._update_status("Starting batch conversion...", 0.1)

        def progress_cb(current, total, filename, success):
            ratio = current / max(total, 1)
            icon = "✅" if success else "❌"
            msg = f"[{current}/{total}] {icon} {filename}"
            self.after(0, lambda: self._batch_log(msg))
            self.after(0, lambda: self._update_status(f"Processing ({current}/{total})...", ratio))

        def worker():
            try:
                self.after(0, lambda: self._batch_log(f"🚀 Processing batch folder: {input_dir}..."))
                res = batch_convert(input_dir, direction=mode, output_dir=output_dir, progress_callback=progress_cb)

                summary = (
                    f"🎉 Batch Conversion Complete!\n\n"
                    f"Total Files Found: {res['total']}\n"
                    f"Successful: {res['successful_count']}\n"
                    f"Failed: {res['failed_count']}"
                )
                self.after(0, lambda: self._batch_log(f"\n{summary}"))
                self.after(0, lambda: messagebox.showinfo("Batch Finished", summary))
                self.after(0, lambda: self._update_status("Batch Completed", 1.0))
            except Exception as e:
                self.after(0, lambda: self._batch_log(f"❌ Batch Error: {e}"))
                self.after(0, lambda: messagebox.showerror("Batch Failed", str(e)))
                self.after(0, lambda: self._update_status("Batch Failed", 0.0))
            finally:
                self.after(0, lambda: self.btn_start_batch.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

def launch_gui():
    app = FileConverterApp()
    app.mainloop()

if __name__ == "__main__":
    launch_gui()

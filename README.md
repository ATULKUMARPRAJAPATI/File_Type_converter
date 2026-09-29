# 📄 Word ⮂ PDF Converter

A python document conversion tool for converting **Microsoft Word files (`.docx`, `.doc`) to PDF** and **PDF files to editable Word documents (`.docx`)**.

Includes a **GUI Application**, a **Command-Line Interface (CLI)**, and a **Python Module**.

---

## ✨ Features

- **Word ➔ PDF (`.docx`/`.doc` to `.pdf`)**:
  - Direct high-fidelity conversion using Microsoft Word COM engine (Windows).
  - Fallback conversion engine using `python-docx` + `ReportLab` if Microsoft Word is not installed.
- **PDF ➔ Word (`.pdf` to `.docx`)**:
  - Recreates document layout, paragraphs, headings, tables, and images accurately using `pdf2docx`.
- **Batch Processing**:
  - Convert entire directories of documents automatically.
- **Modern Graphical User Interface (GUI)**:
  - CustomTkinter UI featuring dark/light mode, file selectors, batch folder processing, live progress bars, and diagnostics.
- **Non-blocking Multithreaded Engine**:
  - UI remains fast and responsive while processing large files.

---

## 🚀 Installation & Requirements

1. **Clone or download this repository** into your project directory.
2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

*(Note: On Windows, for 100% pixel-perfect DOCX to PDF conversion, Microsoft Word is recommended. If MS Word is not present, the tool automatically uses the built-in ReportLab converter fallback).*

---

## 🖥️ Graphical User Interface (GUI)

To open the modern GUI application, run:

```bash
python main.py gui
```
or
```bash
python gui.py
```

### GUI Features:
1. **Single File Tab**: Select any `.docx`, `.doc`, or `.pdf` file, select output location, and click **Convert Document**.
2. **Batch Folder Tab**: Choose an input folder and output destination to convert multiple files at once.
3. **Diagnostics Tab**: Check installed conversion backend engines and system status.

---

## 💻 Command-Line Interface (CLI)

### 1. Convert Word (.docx / .doc) to PDF
```bash
python main.py word2pdf input_document.docx -o output_document.pdf
```

### 2. Convert PDF to Word (.docx)
```bash
python main.py pdf2word input_document.pdf -o output_document.docx
```

### 3. Batch Convert a Folder
```bash
python main.py batch ./my_folder -o ./converted_output --direction auto
```
Options for `--direction`:
- `auto`: Converts `.docx`/`.doc` ➔ PDF and `.pdf` ➔ `.docx`
- `word2pdf`: Converts only Word files to PDF
- `pdf2word`: Converts only PDF files to Word

---

## 🐍 Using as a Python Module

You can import the conversion functions directly into your Python scripts:

```python
from converter import word_to_pdf, pdf_to_word, batch_convert

# 1. Convert Word to PDF
pdf_path = word_to_pdf("contract.docx", "contract.pdf")
print(f"Saved PDF to: {pdf_path}")

# 2. Convert PDF to Word
docx_path = pdf_to_word("invoice.pdf", "invoice.docx")
print(f"Saved DOCX to: {docx_path}")

# 3. Batch Convert Folder
result = batch_convert("./documents", direction="auto", output_dir="./output")
print(f"Converted {result['successful_count']}/{result['total']} files.")
```

---

## 📁 Project Structure

```
File_Type_converter/
├── converter.py      # Core conversion engines & CLI interface
├── gui.py            # CustomTkinter Graphical User Interface
├── main.py           # Entry point script
├── requirements.txt  # Python package dependencies
└── README.md         # Project documentation
```

import os
import sys
import logging
import argparse
from pathlib import Path
from typing import Optional, List, Callable, Dict, Any

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("DocumentConverter")

def word_to_pdf(input_path: str, output_path: Optional[str] = None, fallback_reportlab: bool = True) -> str:
    """
    Converts a Word file (.docx or .doc) to PDF.
    
    Primary engine: docx2pdf (uses Microsoft Word COM on Windows)
    Fallback engine: python-docx + ReportLab (pure Python, cross-platform)
    """
    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")
    
    ext = os.path.splitext(input_path)[1].lower()
    if ext not in ['.docx', '.doc']:
        raise ValueError(f"Unsupported input format: {ext}. Expected .docx or .doc")

    if not output_path:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
    else:
        output_path = os.path.abspath(output_path)

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    # Strategy 1: docx2pdf / win32com (Best formatting output)
    try:
        logger.info(f"Converting '{input_path}' to PDF via docx2pdf...")
        import docx2pdf
        docx2pdf.convert(input_path, output_path)
        if os.path.exists(output_path):
            logger.info(f"Successfully converted: {output_path}")
            return output_path
    except Exception as e:
        logger.warning(f"docx2pdf conversion failed: {e}")

    # Strategy 2: Direct win32com COM automation (Windows specific)
    if sys.platform == "win32":
        try:
            logger.info("Attempting direct win32com COM automation...")
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(input_path)
            # 17 represents wdFormatPDF
            doc.SaveAs(output_path, FileFormat=17)
            doc.Close()
            word.Quit()
            if os.path.exists(output_path):
                logger.info(f"Successfully converted via win32com: {output_path}")
                return output_path
        except Exception as e:
            logger.warning(f"win32com conversion failed: {e}")

    # Strategy 3: ReportLab Fallback (Pure Python)
    if fallback_reportlab and ext == '.docx':
        logger.info("Falling back to pure Python ReportLab converter...")
        try:
            return _word_to_pdf_reportlab(input_path, output_path)
        except Exception as e:
            logger.error(f"ReportLab fallback failed: {e}")
            raise RuntimeError(f"All conversion methods failed for {input_path}") from e

    raise RuntimeError(f"Failed to convert '{input_path}' to PDF. Please ensure Microsoft Word is installed.")


def _word_to_pdf_reportlab(input_path: str, output_path: str) -> str:
    """Fallback converter using python-docx and ReportLab."""
    from docx import Document
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    doc = Document(input_path)
    pdf_doc = SimpleDocTemplate(output_path, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            story.append(Spacer(1, 8))
            continue
        
        style = styles['Normal']
        if p.style.name.startswith('Heading 1'):
            style = styles['Heading1']
        elif p.style.name.startswith('Heading 2'):
            style = styles['Heading2']
        elif p.style.name.startswith('Heading 3'):
            style = styles['Heading3']

        story.append(Paragraph(text, style))
        story.append(Spacer(1, 4))

    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            table_data.append(row_data)
        if table_data:
            t = Table(table_data)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e2e8f0')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#1e293b')),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#cbd5e1')),
            ]))
            story.append(t)
            story.append(Spacer(1, 10))

    pdf_doc.build(story)
    logger.info(f"Successfully converted via ReportLab fallback: {output_path}")
    return output_path


def pdf_to_word(input_path: str, output_path: Optional[str] = None, start_page: int = 0, end_page: Optional[int] = None) -> str:
    """
    Converts a PDF file to Word (.docx).
    Uses pdf2docx library which recreates text layout, tables, fonts, and inline images.
    """
    input_path = os.path.abspath(input_path)
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    ext = os.path.splitext(input_path)[1].lower()
    if ext != '.pdf':
        raise ValueError(f"Unsupported input format: {ext}. Expected .pdf")

    if not output_path:
        output_path = os.path.splitext(input_path)[0] + ".docx"
    else:
        output_path = os.path.abspath(output_path)

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    logger.info(f"Converting '{input_path}' to Word (.docx)...")
    from pdf2docx import Converter

    cv = Converter(input_path)
    cv.convert(output_path, start=start_page, end=end_page)
    cv.close()

    if os.path.exists(output_path):
        logger.info(f"Successfully converted: {output_path}")
        return output_path
    else:
        raise RuntimeError(f"Output file was not created: {output_path}")


def batch_convert(
    folder_path: str,
    direction: str = "auto",
    output_dir: Optional[str] = None,
    progress_callback: Optional[Callable[[int, int, str, bool], None]] = None
) -> Dict[str, Any]:
    """
    Batch converts files in a folder.
    
    direction:
      - 'word2pdf' : Convert all .docx / .doc files to PDF
      - 'pdf2word' : Convert all .pdf files to .docx
      - 'auto'     : Convert .docx/.doc -> .pdf AND .pdf -> .docx
    """
    folder_path = os.path.abspath(folder_path)
    if not os.path.isdir(folder_path):
        raise NotADirectoryError(f"Directory not found: {folder_path}")

    files_to_process = []
    for root, _, files in os.walk(folder_path):
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            file_path = os.path.join(root, f)

            if direction == "word2pdf" and ext in ['.docx', '.doc']:
                files_to_process.append((file_path, "word2pdf"))
            elif direction == "pdf2word" and ext == '.pdf':
                files_to_process.append((file_path, "pdf2word"))
            elif direction == "auto":
                if ext in ['.docx', '.doc']:
                    files_to_process.append((file_path, "word2pdf"))
                elif ext == '.pdf':
                    files_to_process.append((file_path, "pdf2word"))

    total = len(files_to_process)
    successful = []
    failed = []

    logger.info(f"Found {total} file(s) to process in batch mode.")

    for idx, (src_file, mode) in enumerate(files_to_process, 1):
        filename = os.path.basename(src_file)
        if output_dir:
            rel_dir = os.path.relpath(os.path.dirname(src_file), folder_path)
            target_folder = os.path.join(output_dir, rel_dir) if rel_dir != "." else output_dir
            base_name = os.path.splitext(filename)[0]
            target_ext = ".pdf" if mode == "word2pdf" else ".docx"
            dest_file = os.path.join(target_folder, base_name + target_ext)
        else:
            dest_file = None

        success = False
        try:
            if mode == "word2pdf":
                out = word_to_pdf(src_file, dest_file)
            else:
                out = pdf_to_word(src_file, dest_file)
            successful.append((src_file, out))
            success = True
        except Exception as e:
            logger.error(f"Failed converting {filename}: {e}")
            failed.append((src_file, str(e)))

        if progress_callback:
            progress_callback(idx, total, filename, success)

    return {
        "total": total,
        "successful_count": len(successful),
        "failed_count": len(failed),
        "successful": successful,
        "failed": failed
    }


def main():
    parser = argparse.ArgumentParser(
        description="Convert Word documents (.docx/.doc) to PDF and PDF to Word (.docx)."
    )
    subparsers = parser.add_subparsers(dest="command", help="Sub-command to execute")

    # Command: word2pdf
    p_w2p = subparsers.add_parser("word2pdf", help="Convert Word file (.docx/.doc) to PDF")
    p_w2p.add_argument("input", help="Path to input Word document")
    p_w2p.add_argument("-o", "--output", help="Output PDF file path (optional)")

    # Command: pdf2word
    p_p2w = subparsers.add_parser("pdf2word", help="Convert PDF file to Word (.docx)")
    p_p2w.add_argument("input", help="Path to input PDF document")
    p_p2w.add_argument("-o", "--output", help="Output .docx file path (optional)")
    p_p2w.add_argument("--start", type=int, default=0, help="Start page index (default 0)")
    p_p2w.add_argument("--end", type=int, default=None, help="End page index (exclusive)")

    # Command: batch
    p_batch = subparsers.add_parser("batch", help="Batch convert a folder of files")
    p_batch.add_argument("folder", help="Folder containing files to convert")
    p_batch.add_argument("-o", "--output-dir", help="Output directory for converted files")
    p_batch.add_argument("-d", "--direction", choices=["word2pdf", "pdf2word", "auto"], default="auto",
                         help="Conversion direction (default: auto)")

    # Command: gui
    subparsers.add_parser("gui", help="Launch Graphical User Interface")

    args = parser.parse_args()

    if args.command == "word2pdf":
        word_to_pdf(args.input, args.output)
    elif args.command == "pdf2word":
        pdf_to_word(args.input, args.output, start_page=args.start, end_page=args.end)
    elif args.command == "batch":
        res = batch_convert(args.folder, direction=args.direction, output_dir=args.output_dir)
        print(f"\nBatch processing complete: {res['successful_count']}/{res['total']} successful.")
    elif args.command == "gui" or not args.command:
        # Launch GUI
        import gui
        gui.launch_gui()
    else:
        parser.print_help()

if __name__ == "__main__":
    main()

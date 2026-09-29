import os
from docx import Document

def create_sample_docx(filepath="sample.docx"):
    doc = Document()
    doc.add_heading("Sample Document for File Type Converter", level=0)
    
    doc.add_paragraph(
        "This is a sample Word document created to test conversion to PDF and back to Word. "
        "It contains formatted text, lists, and tables."
    )
    
    doc.add_heading("Section 1: Feature List", level=1)
    p = doc.add_paragraph()
    p.add_run("Key Features:").bold = True
    doc.add_paragraph("• Word (.docx) to PDF conversion", style='List Bullet')
    doc.add_paragraph("• PDF to Word (.docx) conversion", style='List Bullet')
    doc.add_paragraph("• Batch file processing", style='List Bullet')
    doc.add_paragraph("• Interactive Graphical User Interface", style='List Bullet')
    
    doc.add_heading("Section 2: Sample Data Table", level=1)
    table = doc.add_table(rows=3, cols=3)
    table.style = 'Table Grid'
    
    headers = ["ID", "Name", "Status"]
    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        
    data = [
        ["101", "Document A", "Processed"],
        ["102", "Report B", "Pending"]
    ]
    for row_data in data:
        row_cells = table.add_row().cells
        for i, val in enumerate(row_data):
            row_cells[i].text = val
            
    doc.save(filepath)
    print(f"Created sample document: {filepath}")

if __name__ == "__main__":
    create_sample_docx()

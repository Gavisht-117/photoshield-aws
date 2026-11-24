"""Exports PhotoShield_Report.docx to PDF with Microsoft Word and prints the page count."""
import os
import sys

import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
src = os.path.join(HERE, "PhotoShield_Report.docx")
dst = os.path.join(HERE, "PhotoShield_Report.pdf")

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(src, ReadOnly=True)
    pages = doc.ComputeStatistics(2)  # wdStatisticPages
    doc.ExportAsFixedFormat(dst, 17)  # wdExportFormatPDF
    doc.Close(False)
finally:
    word.Quit()
print(f"Pages: {pages}\nSaved {dst}")
sys.exit(0)

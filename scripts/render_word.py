"""Windows fallback when the bundled LibreOffice renderer is unavailable.
Uses a separate invisible Word instance, then rasterizes PDF with PyMuPDF.
"""
import sys,json
from pathlib import Path
import win32com.client
import fitz
source=Path(sys.argv[1]).resolve();dest=Path(sys.argv[2]).resolve();dest.mkdir(parents=True,exist_ok=True)
w=win32com.client.DispatchEx('Word.Application');w.Visible=False;w.DisplayAlerts=0
try:
    d=w.Documents.Open(str(source),ReadOnly=True,AddToRecentFiles=False)
    d.ExportAsFixedFormat(str(dest/'render.pdf'),17);d.Close(False)
finally:w.Quit()
with fitz.open(dest/'render.pdf') as pdf:
    for i,p in enumerate(pdf):p.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(dest/f'page-{i+1}.png')
    print(json.dumps({'file':source.name,'pages':len(pdf),'renderer':'Microsoft Word 16.0 COM ExportAsFixedFormat + PyMuPDF'},ensure_ascii=False))

import sys
import os
import comtypes.client
import subprocess

wdFormatPDF = 17

in_file = os.path.abspath(sys.argv[1])
pdf_file = os.path.splitext(in_file)[0] + ".pdf"
png_file = os.path.splitext(in_file)[0] + "%02d.png"

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(pdf_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()

subprocess.call(["c:\Program Files (x86)\ImageMagick-6.8.5-Q16\convert", pdf_file, png_file])

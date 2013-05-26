import sys
import os
import comtypes.client
import subprocess

if sys.version < '3':
    import _winreg as winreg
else:
    import winreg

wdFormatPDF = 17

def get_imagemagick_bin_path():
    return winreg.QueryValueEx(winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE,
                                                "SOFTWARE\ImageMagick\Current"),
                              "BinPath")[0]

in_file = os.path.abspath(sys.argv[1])
pdf_file = os.path.splitext(in_file)[0] + ".pdf"
png_file = os.path.splitext(in_file)[0] + "%02d.png"

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(pdf_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()

subprocess.call([os.path.join(get_imagemagick_bin_path(), "convert"),
                 "-density", "150", pdf_file, png_file])

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

def main():
    in_file = os.path.abspath(sys.argv[1])
    png_file = os.path.splitext(in_file)[0] + "%02d.png"
    
    if in_file.endswith(".doc"):
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        pdf_file = os.path.splitext(in_file)[0] + ".pdf"
        doc.SaveAs(pdf_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
    elif in_file.endswith(".pdf"):
        pdf_file = in_file
    else:
        sys.stderr.write("Unsupported file type: %s" % os.path.splitext(in_file)[1])
        sys.exit(-1)
    
    subprocess.call([os.path.join(get_imagemagick_bin_path(), "convert"),
                     "-density", "150", pdf_file, png_file])

if __name__ == '__main__':
    main()
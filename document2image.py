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

def convert_doc_to_pdf(doc_file, pdf_file):
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(doc_file)
    doc.SaveAs(pdf_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

def convert_pdf_to_png(pdf_file, png_file):
    subprocess.call([os.path.join(get_imagemagick_bin_path(), "convert"),
                     "-density", "150", pdf_file, png_file])

def main():
    in_file = os.path.abspath(sys.argv[1])
    in_file_ext = os.path.splitext(in_file)[1]
    if in_file_ext not in [".doc", ".docx", ".pdf"]:
        sys.stderr.write("Unsupported file type: %s" %
                         os.path.splitext(in_file)[1])
        sys.exit(-1)
        
    # TODO: Check if file exists

    file_base = os.path.splitext(in_file)[0]

    if in_file_ext in [".doc", ".docx"]:
        convert_doc_to_pdf(in_file, file_base + ".pdf")

    convert_pdf_to_png(file_base + ".pdf", file_base + "%02d.png")
    
    # TODO: Delete PDF

if __name__ == '__main__':
    main()
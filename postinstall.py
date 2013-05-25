import sys
import os.path as osp

if sys.version < '3':
    import _winreg as winreg
else:
    import winreg

CONVERT_TO_PNG = "Convert to PNG"
VERB_KEY = r"Software\Classes\%s\shell\%s"
COMMAND_KEY = VERB_KEY + r"\command"

def register_file_type(file_type):
    pythonw = osp.abspath(osp.join(sys.prefix, 'pythonw.exe'))
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, COMMAND_KEY % 
                           (file_type, CONVERT_TO_PNG))
    winreg.SetValueEx(key, "", 0, winreg.REG_SZ,
                      '"%s" "%s\Scripts\doc2png.py" "%%1"' % 
                      (pythonw, sys.prefix))

def unregister_file_type(file_type):
    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, COMMAND_KEY % (file_type,
                                                              CONVERT_TO_PNG))
    winreg.DeleteKey(winreg.HKEY_CURRENT_USER, VERB_KEY % (file_type,
                                                           CONVERT_TO_PNG))

def install():
    register_file_type("Word.Document.8")
    register_file_type("Word.Document.12")
    
def remove():
    unregister_file_type("Word.Document.12")
    unregister_file_type("Word.Document.8")

if __name__ == '__main__':
    if sys.argv[1] == '-install':
        try:
            install()
        except OSError:
            print >> sys.stderr, "Installation failed, "\
                                 "try running installer as administrator."
    elif sys.argv[1] == '-remove':
        remove()

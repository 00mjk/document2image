Prerequisites
=============
* Python 2.7 or 3.1, available [here](http://www.python.org/getit/). Note that Python 3.1 is
  required to support Unicode file names.
* comtypes, available [here](http://sourceforge.net/projects/comtypes/files/comtypes/)
* Ghostscript, available [here](http://www.ghostscript.com/download/gsdnld.html)
* ImageMagick 6.8.5-Q16, available [here](http://www.imagemagick.org/download/binaries/ImageMagick-6.8.5-7-Q16-x86-dll.exe)
* Microsoft Word

Building a Windows installer
============================
`$ python setup.py bdist_wininst --install-script=postinstall.py`

Usage
=====
In File Explorer
----------------------
Right-click the Word Document and select 'Convert to PNG'.

From the command line
---------------------
`$ python doc2png.py <doc file>`

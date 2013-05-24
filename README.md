Prerequisites
=============
1. Python 2.7, available [here](http://www.python.org/getit/)
2. comtypes, available [here](http://sourceforge.net/projects/comtypes/files/comtypes/)
3. Ghostscript, available [here](http://www.ghostscript.com/download/gsdnld.html)
4. ImageMagick, available [here](http://www.imagemagick.org/script/binary-releases.php#windows)
5. Microsoft Word

Building a Windows installer
============================
`$ python setup.py bdist_wininst --install-script=postinstall.py`

Usage
=====
Right-click the Word Document and select 'Convert to PNG'

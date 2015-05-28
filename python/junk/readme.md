# -*- mode: python -*-

"""


Python:
https://www.python.org/downloads/windows/

pywin32:
http://sourceforge.net/projects/pywin32/files/pywin32/

Call python from excel:
http://xlwings.org/examples/
pip install xlwings

Freeze/compile python:
https://github.com/pyinstaller/pyinstaller/wiki (there are alternatives)
pip install pyinstaller


pyinstaller --onedir battery.py
change: PYTHON_FROZEN = ThisWorkbook.Path & "\dist\battery"








https://www.microsoft.com/en-us/download/details.aspx?id=29


1. copy 
C:\users\mike\appdata\local\enthought\canopy\user\lib\site-packages\xlwings-0.3.5-py2.7.egg\xlwings
xlwings_template.xlsm

C:\Windows\winsxs\x86_microsoft.vc90.crt_1fc8b3b9a1e18e3b_9.0.21022.8_none_bcb86ed6ac711f91


msvcm90.dll
msvcp90.dll
msvcr90.dll

pyinstaller --paths=C:\Windows\winsxs --paths=C:\Users\Mike\AppData\Local\Enthought\Canopy\User battery.py 

\Lib\site-packages







Matlab > Apps > Application Compiler

New > Library Compiler Project > Excel Add-in

http://au.mathworks.com/help/compiler/excel/create-a-microsoft-excel-add-in-and-com-component-from-matlab-code.html

Microsoft Windows SDK for Windows 7 and .NET Framework 4
http://www.microsoft.com/en-us/download/details.aspx?id=8279

Excel > Developer > Macro security > "Trust access to the VBA project object model"




Github
Excel (Office business 2010)
Windows 7 SDK and .NET 4
Canopy express
   xlwings


"""



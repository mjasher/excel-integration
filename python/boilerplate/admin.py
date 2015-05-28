
"""
To run
=======================
```python admin.py```
then open use_battery.xlsm


Dependencies
=======================

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

Notes
=======================

pyinstaller --onedir battery.py

to start a new excel notebook with xlwings module included:
	from xlwings import Workbook
	Workbook.open_template()
then change in xlwings VBA module: 
	PYTHON_FROZEN = ThisWorkbook.Path & "\dist\battery"

ensure you have "import ..." at top of python files, not "from ... import"

"""


import shutil
import os
import subprocess
def copytree(src, dst, symlinks=False, ignore=None):
    for item in os.listdir(src):
    	# print item
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)

# base = "C:\\Users\\Mike\\Desktop\\excel\\python\\"
# subprocess.call(["cd", base])
base = ""
subprocess.call(["pyinstaller", "--onedir", "battery.py"])
subprocess.call(["pyinstaller", "--onedir", "battery2.py"])

dest = base+"dist\\battery"
src = base+"dist\\battery2"
copytree(src, dest)


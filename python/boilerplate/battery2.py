import xlwings

from xlwings import Workbook, Sheet, Range, Chart

# for debugging
# Workbook.set_mock_caller("C:\\Users\\Mike\\Desktop\\excel\\python\\use_battery.xlsm")

wb = Workbook.caller()
total = Range('D2').value
print total
vals = Range('A2:C2').value 
Range('A2:C2').value = [v/total for v in vals]


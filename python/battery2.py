import xlwings

from xlwings import Workbook, Sheet, Range, Chart


if __name__ == '__main__':
	Workbook.set_mock_caller("C:\\Users\\Mike\\Desktop\\excel\\python\\use_battery.xlsm")

wb = Workbook.caller()
# wb = Workbook()  # Creates a connection with a new workbook

total = Range('D2').value
print total
vals = Range('A2:C2').value 
Range('A2:C2').value = [v/total for v in vals]

# from xlwings import Workbook
# Workbook.open_template()
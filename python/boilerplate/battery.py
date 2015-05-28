import xlwings
from xlwings import Workbook, Sheet, Range, Chart

# Creates a connection with a new workbook
# wb = Workbook()  

wb = Workbook.caller()
Range('A1').value = 'Foo 1'
print Range('A1').value
Range('A1').value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
print Range('A1').table.value  # or: Range('A1:C2').value
print Sheet(1).name
chart = Chart.add(source_data=Range('A1').table)	


# to open excel template
# from xlwings import Workbook
# Workbook.open_template()
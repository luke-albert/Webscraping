import openpyxl as xl
from openpyxl.styles import Font

wb = xl.Workbook()

ws = wb.active

ws.title = 'First Sheet'

wb.create_sheet(index=1, title='Second Sheet')

ws['A1'] = 'Invoice'

ws['A1'].font = Font(name='Times New Roman', size=24, italic=False, bold=True)

myfont = Font(name='Times New Roman', size=24, italic=False, bold=True)

ws['A1'].font = myfont

ws['A2'] = 'Tires'
ws['A3'] = 'Brakes'
ws['A4'] = 'Alignment'

ws.merge_cells('A1:B1')  # merge and center cells

ws.unmerge_cells('A1:B1')

ws['B2'] = 450
ws['B3'] = 225
ws['B4'] = 150

ws['A8'] = 'Total'
ws['A8'].font = myfont

ws['B8'] = '=SUM(B2:B4)'

ws.column_dimensions['A'].width = 25

wb.save('PythontoExcel.xlsx')


# Read the excel file = 'ProduceReport.xlsx that you created earlier.
# write all the contents of this file to 'Second Sheet' in the current
# workbook

# display the Grand Total and Average of 'Amt Sold' amd 'Total'
# at the bottom of the list along with appropiate labels

write_sheet = wb['Second Sheet']

read_wb = xl.load_workbook('ProduceReport.xlsx')
read_ws = read_wb['ProduceReport']


maxC = read_ws.max_column
maxR = read_ws.max_row


i = 1
for currentrow in read_ws.iter_rows(min_row=1, max_row=read_ws.max_row, max_col=read_ws.max_column):
    fruit = (currentrow[0].value)
    cost = (currentrow[1].value)
    sold = (currentrow[2].value)
    total = (currentrow[3].value)
    write_sheet.cell(i, 1).value = fruit
    write_sheet.cell(i, 2).value = cost
    write_sheet.cell(i, 3).value = sold
    write_sheet.cell(i, 4).value = total
    i += 1
    # print(currentrow[1].value)
    # print(currentrow[2].value)

wb.save('PythontoExcel.xlsx')

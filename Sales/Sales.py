import xlrd
from datetime import datetime
import os
import xlsxwriter

sales = [['S.No.', 'Invoice No.', 'Date', 'Company', 'GSTIN', 'Sub Total', 'CGST', 'SGST', 'IGST', 'Total Tax', 'Grand Total (₹)']]
sNumber = 0

currentMonth = datetime.now().month
currentYear = datetime.now().year

if currentMonth == 1:
    invoicesLocation = 'C:/Users/devas.LAPTOP-5GL18CQ7/Documents/Unique Interiors/Invoices/20-21/12/'
else:
    invoicesLocation = 'C:/Users/devas.LAPTOP-5GL18CQ7/Documents/Unique Interiors/Invoices/20-21/' + str(currentMonth - 1) + '/'

invoices = os.listdir(invoicesLocation)

for invoice in invoices:
    invoiceLocation = invoicesLocation + invoice

    invoice = xlrd.open_workbook(invoiceLocation).sheet_by_index(0)
    
    sNumber = sNumber + 1
    invoiceNumber = int(invoice.cell_value(8, 5))
    date = datetime(*xlrd.xldate_as_tuple(invoice.cell_value(9, 5), 0)).date()
    company = invoice.cell_value(9, 0)
    gstin = invoice.cell_value(14, 0)[8:]
    subTotal = round(invoice.cell_value(46, 7), 2)
    
    if(invoice.cell_value(47, 7)):
        cgst = round(invoice.cell_value(47, 7), 2)
    else:
        cgst = 0.0
        
    if(invoice.cell_value(48, 7)):
        sgst = round(invoice.cell_value(48, 7), 2)
    else:
        sgst = 0.0
        
    if(invoice.cell_value(49, 7)):
        igst = round(invoice.cell_value(49, 7), 2)
    else:
        igst = 0.0
        
    totalTax = cgst + sgst + igst
    grandTotal = round(invoice.cell_value(51, 7), 2)
    
    sales.append([sNumber, invoiceNumber, date, company, gstin, subTotal, cgst, sgst, igst, totalTax, grandTotal])

if currentMonth > 4:
    directory = str(currentYear) + '-' + str(currentYear + 1)
    file = directory + '/' + str(currentMonth - 1) + '.xlsx'
else:
    directory = str(currentYear - 1) + '-' + str(currentYear)
    if currentMonth == 1:
        file = directory + '/12.xlsx'
    else:
        file = directory + '/' + str(currentMonth - 1) + '.xlsx'

if not os.path.exists(directory):
    os.makedirs(directory)

workbook = xlsxwriter.Workbook(file)

bold = workbook.add_format({'bold': 1, 'border': 1, 'align': 'center'})
dateFormat = workbook.add_format({'num_format': 'dd-mm-yyyy', 'border': 1})
numberFormat = workbook.add_format({'num_format': '0.00', 'border': 1})
alignFormat = workbook.add_format({'align': 'center', 'border': 1})
boldNumberFormat = workbook.add_format({'bold': 1, 'num_format': '0.00', 'border': 1})
mergeFormat = workbook.add_format({'bold': 1, 'border': 1, 'align': 'right'})

worksheet = workbook.add_worksheet() 

worksheet.set_column(0, 0, 5)
worksheet.set_column(1, 1, 10)
worksheet.set_column(2, 2, 10)
worksheet.set_column(3, 3, 30)
worksheet.set_column(4, 4, 16)
worksheet.set_column(5, 9, 10)
worksheet.set_column(10, 10, 15)

for row in range(len(sales)):
    for column in range(len(sales[0])):
        if(row == 0):
            worksheet.write(row, column, sales[row][column], bold)
        else:
            if(column == 2):
                worksheet.write(row, column, sales[row][column], dateFormat)
            elif(column > 4):
                worksheet.write(row, column, sales[row][column], numberFormat)
            else:
                worksheet.write(row, column, sales[row][column], alignFormat)
                
row = row +1
mergedCells = 'A' + str(row+1) + ':E' + str(row+1)

worksheet.merge_range(mergedCells, 'Total', mergeFormat)

for column in range(5, 11):
    columnCharacter = chr(65 + column)
    formula = '=SUM(' + columnCharacter + '2:' + columnCharacter + str(row) + ')'
    worksheet.write(row, column, formula, boldNumberFormat)

workbook.close() 

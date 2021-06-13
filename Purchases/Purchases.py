from datetime import datetime
import os
import xlsxwriter

currentMonth = datetime.now().month
currentYear = datetime.now().year
if currentMonth > 3:
    directory = str(currentYear) + '-' + str(currentYear + 1)
    file = directory + '/' + str(currentMonth - 1) + '.xlsx'

if not os.path.exists(directory):
    os.makedirs(directory)

workbook = xlsxwriter.Workbook(file) 

worksheet = workbook.add_worksheet() 

worksheet.write('A1', 'S.No.') 
worksheet.write('B1', 'Invoice No.') 
worksheet.write('C1', 'Date') 
worksheet.write('D1', 'Company')
worksheet.write('E1', 'GSTIN')
worksheet.write('F1', 'Tax Rate')
worksheet.write('G1', 'Taxable Amt.')
worksheet.write('H1', 'CGST')
worksheet.write('I1', 'SGST')
worksheet.write('J1', 'IGST')
worksheet.write('K1', 'Total Tax')
worksheet.write('L1', 'Grand Total (₹)')

workbook.close() 

import xlrd
import csv
from numbers import Number

def excelToCsv(inputFile, sheet, outputFile):
    wb = xlrd.open_workbook(inputFile)

    # Checks if the given sheet is a numer or string
    if isinstance(sheet, Number):
        sh = wb.sheet_by_index(sheet) #Takes the sheet by its numerical order
    else:
        sh = wb.sheet_by_name(sheet) # Takes the sheet by its name

    cvsFile = open(outputFile, 'wb')
    wr = csv.writer(cvsFile, quoting=csv.QUOTE_ALL)

    for rownum in xrange(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    cvsFile.close()

excelToCsv('sample.xlsx', 'Hoja1', 'your_csv_file.csv')

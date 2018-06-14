import xlrd
import csv

def excelToCsv(inputFile, page, outputFile):
    wb = xlrd.open_workbook(inputFile)
    sh = wb.sheet_by_index(page) #filewb.sheet_by_name('Hoja1')
    your_csv_file = open(outputFile, 'wb')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in xrange(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()

excelToCsv('sample.xlsx', 0, 'your_csv_file.csv')

import xlrd
import csv
from numbers import Number
import os.path
import json

''' Configurations, this values will change the behaivor of the script
and the way it interacts with the user
'''
global_allowedExtensions = ['.xls', '.xlsx']
global_extensionNotAllowed = 'The extension %s is not allowed.'

''' Obtains the extension of a file which already exists in disk

@param filename string - the name of the file
@return string
'''
def getFileExtension(fileName):
    return os.path.splitext(fileName)[1]

''' Convers an excel file into a csv

@param inputFile string - includes the extension and it may be .xls or .xlsx
@param sheet string - the name of the sheet, alternatively a number that
represents the order of the pages can be passed where 0 would be the first sheet
@param outputFile string - the desired name of the output file
'''
def excelToCsv(inputFile, sheet, outputFile):
    fileExtension = getFileExtension(inputFile)

    if fileExtension not in global_allowedExtensions:
        raise Exception(global_extensionNotAllowed % fileExtension)

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

    # Output for the user
    output = {}
    output['originalFile'] = inputFile
    output['createdFile'] = outputFile
    output['rows'] = sh.nrows

    print(json.dumps(output, sort_keys=True, indent=4))

excelToCsv('sample.xlsx', 'Hoja1', 'your_csv_file.csv')

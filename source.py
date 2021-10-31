#!/usr/bin/python

from openpyxl import load_workbook
from os import path
import excep as ude
import sys


def main():
    # Ensure proper usage
    if len(sys.argv) != 2:
        print("Usage: python source.py reference")
        sys.exit(1)
    
    # Check correct file extension
    allowableFileTypes = ('.xlsx', '.xlsm', '.xltx', '.xltm')
    if not sys.argv[1].endswith(allowableFileTypes):
        print("Invalid file type used as reference.")
        sys.exit(2)

    # Check reference file in directory
    referenceDir = "tables"
    referenceFile = path.join(referenceDir, sys.argv[1])
    if not path.exists(referenceFile):
        print("Reference file {} missing from tables directory.".format(sys.argv[1]))
        sys.exit(3)

    # Load variables into memory
    wb = load_workbook(referenceFile)
    tbName = getTableName(wb)
    # tbData = getTableData(wb, tbName)
    # rowLength = len(tbData)
    # colLength = len(tbData[0])
    # sSize = getSampleSize()
    # pos = getPosition(rowLength, colLength, tbData)

    print(tbName)
    # print(tbData)
    # print(tbData)
    # print(rowLength)
    # print(colLength)
    # print(sSize)
    # print(pos)

    sys.exit(0)


def getTableName(workbook):
    sheets = tuple(workbook.sheetnames)

    print("Loaded {} sheets from {}".format(len(sheets), sys.argv[1]))
    
    for index, title in enumerate(sheets):
        print("{} - {}".format(index + 1, title))

    while True:
        try:
            usrTable = int(input("Enter number of desired table: "))

            errMsg = "Enter number from available options."
            if (usrTable - 1) < 0 or (usrTable - 1) > len(sheets):
                # FIXME Custom exception does not raise when true. Prints a different error message
                raise ude.OptionError(errMsg)

            break
        
        except ValueError:
            print(errMsg)

        except KeyboardInterrupt:
            sys.exit(0)
    
    table = sheets[usrTable - 1]

    return table


def getTableData(key, value):
    ws = key[value]
    data = []

    for row in ws.iter_rows(min_row=1, values_only=True):
        data.append(row)

    return data

def getSampleSize():
    while True:
        try:
            sampleSize = int(input("Enter sample size: "))

            errMsg = "Value must be a valid positive integer."
            if sampleSize <= 0:
                # FIXME Custom exception does not raise when true. Prints a different error message
                raise ude.InputError(errMsg)

            break

        except ValueError:
            print(errMsg)

        except KeyboardInterrupt:
            sys.exit(0)

    return sampleSize


def getPosition(maxRows, maxColumns, table):
    errMsg = "Value must be a valid positive integer"

    while True:
        try:
            usrRow = int(input("Enter starting row(-y): "))
            if usrRow <= 0 or usrRow > maxRows:
                # FIXME Custom exception does not raise when true. Prints a different error message
                raise ude.InputError("Value exceeds maximum number of rows in table.")

            break

        except ValueError:
            print("Value must be a valid positive integer")

        except KeyboardInterrupt:
            sys.exit(0)

    while True:
        try:
            usrCol = int(input("Enter starting column(+x): "))
            if usrCol < 0 or usrCol > maxColumns:
                # FIXME Custom exception does not raise when true. Prints a different error message
                raise ude.InputError("Value exceeds maximum number of columns in table.")

            break

        except ValueError:
            print(errMsg)

        except KeyboardInterrupt:
            sys.exit(0)

    rStart = table[usrRow - 1][usrCol - 1]

    return rStart


if __name__ == "__main__":
    main()

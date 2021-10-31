#!/usr/bin/python

from openpyxl import load_workbook
from os import path
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

    sep = "----------------------------------------------------------"
    print("{}\n\tRANDOM NUMBER SEQUENCE GENERATOR\n{}".format(sep, sep))

    # Load variables into memory. Wrap this is a loop with a flag to allow reprompt
    wb = load_workbook(referenceFile)
    
    tbName = getTableName(wb, sep)
    tbData = getTableData(wb, tbName)
    pSize = getPopulation(sep)
    sSize = getSampleSize(sep)
    rowLength = len(tbData)
    colLength = len(tbData[0])
    pos = getPosition(rowLength, colLength, tbData, sep)
    swp = getSweep(sep)

    isFinal = finalize(tbName, pSize, sSize, pos, swp, sep)
    
    if isFinal != True:
        print("Exiting program.")
        sys.exit(0)

    sys.exit(0)


def getTableName(workbook, s):
    sheets = tuple(workbook.sheetnames)

    print("INFO: Loaded {} sheets from {}\n{}\n\t\tENTER VARIABLES\nTable Names:".format(len(sheets), sys.argv[1], s))
    
    for index, title in enumerate(sheets):
        print("({}) {}".format(index + 1, title))

    while True:
        try:
            usrTable = int(input("Select desired table: "))

            if (usrTable - 1) < 0 or (usrTable - 1) > len(sheets):
                raise ValueError

            break

        except ValueError:
            print("Enter number from available options.")

        except KeyboardInterrupt:
            sys.exit(0)
    
    print(s)
    table = sheets[usrTable - 1]

    return table


def getTableData(key, value):
    ws = key[value]
    data = []

    for row in ws.iter_rows(min_row = 1, values_only = True):
        data.append(row)

    return data


def getPopulation(s):
    while True:
        try:
            populationLimit = int(input("Enter population limit(inclusive): "))

            if populationLimit <= 0:
                raise ValueError

            break

        except ValueError:
            print("Value must be a valid positive integer")

        except KeyboardInterrupt:
            sys.exit(0)

    print(s)

    return populationLimit


def getSampleSize(s):
    while True:
        try:
            sampleSize = int(input("Enter sample size: "))

            if sampleSize <= 0:
                raise ValueError

            break

        except ValueError:
            print("Value must be a valid positive integer.")

        except KeyboardInterrupt:
            sys.exit(0)

    print(s)

    return sampleSize


def getPosition(maxRows, maxColumns, table, s):
    while True:
        try:
            usrRow = int(input("Enter starting row(-y): "))
            if usrRow <= 0 or usrRow > maxRows:
                raise ValueError

            break

        except ValueError:
            print("Value exceeds maximum number of rows in table.")

        except KeyboardInterrupt:
            sys.exit(0)

    while True:
        try:
            usrCol = int(input("Enter starting column(+x): "))
            if usrCol <= 0 or usrCol > maxColumns:
                raise ValueError

            break

        except ValueError:
            print("Value exceeds maximum number of columns in table.")

        except KeyboardInterrupt:
            sys.exit(0)

    print(s)
    rStart = table[usrRow - 1][usrCol - 1]

    return rStart, usrRow - 1, usrCol - 1


def getSweep(s):
    allowableSweep = ("Left to Right", "Down, Left to Right")
    print("Sweep patterns:")

    for index, sweep in enumerate(allowableSweep):
        print("({}) {}".format(index + 1, sweep))

    while True:
        try:
            usrSweep = int(input("Select desired sweep pattern: "))

            if (usrSweep - 1) < 0 or (usrSweep - 1) > len(allowableSweep):
                raise ValueError

            break

        except ValueError:
            print("Invalid sweep pattern option")

        except KeyboardInterrupt:
            sys.exit(0)

    print(s)

    return allowableSweep[usrSweep - 1]


def finalize(tableName, populationSize, sampleSize, position, sweep, s):
    print("Table: {}\nPopulation Size: {}\nSample Size: {}".format(tableName, populationSize, sampleSize))
    print("RNS Start Value: {}\nRNS Start Position:  row {}, col. {}".format(position[0], position[1] + 1, position[2] + 1))
    print("Sweep: {}".format(sweep))

    while True:
        try:
            response = input("Is this correct(y/n)? ").rstrip().lower()
            
            if response not in ['y', 'n']:
                continue

            else:
                print(s)

                if response == 'y':
                    return True

                else:
                    return False
        
        except KeyboardInterrupt:
            sys.exit(0)


if __name__ == "__main__":
    main()

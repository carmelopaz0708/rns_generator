#!/usr/bin/python

from openpyxl import load_workbook
from os import path
import excep as e
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

    # Load workbook into memory
    wb = load_workbook(referenceFile)
    
    userInput = getInput(wb)
    print(userInput)

    sys.exit(0)


# TODO Must separate getInputs as their own function
def getInput(workbook):
    # Get reference table
    sheet = tuple(workbook.sheetnames)
    print("Loaded {} sheets from {}.".format(len(sheet), sys.argv[1]))
    for i, j in enumerate(sheet):
        print("{} - {}".format(i + 1, j))

    errOptMsg = "Enter number from available options."
    while True:
        try:
            usrTable = int(input("Enter number of desired table: "))

            if (usrTable - 1) < 0 or (usrTable - 1) > len(sheet):
                raise e.OptionError(errOptMsg)

            break
        
        except ValueError:
            print(errOptMsg)

        except KeyboardInterrupt:
            sys.exit(0)
    
    table = sheet[usrTable - 1]

    # Get sample size
    errValMsg = "Value must be a valid positive integer."
    while True:
        try:
            sampleSize = int(input("Enter sample size: "))

            if sampleSize <= 0:
                raise e.InputError(errValMsg)

            break

        except ValueError:
            print(errValMsg)

        except KeyboardInterrupt:
            sys.exit(0)
    
    # Get RNS start position.
    # TODO Iterate rows and column with worksheet.rows / worksheet.columns
    # while True:
    #     try:
    #         column = int(input("Enter RNS column position: "))

    #         # Check if column goes beyond column limit
    #         break

    #     except:
    #         pass

    # Check input if it satisfies table requirements
    # Might need to add some more setup to reference files to indicate metadata. Or use a csv file instead

    # Before returning values, reprompt the user if values are acceptable
    return table, sampleSize


if __name__ == "__main__":
    main()

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


def getInput(workbook):
    # Get reference table
    sheet = tuple(workbook.sheetnames)
    print("Enter number selection to load table (1, 2, ...)")
    for index, title in enumerate(sheet):
        print("{}. {}".format(index + 1, title))

    errorOptionMsg = "Enter selection from available options."
    while True:
        try:
            inputReference = int(input("Enter selected table (1, 2, ...): "))

            if (inputReference - 1) < 0 or (inputReference - 1) > len(sheet):
                raise e.OptionError(errorOptionMsg)

            break
        
        except ValueError:
            print(errorOptionMsg)

        except KeyboardInterrupt:
            sys.exit(0)

    # Get sample size
    errorValueMsg = "Value must be a valid positive integer."
    while True:
        try:
            sampleSize = int(input("Enter sample size: "))

            if sampleSize <= 0:
                raise e.InputError(errorValueMsg)

            break

        except ValueError:
            print(errorValueMsg)

        except KeyboardInterrupt:
            sys.exit(0)
    
    # Get RNS start position

    # Check input if it satisfies table requirements
    # Might need to add some more setup to reference files to indicate metadata. Or use a csv file instead

    # Before returning values, reprompt the user if values are acceptable
    return inputReference, sampleSize


if __name__ == "__main__":
    main()

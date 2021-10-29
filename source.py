#!/usr/bin/python

from openpyxl import load_workbook
from os import path
import sys
import excep


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
    pSize = getPopulation()
    print(pSize)

    sys.exit(0)


# Change this function to return a tuple of input variables
def getPopulation():
    while True:
        try:
            populationSize = int(input("Enter population size: "))
            if populationSize <= 0:
                raise excep.InputError("Population must be a valid positive integer.")

            break

        except ValueError:
            print("Population must be a valid positive integer.")
        
        except KeyboardInterrupt:
            sys.exit(0)
        
    return populationSize


if __name__ == "__main__":
    main()

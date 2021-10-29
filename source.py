from openpyxl import load_workbook
import sys
import os
import helper

def main():
    # Ensure proper usage
    if len(sys.argv) != 2:
        print("Usage: python source.py table.xlsx")
        sys.exit(1)
    
    # Check if file exists in tables directory
    tablesDir = "tables"
    tablesFile = os.path.join(tablesDir, sys.argv[1])
    if not os.path.exists(tablesFile):
        print("{} not in folder".format(sys.argv[1]))
        sys.exit(2)

    # Load file into memory. File must be stored in a multi-dim array. Use classes to separate worksheets
    wb = load_workbook(tablesFile)
    sn = wb.sheetnames
    table = []

    print(sn)

    sys.exit(0)


if __name__ == "__main__":
    main()

# REINITIALIZE REPOSITORY

from openpyxl import load_workbook
import sys

def main():
    if len(sys.argv) != 2:
        print("Usage: python source.py table.xlsx")
        sys.exit(1)

    # Get workbook, sheetnames
    wb = load_workbook(sys.argv[1])
    sn = wb.sheetnames

    # Find a way to extract this key. Use dictionary methods
    table = []
    ws1 = wb[sn[0]]
    # Store as a multi-dim list instead of a list of tuples

    # Use ws.iter_rows()
    for row in ws1.values:
        table.append(row)

    # for row in ws1.values:
    #     print(row)

    print(table)

    sys.exit(0)


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
import os
import sys
import xlwings

__doc__ = """Calling Python script from Excel workbook"""


def main():
    """
    :return: application exit code
    """
    working_dir = os.path.dirname(os.path.realpath(__file__))
    excel_filepath = os.path.join(working_dir, "Accounting.xlsx")
    print("Calling script %s from Excel workbook: %s" % (os.path.realpath(__file__), excel_filepath))
    excel_workbook = xlwings.Book("Accounting.xlsx").caller()
    accounting_sheet = excel_workbook.sheets["Accounting"]

    # Make Total of all accounting expenses
    # Single cells by default return either as float, unicode, None or datetime
    # Range return list of same type values
    # However, once they are in Python, lists lose the information about the orientation
    # Exception: 2 dimensional Ranges are automatically returned as nested lists
    accounting_sheet.range("A4").value = "Total"
    expenses = accounting_sheet.range("B1:B3").value
    print("Expenses list: {}".format(expenses))
    total = sum(expenses)
    print("Total = {}".format(total))
    accounting_sheet.range("B4").value = total

    excel_workbook.save(excel_filepath)

    return 0


if __name__ == '__main__':
    sys.exit(main())

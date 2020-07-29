# -*- coding: utf-8 -*-
import os
import sys
import xlwings

__doc__ = """Creating Excel workbook"""


def main():
    """
    :return: application exit code
    """
    working_dir = os.path.dirname(os.path.realpath(__file__))
    excel_filepath = os.path.join(working_dir, "Example.xlsx")
    print("Creating Excel workbook: %s" % excel_filepath)
    excel_workbook = xlwings.Book()
    accounting_sheet = excel_workbook.sheets.add()
    accounting_sheet.name = "Accounting"
    tm_sheet = excel_workbook.sheets.add()
    tm_sheet.name = "Time Management"

    accounting_sheet.range("A1").value = "Rent"
    accounting_sheet.range("A2").value = "Transport"
    accounting_sheet.range("A3").value = "Home"
    accounting_sheet.range("B1").value = 10500
    accounting_sheet.range("B2").value = 3600
    accounting_sheet.range("B3").value = 5400

    tm_sheet.range("A1").value = "Programming"
    tm_sheet.range("A2").value = "Study"
    tm_sheet.range("A3").value = "Rest"
    tm_sheet.range("B1").value = 8.45
    tm_sheet.range("B2").value = 4.5
    tm_sheet.range("B3").value = 1.8

    excel_workbook.save(excel_filepath)
    excel_workbook.app.quit()

    return 0


if __name__ == '__main__':
    sys.exit(main())

import xlwings


__doc__ = """Calling Python function from Excel workbook
1.Enable Trust access to the VBA project object model 
under File > Options > Trust Center > Trust Center Settings > Macro Settings
2.Install the add-in via command prompt: xlwings addin install
3.xlwings quickstart Functional
The default addin settings expect a Python source file in the way it is created by quickstart:
- in the same directory as the Excel file
- with the same name as the Excel file, but with a .py ending instead of .xlsm
"""


@xlwings.func
def double_square(x):
    """Returns twice the square of the argument"""
    return 2 * x * x


@xlwings.func
def double_sum(x, y):
    """Returns twice the sum of the two arguments"""
    return 2 * (x + y)


@xlwings.sub  # only required if you want to import it or run it via UDF Server
def main():
    wb = xlwings.Book.caller()
    sheet = wb.sheets[0]
    print("Worksheet: %s" % sheet)


if __name__ == "__main__":
    xlwings.Book("Functional.xlsm").set_mock_caller()
    main()

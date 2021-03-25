import win32com.client as win32
import pandas as pd

# excel = win32.gencache.EnsureDispatch("Excel.Application")
# wb = excel.Workbooks.Open(r"D:\경기과학고\상장인쇄\test.xls")
# ws = wb.Worksheets(1)
# excel.Visible = True
excel = pd.read_excel(r"D:\경기과학고\상장인쇄\test.xls")
print(excel)
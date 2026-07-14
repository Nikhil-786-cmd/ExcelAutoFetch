# ExcelAutoFetch
How to create new Excel Files from Python, as well as replacing data, saving Workbooks, automatically opening Excel Files and closing the Excel Programme, as well as creating Pandas DataFrames and automatically storing these within Excel!
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from win32com.client import Dispatch
#sheet["A1"] = "hello"
#sheet["B1"] = "YouTube"

workbook.save(filename="hello_youtube.xlsx")
#xl = Dispatch("Excel.Application")
#xl.Visible = True

#wb = xl.Workbooks.Open(r'"C:\Users\Lenovo\Desktop\hello_youtube.xlsx')

#cell = sheet["A1"]

#cell.value = "Greetings"
#print(cell.value)

workbook.save(filename="hello_youtube.xlsx")
x1=Dispatch("Excel.Application")
x1.Visible = True

wb = x1.workbook.open(r'C:\Users\Lenovo\Desktop\hello_youtube.xlsx')
data = {
    "Asset Name": ["Asset 1", "Asset 2"],
    "Month 1": [15, 30],
    "Month 2": [5, 35],
}

df = pd.DataFrame(data)

workbook = Workbook()
sheet = workbook.active

for row in dataframe_to_rows(df, index=False, header=True):
    sheet.append(row)

pandas_file = r"C:\Users\Lenovo\Desktop\pandas.xlsx"

workbook.save(pandas_file)
workbook.close()

# -------------------------
# Open pandas.xlsx
# -------------------------

xl = Dispatch("Excel.Application")
xl.Visible = True

wb = xl.Workbooks.Open(pandas_file)

# Keep Excel open until Enter is pressed
input("Press Enter to close Excel...")

wb.Close(SaveChanges=True)
xl.Quit()

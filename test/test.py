import pandas as pd

file = "50-sample-contacts.xlsx"
xls = pd.ExcelFile(file)
sheets = xls.sheet_names
if len(sheets) == 1:
    print(sheets[0], "is selected as contacts sheet since excel contains one sheet: ")
    cont = xls.parse(sheets[0])
print(cont.iloc[1, 1])
print(type(""))
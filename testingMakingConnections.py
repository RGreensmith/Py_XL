import openpyxl
import pandas as pd
from pathlib import Path

# Setting the path to the xlsx file
xlsx_file = Path('to_ignore','test_xlsx.xlsx')
#print(xlsx_file)

wb_obj = openpyxl.load_workbook(xlsx_file)
#print(wb_obj)

#Read active sheet
sheet = wb_obj.active
#print(sheet)

print(sheet["B2"].value)

#populate pandas df with content of test_xlsx
df = pd.DataFrame(columns=['A','B','C'],index=[1,2,3])
df.loc['1','A'] = sheet["B2"].value

print(df)
df.to_excel("to_ignore/df.xlsx")

### End ###
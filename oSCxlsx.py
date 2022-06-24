import openpyxl
from pathlib import Path
######## notes #########
# loading values only
#wb = openpyxl.load_workbook(filename, data_only=True)

#Read active sheet
#sheet = wb_obj.active

########################
def oSCxlsx(
    openSaveClose = ["open","save", "saveClose","close"],
    wbName = ["wb"],
    wbPath = ["to_ignore"],
    wbNm = ["BNG_no_macros_3_1.xlsx"]
    ):
    filePath = [wbPath,wbName]

    if openSaveClose == "open":

        # Setting the path to the xlsx file (BNG calculator v3.1) and Read specific sheet (in this case A-2 Site Habitat Creation)
        xlsx_file = Path(*wbPath,*wbNm)
        wb = openpyxl.load_workbook(xlsx_file)
        
        return wb

    elif openSaveClose == "save":
        wb.save(filePath)

    elif openSaveClose == "saveClose":
        wb.save(filePath)
        wb.close(wb)

    elif openSaveClose == "close":
        wb.close(wb)
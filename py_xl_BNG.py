import openpyxl
from sqlalchemy import true
from oSCxlsx import oSCxlsx
import pandas as pd
from pathlib import Path

def py_xl_BNG(habTypes):
    """
    For controlling inputting habitat types into BNG calculator (V3.1)
    (Site Habitat Creation) to calculate Biodiversity Units for 1 Ha of habitat.


    Habitat condition is defaulted to 'Good' (if not NA)

    Strategic significance is defaulted to 'Location ecologically desirable but not in
    local strategy'

    'Habitat created in advance/years' and 'Delay in starting habitat creation/years' are
    both defaulted to '0'

    Returns:
    pandas dataframe of habitat types and their corresponding Biodiversity Units
    and saves to specified location.
    """
    # Create  empty df for results 
    results = pd.DataFrame()

    # Cell references for BNG calc spreadsheet
    cellRefs = pd.read_csv('to_ignore/habitat_creation_cell_references.csv',header=0)
    bioUnit = "Y11"
    print(habTypes.to_string())

    xlsx_file = Path('to_ignore',"BNG_no_macros_3_1.xlsx")
    wb = openpyxl.load_workbook(xlsx_file)
    sheet = wb.get_sheet_by_name('A-2 Site Habitat Creation')

    # Loop through habitats to calculate Biodiversity Units
    
    for habIndex,eachHab in habTypes.iterrows():
   
        print(eachHab)
        
        # Populate BNG calculator cells
        for index,cellRef in enumerate(cellRefs.loc[:,"cellRef"]):
            sheet[cellRef+str(11+habIndex)].value = eachHab[index]
            
    wb.save("to_ignore/BNG_no_macros_3_1.xlsx")
    wb.close()

### End ###
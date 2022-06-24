import openpyxl
import pandas as pd
from pathlib import Path

def py_xl_BNG(habTypes,save_path):
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
    results = pd.DataFrame(columns=['A','B','C'],index=[1,2,3])

    #Cell referencers
    cellRefs = pd.read_csv('to_ignore/habitat_creation_cell_references') # cell referencers for BNG calc spreadsheet
    bioUnit = "Y11"

    # Loop through habitats to calculate Biodiversity Units
    for eachHab in habTypes.iter_rows():
        wb = oSCxlsx('open')
        sheet = wb.get_sheet_by_name('A-2 Site Habitat Creation')

        # Populate BNG calculator cells
        for eachCell in len(cellRefs):
            sheet[cellRefs[3,eachCell]].value = habTypes[eachHab,eachCell]

        oSCxlsx('saveClose')

        #populate results df with content of test_xlsx
        wb = oSCxlsx('open')
        sheet = wb.get_sheet_by_name('A-2 Site Habitat Creation')
        results.loc['1','A'] = sheet[bioUnit].value
        print(habTypes[eachHab,2]," ",sheet[bioUnit].value)


    #export xlsx
    results.to_excel(save_path,"habitat_Bio_Units.xlsx")
    print(results)

    return results
### End ###
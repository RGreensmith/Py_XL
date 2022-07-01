for habIndex,eachHab in habTypes.iterrows():
   
        print(eachHab)
        
        # Populate BNG calculator cells
        for index,cellRef in enumerate(cellRefs.loc[:,"cellRef"]):
            sheet[cellRef+str(11+habIndex)].value = eachHab[index]
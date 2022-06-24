import pandas as pd

habTypes = pd.read_csv('to_ignore/habitat_types') # habitat lists to loop through
save_path = 'to_ignore'

py_xl_BNG(habTypes,save_path)

### End ###
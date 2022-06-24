import pandas as pd
from sqlalchemy import true
from py_xl_BNG import py_xl_BNG
habTypes = pd.read_csv('to_ignore/habitat_types.csv',header=0) # habitat lists to loop through

py_xl_BNG(habTypes)

### End ###
# -*- coding: utf-8 -*-
"""
@author: ShivMakh
"""


import pandas as pd 
import numpy as np

#dummy data
df = pd.DataFrame({'Header1': ['val1','val2','val3'],
                   'Header2': ['val1','val2','val3']})

# now lets try and reset the column names, by shifting the Header values into the first row
# without overwriting any data

## using iloc will get incorrect results
df.index+=1 # shift index up (to try and move all the data "down" one row, making space for column headers)
df.iloc[0] = df.columns #now place column names in 0th index position
df.columns = [x for x in range(len(df.columns))] #replace column names

#we see here we overwrote val1 row and replaced with the column headers, rather than shift everything down
#this is because the df.iloc[0] just replaces the first row, as iloc is positional

#now lets do this using loc
#dummy data
df = pd.DataFrame({'Header1': ['val1','val2','val3'],
                   'Header2': ['val1','val2','val3']})

df.index +=1
df.loc[0] = df.columns
df.columns = [x for x in range(len(df.columns))]




# this is a way of demonstrating 

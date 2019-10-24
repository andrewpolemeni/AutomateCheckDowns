import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile



excelFile = input("Enter the path to the file name: ")
df1 = pd.read_excel(excelFile, sheet_name='Sheet1', usecols="A, M, V, P, AI, AJ, AR") #df = dataframe variable and import file here
df1.dropna(inplace=True) #drop cells with NaN
df1.drop(df1.loc[df1['Grade']=='F'].index, inplace=True) # Delete grades that are equal to F


print(df1)
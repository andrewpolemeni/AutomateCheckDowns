#==========================================================================================================================
# Programmer: Andrew Polemeni
# Organization: Daytona State College
# Date Written: October 24, 2019
# Program Purpose: To automate course check downs for student progress.
#==========================================================================================================================
# Import statements
import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
#==========================================================================================================================
# will add GUI later using QT


#==========================================================================================================================
#Grab the first excel file which will be the Query Report
queryReport = input("Enter the file name based on the path of the query report: ") #Input the file name
df1 = pd.read_excel(queryReport, sheet_name='Sheet1', usecols="A, M, V, P, AI, AJ, AR") #df = dataframe variable and import file here
df1.dropna(inplace=True) #drop cells with NaN
df1.drop(df1.loc[df1['Grade']=='F'].index, inplace=True) # Delete grades that are equal to F

#==========================================================================================================================
#Grab the second excel file which is the check down list
#checkDown = input("Enter the file name based on the path of the check down: ")
#df2 = pd.read_excel(checkDown, sheet_name="Sheet1")

print(df1)

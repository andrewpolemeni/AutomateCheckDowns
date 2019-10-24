#==========================================================================================================================
# Programmer: Andrew Polemeni
# Organization: Daytona State College
# Date Written: October 24, 2019
# Program Purpose: To automate course check downs for student progress.
#==========================================================================================================================
# Import statements
import pandas as pd
import openpyxl
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
df1.to_excel('queryReport.xlsx', index = None, header=True) #output to CSV
print(df1)
#==========================================================================================================================
# Open the query report file
wb1 = openpyxl.load_workbook('queryReport.xlsx')
ws1 = wb1.active
# File to be pasted into
wb2 = openpyxl.load_workbook('checkDown.xlsx')
ws2 = wb2.active
#==========================================================================================================================



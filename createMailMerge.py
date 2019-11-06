#============================================================================================================================
# Generate Mail Merge File from Data frame
#============================================================================================================================
import pandas as pd
import numpy as np
import sys, os
import openpyxl

#============================================================================================================================
# 1) OS walk the directory to find path names.
#============================================================================================================================
# path variable will change based on where your student check down files are.
targetdir = "ATC_STUDENT_CHECKDOWNS/"
data = list()
for root, dirs, files in os.walk(targetdir):
    for filename in files:
        nm, ext = os.path.splitext(filename)
        if ext.lower().endswith(('.xlsx')):
            fullpath = os.path.join(os.path.abspath(root), filename)
            data.append((filename, fullpath))
df2 = pd.DataFrame(data, columns=['Filename', 'Fullpath'])
print(df2)
            
# create a data frame for the student report array

# or you can do it this way with multiple paths
# https://stackoverflow.com/questions/7588620/os-walk-multiple-directories-at-once

#============================================================================================================================
# 2) create data frame from query file.
#============================================================================================================================
def getMailMergeData(query_filename):
    df1 = pd.read_excel(query_filename, sheet_name='sheet1', usecols="A, B, C, D, E") #df = dataframe variable and import file here
    df1.drop_duplicates(subset ="ID", inplace = True)
    #print(df1)


#def createMailMergeFile():


df1 = getMailMergeData("Query.xlsx")

#def GenerateMailMerge():

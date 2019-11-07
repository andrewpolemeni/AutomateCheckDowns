#============================================================================================================================
# Generate Mail Merge File from Data frame
#============================================================================================================================
import pandas as pd
import numpy as np
import sys, os
import openpyxl
import time, fnmatch, shutil


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
df1 = pd.DataFrame(data, columns=['Filename', 'Fullpath'])
#print(df1)
            
# create a data frame for the student report array

# or you can do it this way with multiple paths
# https://stackoverflow.com/questions/7588620/os-walk-multiple-directories-at-once

#============================================================================================================================
# 2) create data frame from query file.
#============================================================================================================================
def getMailMergeData(query_filename):

    df2 = pd.read_excel(query_filename, sheet_name='sheet1', usecols="A, B, C, D, E") #df = dataframe variable and import file here
    df2.drop_duplicates(subset ="ID", inplace = True) # drop duplicates of students ID
    # df2["Fullpath"] = [""]*len(df2) # create a row for full path

    # create another dataframe to merge df1 into df2
    #df3 = df1.merge(df2, left_on="Filename", right_on="Campus Email", how="right")

    #for index, row in df2.iterrows():
    #    if df2['Acad Plan'].iloc[0] == 633100:
    #        df2.at[index, "Fullpath"] = "BSET_" + df2.loc[index]["First Name"] + "_" + df2.loc[index]["Last"] + ".xlsx"
    #   if df2['Acad Plan'].iloc[0] == 633400:
    #        df2.at[index, "Fullpath"] = "BSIT_" + df2.loc[index]["First Name"] + "_" + df2.loc[index]["Last"] + ".xlsx"
    #    if df2['Acad Plan'].iloc[0] == 633300:
    #        df2.at[index, "Fullpath"] = "BSEET_" + df2.loc[index]["First Name"] + "_" + df2.loc[index]["Last"] + ".xlsx"
    #print(df2)
    
    
    # create file to save with timestamp
    t = time.localtime()
    timestamp = time.strftime("%b-%d-%Y", t)
    # write dataframes to excel file for mail merge
    writer = pd.ExcelWriter("MailMerge_" + timestamp + ".xlsx", engine='xlsxwriter')
    df2.to_excel(writer, sheet_name="Sheet1", index=False)
    df1.to_excel(writer, sheet_name="Sheet1", index=False, startcol=5)
    writer.save()
#def createMailMergeFile():


df1 = getMailMergeData("Query.xlsx")

#def GenerateMailMerge():

#============================================================================================================================
# Generate Mail Merge File from Data frame
#============================================================================================================================
import pandas as pd
import numpy as np
import sys, os
import openpyxl
import time, fnmatch, shutil
#============================================================================================================================
# Mail Merge
#============================================================================================================================
def getMailMergeData(query_filename):
    absolutePath = os.path.abspath("ATC_STUDENT_CHECKDOWNS/") # create a varaible for the absolute path
    print(absolutePath)
    df2 = pd.read_excel(query_filename, sheet_name='sheet1', usecols="A, B, C, D, E") #df = dataframe variable and import file here
    df2.drop_duplicates(subset ="ID", inplace = True) # drop duplicates from the sheet

    df2["Filename"] = [""]*len(df2) # add blank dataframe filename

    for index, row in df2.iterrows(): # here we add the path to the file name for that student.
        pre = "BSET"
        if df2.loc[index]["Acad Plan"] == 633300:
            pre = "BSEET"
        if df2.loc[index]["Acad Plan"] == 633400:
            pre = "BSIT"
        df2.at[index, "Filename"] = absolutePath + '\\' + pre + "_" + df2.loc[index]["First Name"] + "_" + df2.loc[index]["Last"] + ".xlsx"
    
    
    
    t = time.localtime() # create file to save with timestamp 
    timestamp = time.strftime("%b-%d-%Y", t) 

    # write dataframes to excel file for mail merge
    writer = pd.ExcelWriter("MailMerge_" + timestamp + ".xlsx", engine='xlsxwriter')
    df2.to_excel(writer, sheet_name="Sheet1", index=False)
    # save the file and return df2
    writer.save()
    return df2


# Call the function to open the query file
df1 = getMailMergeData("Query.xlsx")


#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import os
import openpyxl


# In[2]:


def get_file_name(file_path):
    # Use os.path.basename to get the file name from the file path
    file_name = os.path.basename(file_path)
    return file_name


# In[3]:


def Name(f_name):
    Name = f_name.split("_")
    Name.pop()  # Remove the last part
    Name.append("Result")
    File_Name = "_".join(Name) + ".xlsx"
    Title = "_".join(Name)
    return File_Name, Title


# In[4]:


def Sheet_1(source_File,Target_File):
    ''' Comparing two Files Sourec and Target column by column'''
    
    #Loading Source File as df1
    df1 = pd.read_csv(source_File)
    #Loading Target File as df2
    df2 = pd.read_csv(Target_File)  

    
    #file name From given path
    f_name=get_file_name(source_File) # returns f_name
    Names = Name(f_name) # returns FileName and Title for result 
    File_Name = Names[0]
    Title = Names[1]
    
    #common_columns = df1.columns.intersection(df2.columns)
    
   
    #getting column names
    Headdings = df1.columns.intersection(df2.columns)#df1.columns
    if(len(Headdings)<1):
        df_empty = pd.DataFrame()
        df_empty["There are no common columns between two files "]=""
        return df_empty
    else: 
        subheaddings="Source,Target,Result".split(',')

        # Create a Empty Data Frame
        df_new = pd.DataFrame()
        length =len(Headdings)

        # Creating a temperory data Frame
        df_temp =pd.DataFrame()

        for i in range(length):
            df_temp["Source"] = df1.iloc[:,i]
            df_temp["Target"] = df2.iloc[:,i]
            df_temp["Result"] = (df1.iloc[:, i] == df2.iloc[:, i]).astype(str).apply(str.upper)
            df_temp["Result"] = ((df1.iloc[:, i].isnull() & df2.iloc[:, i].isnull()) | (df1.iloc[:, i] == df2.iloc[:, i])).astype(str).apply(str.upper)

            df_new=pd.concat([df_new, df_temp], axis=1,ignore_index=True)


        # Assigning multilevel  columns to df_new
        df_new.columns =pd.MultiIndex.from_product([[Title],Headdings,subheaddings])
        df_new.dropna(how='all', inplace=True)
    
    
    # Exporting df_new as excelfile
    #df_new.to_excel(File_Name,sheet_name='Sheet_1')
   
    return df_new
   


# In[5]:


import pandas as pd

def Sheet_2(source_File, Target_File):
    ''' To check whether the columns of both target and source are the same '''
    # Read the source and target files
    df1 = pd.read_csv(source_File)
    df2 = pd.read_csv(Target_File)
    
    # Get column headers from both dataframes
    Hs = list(df1.columns)
    Ht = list(df2.columns)
    
    # Create lists for headers and subheaders
    Headdings = "Source,Target".split(',')
    subheaddings = "Columns,Data_type".split(',')
    
    # Get data types for source columns
    Hsdt = []
    for col in Hs:
        if col in df1.columns:
            Hsdt.append(str(df1[col].dtype))
        else:
            Hsdt.append("Nan")
    
    # Get data types for target columns
    Htdt = []
    for col in Ht:
        if col in df2.columns:
            Htdt.append(str(df2[col].dtype))
        else:
            Htdt.append("Nan")
    
    # Balance the lengths of the lists by adding "Nan" where necessary
    if len(Hs) < len(Ht):
        for _ in range(len(Ht) - len(Hs)):
            Hs.append("Nan")
            Hsdt.append("Nan")
    elif len(Hs) > len(Ht):
        for _ in range(len(Hs) - len(Ht)):
            Ht.append("Nan")
            Htdt.append("Nan")
    
    # Create a DataFrame to hold the comparison
    s2 = pd.DataFrame({'SColumns': Hs,'Sdtype': Hsdt,'TColumns': Ht,'Tdtype': Htdt})
    
    # Set multi-index columns
    s2.columns = pd.MultiIndex.from_product([Headdings, subheaddings])
    
    return s2


# In[6]:


def Sheet_3(source_File,Target_File):
    '''to count count,mean,std,min,max,25%,50%'''
    #Loading Source File as df1
    df_1 = pd.read_csv(source_File)
    #Loading Target File as df2
    df_2 = pd.read_csv(Target_File)
    
    df1 = df_1.describe()
    df1_c = list(df1.columns)
    df1.columns =pd.MultiIndex.from_product([["Source"],df1_c])
    df2 = df_2.describe()
    df2_c = list(df2.columns)
    df2.columns =pd.MultiIndex.from_product([["Target"],df2_c])

    
    # Create a Empty Data Frame
    df_new = pd.DataFrame()
    
 
    df_new=pd.concat([df1, df2], axis=1)
    
   
    
    return df_new
    


# In[11]:


source_File = input("Enter the Name/Path of the Source File : ")
Target_File = input("Enter the Name/Path of the Target File : ")


# In[8]:


def Create_file(source_File,Target_File):
    df_S1= Sheet_1(source_File,Target_File)
    df_S2 = Sheet_2(source_File,Target_File)
    df_S3 = Sheet_3(source_File,Target_File)
    f_name=get_file_name(source_File) 
    Names = Name(f_name)
    File_Name = Names[0]
    with pd.ExcelWriter(File_Name, engine='openpyxl') as writer:
        # Write each DataFrame to a specific sheet
        df_S1.to_excel(writer, sheet_name='Sheet1')
        df_S2.to_excel(writer, sheet_name='Sheet2')
        df_S3.to_excel(writer, sheet_name='Sheet3')

    # Closing Statement
    print("\nThe source and target files have been successfully compared.\n ")
    


# In[12]:


Create_file(source_File,Target_File)


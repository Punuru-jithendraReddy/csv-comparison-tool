# CSV File Comparison Tool

## Overview

This tool allows users to compare two CSV files and generate a comprehensive Excel report with three sheets. The sheets detail the differences between the files column by column, compare the columns and their data types, and provide statistical summaries of the data in each file.

## Features

1. **Sheet 1**: Compares the source and target files column by column, showing the source values, target values, and whether they match.
2. **Sheet 2**: Compares the columns of both files, listing the column names and their data types.
3. **Sheet 3**: Provides descriptive statistics (count, mean, standard deviation, min, max, 25%, 50%, 75%) for both files.

## Requirements

- Python 3.x
- Pandas
- NumPy
- Openpyxl

## Installation

1. **Clone the repository:**
    ```bash
    git clone https://github.com/Punuru-jithendraReddy/csv-comparison-tool.git
    cd csv-comparison-tool
    ```

2. **Install the required Python packages:**
    ```bash
    pip install pandas numpy openpyxl
    ```

## Usage

1. **Run the script:**
    ```bash
    python compare_csv.py
    ```

2. **Enter the file paths when prompted:**
    ```text
    Enter the Name/Path of the Source File: path/to/source_file.csv
    Enter the Name/Path of the Target File: path/to/target_file.csv
    ```

3. **The program will generate an Excel file with the comparison results.**

## Code Explanation

### Functions

1. **get_file_name(file_path)**: Extracts the file name from the provided file path.
2. **Name(f_name)**: Processes the file name to generate a result file name and title.
3. **Sheet_1(source_File, Target_File)**: Compares the source and target files column by column.
4. **Sheet_2(source_File, Target_File)**: Checks whether the columns of both files are the same and compares their data types.
5. **Sheet_3(source_File, Target_File)**: Provides descriptive statistics for both files.
6. **Create_file(source_File, Target_File)**: Orchestrates the creation of the Excel file, combining the results from the other functions.

### Example

```python
import pandas as pd
import numpy as np
import os
import openpyxl

def get_file_name(file_path):
    # Use os.path.basename to get the file name from the file path
    file_name = os.path.basename(file_path)
    return file_name

def Name(f_name):
    Name = f_name.split("_")
    Name.pop()  # Remove the last part
    Name.append("Result")
    File_Name = "_".join(Name) + ".xlsx"
    Title = "_".join(Name)
    return File_Name, Title

def Sheet_1(source_File, Target_File):
    # Comparing two Files Source and Target column by column
    df1 = pd.read_csv(source_File)
    df2 = pd.read_csv(Target_File)

    f_name = get_file_name(source_File)
    Names = Name(f_name)
    File_Name = Names[0]
    Title = Names[1]

    Headdings = df1.columns.intersection(df2.columns)
    if len(Headdings) < 1:
        df_empty = pd.DataFrame()
        df_empty["There are no common columns between two files"] = ""
        return df_empty
    else:
        subheaddings = "Source,Target,Result".split(',')

        df_new = pd.DataFrame()
        length = len(Headdings)
        df_temp = pd.DataFrame()

        for i in range(length):
            df_temp["Source"] = df1.iloc[:, i]
            df_temp["Target"] = df2.iloc[:, i]
            df_temp["Result"] = (df1.iloc[:, i] == df2.iloc[:, i]).astype(str).apply(str.upper)
            df_temp["Result"] = ((df1.iloc[:, i].isnull() & df2.iloc[:, i].isnull()) | (df1.iloc[:, i] == df2.iloc[:, i])).astype(str).apply(str.upper)
            df_new = pd.concat([df_new, df_temp], axis=1, ignore_index=True)

        df_new.columns = pd.MultiIndex.from_product([[Title], Headdings, subheaddings])
        df_new.dropna(how='all', inplace=True)
    
    return df_new

def Sheet_2(source_File, Target_File):
    # Check whether the columns of both target and source are the same
    df1 = pd.read_csv(source_File)
    df2 = pd.read_csv(Target_File)
    
    Hs = list(df1.columns)
    Ht = list(df2.columns)
    
    Headdings = "Source,Target".split(',')
    subheaddings = "Columns,Data_type".split(',')
    
    Hsdt = [str(df1[col].dtype) if col in df1.columns else "Nan" for col in Hs]
    Htdt = [str(df2[col].dtype) if col in df2.columns else "Nan" for col in Ht]
    
    if len(Hs) < len(Ht):
        for _ in range(len(Ht) - len(Hs)):
            Hs.append("Nan")
            Hsdt.append("Nan")
    elif len(Hs) > len(Ht):
        for _ in range(len(Hs) - len(Ht)):
            Ht.append("Nan")
            Htdt.append("Nan")
    
    s2 = pd.DataFrame({'SColumns': Hs, 'Sdtype': Hsdt, 'TColumns': Ht, 'Tdtype': Htdt})
    s2.columns = pd.MultiIndex.from_product([Headdings, subheaddings])
    
    return s2

def Sheet_3(source_File, Target_File):
    # Count, mean, std, min, max, 25%, 50%
    df_1 = pd.read_csv(source_File)
    df_2 = pd.read_csv(Target_File)
    
    df1 = df_1.describe()
    df1.columns = pd.MultiIndex.from_product([["Source"], list(df1.columns)])
    df2 = df_2.describe()
    df2.columns = pd.MultiIndex.from_product([["Target"], list(df2.columns)])

    df_new = pd.concat([df1, df2], axis=1)
    
    return df_new

def Create_file(source_File, Target_File):
    df_S1 = Sheet_1(source_File, Target_File)
    df_S2 = Sheet_2(source_File, Target_File)
    df_S3 = Sheet_3(source_File, Target_File)
    
    f_name = get_file_name(source_File)
    Names = Name(f_name)
    File_Name = Names[0]
    
    with pd.ExcelWriter(File_Name, engine='openpyxl') as writer:
        df_S1.to_excel(writer, sheet_name='Sheet1')
        df_S2.to_excel(writer, sheet_name='Sheet2')
        df_S3.to_excel(writer, sheet_name='Sheet3')

    print("\nThe source and target files have been successfully compared.\n")

if __name__ == "__main__":
    source_File = input("Enter the Name/Path of the Source File: ")
    Target_File = input("Enter the Name/Path of the Target File: ")
    Create_file(source_File, Target_File)
```


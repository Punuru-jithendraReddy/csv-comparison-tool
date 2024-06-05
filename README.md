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

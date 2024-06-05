# CSV File Comparison Tool

## Overview

This tool is designed to simplify the process of comparing two CSV files and generating a comprehensive Excel report summarizing the differences between them. It provides users with detailed insights into the data consistency, column structure, and statistical characteristics of the files.

## Features

1. **Sheet 1 - Column Comparison**: This sheet compares the source and target files column by column, highlighting differences in values and indicating whether they match. Users can easily identify discrepancies and inconsistencies between corresponding columns in the two files.

2. **Sheet 2 - Column Metadata**: Here, users can compare the metadata of columns in both files, including their names and data types. This sheet helps ensure that the column structure remains consistent across the files, facilitating data integrity and compatibility.

3. **Sheet 3 - Statistical Summary**: This sheet provides a statistical summary of the data in each file, including count, mean, standard deviation, minimum, maximum, and percentile values. Users can gain valuable insights into the distribution and characteristics of the data, aiding in data analysis and decision-making processes.

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
    pip install pandas
    pip install numpy
    pip install openpyxl
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

3. **Review the generated Excel report:**  
   Once the script completes execution, you will find an Excel file named according to the source file with "_Result" appended. This file contains the comparison results across the three sheets, providing a comprehensive overview of the differences between the source and target files.

## Additional Notes

- This tool is developed to address common challenges in data comparison tasks, such as verifying data consistency, ensuring column alignment, and analyzing data distributions.
- It offers a user-friendly interface, allowing users to quickly identify and address discrepancies between files.
- The tool is highly customizable and can be easily extended to accommodate additional comparison criteria or data processing requirements.

## Contributing

Contributions are welcome! If you have any suggestions, feature requests, or bug reports, please feel free to open an issue or submit a pull request. Your feedback is valuable in improving the functionality and usability of the tool for the community.


## About the Author

This tool is developed and maintained by Punuru Jithendra Reddy, a passionate software engineer with expertise in data analysis and automation. You can find more of my projects and contributions on [GitHub](https://github.com/Punuru-jithendraReddy).

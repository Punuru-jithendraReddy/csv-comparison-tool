# CSV & Excel Comparison Tool

## 🔍 Overview

This tool is designed to **compare two CSV or Excel files** and generate a detailed Excel report summarizing their differences. It supports value-by-value comparison, column metadata inspection, statistical summaries, unique row detection, and identification of unique values in non-numeric columns.

## ✅ Key Features

### 📄 Sheet 1 – **Value Comparison**

* Compares matching columns **cell-by-cell** between the source and target files.
* Flags mismatches as `FALSE` and highlights them in **yellow** for easy review.

### 🧱 Sheet 2 – **Column Metadata**

* Displays column names and data types from both files.
* Highlights missing or mismatched columns with a **red background**.

### 📊 Sheet 3 – **Summary Stats**

* Provides **statistical summaries** (count, mean, std, min, max, percentiles) for numeric columns.
* Helps quickly analyze data distributions across files.

### 📌 Sheet 4 – **Row Differences**

* Performs a **row-level comparison** on all common columns.
* Flags rows unique to either source or target and those common to both.

### 🔠 Sheet 5 – **Unique Non-Numeric Values**

* Compares only **non-numeric columns**.
* Shows values **unique to each file**, grouped by column.
* Automatically **removes the third row** from this sheet after creation (custom rule).

## ⚙️ Requirements

* Python 3.x
* [pandas](https://pypi.org/project/pandas/)
* [numpy](https://pypi.org/project/numpy/)
* [openpyxl](https://pypi.org/project/openpyxl/)

## 💾 Installation

```bash
git clone https://github.com/Punuru-jithendraReddy/csv-comparison-tool.git
cd csv-comparison-tool
pip install -r requirements.txt
```

> Or install individually:

```bash
pip install pandas numpy openpyxl
```

## 🚀 Usage

Run the script using Python:

```bash
python compare_csv.py
```

When prompted, enter the required inputs:

```
Enter the header row number (e.g., 2 for row 2) in the source file: 2
Enter the header row number (e.g., 2 for row 2) in the target file: 2
```

Once completed, an output Excel file will be generated in the current directory, named:

```
<SourceFileName>_Result.xlsx
```

## 📘 Output Overview

| Sheet Name       | Description                                  |
| ---------------- | -------------------------------------------- |
| Value Comparison | Cell-by-cell comparison for shared columns   |
| Column Metadata  | Column name and data type comparison         |
| Summary Stats    | Descriptive statistics of numeric columns    |
| Row Differences  | Unique rows from both files                  |
| Unique Values    | Column-wise unique values (non-numeric only) |

## 🧠 Notes

* Supports both `.csv` and `.xlsx` files.
* Automatically deletes any existing result file to avoid overwrite prompts.
* Handles flexible header row positions via user input.
* Highlights mismatches and missing fields using color-coded styles.
* Skips entirely blank or perfectly matched columns and rows.

## 🤝 Contributing

Feel free to fork the repository and submit pull requests. Suggestions, bug fixes, and feature enhancements are welcome.

## 👨‍💻 About the Author

Developed and maintained by **Punuru Jithendra Reddy**, a data automation enthusiast.
Explore more projects and tools at [GitHub - @Punuru-jithendraReddy](https://github.com/Punuru-jithendraReddy)

---




---

# Excel File Comparison Tool (Tkinter GUI)

This Python project is a GUI-based tool for comparing two Excel files. It supports multiple comparison modes:

* Column names
* Row-by-row content
* Unique values in each column
* Summary statistics (for numerical data)

The tool produces an Excel file with individual sheets for each comparison type.

---

## Features

1. Column Comparison

   * Displays column names side-by-side.
   * Highlights mismatched column names in red.

2. Row-by-Row Comparison

   * Compares rows based on content.
   * Adds a `Status` column with: Match, Only in Source, or Only in Target.

3. Unique Value Comparison

   * Lists unique values from both files, separated into Only in Source and Only in Target.

4. Summary Statistics

   * Provides statistical comparison (mean, median, min, max, std) for numeric columns.

---

## How to Use

### Step 1: Clone or Download the Project

Download this project or clone it from GitHub:

```
git clone https://github.com/your-username/excel-comparison-tool.git
cd excel-comparison-tool
```

---

### Step 2: Install Required Libraries

Make sure Python 3.8 or above is installed.
Install the required libraries using pip:

```
pip install pandas openpyxl numpy ttkbootstrap
```

Or using a `requirements.txt` file:

```
pandas
openpyxl
numpy
ttkbootstrap
```

Then run:

```
pip install -r requirements.txt
```

---

### Step 3: Run the GUI Tool

```
python excel_comparator_gui.py
```

---

### Step 4: Use the GUI

* Select Source and Target Excel files (.xlsx)
* Enter output filename (e.g., comparison\_result.xlsx)
* Select checkboxes for sheets to generate
* Click "Compare Files"

The result Excel file will be saved in the same folder.

---

## Output Excel File

| Sheet Name              | Description                                                             |
| ----------------------- | ----------------------------------------------------------------------- |
| Column Comparison       | Side-by-side list of column names with mismatches highlighted           |
| Row-by-Row Comparison   | Entire content row-by-row with status flags                             |
| Unique Value Comparison | Unique values present in either file but not both, by column            |
| Stats Comparison        | Statistical metrics comparison for numeric columns (mean, median, etc.) |

---

## Built With

* Python 3.8+
* Tkinter / ttkbootstrap - GUI
* Pandas - Data handling
* Openpyxl - Excel writing
* NumPy - Statistics

---

## Notes

* Both files must be in `.xlsx` format
* Column headers are assumed to be in the first row
* Blank or missing values are handled gracefully

---

## License

This project is open-source and free to use for educational or professional purposes.

---



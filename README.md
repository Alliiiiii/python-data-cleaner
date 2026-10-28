# Python Data Cleaner

A cross-platform Python command-line tool for cleaning messy Excel and CSV files. It standardizes data, removes duplicate records, identifies potential data-quality issues, and generates cleaned files with a readable report.

Works on **Windows, Linux, and macOS**.

## Features

- **Header detection:** Finds the actual column header when files contain title rows.
- **Row cleanup:** Removes blank rows and recognized footer rows.
- **Column cleanup:** Removes empty columns.
- **Text normalization:** Trims extra spaces and standardizes common text values.
- **Email and phone normalization:** Standardizes email addresses and phone number formats.
- **Date and number parsing:** Converts recognized dates and numeric values into consistent formats.
- **Duplicate detection:** Identifies exact duplicates and records sharing configured duplicate keys.
- **Missing-value handling:** Retains the most complete duplicate record and fills available gaps.
- **Issue reporting:** Flags invalid email addresses and unreadable dates for manual review.
- **Multiple outputs:** Generates cleaned CSV and Excel files and a Markdown report.

The original input file is not modified.

## Requirements

- Python 3
- pandas
- openpyxl
- python-dateutil

No operating-system-specific Python code or paid service is required.

## Installation

### 1. Install Python

Download Python for your operating system:

- [Windows](https://www.python.org/downloads/windows/)
- [Linux](https://www.python.org/downloads/source/)
- [macOS](https://www.python.org/downloads/macos/)

Verify the installation in your terminal:

```bash
python --version
```

On some Linux distributions and macOS installations, use `python3` instead of `python`.

On Windows, enable the option to add Python to PATH during installation if available.

### 2. Download the project

Clone the repository using Git:

```bash
git clone https://github.com/Alliiiiii/python-data-cleaner.git
cd python-data-cleaner
```

Alternatively, download the ZIP archive from the GitHub repository and extract it.

### 3. Create a virtual environment

A virtual environment keeps the project's dependencies separate from other Python projects.

**Windows — PowerShell:**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

**Windows — Command Prompt:**

```bat
python -m venv .venv
.venv\Scripts\activate.bat
```

**Linux and macOS:**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

If your system uses `python` instead of `python3`, use that command.

### 4. Install dependencies

With the virtual environment activated, run:

```bash
python -m pip install pandas openpyxl python-dateutil
```

If `python` is not the Python command on your system, substitute `python3` where appropriate.

## Usage

Run commands from the project directory with the virtual environment activated.

### Generate sample files

```bash
python make_sample.py
```

This generates example Excel and CSV files in the `sample/` directory.

### Clean an Excel file

```bash
python datacleaner.py sample/messy_customers.xlsx
```

### Clean a CSV file

```bash
python datacleaner.py sample/messy_customers.csv
```

### Normalize Pakistani phone numbers

```bash
python datacleaner.py sample/messy_customers.xlsx --phone-prefix +92
```

### Configure duplicate matching

Use email as the duplicate key:

```bash
python datacleaner.py sample/messy_customers.xlsx --keys email
```

Use email or a combination of first and last name:

```bash
python datacleaner.py sample/messy_customers.xlsx --keys "email,first_name+last_name"
```

The column names must match the input file.

### Handle ambiguous dates

Dates are interpreted day-first by default where possible. To interpret ambiguous dates month-first:

```bash
python datacleaner.py sample/messy_customers.xlsx --monthfirst
```

### Select an Excel sheet

```bash
python datacleaner.py sample/messy_customers.xlsx --sheet Sales
```

### Choose an output directory

```bash
python datacleaner.py sample/messy_customers.xlsx --outdir results/
```

## Output

By default, generated files are written to the `output/` directory.

| File | Description |
|---|---|
| `messy_customers_clean.xlsx` | Excel workbook containing Clean Data, Removed Duplicates, Issues, and Cleaning Summary sheets |
| `messy_customers_clean.csv` | Cleaned records in CSV format |
| `messy_customers_report.md` | Summary of the cleaning operation |

Output filenames are based on the input filename.

## Run tests

Run the automated test suite from the project directory:

```bash
python -m unittest -v test_datacleaner
```

The current tests cover date parsing, number parsing, phone and email normalization, and end-to-end sample processing.

## Limitations

- Duplicate detection uses configured keys. Similar-looking names are not automatically identified as duplicates.
- Currency symbols are removed, but the tool assumes one currency per column.
- Ambiguous dates may be interpreted incorrectly. Confirm date conventions before using the output.
- Records listed in the Issues sheet require manual review.
- Review the results before using them for important decisions.

## Project structure

```text
python-data-cleaner/
├── datacleaner.py
├── make_sample.py
├── test_datacleaner.py
├── README.md
├── sample/
└── output/
```

## License

No license has been specified. Contact the repository owner for permission before redistributing or reusing this software.

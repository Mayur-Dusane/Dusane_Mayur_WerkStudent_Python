# Invoice Data Extraction Tool for WerkStudent_Python

This Python program extracts invoice data (e.g., dates, values, and tables) from PDF files and saves the results into Excel and CSV files. It uses libraries like pandas, tabula-py, openpyxl and JPype1.

## Features

- Extracts specific data fields from PDF invoices.
- Saves extracted data into Excel and CSV files.

## Requirements

- Python 3.9 or higher
- pandas
- tabula-py
- openpyxl
- JPype1

## Installation

1. **Clone the repository:**

   ```bash
   git clone https://github.com/Mayur-Dusane/Dusane_Mayur_WerkStudent_Python.git
   cd Dusane_Mayur_WerkStudent_Python
   ```

2. **Create and activate a virtual environment:**

   ```bash
   python -m venv myenv
   myenv\Scripts\activate  # Windows
   source myenv/bin/activate  # Linux/macOS
   ```

3. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

4. **Ensure Java is installed:**

   - Download and install Java from [Oracle](https://www.oracle.com/java/technologies/javase-downloads.html) or [OpenJDK](https://openjdk.org/).
   - Add Java to your system's PATH environment variable.

5. **Run the program:**

   ```bash
   python Interview_Problem.py
   ```

## Usage

1. Place your PDF files (`sample_invoice_1.pdf` and `sample_invoice_2.pdf`) in the same directory as the script or executable.

2. Run the program:

   ```bash
   python Interview_Problem.py
   ```

3. Output files:

   - `sheet_1.xlsx`: Summary of extracted data.
   - `sheet_2.xlsx`: Pivot table of extracted data.
   - `Invoice_data.csv`: Combined table data from both PDFs.

## Troubleshooting

- **Java Error:**
  Ensure Java is installed and its path is added to the `PATH` environment variable.

- **Empty Output Files:**
  Verify the structure of the input PDF matches the expected format.

- **Module Not Found Errors:**
  Ensure all dependencies are installed by running `pip install -r requirements.txt`.

## Acknowledgments

- `tabula-py` for PDF data extraction.
- `pandas` for data manipulation.
- `openpyxl` for Excel file creation.
- `JPype1` for enabling interaction with Java-based libraries.

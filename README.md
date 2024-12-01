# Class Delivery List Automation

This repository contains a Python script to automate the processing and printing of daily delivery lists for classes. The system efficiently handles 1,000+ records daily, generating filtered Excel files, Word documents, and PDFs for class deliveries.

## Features
- Automated Data Processing: Reads and processes large CSV files of student data.
- Class-Specific Filtering: Creates separate delivery lists for each class while skipping predefined categories.
- Customizable Labels: Generates printable labels with detailed class and student information.
- Batch File Creation: Produces Excel, Word, and PDF files for easy distribution.
- High Efficiency: Handles over 1,000 entries daily, formatted into batches for optimized printing.

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/irugallbandara/AP-ExcelToDocLabelGenerator.git
   cd physics-class-delivery-list
   ```

2. Install required Python packages:
   ```
   pip install -r requirements.txt
   - pandas
   - openpyxl
   - python-docx
   - pywin32

   ```

## Usage

1. Prepare Input:
   - Save your data as a CSV file.
   - Ensure the file contains columns like `Date - Time`, `Delivery Method`, `Paper Type`, `Class`, `Name`, etc.

2. Run the Script:
   Update the `csv_file_path` in the script with the path to your CSV file, then execute:
   ```
   python process_class_data.py
   ```

3. Output:
   - Filtered delivery lists are saved as Excel files.
   - Labels are generated as Word documents and converted to PDFs, ready for printing.

## Example CSV Structure

| Date - Time        | Delivery Method | Paper Type | Class     | Name       | Address           | District  | Phone      |
|--------------------|----------------|-----------|-----------|------------|-------------------|-----------|------------|
| 2024-11-29 21:00  | Delivery       | Paper 1   | Physics A | John Doe   | 123 Main Street   | Colombo   | 1234567890 |
| 2024-11-29 21:00  | Speed Post     | Paper 2   | Physics B | Jane Smith | 456 Elm Street    | Galle     | 9876543210 |

## Output Files

- **Excel Files**: Filtered delivery lists for each class.
- **PDF Files**: Printable labels for `Delivery` and `Speed Post` methods.

## Repository Details

- Use Case: Automates daily delivery list generation and printing for Physics classes.

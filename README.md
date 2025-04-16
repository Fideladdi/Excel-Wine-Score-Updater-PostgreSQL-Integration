# 🍷 Excel Wine Score Updater with PostgreSQL

This Python script enriches an Excel file of wine data by connecting to a PostgreSQL database and retrieving missing wine scores, reviewers, and tasting notes. It’s designed to automate the process of updating historical wine data files with detailed scoring information.


## 🔍 What It Does

- Loads an Excel file containing wine entries
- Iterates through each row
- Skips non-vintage entries (`NV`)
- Queries a PostgreSQL database using `wine_name` and `vintage`
- If a match is found:
  - Updates the `score`, `reviewer`, and `wine_notes` columns
  - Flags updated rows in column 17

## 📥 Input

- Excel file: `Scores Only_List 3.xlsx`
- Required columns:
  - Column 3 – Vintage
  - Column 4 – Wine Name
  - Columns 8–10 – (Score, Reviewer, Notes)

## 📤 Output

- Excel file:  
  `Scores and notes - Wine April  2025 TEST.xlsx`  
  (saved in the same directory or custom output path)



## 🛠 Requirements

Install the required Python packages:

pip install pandas sqlalchemy openpyxl psycopg2-binary

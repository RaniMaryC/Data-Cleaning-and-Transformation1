# Data-Cleaning-and-Transformation1
markdown# Excel Assignment 2 - Data Cleaning and Transformation

Cleaned and transformed raw product data in Microsoft Excel.
File: `Excel_Assignment_2.xlsx`

## Files / Sheets

| Sheet | What it contains |
| --- | --- |
| Product_DB2 | Raw data - 47 products |
| Task | 6 assignment tasks |
| Answer | Cleaned data with formulas |
| Product_DB24 | Cleaned table for Power Query |
| Copy for Power Query | Copy used for Power Query Remove Duplicates |

## Raw Data Columns

| Column | Example |
| --- | --- |
| Product ID | 28-JAN-US |
| Product Name | laptop, smartphone, headphones |
| Brand Name | Dell, Nike, Samsung |
| Price ($) | 1000, 80, 130 |
| Quantity | 30, 15, 40 |
| Category | Electronics, Fashion, Kitchen |

Product ID format is: `DD-MMM-CC`
Example: `28-JAN-US` = 28 Jan, Country US

## What I Did

### 1. Handling Missing Values

- **Missing Price:** Filled with Median Price
    - Formula: `=MEDIAN(I2:I35)`
    - Example: Headphones (07-JUN-UK), Camping Tent, Sunglasses had no price
- **Missing Category:** Filled with `Unknown`
    - Formula: `=IF(ISBLANK(Category),"Unknown",Category)`
    - Example: Backpack, Sneakers, Coffee Maker had no category

### 2. Correcting Inconsistent Data

- **Product Name:** `laptop` -> `Laptop`, `smartphone` -> `Smartphone`, `headphones` -> `Headphones`
    - Formula: `=PROPER(TRIM(Product Name))`
- **Category Typo:** `Electroni` -> `Electronics`
    - Formula: `=IF(Category="Electroni","Electronics",Category)`
    - Also fixed with Find and Replace

### 3. Removing Duplicates

- Used Power Query: `Data > From Table / Range > Remove Duplicates`
- Duplicate full rows found:
    - `21-AUG-CA - Laptop Bag - Samsonite` (2 times)
    - `16-APR-ES - headphones - Bose` (2 times)
- Removed duplicates and loaded cleaned data back to Excel

### 4. Splitting and Merging Data

- **Manufacturing Date from Product ID:**
    - `=LEFT(Product ID,6)` -> `28-JAN`
    - Then `=DATEVALUE(TRIM(Manufacturing Date)&"-2026")`
- **Country Code from Product ID:**
    - `=RIGHT(Product ID,2)` -> `US`
- **Product Brand - Merge Product Name + Brand Name:**
    - `=Product Name & " - " & Brand Name`
    - Example: `Laptop - Dell`, `Sneakers - Nike`

### 5. Number Formatting

- **Price to Currency:** Format Cells > Currency > `$1,000.00`
- **Manufacturing Date to DD-MM-YYYY:** Format Cells > Custom > `DD-MM-YYYY`
    - Example: `28-01-2026`

### 6. Conditional Formatting

- **Price column:** Data Bars
    - Highest price `$1,000.00` = Longest bar
    - Lowest price `$30.00` = Shortest bar
- **Category column:** Custom Rule
    - Formula: `=$P2="Electronics"`
    - Electronics cells highlighted in Light Blue

## Key Formulas Used

- `LEFT`, `RIGHT`, `DATEVALUE`, `TRIM`, `PROPER`, `IF`, `ISBLANK`, `MEDIAN`


## Tools Used

- Microsoft Excel
- Power Query
- Excel Formulas + Conditional Formatting

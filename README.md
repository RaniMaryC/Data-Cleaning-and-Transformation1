# Excel Capstone Project - E-commerce Sales Analysis

## 1. Project Overview
Analysis of 2,000 e-commerce orders to find which Category, Region, and Loyalty Tier drive the most sales, and to identify data quality issues.

**Total Sales:** ₹21,74,024
**Total Orders:** 2,000
**Average Order Value:** ₹1,087.01

## 2. Tools Used
- Microsoft Excel
- Pivot Tables
- VLOOKUP / XLOOKUP
- Descriptive Statistics (Analysis ToolPak)
- Charts - Bar, Column, Doughnut
- Data Cleaning & Validation

## 3. Files in This Project
| File / Sheet | Description |
| --- | --- |
| Raw Data | Original sales data - 2000 rows |
| Category Master | Category_ID to Category_Name mapping |
| Region Master | Region_ID to Region_Name mapping |
| Loyalty Master | Tier_ID to Tier_Name mapping |
| Final Sheet | Cleaned data after VLOOKUP - all masters joined |
| Pivot - Category | Sales by Category |
| Pivot - Region | Sales by Region |
| Pivot - Loyalty | Sales by Loyalty Tier |
| Descriptive Stats | Mean, Median, Mode, Std Dev for Total_Amount |
| Dashboard | 3 Charts in one page |
| Summary | Key insights & recommendations |

## 4. Data Cleaning Done
1.  Joined 3 master tables using VLOOKUP on Final Sheet
2.  Created Total_Amount = Quantity * Unit_Price
3.  Found Missing Values:
    - ₹46,059 (2.1%) sales as Unknown Category - Category_ID not in master
    - ₹17,513 (0.8%) sales as Unknown Tier - Tier_ID not in master
4.  Checked duplicates, blank rows, and negative values - none found

## 5. Key Insights

### Category
Sports is top at ₹5,42,003 (24.9%), followed by Electronics ₹5,05,875 (23.3%). Together they are 48% of sales. Beauty is lowest at ₹3,01,118 (13.9%).

### Region
North ₹6,43,565 (29.6%), South ₹6,34,885 (29.2%) and East ₹6,24,868 (28.7%) are almost equal. West is very low at ₹2,70,706 (12.5%) - 57% below others. Needs investigation.

### Loyalty Tier
Platinum is highest at ₹6,55,317 (30.1%). Bronze ₹5,23,700 (24.1%), Gold ₹5,18,192 (23.8%), Silver lowest at ₹4,59,302 (21.1%). Premium customers drive the most revenue.

## 6. Descriptive Statistics - Total_Amount
- Mean: ₹1,087.01
- Median: ₹868.50
- Mode: ₹571
- Standard Deviation: ₹818.88
- Min: ₹18 | Max: ₹3,908
- Skewness: 0.88

**Interpretation:** Mean > Median, so distribution is right-skewed. Most orders are small, few high-value orders pull the average up.

## 7. Data Visualization
1.  Bar Chart - Sales by Category
2.  Doughnut Chart - Sales by Region (%)
3.  Column Chart - Sales by Loyalty Tier

All charts are in the `Dashboard` sheet.

## 8. Recommendations
1.  Focus marketing on Sports & Electronics - top 2 categories.
2.  Retain Platinum customers with exclusive offers - they give 30% revenue.
3.  Investigate West Region - low sales may be due to logistics, pricing, or missing data.
4.  Fix Master Data - add missing Category_ID and Tier_ID to reduce Unknown sales.
5.  Promote Beauty category with bundles to increase its 13.9% share.

## 9. How to Use This File
1.  Open `Excel Capstone Project - E-commerce Analysis.xlsx`
2.  Start from `Summary` sheet for insights
3.  Go to `Dashboard` for charts
4.  Check `Final Sheet` for cleaned data
5.  Check `Descriptive Stats` for statistical summary

## 10. Author
Capstone Project - E-commerce Sales Analysis using Excel

# E-Commerce Sales Data Cleaning & Reporting Project

## 📌 Project Overview
This project demonstrates an automated data cleaning and reporting pipeline designed for e-commerce and retail small businesses. It takes unformatted, duplicate-ridden, and messy CSV exports from point-of-sale (POS) systems or online store databases, cleans the data, standardizes formats, and generates an executive summary report in Excel.

---

## 🎯 Business Problem
Small online stores and retail agencies often encounter messy transaction exports:
1. **Duplicate Transactions:** Repeated order records skew sales figures and inventory count.
2. **Inconsistent Naming & Casing:** Unstandardized text formats (e.g., `rahul sharma`, `AMIT VERMA`, mixed city cases `PUNE` vs `Pune`) prevent reliable filtering.
3. **Leading / Trailing Whitespaces:** Extra spaces corrupt string matching and lookup functions (`VLOOKUP` / `XLOOKUP`).
4. **Invalid Number Formats:** Currency symbols (`$`, `₹`) embedded inside numeric values prevent mathematical aggregation.
5. **Irregular Date Formats:** Mixed formats (ISO, DD/MM/YYYY, Month Day Year) break time-series reporting.

---

## 🛠️ Solution & Features
The pipeline processes the data using Python (`pandas` & `openpyxl`) and outputs an audit-ready multi-tab Excel workbook:

- **Duplicate Removal:** Automatically isolates distinct orders based on unique `Order_ID`.
- **String Cleaning & Title Casing:** Strips extra whitespace and standardizes customer and city names to Proper Case.
- **Price Sanitization:** Strips non-numeric characters and converts amounts to floating-point numbers.
- **Date Standardization:** Standardizes all dates into `YYYY-MM-DD` ISO format.
- **Metric Computation:** Automatically calculates `Total_Amount = Quantity * Unit_Price`.
- **Multi-Tab Excel Output:**
  - `Raw_Data`: Untouched client source data for audit compliance.
  - `Clean_Data`: Cleaned, structured, and calculated transaction records.
  - `Executive_Summary`: High-level business KPIs, product-level sales breakdown, and city-level revenue distribution.

---

## 📊 Summary KPIs Generated
- **Total Cleaned Orders:** 10
- **Total Units Sold:** 18
- **Total Gross Revenue:** ₹10,950
- **Average Order Value (AOV):** ₹1,095
- **Top Revenue Products:** Laptop Stand, Wireless Mouse, Keyboard

---

## 📁 Repository Structure
```text
├── raw_sales_data.csv          # Original messy sample data
├── clean_data.py               # Automation script
├── sales_report_cleaned.xlsx   # Cleaned multi-tab Excel workbook
└── README.md                   # Project documentation
```

---

## 🚀 How to Run Locally
1. Clone this repository or copy the project files.
2. Install required dependencies:
   ```bash
   pip install pandas openpyxl
   ```
3. Run the cleaning script:
   ```bash
   python clean_data.py
   ```
4. Open `sales_report_cleaned.xlsx` to review the results.

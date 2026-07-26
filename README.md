# Executive Sales Report Automation

## Overview

This project automates the generation of an **Employee Status Report** by consolidating data from multiple business sources, including Targets, Sales, Human Resources Master Data, Customer Master, Billing, Booking Order, and Organizational Hierarchy.

## Python data processing pipeline:
The script calculates Year-to-Date (YTD) targets and sales, validates employees, computes achievement percentages, maps organizational information, and produces a formatted Excel report.

---

## Features

- Calculates **Financial Year YTD Target** (April–March).
- Automatically **prorates the current month's target** based on the current date.
- Excludes vacant positions.
- Validates employees using HRMS records.
- Retrieves:
  - Employee Email
  - Reporting Manager
  - State
  - Sales Organization
- Calculates:
  - Sales YTD
  - Achievement %
  - Pending COB
  - Outstanding (<90 Days)
  - Outstanding (>90 Days)
- Maps Customer Names.
- Generates a professionally formatted Excel report.

---

# Project Structure

```text
project/
│
├── Employee_Status.py
├── Inputs/
│   ├── Target.xlsx
│   ├── MiniHR.xlsx
│   ├── HRMS.xlsx
│   ├── Customer_Master.xlsx
│   ├── Sales_GT.xlsx
│   ├── Pending_COB.xlsx
│   ├── Outstanding.xlsx
│   └── Rollup.xlsx
│
├── Output/
│   └── Employee_Status_Report.xlsx
│
└── README.md
```

---

# Input Files

The script expects the following input files.

| Variable | Sample File | Purpose |
|-----------|-------------|---------|
| `Input1` | `Target.xlsx` | Employee Target Data |
| `Input2` | `MiniHR.xlsx` | Employee Email Details |
| `Input3` | `HRMS.xlsx` | Employee Validation & State |
| `Input4` | `Customer_Master.xlsx` | Customer Information |
| `Input5` | `Sales_GT.xlsx` | Sales Transactions |
| `Input6` | `Pending_COB.xlsx` | Pending Collection on Billing |
| `Input7` | `Outstanding.xlsx` | Outstanding Ageing |
| `Input8` | `Rollup.xlsx` | Reporting Hierarchy |

---

# Output

The script generates the following Excel file:

```
Employee_Status_Report.xlsx
```

Worksheet:

```
Employee_status
```

---

# Required Libraries

Install the required packages:

```bash
pip install pandas
pip install numpy
pip install openpyxl
pip install XlsxWriter
```

---

# Workflow

## Step 1 – Target YTD

- Load Target file.
- Convert Date column to datetime.
- Remove vacant employees.
- Determine the Financial Year start (April).
- Filter records up to yesterday.
- Prorate the current month's target.
- Aggregate YTD Target by:
  - Position ID
  - Employee
  - Customer

---

## Step 2 – Employee Validation & Lookups

- Map Sales Organization.
- Validate employees against HRMS.
- Remove invalid employees.
- Retrieve:
  - Official Email
  - Reporting Manager

---

## Step 3 – Sales YTD

Aggregate sales by:

- Position ID
- Customer ID

Merge sales with target data.

---

## Step 4 – Achievement %

Calculate:

```text
Achievement % = (Sales YTD / Target YTD) × 100
```

If Target YTD equals zero, Achievement % is set to 0.

---

## Step 5 – Pending COB

Aggregate Pending COB by Customer ID and merge with the main dataset.

---

## Step 6 – Outstanding (<90 Days)

Sum the following ageing buckets:

- 1–15 Days
- 16–30 Days
- 31–45 Days
- 46–60 Days
- 61–75 Days
- 76–90 Days

Aggregate by:

- Customer ID
- Sales Organization

---

## Step 7 – Outstanding (>90 Days)

Sum the following ageing buckets:

- 91–120 Days
- 121–150 Days
- 151–180 Days
- 181–270 Days
- 271–360 Days
- 360+ Days

Aggregate by:

- Customer ID
- Sales Organization

---

## Step 8 – Employee State

Retrieve employee work location state from HRMS.

---

## Step 9 – Customer Mapping

Retrieve Customer Name using Customer ID.

---

## Step 10 – Final Dataset

The generated report contains the following columns:

| Column |
|----------|
| Sr. no. |
| Position ID |
| Employee Code |
| Employee Name |
| Org |
| State |
| Reporting Manager |
| Customer Name |
| Customer ID |
| Target YTD |
| Sales YTD |
| Achievement % |
| Pending COB |
| Outstanding <90 Days |
| Outstanding >90 Days |
| Employee Email |

---

# Excel Formatting

The output workbook includes:

- Thousand separator formatting for numeric columns.
- Rounded integer values.
- Fixed column widths.
- Right-aligned numeric fields.
- Ready-to-share report formatting.

---

# Business Logic

## Financial Year

```
April → March
```

## Current Month Target

The current month's target is prorated as:

```text
(Target ÷ Days in Month) × Elapsed Days
```

## Employee Validation

Only employees present in the HRMS master are included in the final report.

## Vacant Positions

Employees whose **Employee Code** starts with:

```
VACANT
```

are excluded from processing.

---

# Running the Script

```python
python Employee_Status.py
```

Upon successful execution:

```
Employee Status file created successfully.
```

---

# Technologies Used

- Python
- Pandas
- NumPy
- XlsxWriter
- OpenPyXL
- Google Colab

---

# Notes

- Ensure all input files follow the expected column structure.
- Sheet names must match those referenced in the script.
- Missing numeric values are automatically replaced with zero.
- Numeric columns are rounded before export.
- The report is generated up to **yesterday's date** for accurate YTD calculations.

---

# Future Enhancements

- Configuration file for input/output paths.
- Command-line execution support.
- Detailed logging.
- Automated email distribution.
- Power BI integration.
- Scheduled execution using Cron or Windows Task Scheduler.
- Enhanced exception handling and validation reports.

---

## Author

Developed using **Python**, **Pandas**, and **XlsxWriter** to automate Employee Performance and Status reporting from multiple enterprise data sources.

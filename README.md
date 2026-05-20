# Cohort Counter

A lightweight automation tool that processes weekly Excel shipping reports and summarizes order status counts with percentages — no manual filtering needed.

## What it does

- Auto-detects the latest `.xlsx` file in the folder
- Scans all tabs named `W#` (e.g. W18, W19, W20, W21)
- For each week, counts orders in column I by status:
  - **Total Orders**
  - **Delivered**
  - **Awaiting Pickup**
  - **Out for Delivery**
- Prints a clean summary table in the terminal
- Writes a **Cohorts** tab to the Excel file with counts + percentages

## Setup

Install dependencies:
```bash
pip3 install pandas openpyxl
```

## Usage

1. Drop your `.xlsx` report into this folder
2. Double-click **`Run Order Counter.command`**

Results print instantly in the terminal and a Cohorts tab is added to your file.

## Output example

```
WEEK       TOTAL    DELIVERED    AWAITING PICKUP    OUT FOR DELIVERY
-----------------------------------------------------------------------
W21         2497      287 (11.5%)       100 ( 4.0%)        50 ( 2.0%)
W20         9566      812 ( 8.5%)       430 ( 4.5%)       210 ( 2.2%)
```

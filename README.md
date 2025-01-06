# Consolidated Excel Filtering Script

This Python script processes multiple Excel files containing transaction data, filters transactions based on specific time thresholds (`1.8 to <2 seconds` and `>=2 seconds`), and consolidates the results into a **single CSV file** with clear headings and separate tables for each report.

---

## **Features**
- Filters transactions from Excel files based on specific conditions.
- Outputs results for all input files into **one consolidated CSV file**.
- Adds headings and separates tables for better readability.
- Handles non-numeric or missing data gracefully.

---

## **Approach**

1. **Input Excel Files**:
   - The script reads all `.xlsx` files from a specified `input_reports` folder.

2. **Processing**:
   - Ensures the `time(90%)` column contains numeric values.
   - Filters the transactions into two categories:
     - Transactions with **`1.8 to <2 seconds`**.
     - Transactions with **`>=2 seconds`**.

3. **Output**:
   - Saves all results into one **consolidated CSV file**.
   - For each input file:
     - Includes a heading like `Report 1: Filtered Data for <filename>`.
     - Contains two separate tables for the above categories.
     - Blank rows separate different tables and reports for readability.

---

## **Requirements**
- Python 3.x
- Libraries: `pandas`, `os`

Install the required Python libraries using:
```bash
pip install pandas

<img width="379" alt="su1" src="https://github.com/user-attachments/assets/444e1c48-b769-45ae-ad3a-97f5f55b5c77" />






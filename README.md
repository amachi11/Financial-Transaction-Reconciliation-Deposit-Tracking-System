Financial Transaction Reconciliation & Deposit Tracking System (Excel)

This project is a finance operations simulation workbook designed to demonstrate practical experience in transaction processing, reconciliation, vendor management, and collections tracking. It mirrors real-world financial operations responsibilities such as processing deposits, reviewing credit card payments, maintaining vendor records, and identifying discrepancies.

The system is built using structured Excel tables, advanced formulas, lookup logic, and automated KPI summaries to improve accuracy and reduce manual review time.

Workbook Structure
1️⃣ Bank_Deposits

Tracks daily deposit activity and reconciliation status.
Includes:

Variance calculations between expected and actual amounts

Match / Investigate logic using IF statements

Summary calculations using COUNTIFS and SUMIFS

2️⃣ Direct_Deposits

Monitors payment processing status and timing.
Includes:

Payment status tracking

Date-based calculations

Automated summary counts

3️⃣ CreditCard_Payments

Reconciles card processor records with internal records.
Includes:

Variance calculations

Exception detection (variance ≠ 0)

Automated totals and exception counts

4️⃣ Vendor_Maintenance

Maintains vendor master records.
Includes:

Structured table formatting

Central reference table for lookups

Supports transaction sheet validation

5️⃣ Collections

Tracks outstanding balances and aging categories.
Includes:

Days Outstanding calculation using TODAY()

Aging bucket categorization using IFS

Summary totals by aging group

6️⃣ Pivot_Summaries (Formula-Based Pivot Logic)

This sheet acts as a Pivot-style dashboard using structured references and aggregation formulas instead of manual pivot tables.

Includes:

Deposit reconciliation summary

Credit card variance summary

Collections aging summary

Uses:

SUMIFS

COUNTIFS

Structured table references (tblBankDeposits, tblCardPayments, tblCollections)

7️⃣ Vendor Lookup Integration

Transaction sheets automatically pull vendor names using:

XLOOKUP (or VLOOKUP alternative)

IFERROR for clean outputs

This ensures data consistency and reduces manual cross-referencing.

Advanced Excel Features Used

Structured Tables

SUMIFS / COUNTIFS (conditional aggregation)

IF / IFS logic

IFERROR (error handling)

TODAY() (dynamic date aging)

XLOOKUP / VLOOKUP (data validation and enrichment)

Conditional formatting for exception highlighting

Optional VBA Macros (Included in Sheet)

Two ready-to-use macros are provided for:

Clearing transaction input rows

Refreshing all data

These can be copied into the VBA editor if the file is saved as .xlsm.

Skills Demonstrated

Financial transaction reconciliation

Accounts receivable and payable tracking

Exception reporting and variance analysis

Vendor data management

Collections aging analysis

Dashboard-style KPI reporting

Process automation using advanced Excel formulas

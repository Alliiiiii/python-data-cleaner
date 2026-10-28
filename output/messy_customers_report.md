# Data Cleaning Report - messy_customers.xlsx

_Generated 2026-10-28 10:06 by datacleaner v1.0_

## Summary

- Data rows in source: **26**
- Blank / footer rows removed: **4**
- Duplicates removed: **7** (4 exact, 3 by key)
- Gaps filled from duplicate records: **1**
- Clean rows delivered: **15**
- Cells flagged for review: **2**

## Columns

| Column | Type | Cells changed | Missing | Example |
|---|---|---|---|---|
| customer_name | title | 7 | 0 | 'ayesha khan' -> 'Ayesha Khan'; '  Bilal   Ahmed ' -> 'Bilal Ahmed' |
| email | email | 4 | 0 | 'Ayesha.Khan@gmail.com' -> 'ayesha.khan@gmail.com'; 'ayesha.khan@gmail.com ' -> 'ayesha.khan@gmail.com' |
| phone | phone | 19 | 0 | '0300-1234567' -> '+923001234567'; '0321 7654321' -> '+923217654321' |
| signup_date | date | 15 | 0 | '13/05/2025' -> '2025-05-13'; '5 Jan 2025' -> '2025-01-05' |
| city | title | 4 | 0 | 'karachi' -> 'Karachi'; 'LAHORE' -> 'Lahore' |
| country | country | 5 | 0 | 'PK' -> 'Pakistan'; 'pakistan' -> 'Pakistan' |
| order_amount_rs | number | 15 | 2 | 'Rs 12,500' -> '12500'; '8,000' -> '8000' |
| status | category | 4 | 0 | 'active' -> 'Active'; 'ACTIVE' -> 'Active' |

## Duplicates removed

| Source row | Reason | Kept row |
|---|---|---|
| 22 | Exact duplicate | 5 |
| 24 | Exact duplicate | 8 |
| 26 | Exact duplicate | 12 |
| 28 | Exact duplicate | 10 |
| 23 | Duplicate email | 6 |
| 25 | Duplicate email | 7 |
| 27 | Duplicate email | 15 |

## Needs a human look

| Source row | Column | Value | Issue |
|---|---|---|---|
| 20 | email | kamran.abbasi.gmail.com | Invalid email format |
| 18 | signup_date | not sure | Could not read as a date |

## Assumptions

- Dates read day-first unless `--monthfirst` was used.
- Currency symbols were stripped from number columns.
- Duplicate rules: exact row match; email.

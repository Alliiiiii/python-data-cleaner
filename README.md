# datacleaner

Cleans messy Excel/CSV files, removes duplicate records, and produces a report the client can read in two minutes. Your original file is never changed.

## What it does

| Step | Example |
|---|---|
| Finds the real header | Skips title lines like "Customer Export - Oct 2026" |
| Removes junk rows | Blank rows, "Total records: 22" footers, empty columns |
| Fixes text | Extra spaces, `ayesha khan` -> `Ayesha Khan`, `PK` -> `Pakistan`, `active/ACTIVE` -> `Active` |
| Standardises emails and phones | `0300-1234567` -> `+923001234567` (with `--phone-prefix +92`) |
| Standardises dates and numbers | `13/05/2025` -> `2025-05-13`, `Rs. 12,500` -> `12500` |
| Removes duplicates | Exact copies and same-email copies; keeps the most complete record and fills its gaps |
| Flags problems | Bad emails, unreadable dates (highlighted yellow, listed in `Issues`) |

Every source row is accounted for: `source rows - blank - duplicates = clean rows` (the summary sheet checks this and shows 0).

## Setup on Arch Linux

```bash
sudo pacman -S python python-pandas python-openpyxl python-dateutil
# or, without touching system packages:
python -m venv .venv && source .venv/bin/activate
pip install pandas openpyxl python-dateutil
```

All free. No internet needed after install.

## Try it

```bash
python make_sample.py                                   # creates sample/messy_customers.xlsx + .csv
python datacleaner.py sample/messy_customers.xlsx --phone-prefix +92
python -m unittest test_datacleaner -v                  # 4 tests should pass
```

Output goes to `./output/`:

- `messy_customers_clean.xlsx` - sheets: Clean Data, Removed Duplicates, Issues, Cleaning Summary
- `messy_customers_clean.csv` - clean rows only
- `messy_customers_report.md` - plain-text report

## Options

| Option | Meaning |
|---|---|
| `--keys "email,first_name+last_name"` | Duplicate rules. Comma = separate rules, `+` = combined columns. Default: the email column |
| `--phone-prefix +92` | Turns local numbers (03xx) into international format |
| `--monthfirst` | Read 03/04/2025 as March 4 (default is day first) |
| `--sheet Sales` | Excel sheet name or number |
| `--outdir folder/` | Where results go |

## Known limits (tell clients honestly)

- Duplicates are found by exact key (e.g. same email), not by "similar-looking names". Near-matches are not merged automatically.
- Currency symbols are removed; it assumes one currency per column.
- Day/month order is a guess: confirm it with the client before delivering.
- Always send the client the `Issues` sheet. Those rows need a human decision.
- Never process files containing data you are not allowed to hold. Ask the client to remove sensitive columns (CNIC, bank details) first, or work on a copy they sanitise.

## Your first-job checklist

1. Client sends a file. Run the tool on a copy.
2. Open `Cleaning Summary` and `Issues`. Check 5-10 rows by eye.
3. Send: clean file + summary + short note listing the assumptions.
4. Ask for payment per the agreed milestone (see below).

## Price (estimate only, not verified market data)

Start low to get your first reviews, then raise it:

- Small file (under ~1,000 rows, one sheet): about PKR 1,500-3,000
- Medium file or several sheets: about PKR 3,000-6,000
- Monthly repeat cleanup for a small business: about PKR 5,000-10,000

Take 50% before starting for any job above PKR 3,000 (never share the full clean file before the deposit lands). Do a free demo on *their* first 20 rows only.

## Getting your first client (the tool alone will not bring customers)

Do these in order; each is free:

1. Make a 1-page demo: screenshots of before/after from the sample file (nothing from real clients).
2. Ask 10 people you already know who run something small (a shop, a coaching centre, a hostel, a society, a clinic front desk) if they keep customer or student lists in Excel.
3. Send the message below to 10 of them. Track replies in a sheet.
4. Offer to clean the first 20 rows free as a sample.
5. After the first paid job, ask for a one-line testimonial (with permission to show it).

### Message you can adapt

> Assalam o Alaikum [Name]. I'm an IT student at the University of Sindh. I clean messy Excel/CSV lists: remove duplicates, fix phone numbers, dates and spellings, and give you a short report of what changed. If you have a customer, student or stock list that is hard to use, I can clean the first 20 rows free so you can see the result. Would that be useful?

## Improve it after the first paid job

- Add a Pakistani CNIC format check (13 digits) as an option.
- Add fuzzy name matching with a "review" sheet instead of auto-merging.
- Add a small desktop or web upload page.

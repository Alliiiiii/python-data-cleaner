#!/usr/bin/env python3
"""Builds sample/messy_customers.xlsx and .csv - deliberately dirty data to test datacleaner.py."""
import csv
from pathlib import Path

from openpyxl import Workbook

out = Path(__file__).parent / "sample"
out.mkdir(exist_ok=True)

HEADER = ["Customer Name ", "E-mail", "Phone #", "Signup Date", "City", "Country", "Order Amount (Rs)", "Status", ""]

R = [  # name, email, phone, date, city, country, amount, status
    ["ayesha khan", "Ayesha.Khan@gmail.com", "0300-1234567", "13/05/2025", "karachi", "PK", "Rs 12,500", "Active"],
    ["  Bilal   Ahmed ", "bilal.ahmed@outlook.com", "0321 7654321", "2025-02-11", "Hyderabad", "Pakistan", "8,000", "active"],
    ["SANA MALIK", "sana.malik@yahoo.com", "+92 333 1112233", "5 Jan 2025", "LAHORE", "pakistan", "Rs. 25,750.50", "ACTIVE"],
    ["Usman Ali", "usman.ali@gmail.com", "03451239876", "22/07/2025", "Jamshoro", "Pakistan", "N/A", "Inactive"],
    ["fatima noor", "fatima.noor@gmail.com", "0312-9988776", "30-08-2025", "Sukkur", "PK", "15000", "Active"],
    ["Hassan Raza", "hassan.raza@company.pk", "0301 5554433", "2024-12-01", "Karachi", "Pakistan", "(2,500)", "Inactive"],
    ["Zainab Shah", "zainab.shah@gmail.com", "0333-4445566", "17/03/2025", "Larkana", "Pakistan", "9,999", "Active"],
    ["omar farooq", "omar.farooq@hotmail.com", "0322 1239870", "09/09/2025", "Lahore", "Pakistan", "31,200", "active"],
    ["Mariam Siddiqui", "mariam.s@gmail.com", "0315 6677889", "14/11/2025", "Karachi", "Pakistan", "7,450", "Active"],
    ["Imran Qureshi", "imran.qureshi@gmail.com", "", "01/06/2025", "Hyderabad", "Pakistan", "18,000", "Active"],
    ["Hina Baloch", "hina.baloch@gmail.com", "0345 8889900", "28/02/2025", "Quetta", "Pakistan", "N/A", "Inactive"],
    ["Danish Memon", "danish.memon@gmail.com", "0331-7776655", "19/04/2025", "Jamshoro", "Pakistan", "42,300", "Active"],
    ["Noor Jehan", "noor.jehan@gmail.com", "0302 1010101", "not sure", "Karachi", "Pakistan", "5,000", "Active"],
    ["Kamran Abbasi", "kamran.abbasi.gmail.com", "0311 2020202", "12/12/2024", "Islamabad", "Pakistan", "6,100", "Active"],
    ["Sara Lund", "sara.lund@example.co.uk", "+44 7700 900123", "2025-03-04", "london", "UK", "14,800", "Active"],
]

rows = [r for r in R]
# --- duplicates in disguise ---
rows.append(["AYESHA KHAN", "ayesha.khan@gmail.com ", "+923001234567", "13/05/2025", "Karachi", "Pakistan", "12500", "active"])  # same person, different formatting
rows.append(["Bilal Ahmed", "BILAL.AHMED@OUTLOOK.COM", "", "2025-02-11", "", "Pakistan", "8000", "Active"])                      # less complete copy
rows.append(["Usman Ali", "usman.ali@gmail.com", "03451239876", "22/07/2025", "Jamshoro", "Pakistan", "N/A", "Inactive"])         # exact copy
rows.append(["Sana Malik", "sana.malik@yahoo.com", "0333 1112233", "", "", "", "", ""])                                           # sparse copy
rows.append(["Zainab Shah", "zainab.shah@gmail.com", "0333-4445566", "17/03/2025", "Larkana", "Pakistan", "9,999", "Active"])    # exact copy
rows.append(["Imran Qureshi", "imran.qureshi@gmail.com", "0300 2223344", "01/06/2025", "", "Pakistan", "18,000", "Active"])    # has the phone the first copy lacks
rows.append(["Hassan  Raza", "Hassan.Raza@company.pk", "0301-5554433", "2024-12-01", "karachi", "PK", "(2,500)", "Inactive"])   # variant

wb = Workbook()
ws = wb.active
ws.title = "Customers"
ws.append(["Customer Export - October 2026"])
ws.append(["Exported by: sales-crm v2.3"])
ws.append([])
ws.append(HEADER)
for i, r in enumerate(rows):
    ws.append(r + [""])
    if i in (5, 12):
        ws.append([])  # stray blank rows
ws.append([])
ws.append(["Total records: %d" % len(rows)])
wb.save(out / "messy_customers.xlsx")

with open(out / "messy_customers.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["Customer Export - October 2026"])
    w.writerow([])
    w.writerow(HEADER)
    w.writerows(rows)
    w.writerow([])
    w.writerow(["Total records: %d" % len(rows)])

print(f"Wrote {len(rows)} messy rows to {out}")

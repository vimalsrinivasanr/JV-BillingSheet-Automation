# GCC SAP JV Automation - KT Summary

## 1. Project Goal
Automate SAP Journal Voucher (JV) upload files from raw monthly billing Excel sheets for Randstad GCC.

TWO STAGE PIPELINE (must stay separate):
- Stage 1 - normalizer.py: Converts messy monthly Excel -> clean February Golden Format Excel -> user reviews
- Stage 2 - engine.py: Takes approved normalized file -> generates SAP JV upload Excel

## 2. Active Directory
/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/

OLD/STALE directory (ignore):
/Users/macbook/Downloads/Library/PROJECTS/Randstad/Billinf Sheet->JV automation/

## 3. Files
- main_gui.py: CustomTkinter UI (needs Stage 1 button)
- engine.py: Stage 2 JV generation (old hardcoded indices - needs update)
- ai_mapper.py: Gemini column discovery (working)
- normalizer.py: Stage 1 normalizer (NOT YET WRITTEN - IMMEDIATE TASK)
- build_windows.py: PyInstaller packaging (done)
- Input Data.xlsx: February Golden Template
- Input Data March.xlsx: March messy input

## 4. February Golden Template (Input Data.xlsx, sheet: Billing sheet)
Header: Row 0, Data starts: Row 1, Total: 93 columns

Key columns:
Col 0  Workday ID         - Employee ID W000xxxx
Col 1  EmpNo              - Oracle HR number GCCxxxx
Col 2  Name
Col 5  Grade
Col 6  Capability Center
Col 7  ELT                - Manager name
Col 8  Location
Col 14 Fixed CTC          - GL 742234
Col 16 Bonus 2026         - GL 742234
Col 17 Gratuity           - GL 742237
Col 18 Leave Encashment   - GL 742237
Col 19 PF admin Monthly   - GL 742234
Col 20 EDLI Monthly       - GL 742234
Col 25 Leadership cost    - GL 742235
Col 26 Managers desk fee  - GL 742238
Col 30 Manager cost       - GL 742238
Col 34 TP on total cost   - GL 842028
Col 35 Gross Desk Cost    - GL 742236
Col 36 TP on desk cost    - GL 842028
Col 41 Legal Entity
Col 42 Classification     - FILTER: Billable / Non Billable
Col 79 Billed/ Unbilled   - FILTER: Billed / Accruals / Unbillled
Col 80 IC Code            - e.g. NL_RSH, US_SAP_L -> JV Ref Key 1
Col 82 Invoice No.        - -> JV Reference.1
Col 84 Capability Center  - -> JV Ref Key 2
Col 85 Recharge - Payroll - GL 742234 (formula-derived)
Col 86 Recharge - Manager - GL 742238 (formula-derived)
Col 87 Recharge - Leadership - GL 742235 (formula-derived)
Col 88 Recharge - Desk Cost  - GL 742236 (formula-derived)
Col 89 Recharge - Retirals   - GL 742237 (formula-derived)
Col 90 Mark up               - GL 842028 (formula-derived)
Col 92 Diff                  - Balance check must be 0

IMPORTANT: Cols 85-90, 92 are NOT in raw data. User adds them manually as formulas.
The normalizer must auto-calculate them.

## 5. March File Inconsistencies (Input Data March.xlsx, sheet: Billing sheet)
- Header at Row 1 (not Row 0)
- Data starts at Row 2 (not Row 1)
- Workday ID at col 4 (not col 0)
- "Grade" renamed to "Oracle Level"
- "Fixed CTC" renamed to "Monthly Fixed Pay"
- "Leave Encashment" renamed to "Leave Enc"
- "Classification" renamed to "Billable/ Non Billable"
- "Invoice No." renamed to "Invoice No/ JV no"
- No Billed/Unbilled column - must derive from Invoice col 58
- IC Code col 70 exists but data is empty
- GL recharge columns missing - must be calculated

## 6. Feb -> March Exact Column Mapping (confirmed by data analysis)
Feb 0  Workday ID              -> Mar 4  Workday ID
Feb 1  EmpNo                   -> Mar 5  Oracle ID
Feb 2  Name                    -> Mar 6  Name
Feb 3  Date of joining         -> Mar 7  Date of joining
Feb 4  Last working day        -> Mar 8  Last working day
Feb 5  Grade                   -> Mar 9  Oracle Level
Feb 6  Capability Center       -> Mar 11 Capability Center
Feb 7  ELT                     -> Mar 12 ELT
Feb 8  Location                -> Mar 13 Core Location
Feb 13 Billing days            -> Mar 16 Billing days
Feb 14 Fixed CTC               -> Mar 18 Monthly Fixed Pay
Feb 15 % of Bonus              -> Mar 19 % of Bonus
Feb 16 Bonus 2026              -> Mar 20 Bonus 2026
Feb 17 Gratuity                -> Mar 22 Gratuity
Feb 18 Leave Encashment        -> Mar 23 Leave Enc
Feb 19 PF admin Monthly        -> Mar 24 PF admin Monthly
Feb 20 EDLI Monthly            -> Mar 25 EDLI Monthly
Feb 21 Commision/bonus         -> Mar 26 Commision/ Other bonus
Feb 22 SBP Accrual             -> Mar 27 SBP Cost
Feb 23 Insurance               -> Mar 28 Insurance
Feb 24 Total Payroll           -> Mar 29 Total Payroll
Feb 25 Leadership cost         -> Mar 30 Leadership cost
Feb 26 Managers desk fee       -> Mar 31 VSL desk fee
Feb 30 Manager cost            -> Mar 35 VSL's Cost
Feb 33 Total Cost excl desk    -> Mar 39 Total Cost other than desk cost
Feb 34 TP on total cost        -> Mar 40 TP on total cost
Feb 35 Gross Desk Cost         -> Mar 41 Gross Desk Cost
Feb 36 TP on desk cost         -> Mar 42 TP on desk cost
Feb 37 Total billable amount   -> Mar 43 Total billable amount
Feb 38 Client                  -> Mar 45 Client
Feb 39 Client for sumifs       -> Mar 46 Client for sumifs
Feb 41 Legal Entity            -> Mar 47 Legal Entity
Feb 40 Cost center             -> Mar 48 Cost center
Feb 42 Classification          -> Mar 55 Billable/ Non Billable
Feb 43 FTE count               -> Mar 56 FTE count
Feb 79 Billed/ Unbilled        -> NOT PRESENT (derive from col 58)
Feb 80 IC Code                 -> Mar 70 IC CODE (data is empty)
Feb 82 Invoice No.             -> Mar 58 Invoice No/ JV no

## 7. GL Formula Logic (for normalizer to compute cols 85-90)
GL 742234 = Fixed CTC + Bonus 2026 + PF admin Monthly + EDLI Monthly
GL 742238 = Managers desk fee + Manager cost
GL 742235 = Leadership cost
GL 742236 = Gross Desk Cost
GL 742237 = Gratuity + Leave Encashment
GL 842028 = TP on total cost + TP on desk cost

## 8. Synthesizing Billed/Unbilled from Invoice Column
March col 58 values:
"RG30001390" (actual invoice ID) -> Billed
"Not billed"                     -> Unbilled
"Non Billable"                   -> Non Billable (skip in JV)
empty                            -> skip

## 9. normalizer.py Requirements (IMMEDIATE TASK)
Location: /Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation/normalizer.py

Must:
1. Read any monthly billing Excel (sheet: Billing sheet)
2. Auto-detect header row by scanning top 20 rows for Workday ID anchor
3. Map source cols to February Golden Format using:
   a. Exact name match (case-insensitive)
   b. Alias/synonym match (using the table above)
   c. Data-value fingerprinting (W000xxxx = Workday ID, GCCxxxx = EmpNo)
4. Synthesize Billed/Unbilled from Invoice column when missing
5. Calculate GL recharge columns using formula logic above
6. Output *_NORMALIZED.xlsx with 3 sheets:
   a. Normalized: Clean data in Feb column order. Header color coding:
      Green = mapped from source, Orange = synthesized, Red = missing/blank
   b. Mapping Report: Each column, source name, source index, match method
   c. Warnings & Actions: Missing columns and user instructions
7. Accept log_callback parameter for GUI integration

DO NOT generate JV file. Stage 1 only.

## 10. SAP JV Rules (Stage 2 engine.py)
- Filter: Classification == Billable AND Billed/Unbilled == Billed
- Group by Invoice No. (one JV per invoice)
- Max 999 lines per batch
- 1 credit row per batch (Account 500003, Key "40") + N debit rows (Key "50")
- 3-row header in output Excel
- Balance check formula =ROUND(SUM(J4:Jn),2) at bottom
- Decimal arithmetic for precision

37 output columns: Reference, Document Date, Document Type, Company Code, Posting Date,
Reference.1, Document Header Text, Currency, Exchange rate, Amount, Posting Key, Account,
Special G/L ind., Cost Center, Internal Order, Profit Center, Business Area,
Assignment Number (20), Item Text (50), Ref Key 1, Ref Key 2, Ref Key 3 (20), Material,
Trading Partner, Tax Code, Withholding tax code, Withholding tax base amount in document currency,
Customer, Contracts, Revenue Period, Core Consultant, Revenue Month, Reversal Date, LEDGER,
WT CODE1, WT Amount, Inovice Receipt Date

## 11. Tech Stack
Python: /opt/anaconda3/bin/python3 (3.12)
Packages: pandas, openpyxl, customtkinter, google-generativeai, pyinstaller
GUI: CustomTkinter dark mode sidebar layout
AI: Gemini 1.5 Flash

## 12. Key Decisions
- No hardcoded column numbers - name/alias/data-value driven only
- Two-stage pipeline with human review checkpoint between stages
- February format is immutable reference standard
- engine.py has old hardcoded logic - update AFTER normalizer.py is complete

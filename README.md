# GCC SAP JV Automation Desktop Hub

A professional desktop automation tool for generating SAP-ready Journal Voucher files from raw billing spreadsheets.

## 🌟 Key Features
- **AI-Powered Mapping:** Uses Gemini AI to intelligently discover spreadsheet columns (even if the layout changes).
- **Auto-Calculation Engine:** Automatically calculates missing GL columns (Payroll, Markups, etc.) from raw data columns A-AQ.
- **Robust Fallback:** Works 100% offline using hardcoded business rules if AI or internet is unavailable.
- **Modern UI:** Sleek, dark-mode "Slate-style" interface built with CustomTkinter.
- **SAP Ready:** Generates perfectly balanced JV entries with 999-line batching logic.

## 🚀 Getting Started

### Prerequisites
- Python 3.10+
- Install dependencies:
  ```bash
  pip install customtkinter pandas openpyxl google-generativeai pyinstaller
  ```

### Running the App
```bash
python scripts/main_gui.py
```

### Sheet selection and manual inputs
- If the workbook contains only one sheet, the app will auto-select it (no specific sheet name required).
- If the workbook contains multiple sheets, the app will prompt you to choose which sheet to use.
- You may also provide a manually prepared "filter-check" workbook containing only the final columns required for JV generation (for example: Workday ID, Classification, Billed/ Unbilled, Invoice No., EmpNo, Capability Center, Recharge columns, Mark up). The engine will accept this trimmed sheet and generate the JV without running full normalization.


### Packaging as Windows EXE
To create a standalone `.exe` for distribution on a Windows machine:
1.  Open your command prompt (CMD) on Windows.
2.  Navigate to this folder.
3.  Run the build script:
    ```bash
    python build_windows.py
    ```
4.  Once finished, your single-file application `GCC_JV_Automation_Hub.exe` will be in the **`dist/`** folder.

> ⚠️ If the build fails with a Windows permission error, make sure the previous `GCC_JV_Automation_Hub.exe` is not currently running or held open by another process.

> [!TIP]
> **Pro Tip:** If your client uses a very specific billing format that changes, just provide the Gemini API key in the UI. The AI Mapping engine will automatically handle the new column layout without you having to re-code the application.

---
Produced for Randstad GCC Automation.

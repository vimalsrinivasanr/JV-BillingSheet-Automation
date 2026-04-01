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
python main_gui.py
```

### Packaging as Windows EXE
To create a standalone `.exe` for distribution on a Windows machine:
1.  Open your command prompt (CMD) on Windows.
2.  Navigate to this folder.
3.  Run the build script:
    ```bash
    python build_windows.py
    ```
4.  Once finished, your single-file application `GCC_JV_Automation_Hub.exe` will be in the **`dist/`** folder.

> [!TIP]
> **Pro Tip:** If your client uses a very specific billing format that changes, just provide the Gemini API key in the UI. The AI Mapping engine will automatically handle the new column layout without you having to re-code the application.

---
Produced for Randstad GCC Automation.

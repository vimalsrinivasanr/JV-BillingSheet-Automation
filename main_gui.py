import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
import pandas as pd
from engine import JVEngine
from decimal import Decimal

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("GCC SAP JV Automation Hub")
        self.geometry("850x650")

        # Configuration
        self.engine = None
        self.input_path = ""
        self.output_dir = os.path.dirname(os.path.abspath(__file__))

        # UI Elements
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar for settings
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="GCC Automation", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Month Config
        self.month_entry = self.add_sidebar_entry("Month Label:", "Feb'26", 1)
        self.date_entry = self.add_sidebar_entry("End Date (DDMMYYYY):", "28022026", 3)
        self.cc_entry = self.add_sidebar_entry("Company Code:", "6000", 5)
        self.apikey_entry = self.add_sidebar_entry("Gemini API Key:", "", 7)
        self.apikey_entry.configure(show="*")

        self.btn_save = ctk.CTkButton(self.sidebar_frame, text="Save Config", command=self.save_config)
        self.btn_save.grid(row=9, column=0, padx=20, pady=10)

        # Main Area
        self.main_frame = ctk.CTkFrame(self, corner_radius=15)
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.header = ctk.CTkLabel(self.main_frame, text="SAP Journal Voucher Processing", font=ctk.CTkFont(size=24, weight="bold"))
        self.header.grid(row=0, column=0, padx=20, pady=(40, 20))

        # File Drop/Select
        self.drop_frame = ctk.CTkFrame(self.main_frame, height=150, border_width=2, border_color="#333")
        self.drop_frame.grid(row=1, column=0, padx=40, pady=20, sticky="ew")
        self.drop_frame.grid_propagate(False)
        self.drop_frame.grid_columnconfigure(0, weight=1)
        self.drop_frame.grid_rowconfigure(0, weight=1)

        self.btn_browse = ctk.CTkButton(self.drop_frame, text="Select Billing Sheet (Excel)", command=self.browse_file)
        self.btn_browse.grid(row=0, column=0, padx=20, pady=20)

        self.lbl_file = ctk.CTkLabel(self.main_frame, text="No file selected", text_color="gray")
        self.lbl_file.grid(row=2, column=0, padx=20, pady=0)

        # Process Button
        self.btn_run = ctk.CTkButton(self.main_frame, text="GENERATE SAP JV", height=50, 
                                     font=ctk.CTkFont(size=16, weight="bold"),
                                     fg_color="#285", hover_color="#274",
                                     command=self.start_processing)
        self.btn_run.grid(row=3, column=0, padx=40, pady=40, sticky="ew")

        # Console
        self.textbox = ctk.CTkTextbox(self.main_frame, height=150)
        self.textbox.grid(row=4, column=0, padx=20, pady=10, sticky="nsew")
        self.textbox.insert("0.0", "System ready.\n")

    def add_sidebar_entry(self, label, default, row):
        lbl = ctk.CTkLabel(self.sidebar_frame, text=label, anchor="w")
        lbl.grid(row=row, column=0, padx=20, pady=(10, 0), sticky="w")
        entry = ctk.CTkEntry(self.sidebar_frame, placeholder_text=default)
        entry.insert(0, default)
        entry.grid(row=row+1, column=0, padx=20, pady=(0, 10))
        return entry

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.input_path = filename
            self.lbl_file.configure(text=os.path.basename(filename), text_color="white")
            self.log(f"Selected: {filename}")

    def log(self, msg):
        self.textbox.insert("end", f"[{threading.current_thread().name}] {msg}\n")
        self.textbox.see("end")

    def save_config(self):
        messagebox.showinfo("Config", "Configuration updated locally.")

    def start_processing(self):
        if not self.input_path:
            messagebox.showwarning("Error", "Please select a file first.")
            return
        
        self.btn_run.configure(state="disabled")
        threading.Thread(target=self.run_engine, name="Engine").start()

    def run_engine(self):
        try:
            m_label = self.month_entry.get()
            m_date = self.date_entry.get()
            c_code = self.cc_entry.get()
            a_key = self.apikey_entry.get()

            config = {
                "MONTH_LABEL": m_label,
                "MONTH_END_DATE": m_date,
                "COMPANY_CODE": int(c_code),
                "API_KEY": a_key
            }
            
            engine = JVEngine(config)
            rows = engine.run_processing(self.input_path, log_callback=self.log, api_key=a_key)
            
            self.log("Finalizing Excel file structure...")
            safe_label = m_label.replace("'", "").replace(" ", "_")
            out_name = f"SAP_JV_Upload_{safe_label}.xlsx"
            out_path = os.path.join(os.path.dirname(self.input_path), out_name)
            
            # Use engine's specialized writer
            engine.write_excel(rows, out_path, log_callback=self.log)
            
            self.log(f"SUCCESS: {out_name} generated.")
            # Thread-safe UI update
            self.after(0, lambda: messagebox.showinfo("Success", f"JV File Generated!\n\nLocation: {out_path}"))
            
        except Exception as e:
            err_msg = str(e)
            self.log(f"ERROR: {err_msg}")
            # Thread-safe UI update
            self.after(0, lambda: messagebox.showerror("Processing Error", f"Failed to process file:\n{err_msg}"))
        finally:
            self.after(0, lambda: self.btn_run.configure(state="normal"))

if __name__ == "__main__":
    app = App()
    app.mainloop()

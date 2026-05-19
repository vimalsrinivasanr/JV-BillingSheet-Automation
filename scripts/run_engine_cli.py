import sys
import os
import importlib.util
from pathlib import Path

if len(sys.argv) < 2:
    print("Usage: python scripts/run_engine_cli.py <input_xlsx> [out_dir]")
    sys.exit(1)

input_path = sys.argv[1]
out_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(input_path))

# Load engine module by path
spec = importlib.util.spec_from_file_location("engine_mod", os.path.join(os.path.dirname(__file__), "engine.py"))
engine_mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine_mod)
JVEngine = engine_mod.JVEngine

# Load normalizer module by path
spec2 = importlib.util.spec_from_file_location("norm_mod", os.path.join(os.path.dirname(__file__), "normalizer.py"))
norm_mod = importlib.util.module_from_spec(spec2)
spec2.loader.exec_module(norm_mod)
BillingNormalizer = norm_mod.BillingNormalizer

print(f"Running engine on: {input_path}")
# Config defaults
config = {"MONTH_LABEL": "Run", "MONTH_END_DATE": "01012000", "COMPANY_CODE": 6000}

# Auto-normalize when needed
xls = None
try:
    xls = __import__("pandas").ExcelFile(input_path)
except Exception:
    pass

norm_path = input_path
if xls is None or "Normalized" not in xls.sheet_names:
    print("Normalizing input (Stage 1)...")
    normalizer = BillingNormalizer(log_callback=print)
    norm_path, _ = normalizer.normalize(input_path)
    print(f"Normalized -> {norm_path}")
else:
    print("Detected Normalized sheet; using input directly.")

eng = JVEngine(config)
rows = eng.run_processing(norm_path, log_callback=print)

safe_label = config["MONTH_LABEL"].replace("'", "").replace(" ", "_")
out_name = f"SAP_JV_Upload_{safe_label}.xlsx"
out_path = os.path.join(out_dir, out_name)
eng.write_excel(rows, out_path, log_callback=print)
print(f"Output written: {out_path}")

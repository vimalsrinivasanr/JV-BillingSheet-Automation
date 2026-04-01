import os
import json
import pandas as pd
from difflib import get_close_matches
from engine import JVEngine

base = '/Users/macbook/Downloads/Library/PROJECTS/Randstad/JV-BillingSheet-Automation'
feb_path = os.path.join(base, 'Input Data.xlsx')
normalized_path = os.path.join(base, 'Input Data March_NORMALIZED.xlsx')
final_out = os.path.join(base, 'SAP_JV_Upload_Mar26_FINAL.xlsx')

# 1) Compare normalized headers to February reference headers
feb_raw = pd.read_excel(feb_path, sheet_name='Billing sheet', header=None, dtype=str)
anchors = {'workday id', 'empno', 'name'}
feb_hdr = 0
for i in range(min(20, len(feb_raw))):
    vals = {str(v).strip().lower() for v in feb_raw.iloc[i] if pd.notna(v)}
    if len(anchors & vals) >= 2:
        feb_hdr = i
        break
feb_headers = [str(h).strip() if pd.notna(h) else '' for h in feb_raw.iloc[feb_hdr]]

norm_df = pd.read_excel(normalized_path, sheet_name='Normalized', dtype=str)
norm_headers = [str(c).strip() for c in norm_df.columns]

engine_required = [
    'Workday ID', 'Capability Center', 'Legal Entity', 'Classification', 'Billed/ Unbilled', 'IC Code',
    'Invoice No.', 'EmpNo (ref)', 'Capability Center (ref)',
    'Recharge - Payroll', 'Recharge - Manager', 'Recharge - Leadership',
    'Recharge - Desk Cost', 'Recharge - Retirals', 'Mark up'
]
missing_engine = [c for c in engine_required if c not in norm_headers]

exact_overlap = sum(1 for h in norm_headers if h in feb_headers)
missing_vs_feb = [h for h in feb_headers if h and h not in norm_headers]

report = {
    'feb_header_row': feb_hdr,
    'feb_header_count': len(feb_headers),
    'normalized_header_count': len(norm_headers),
    'exact_name_overlap_count': exact_overlap,
    'engine_required_missing': missing_engine,
    'sample_missing_vs_feb': missing_vs_feb[:12],
}

# 2) Optional Gemini check (if API key exists)
api_key = os.getenv('GEMINI_API_KEY', '').strip()
if api_key:
    try:
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash-latest')
        prompt = (
            'Compare these two header lists and return strict JSON with keys: '
            'is_compatible_for_stage2(boolean), critical_missing(list), notes(list).\\n'
            f'Normalized headers: {json.dumps(norm_headers)}\\n'
            f'Required headers: {json.dumps(engine_required)}'
        )
        resp = model.generate_content(prompt)
        text = (resp.text or '').strip()
        if '```json' in text:
            text = text.split('```json', 1)[1].split('```', 1)[0].strip()
        gemini = json.loads(text)
        report['gemini_available'] = True
        report['gemini_result'] = gemini
    except Exception as e:
        report['gemini_available'] = True
        report['gemini_error'] = str(e)
else:
    report['gemini_available'] = False
    report['gemini_note'] = 'GEMINI_API_KEY not set; skipped LLM validation.'

# 3) Produce final JV file if Stage-2 required columns exist
if missing_engine:
    report['final_jv_generated'] = False
    report['final_jv_reason'] = 'Missing required Stage-2 columns.'
else:
    cfg = {
        'MONTH_LABEL': 'Mar26',
        'MONTH_END_DATE': '31032026',
        'COMPANY_CODE': 6000,
    }
    eng = JVEngine(cfg)
    rows = eng.run_processing(normalized_path, log_callback=print)
    eng.write_excel(rows, final_out, log_callback=print)
    jv = pd.read_excel(final_out, sheet_name='JV', header=2)
    amount_sum = pd.to_numeric(jv['Amount'], errors='coerce').fillna(0).sum()
    report['final_jv_generated'] = True
    report['final_jv_path'] = final_out
    report['final_jv_rows'] = int(len(rows))
    report['final_jv_amount_sum'] = round(float(amount_sum), 2)

print(json.dumps(report, indent=2, default=str))

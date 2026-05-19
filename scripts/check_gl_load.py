import importlib.util, sys
import os
spec = importlib.util.spec_from_file_location('engine_mod', os.path.join('scripts','engine.py'))
engine_mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(engine_mod)
JVEngine = engine_mod.JVEngine
eng = JVEngine()
path = sys.argv[1]
df = eng._load_normalized_data(path)
import pandas as pd
for c in ['gl_742234','gl_742238','gl_742235','gl_742236','gl_742237','gl_842028']:
    s = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
    print(c, 'non-zero count=', (s.abs()>0.009).sum(), 'sum=', s.sum())
print('Sample values (first 15)')
print(df[['gl_742234','gl_742238','gl_742235','gl_742236','gl_742237','gl_842028']].head(15))

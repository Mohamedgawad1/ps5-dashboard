import re,os,glob
for pat in ['ovPunchlist','ovTasks','Punchlist_1399','TestsPlanned_1369','DETAILED PUNCH','DETAILED ITR','DETAILED ITR LIST']:
    hits=[]
    for f in glob.glob(r'*.py'):
        try: t=open(f,encoding='utf-8',errors='replace').read()
        except: continue
        if pat.lower() in t.lower(): hits.append(f)
    if hits: print(pat,'->',hits)

import io
txt=open('dpr_dashboard.py',encoding='utf-8').read()
i=txt.find('def extract_rfc')
j=txt.find('\ndef ',i+10)
print(txt[i:j])

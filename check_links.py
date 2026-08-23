import json, os, urllib.parse
d = json.load(open('data.json','r',encoding='utf-8'))
master = r'C:\Users\mylap\OneDrive\Desktop\dashboard\WIRING - MASTER'

broken_pdf = 0
ok_pdf = 0
for item in d:
    link = item.get('link','')
    if not link: continue
    decoded = urllib.parse.unquote(link)
    if decoded.startswith('/pdf/'):
        decoded = decoded[5:]
    fpath = os.path.join(master, decoded)
    if os.path.exists(fpath):
        ok_pdf += 1
    else:
        broken_pdf += 1
        if broken_pdf <= 5:
            tag = item['asset_tag']
            print(f'BROKEN PDF: {tag} -> {os.path.basename(decoded)}')

broken_rfi = 0
ok_rfi = 0
for item in d:
    rfi_pdf = item.get('rfi_pdf','')
    if not rfi_pdf: continue
    if rfi_pdf.startswith('/pdf/'):
        fname = rfi_pdf[5:]
    else:
        fname = rfi_pdf
    fpath = os.path.join(master, fname)
    if os.path.exists(fpath):
        ok_rfi += 1
    else:
        broken_rfi += 1
        if broken_rfi <= 5:
            tag = item['asset_tag']
            print(f'BROKEN RFI: {tag} -> {fname}')

print()
print(f'Wiring PDFs: {ok_pdf} OK, {broken_pdf} BROKEN (out of {ok_pdf+broken_pdf} with link)')
print(f'RFI PDFs:    {ok_rfi} OK, {broken_rfi} BROKEN (out of {ok_rfi+broken_rfi} with rfi_pdf)')
print(f'Total assets: {len(d)}')

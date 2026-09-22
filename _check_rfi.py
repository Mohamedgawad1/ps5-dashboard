import fitz,os,re,json
master=r'C:\Users\mylap\OneDrive\Desktop\dashboard\WIRING - MASTER'
TAG_RE=re.compile(r'PS5-[\w]+-[\w]+-[\w]+(?:-[\w]+)?')
d=json.load(open('data.json','r',encoding='utf-8'))
all_tags=set(i['asset_tag'] for i in d)

doc=fitz.open(os.path.join(master,'CPP-RFI-68-64-11-0195.pdf'))
text=''.join(p.get_text() for p in doc);doc.close()
pdf_tags=set(t.replace(' ','') for t in TAG_RE.findall(text))
matched=pdf_tags & all_tags
print('Tags in 0195 PDF:',len(pdf_tags))
print('All tags from PDF:',sorted(pdf_tags))
print('Matched to data.json assets:',len(matched))
for t in sorted(matched)[:10]:
    item=[i for i in d if i['asset_tag']==t][0]
    rfi_pdf=item.get('rfi_pdf','')
    print(f'  {t} -> rfi_pdf: {rfi_pdf}')

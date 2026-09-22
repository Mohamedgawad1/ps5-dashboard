import os, re

WIRING_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'WIRING - MASTER')
rfi_pdfs = sorted(f for f in os.listdir(WIRING_DIR) if f.startswith('CPP-RFI'))
print(f'CPP-RFI PDFs: {len(rfi_pdfs)}')

test_rfis = [
    'CPP-RFI-68-33/1-06-0360',
    'CPP-RFI-68-331-06-0379',
    'CPP-RFI-68-332-10-0324',
    'CPP-RFI-68-64-11-0140',
    'CPP-RFI-68-33/1-06-0482',
    'CPP-RFI-68-65-20-0039',
]
for rfi in test_rfis:
    rfi_clean = rfi.replace('/', '').replace(' ', '')
    found = None
    for f in rfi_pdfs:
        fn = f.replace('.pdf', '').replace('.PDF', '').replace(' ', '')
        if rfi_clean == fn:
            found = f
            break
    if not found:
        num_m = re.search(r'(\d{4})$', rfi_clean)
        if num_m:
            num = num_m.group(1)
            for f in rfi_pdfs:
                fn = f.replace('.pdf', '').replace('.PDF', '').replace(' ', '')
                if fn.endswith(num):
                    found = f
                    break
    status = 'OK' if found else 'NOT FOUND'
    print(f'  {rfi} -> {status} ({found})')

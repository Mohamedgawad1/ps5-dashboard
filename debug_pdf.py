import os, urllib.parse

WIRING = r'C:\Users\mylap\OneDrive\Desktop\dashboard\WIRING - MASTER'
test_url = '/pdf/CPP-RFI-68-331-06-0360.pdf'
filename = urllib.parse.unquote(test_url[5:])
filepath = os.path.join(WIRING, filename)
print('filename:', repr(filename))
print('filepath:', repr(filepath))
print('exists:', os.path.exists(filepath))

# List actual files starting with CPP-RFI-68-331-06-03
for f in os.listdir(WIRING):
    if 'CPP-RFI-68-331-06-03' in f:
        print('  actual:', repr(f))

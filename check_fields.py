import re
t = open('phone_sms.py').read()
for m in re.finditer(r"data\.get\('(\w+)'", t):
    print(m.group(1))

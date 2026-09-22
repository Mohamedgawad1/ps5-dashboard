import re
with open('index.html', 'r', encoding='utf-8') as f:
    c = f.read()

section_ids = re.findall(r'id="([^"]*sec[^"]*)"', c)
print("Section IDs:", section_ids)

sidebar = re.findall(r'<label[^>]*>.*?</label>', c, re.DOTALL)
print(f"\nSidebar labels ({len(sidebar)}):")
for l in sidebar:
    txt = re.sub(r'<[^>]+>', '', l).strip()
    print(f"  {txt[:100]}")

# Check for specific sections
checks = ['Project Overview', 'ITR Daily', 'ITR Weekly', 'ITR Monthly',
          'E&I&T', 'Milestone Progress', 'ITR Description Table',
          'CMT & QC Punch', 'Punch List', 'RFI Status', 'Search Index',
          'Cable CMT', 'Cable Schedule', 'Master Tracker']
for s in checks:
    count = c.count(s)
    print(f"'{s}': {count} occurrences")

import json

with open(r'C:\Users\mylap\OneDrive\Desktop\dashboard\dashboard_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

data_json = json.dumps(data, ensure_ascii=False)

with open(r'C:\Users\mylap\OneDrive\Desktop\dashboard\dashboard.html', 'r', encoding='utf-8') as f:
    html = f.read()

# Replace the loadData function to use embedded data
old = """async function loadData(){try{const r=await fetch('dashboard_data.json?t='+Date.now());return await r.json()}catch(e){return null}}"""
new = f"""function loadData(){{return {data_json} }}"""
html = html.replace(old, new)

with open(r'C:\Users\mylap\OneDrive\Desktop\dashboard\dashboard.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("Data embedded into dashboard.html!")

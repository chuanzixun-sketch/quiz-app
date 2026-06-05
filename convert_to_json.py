"""Convert questions.xls (HTML table) to questions.json"""
import re
import json

with open('questions.xls', 'r', encoding='utf-8-sig') as f:
    content = f.read()

# Extract all table rows
rows = re.findall(r'<tr>(.*?)</tr>', content, re.DOTALL)
print(f'Found {len(rows)} rows (including header)')

# Parse header from first row
header_cells = re.findall(r'<t[hd]\s+class="text">(.*?)</t[hd]>', rows[0], re.DOTALL)
if not header_cells:
    # Try simpler pattern
    header_cells = re.findall(r'<t[hd][^>]*>(.*?)</t[hd]>', rows[0], re.DOTALL)

# Clean header: strip HTML tags and whitespace
header = []
for c in header_cells:
    clean = re.sub(r'<[^>]+>', '', c).strip()
    # Remove all spaces (Chinese and English)
    clean = clean.replace(' ', '').replace('　', '')
    header.append(clean)
print(f'Header ({len(header)}): {header}')

# Map Chinese column names to internal keys
COL_MAP = {
    '题目类型': 'qtype',
    '选择题题干': 'stem',
    '正确答案': 'answer',
    '答案解析': 'explanation',
    '难易度': 'difficulty',
    '知识点': 'topic',
    '标签': 'tags',
    '选项数': 'opt_count',
    '选项A': 'opt_a',
    '选项B': 'opt_b',
    '选项C': 'opt_c',
    '选项D': 'opt_d',
}
mapped = [COL_MAP.get(h, h) for h in header]
print(f'Mapped: {mapped}')

# Parse data rows
data = []
for row in rows[1:]:
    cells = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL)
    if not cells:
        continue

    vals = []
    for c in cells:
        clean = re.sub(r'<[^>]+>', '', c).strip()
        vals.append(clean)

    # Skip empty rows
    has_content = any(v for v in vals)
    if not has_content:
        continue

    obj = {}
    for i, h in enumerate(mapped):
        if i < len(vals):
            obj[h] = vals[i]
        else:
            obj[h] = ''

    # Only keep valid rows: must have stem and valid answer (A/B/C/D)
    stem = obj.get('stem', '')
    answer = obj.get('answer', '').strip().upper()
    if stem and answer in ('A', 'B', 'C', 'D'):
        obj['answer'] = answer
        data.append(obj)

print(f'Valid questions: {len(data)}')

# Write JSON
with open('questions.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False)
print('Written to questions.json')

# Print sample
if data:
    print(f'\nSample question:\n  stem: {data[0].get("stem","")[:80]}\n  answer: {data[0].get("answer","")}\n  options: A={data[0].get("opt_a","")[:40]} B={data[0].get("opt_b","")[:40]}')

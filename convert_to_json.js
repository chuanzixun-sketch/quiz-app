/**
 * Convert questions.xls (HTML table) to questions.json
 * Node.js script — no external dependencies
 */
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync('questions.xls', 'utf-8');

// Column name mapping
const COL_MAP = {
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
};

// Parse HTML table rows
const rowRegex = /<tr>(.*?)<\/tr>/gs;
const cellRegex = /<t[hd][^>]*>(.*?)<\/t[hd]>/gs;

const rows = [...html.matchAll(rowRegex)].map(m => m[1]);
console.log(`Found ${rows.length} rows (including header)`);

if (rows.length < 2) {
  console.error('Not enough rows found in HTML table');
  process.exit(1);
}

// Parse header
function parseCells(rowHtml) {
  const cells = [...rowHtml.matchAll(cellRegex)].map(m => {
    // Strip all HTML tags and trim
    return m[1].replace(/<[^>]+>/g, '').trim();
  });
  return cells;
}

const headerCells = parseCells(rows[0]);
const header = headerCells.map(c => c.replace(/[\s　]/g, ''));
console.log('Header:', header);

const mapped = header.map(h => COL_MAP[h] || h);
console.log('Mapped:', mapped);

// Parse data
const data = [];
for (let i = 1; i < rows.length; i++) {
  const cells = parseCells(rows[i]);
  if (cells.length === 0) continue;

  // Skip rows with no content
  if (!cells.some(c => c.length > 0)) continue;

  const obj = {};
  for (let j = 0; j < mapped.length; j++) {
    obj[mapped[j]] = j < cells.length ? cells[j] : '';
  }

  // Validate: must have stem and valid answer (A/B/C/D)
  const stem = obj.stem || '';
  const answer = (obj.answer || '').toUpperCase().trim();
  if (stem && /^[ABCD]$/.test(answer)) {
    obj.answer = answer;
    data.push(obj);
  }
}

console.log(`Valid questions: ${data.length}`);

// Write JSON
fs.writeFileSync('questions.json', JSON.stringify(data, null, 2), 'utf-8');
console.log('Written to questions.json');

const jsonSize = JSON.stringify(data).length;
console.log(`JSON size: ${(jsonSize / 1024).toFixed(1)} KB`);

// Show sample
if (data.length > 0) {
  const q = data[0];
  console.log(`\nSample question:`);
  console.log(`  stem: ${(q.stem || '').substring(0, 80)}`);
  console.log(`  answer: ${q.answer}`);
  console.log(`  opt_a: ${(q.opt_a || '').substring(0, 40)}`);
  console.log(`  opt_b: ${(q.opt_b || '').substring(0, 40)}`);
}

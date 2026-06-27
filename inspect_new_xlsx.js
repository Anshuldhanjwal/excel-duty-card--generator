const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, 'FOR DUTY CARD.xlsx');
const workbook = XLSX.read(require('fs').readFileSync(filePath), { type: 'buffer' });

const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });

// Print rows 0-8 (header area)
for (let r = 0; r < 9; r++) {
  const row = rows[r];
  console.log(`\nRow ${r}:`);
  for (let c = 0; c < row.length; c++) {
    if (row[c] !== '') console.log(`  Col ${c}: "${row[c]}"`);
  }
}

// Show merges for header
if (sheet['!merges']) {
  console.log('\n=== MERGES ===');
  for (const m of sheet['!merges']) {
    console.log(`  ${XLSX.utils.encode_range(m)}`);
  }
}

// Count total blocks (serial numbers)
let blockCount = 0;
for (let r = 0; r < rows.length; r++) {
  const val = rows[r][0];
  if (typeof val === 'number' || (typeof val === 'string' && /^\d+$/.test(val.trim()) && val.trim() !== '')) {
    blockCount++;
  }
}
console.log(`\nTotal blocks/locations: ${blockCount}`);

// Count how many blocks have nightshift data (col 8-12)
let dayOnlyBlocks = 0;
let dayNightBlocks = 0;
for (let r = 0; r < rows.length; r++) {
  const val = rows[r][0];
  if (typeof val === 'number' || (typeof val === 'string' && /^\d+$/.test(val.trim()) && val.trim() !== '')) {
    const hasNight = String(rows[r][8] || '').trim() !== '';
    if (hasNight) dayNightBlocks++; else dayOnlyBlocks++;
  }
}
console.log(`Blocks with both day+night: ${dayNightBlocks}`);
console.log(`Blocks with day only: ${dayOnlyBlocks}`);

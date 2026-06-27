/**
 * Test script to verify the parser output from "FOR DUTY CARD.xlsx"
 * Run with: node test_parser.js
 */
const XLSX = require('xlsx');
const path = require('path');

// Simulate the cleanAndConvert and parseForDutyCardSheet logic
const filePath = path.join(__dirname, 'FOR DUTY CARD.xlsx');
const workbook = XLSX.read(require('fs').readFileSync(filePath), { type: 'buffer' });

const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });

console.log(`Sheet: "${workbook.SheetNames[0]}", Rows: ${rawRows.length}`);

// Find header row
let headerRowIdx = -1;
for (let r = 0; r < Math.min(5, rawRows.length); r++) {
  const col0 = String(rawRows[r][0] || '').toLowerCase();
  if (col0.includes('dz') || col0.includes('क्र') || col0.includes('सं')) {
    headerRowIdx = r;
    break;
  }
}
console.log(`Header row: ${headerRowIdx}`);

// Find first data row
let dataStartRow = headerRowIdx + 1;
for (let r = dataStartRow; r < rawRows.length; r++) {
  const val = String(rawRows[r][0] || '').trim();
  if (/^\d+$/.test(val)) {
    dataStartRow = r;
    break;
  }
}
console.log(`Data starts at row: ${dataStartRow}`);

// Parse blocks
let blocks = [];
let currentBlock = null;

for (let r = dataStartRow; r < rawRows.length; r++) {
  const row = rawRows[r];
  const col0 = String(row[0] || '').trim();
  const isSerial = /^\d+$/.test(col0);

  if (isSerial) {
    currentBlock = {
      serial: col0,
      place: String(row[1] || '').trim().substring(0, 40),
      dayOfficers: [],
      nightOfficers: [],
    };
    blocks.push(currentBlock);

    const dayName = String(row[3] || '').trim();
    if (dayName) currentBlock.dayOfficers.push(dayName.substring(0, 30));

    const nightName = String(row[8] || '').trim();
    if (nightName) currentBlock.nightOfficers.push(nightName.substring(0, 30));
  } else if (currentBlock) {
    const dayName = String(row[3] || '').trim();
    if (dayName) currentBlock.dayOfficers.push(dayName.substring(0, 30));

    const nightName = String(row[8] || '').trim();
    if (nightName) currentBlock.nightOfficers.push(nightName.substring(0, 30));
  }
}

let totalDayCards = 0;
let totalNightCards = 0;

console.log(`\n=== BLOCKS SUMMARY ===`);
for (const block of blocks) {
  const dayCount = block.dayOfficers.length;
  const nightCount = block.nightOfficers.length;
  totalDayCards += dayCount;
  totalNightCards += nightCount;
  console.log(`Block ${block.serial}: "${block.place}" => Day: ${dayCount} cards, Night: ${nightCount} cards, Total: ${dayCount + nightCount}`);
}

console.log(`\n=== TOTALS ===`);
console.log(`Total blocks: ${blocks.length}`);
console.log(`Total day cards: ${totalDayCards}`);
console.log(`Total night cards: ${totalNightCards}`);
console.log(`GRAND TOTAL: ${totalDayCards + totalNightCards} individual duty cards`);

// Verify block 1 specifically
console.log(`\n=== BLOCK 1 VERIFICATION ===`);
const b1 = blocks[0];
console.log(`Place: ${b1.place}`);
console.log(`Day officers (${b1.dayOfficers.length}):`);
b1.dayOfficers.forEach((o, i) => console.log(`  ${i+1}. ${o}`));
console.log(`Night officers (${b1.nightOfficers.length}):`);
b1.nightOfficers.forEach((o, i) => console.log(`  ${i+1}. ${o}`));

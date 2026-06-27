const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, 'FOR DUTY CARD.xlsx');
const workbook = XLSX.read(require('fs').readFileSync(filePath), { type: 'buffer' });

const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });

console.log('Searching for any date pattern like dd-mm-yyyy or similar in Sheet 1:');
const dateRe = /(\d{1,2}[-./]\d{1,2}[-./]\d{4})/g;

rawRows.forEach((row, rIdx) => {
  row.forEach((cell, cIdx) => {
    const str = String(cell);
    const matches = str.match(dateRe);
    if (matches) {
      console.log(`Found match in Row ${rIdx}, Col ${cIdx}: "${str}" ->`, matches);
    }
    // Also look for Krutidev date patterns or Devanagari numbers if any
    if (str.toLowerCase().includes('fnukad') || str.includes('दिनांक') || str.includes('fnua%') || str.includes('दि०')) {
      console.log(`Found date keyword in Row ${rIdx}, Col ${cIdx}: "${str}"`);
    }
  });
});

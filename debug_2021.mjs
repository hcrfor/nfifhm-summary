import XLSX from 'xlsx';

const filename = '2021자료.xlsx';
let wb;
try {
    wb = XLSX.readFile(filename);
} catch (e) {
    console.error('Error reading workbook:', e);
    process.exit(1);
}

const sheetName = wb.SheetNames.find(n => n.includes('임목조사표(2021)')) || 
                wb.SheetNames.find(n => n.includes('2021')) || 
                wb.SheetNames[0];

const ws = wb.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(ws);

const targetCluster = '431031518958';
const matches = data.filter(row => String(row['집락번호'] || '').trim() === targetCluster);

console.log(`Sheet Name: ${sheetName}`);
console.log(`Searching for Cluster: ${targetCluster}`);
console.log(`Found matches: ${matches.length}`);

if (matches.length > 0) {
    console.log('Sample match:', JSON.stringify(matches[0], null, 2));
} else {
    // try searching by partial match or different keys
    console.log('No exact matches. Checking unique cluster IDs...');
    const clusters = [...new Set(data.slice(0, 100).map(row => String(row['집락번호'] || '').trim()))];
    console.log('First 20 unique clusters:', clusters.slice(0, 20));
    
    const partialMatch = data.find(row => String(row['표본점번호'] || '').includes('431031518958'));
    if (partialMatch) {
         console.log('Found partial match in 표본점번호:', JSON.stringify(partialMatch, null, 2));
    }
}

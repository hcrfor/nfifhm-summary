import XLSX from 'xlsx';

const filename = '2021자료.xlsx';
const wb = XLSX.readFile(filename);
const sheetName = wb.SheetNames.find(n => n.includes('임목조사표(2021)')) || wb.SheetNames[0];
const ws = wb.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(ws);

const targetCluster = '8958'; // Using the short ID I saw in debug output
const matches = data.filter(row => String(row['집락번호'] || '').trim() === targetCluster);

console.log(`Searching for Cluster Serial: ${targetCluster}`);
console.log(`Found points: ${matches.length}`);

matches.forEach(m => {
    console.log(`New ID: ${m['표본점번호']} -> Old ID: ${m['구표본점번호']}`);
});

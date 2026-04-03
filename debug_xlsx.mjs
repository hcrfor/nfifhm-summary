import * as XLSX from 'xlsx';
import fs from 'fs';

const filePath = 'c:\\Users\\han\\development\\antigraviy\\nfifhm-summary\\public\\2021자료.xlsx';
const buf = fs.readFileSync(filePath);
const wb = XLSX.read(buf, {type: 'buffer'});

const result = {
    sheets: wb.SheetNames,
    firstSheetData: []
};

for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws);
    if  (data.length > 0) {
        const searchId = '471043218357';
        const match = data.find(row => String(row['집락번호'] || '').includes(searchId));
        if (match) {
            result.matchingSheet = sheetName;
            result.matchSample = match;
            break;
        }
    }
}

fs.writeFileSync('debug_result.json', JSON.stringify(result, null, 2));
console.log('Done!');

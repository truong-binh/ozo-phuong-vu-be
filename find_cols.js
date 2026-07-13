const XLSX = require('xlsx');

function findColumns(filePath) {
  const wb = XLSX.readFile(filePath);
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const row = rows[r] || [];
      if (row.includes('SL đặt') || row.includes('SL nhận')) {
        console.log(`Found columns in Sheet "${name}", Row ${r + 1}:`);
        console.log(row.slice(0, 20));
      }
    }
  }
}

try {
  findColumns('D:\\CODE\\ozo-chi-phuong\\File mẫu - Theo dõi kho theo cấu trúc dài.xlsx');
} catch (e) {
  console.error(e);
}

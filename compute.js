// Shared calculation: from a 2D array of Lark rows -> Map(sku -> Σ(SL đặt − SL nhận))
// counting only rows whose status == processValue ("On process").
const { colToNum, normSku } = require('./excel');

function colIdx(letter) {
  return colToNum(String(letter).toUpperCase()) - 1;
}

function toNum(x) {
  if (x == null || x === '') return 0;
  const n = Number(String(x).replace(/[,\s]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

// Chuyển 1 giá trị ô ngày -> "YYYY-MM" (hoặc null nếu không đọc được).
// Ưu tiên serial number (khi đọc UnformattedValue); có fallback cho vài dạng chuỗi.
function serialToYM(n) {
  const ms = Math.round((n - 25569) * 86400 * 1000); // serial 1899-12-30 -> Unix
  const d = new Date(ms);
  if (isNaN(d.getTime())) return null;
  return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
}
function valueToYM(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return serialToYM(v);
  const s = String(v).trim();
  if (/^\d+(\.\d+)?$/.test(s)) return serialToYM(Number(s)); // chuỗi số = serial
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/); // dd/mm/yy hoặc dd/mm/yyyy
  if (m) {
    let yr = Number(m[3]);
    if (yr < 100) yr += 2000;
    return yr + '-' + String(Number(m[2])).padStart(2, '0');
  }
  const d = new Date(s); // fallback: để JS thử
  if (!isNaN(d.getTime())) return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
  return null;
}

// values: array of rows (each row = array indexed by column, A=0)
function computeValuesBySku(values, cfg) {
  const {
    headerRow = 1, // 1-based; data starts after it
    skuCol = 'B',
    orderCol = 'I',
    recvCol = 'J',
    statusCol = 'AP',
    processValue = 'On process',
    monthFilter = null, // "YYYY-MM": chỉ tính dòng có ngày (dateCol) thuộc tháng này
    dateCol = 'AL',
  } = cfg;

  const iSku = colIdx(skuCol);
  const iOrder = colIdx(orderCol);
  const iRecv = colIdx(recvCol);
  const iStatus = colIdx(statusCol);
  const iDate = colIdx(dateCol);
  const want = String(processValue).trim().toLowerCase();

  const valuesBySku = new Map();
  let processRows = 0;
  let unparsedDates = 0;
  for (let r = headerRow; r < values.length; r++) {
    const row = values[r] || [];
    const sku = normSku(row[iSku]);
    if (!sku) continue;
    const status = row[iStatus] != null ? String(row[iStatus]).trim().toLowerCase() : '';
    if (status !== want) continue; // only "On process"; "Done" skipped
    if (monthFilter) {
      const ym = valueToYM(row[iDate]);
      if (ym == null) unparsedDates++;
      if (ym !== monthFilter) continue; // khác tháng (hoặc không đọc được ngày) -> bỏ
    }
    const transit = toNum(row[iOrder]) - toNum(row[iRecv]);
    processRows++;
    valuesBySku.set(sku, (valuesBySku.get(sku) || 0) + transit);
  }
  return { valuesBySku, processRows, unparsedDates };
}

module.exports = { computeValuesBySku };

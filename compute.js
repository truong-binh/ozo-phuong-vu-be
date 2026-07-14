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

// Chuyển 1 giá trị ô ngày -> "YYYY-MM-DD" (hoặc null nếu không đọc được). Dùng cho lũy kế THEO NGÀY.
function serialToYMD(n) {
  const ms = Math.round((n - 25569) * 86400 * 1000);
  const d = new Date(ms);
  if (isNaN(d.getTime())) return null;
  return (
    d.getUTCFullYear() +
    '-' + String(d.getUTCMonth() + 1).padStart(2, '0') +
    '-' + String(d.getUTCDate()).padStart(2, '0')
  );
}
function valueToYMD(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return serialToYMD(v);
  const s = String(v).trim();
  if (/^\d+(\.\d+)?$/.test(s)) return serialToYMD(Number(s)); // chuỗi số = serial
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/); // dd/mm/yy hoặc dd/mm/yyyy
  if (m) {
    let yr = Number(m[3]);
    if (yr < 100) yr += 2000;
    return yr + '-' + String(Number(m[2])).padStart(2, '0') + '-' + String(Number(m[1])).padStart(2, '0');
  }
  const d = new Date(s); // fallback: để JS thử
  if (!isNaN(d.getTime()))
    return (
      d.getUTCFullYear() +
      '-' + String(d.getUTCMonth() + 1).padStart(2, '0') +
      '-' + String(d.getUTCDate()).padStart(2, '0')
    );
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
    monthFilter = null, // "YYYY-MM": LŨY KẾ THÁNG — tính dòng có ngày (dateCol) <= hết tháng này
    dateFilter = null,  // "YYYY-MM-DD": LŨY KẾ NGÀY — tính dòng có ngày (dateCol) <= đúng ngày này (ưu tiên hơn monthFilter)
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
    if (dateFilter) {
      // LŨY KẾ THEO NGÀY: chỉ lấy dòng có ngày <= đúng ngày cutoff. Vd cutoff 2026-03-12
      // thì dòng ngày 2026-03-13 bị bỏ.
      const ymd = valueToYMD(row[iDate]);
      if (ymd == null) {
        unparsedDates++;
        continue; // không đọc được ngày -> bỏ (báo về để đối chiếu)
      }
      if (ymd > dateFilter) continue; // ngày SAU cutoff -> bỏ; <= thì tính (lũy kế)
    } else if (monthFilter) {
      const ym = valueToYM(row[iDate]);
      if (ym == null) {
        unparsedDates++;
        continue; // không đọc được ngày -> bỏ (báo về để đối chiếu)
      }
      if (ym > monthFilter) continue; // đặt SAU tháng cutoff -> bỏ; <= thì tính (lũy kế)
    }
    const transit = toNum(row[iOrder]) - toNum(row[iRecv]);
    processRows++;
    valuesBySku.set(sku, (valuesBySku.get(sku) || 0) + transit);
  }
  return { valuesBySku, processRows, unparsedDates };
}

module.exports = { computeValuesBySku };

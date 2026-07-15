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

// Chuyển 1 giá trị ô "Thời gian đặt hàng" -> "YYYY-MM" (null nếu không đọc được).
//
// Cột này trên Lark chỉ có THÁNG-NĂM ("Aug-25", "Jun-26") nên mọi so sánh đều ở
// mức tháng — không có ngày để mà lũy kế theo ngày.
//
// KHÔNG dùng `new Date(s)` làm fallback: nó đọc "Aug-25" thành ngày 25/08/2001
// (25 = ngày, năm mặc định 2001) -> mọi dòng lọt qua mọi cutoff mà không báo lỗi.
// Không khớp mẫu nào -> trả null để dòng bị đếm vào unparsedDates và lộ ra ngay.
const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

function ym(year, month) {
  if (!(month >= 1 && month <= 12)) return null;
  return year + '-' + String(month).padStart(2, '0');
}

function serialToYM(n) {
  const d = new Date(Math.round((n - 25569) * 86400 * 1000)); // serial 1899-12-30 -> Unix
  if (isNaN(d.getTime())) return null;
  return ym(d.getUTCFullYear(), d.getUTCMonth() + 1);
}

function valueToYM(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return serialToYM(v);
  const s = String(v).trim();
  if (s === '') return null;
  if (/^\d+(\.\d+)?$/.test(s)) return serialToYM(Number(s)); // chuỗi số = serial (UnformattedValue)

  let m = s.match(/^([A-Za-z]{3,9})[-/ ](\d{2}|\d{4})$/); // "Aug-25", "Aug-2025", "Aug 25"
  if (m) {
    const mon = MONTHS[m[1].slice(0, 3).toLowerCase()];
    if (!mon) return null;
    const yr = Number(m[2]);
    return ym(yr < 100 ? yr + 2000 : yr, mon);
  }
  m = s.match(/^(\d{4})-(\d{1,2})(-\d{1,2})?$/); // "2026-07" hoặc "2026-07-15"
  if (m) return ym(Number(m[1]), Number(m[2]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/); // dd/mm/yy hoặc dd/mm/yyyy
  if (m) {
    const yr = Number(m[3]);
    return ym(yr < 100 ? yr + 2000 : yr, Number(m[2]));
  }
  return null;
}

// values: array of rows (each row = array indexed by column, A=0)
function computeValuesBySku(values, cfg) {
  const {
    headerRow = 2, // 1-based; data starts after it (Lark có 2 dòng header gộp)
    skuCol = 'B', // Mã SP cần đặt
    orderCol = 'J', // SL đặt
    recvCol = 'K', // SL nhận
    statusCol = 'AQ', // Tình trạng lô hàng
    processValue = 'On process',
    // Cột "Thời gian đặt hàng" chỉ có tháng-năm -> MỌI cutoff đều so ở mức THÁNG.
    monthFilter = null, // "YYYY-MM": lũy kế — tính dòng đặt <= hết tháng này
    dateFilter = null,  // "YYYY-MM-DD": chỉ lấy phần tháng; giữ để .xlsm cũ gửi &date= vẫn chạy
    dateCol = 'AM', // Thời gian đặt hàng
  } = cfg;

  // Nguồn chỉ có tháng -> cutoff ngày 15/07/2026 = cutoff tháng 2026-07.
  const cutoffYM = dateFilter ? String(dateFilter).slice(0, 7) : monthFilter;

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
    if (cutoffYM) {
      // LŨY KẾ THEO THÁNG: lấy dòng đặt <= hết tháng cutoff. Vd cutoff 2026-03
      // thì dòng "Apr-26" bị bỏ, "Mar-26" và cũ hơn thì tính.
      const rowYM = valueToYM(row[iDate]);
      if (rowYM == null) {
        unparsedDates++;
        continue; // trống / không đọc được -> bỏ (trả về qua X-Unparsed-Dates để đối chiếu)
      }
      if (rowYM > cutoffYM) continue; // so sánh chuỗi "YYYY-MM" = so sánh thời gian
    }
    const transit = toNum(row[iOrder]) - toNum(row[iRecv]);
    processRows++;
    valuesBySku.set(sku, (valuesBySku.get(sku) || 0) + transit);
  }
  return { valuesBySku, processRows, unparsedDates };
}

module.exports = { computeValuesBySku };

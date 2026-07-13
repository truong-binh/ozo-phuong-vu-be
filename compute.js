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

// values: array of rows (each row = array indexed by column, A=0)
function computeValuesBySku(values, cfg) {
  const {
    headerRow = 1, // 1-based; data starts after it
    skuCol = 'B',
    orderCol = 'I',
    recvCol = 'J',
    statusCol = 'AP',
    processValue = 'On process',
  } = cfg;

  const iSku = colIdx(skuCol);
  const iOrder = colIdx(orderCol);
  const iRecv = colIdx(recvCol);
  const iStatus = colIdx(statusCol);
  const want = String(processValue).trim().toLowerCase();

  const valuesBySku = new Map();
  let processRows = 0;
  for (let r = headerRow; r < values.length; r++) {
    const row = values[r] || [];
    const sku = normSku(row[iSku]);
    if (!sku) continue;
    const status = row[iStatus] != null ? String(row[iStatus]).trim().toLowerCase() : '';
    if (status !== want) continue; // only "On process"; "Done" skipped
    const transit = toNum(row[iOrder]) - toNum(row[iRecv]);
    processRows++;
    valuesBySku.set(sku, (valuesBySku.get(sku) || 0) + transit);
  }
  return { valuesBySku, processRows };
}

module.exports = { computeValuesBySku };

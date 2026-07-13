// Surgical injection of values into ONE column of ONE worksheet inside an .xlsx,
// leaving Power Query / Pivot / formulas / styles 100% untouched.
//
// We only rewrite the target worksheet XML: for each data row we read the SKU
// cell (key column) from its cached <v>, and if that SKU is in `valuesBySku`
// we set the target column cell's value — preserving its style attribute.

const JSZip = require('jszip');

// Normalize a SKU for matching: trim, collapse inner whitespace, uppercase.
// Both sides (Excel key cell + Lark data) go through this so tiny formatting
// differences (a stray space, lower/upper case) don't cause "0 matches".
function normSku(s) {
  return String(s == null ? '' : s).trim().replace(/\s+/g, ' ').toUpperCase();
}

// column letter <-> index (A=1)
function colToNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function unescapeXml(s) {
  return String(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function parseRels(xml) {
  const map = {};
  const re = /<Relationship\b([^>]*)\/>/g;
  let m;
  while ((m = re.exec(xml))) {
    const attrs = {};
    for (const a of m[1].matchAll(/(\w+)="([^"]*)"/g)) attrs[a[1]] = a[2];
    if (attrs.Id && attrs.Target) map[attrs.Id] = attrs.Target;
  }
  return map;
}

// Find "xl/worksheets/sheetN.xml" for a given sheet display name.
function resolveSheetPath(workbookXml, relsXml, sheetName) {
  const rels = parseRels(relsXml);
  const wantName = sheetName.replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const re = /<sheet\b([^>]*)\/>/g;
  let m;
  while ((m = re.exec(workbookXml))) {
    const attrs = {};
    for (const a of m[1].matchAll(/([\w:]+)="([^"]*)"/g)) attrs[a[1]] = a[2];
    if (attrs.name === wantName || attrs.name === sheetName) {
      const rid = attrs['r:id'] || attrs.id;
      let target = rels[rid];
      if (!target) return null;
      if (!target.startsWith('xl/')) target = 'xl/' + target.replace(/^\/?/, '');
      return target;
    }
  }
  return null;
}

// Escape a number/string for an XML <v> body.
function xmlNum(v) {
  // We only ever write numbers into the target column.
  const n = Number(v);
  if (!Number.isFinite(n)) return null;
  return String(n);
}

/**
 * @param {Buffer} xlsxBuffer  the uploaded template
 * @param {object} opts
 * @param {string} opts.sheetName   display name, e.g. "Tồn kho & Kế hoạch đặt hàng"
 * @param {string} opts.keyCol      key column letter, e.g. "A"
 * @param {string} opts.targetCol   column to write, e.g. "G"
 * @param {number} opts.firstRow    first data row, e.g. 3
 * @param {Map<string,number>} opts.valuesBySku  sku -> value
 * @returns {Promise<{buffer:Buffer, written:number, matched:string[], missing:string[]}>}
 */
async function injectColumn(xlsxBuffer, opts) {
  const { sheetName, keyCol, targetCol, firstRow } = opts;
  // Normalize incoming keys so matching is robust regardless of caller.
  const valuesBySku = new Map();
  for (const [k, v] of opts.valuesBySku) valuesBySku.set(normSku(k), v);
  const zip = await JSZip.loadAsync(xlsxBuffer);

  const workbookXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const sheetPath = resolveSheetPath(workbookXml, relsXml, sheetName);
  if (!sheetPath || !zip.file(sheetPath)) {
    throw new Error('Không tìm thấy sheet "' + sheetName + '" trong file Excel.');
  }

  let xml = await zip.file(sheetPath).async('string');

  const keyColNum = colToNum(keyCol);
  const targetColNum = colToNum(targetCol);
  const matched = [];
  const usedSku = new Set();

  // Process row by row so we can read the key cell then set the target cell.
  xml = xml.replace(/<row\b([^>]*)>([\s\S]*?)<\/row>/g, (rowFull, rowAttrs, rowBody) => {
    const rm = /\br="(\d+)"/.exec(rowAttrs);
    const rowNum = rm ? parseInt(rm[1], 10) : null;
    if (rowNum == null || rowNum < firstRow) return rowFull;

    // read the key (SKU) cell cached value
    const keyRef = keyCol + rowNum;
    const keyCellRe = new RegExp('<c\\b[^>]*\\br="' + keyRef + '"[^>]*>([\\s\\S]*?)<\\/c>');
    const kc = keyCellRe.exec(rowBody);
    let sku = null;
    if (kc) {
      const vm = /<v>([\s\S]*?)<\/v>/.exec(kc[1]);
      if (vm) sku = normSku(unescapeXml(vm[1]));
    }
    if (!sku) return rowFull;

    // Mã có dữ liệu tính toán -> ghi giá trị; mã không có -> điền 0 (thay vì bỏ trống).
    const raw = valuesBySku.get(sku);
    const val = raw == null ? 0 : raw;
    const num = xmlNum(val);
    if (num == null) return rowFull;

    if (raw != null) {
      usedSku.add(sku);
      matched.push(sku);
    }

    const tRef = targetCol + rowNum;
    let newBody = rowBody;

    // Case 1: self-closing empty cell  <c r="G3" s="10"/>
    const emptyRe = new RegExp('<c\\b([^>]*?)\\br="' + tRef + '"([^>]*?)\\/>');
    // Case 2: cell with content        <c r="G3" s="10">...</c>
    const fullRe = new RegExp('<c\\b([^>]*?)\\br="' + tRef + '"([^>]*?)>[\\s\\S]*?<\\/c>');

    if (emptyRe.test(newBody)) {
      newBody = newBody.replace(emptyRe, (_m, a, b) => {
        // strip any t="..." (make it a number cell), keep everything else (style s=)
        const attrs = (a + b).replace(/\s*t="[^"]*"/g, '');
        return '<c ' + attrs.trim() + ' r="' + tRef + '"><v>' + num + '</v></c>';
      });
    } else if (fullRe.test(newBody)) {
      newBody = newBody.replace(fullRe, (_m, a, b) => {
        const attrs = (a + b).replace(/\s*t="[^"]*"/g, '');
        return '<c ' + attrs.trim() + ' r="' + tRef + '"><v>' + num + '</v></c>';
      });
    } else {
      // Target cell doesn't exist yet — insert it in column order.
      const cellXml = '<c r="' + tRef + '"><v>' + num + '</v></c>';
      newBody = insertCellInOrder(newBody, targetColNum, cellXml, rowNum);
    }

    return '<row' + rowAttrs + '>' + newBody + '</row>';
  });

  const missing = [];
  for (const sku of valuesBySku.keys()) if (!usedSku.has(sku)) missing.push(sku);

  zip.file(sheetPath, xml);
  // Drop folder marker entries so the archive stays byte-clean like the original.
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir) delete zip.files[name];
  }
  const out = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  });
  return { buffer: out, written: matched.length, matched, missing };
}

// Insert a <c> in the correct column position within a row body.
function insertCellInOrder(rowBody, targetColNum, cellXml, rowNum) {
  const cells = [...rowBody.matchAll(/<c\b[^>]*\br="([A-Z]+)\d+"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g)];
  if (cells.length === 0) return cellXml;
  for (const c of cells) {
    const colNum = colToNum(c[1]);
    if (colNum > targetColNum) {
      const idx = c.index;
      return rowBody.slice(0, idx) + cellXml + rowBody.slice(idx);
    }
  }
  return rowBody + cellXml;
}

module.exports = { injectColumn, colToNum, normSku };

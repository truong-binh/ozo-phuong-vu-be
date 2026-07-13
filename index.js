require('dotenv').config();
const express = require('express');
const cookieSession = require('cookie-session');
const cors = require('cors');
const multer = require('multer');
const crypto = require('crypto');
const path = require('path');
const XLSX = require('xlsx');
const lark = require('./lark');
const { injectColumn } = require('./excel');
const { computeValuesBySku } = require('./compute');

const app = express();
app.set('trust proxy', 1);
const PORT = process.env.PORT || 3000;
const CLIENT_URL = process.env.CLIENT_URL || 'http://localhost:5173';

app.use(cors({
  origin: CLIENT_URL,
  credentials: true,
}));

app.use(express.json({ limit: '2mb' }));
app.use(
  cookieSession({
    name: 'sess',
    keys: [process.env.SESSION_SECRET || 'dev-secret-change-me'],
    maxAge: 24 * 60 * 60 * 1000,
    httpOnly: true,
    sameSite: 'none',
    secure: true,
  })
);

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 30 * 1024 * 1024 } });

// ---- helpers ------------------------------------------------------------
function requireLogin(req, res, next) {
  if (!req.session || !req.session.token) {
    return res.status(401).json({ error: 'not_logged_in' });
  }
  next();
}

// Build a friendly summary + response headers, then send the xlsx.
function sendResult(res, result, processRows, valuesBySku, filename) {
  res.setHeader(
    'X-Summary',
    encodeURIComponent(
      JSON.stringify({
        processRows,
        distinctSku: valuesBySku.size,
        written: result.written,
        missingCount: result.missing.length,
        missing: result.missing.slice(0, 50),
      })
    )
  );
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="' + (filename || 'ket-qua.xlsx') + '"');
  res.send(result.buffer);
}

// ---- auth ---------------------------------------------------------------
app.get('/api/status', (req, res) => {
  res.json({
    hasCreds: lark.hasCreds(),
    loggedIn: Boolean(req.session && req.session.token),
    scopes: lark.SCOPES,
    redirectUri: lark.REDIRECT_URI,
  });
});

app.get('/auth/login', (req, res) => {
  if (!lark.hasCreds()) return res.status(500).send('Chưa cấu hình LARK_APP_ID / LARK_APP_SECRET trong .env');
  const state = crypto.randomBytes(8).toString('hex');
  req.session.oauthState = state;
  res.redirect(lark.buildAuthorizeUrl(state));
});

app.get('/auth/callback', async (req, res) => {
  try {
    const { code, state } = req.query;
    if (!code) return res.status(400).send('Thiếu code');
    if (state !== req.session.oauthState) return res.status(400).send('State không khớp');
    const tok = await lark.exchangeCodeForToken(code);
    req.session.token = tok.access_token;
    req.session.refresh = tok.refresh_token;
    res.redirect(CLIENT_URL);
  } catch (e) {
    res.status(500).send('Đăng nhập Lark lỗi: ' + e.message);
  }
});

app.post('/api/logout', (req, res) => {
  req.session = null;
  res.json({ ok: true });
});

// ---- Lark data ----------------------------------------------------------
// Resolve a pasted Lark URL -> spreadsheet token + list of sheets.
app.post('/api/resolve', requireLogin, async (req, res) => {
  try {
    const { url } = req.body;
    const parsed = lark.parseLarkUrl(url || process.env.LARK_SHEET_URL || '');
    if (!parsed) return res.status(400).json({ error: 'URL Lark không hợp lệ' });
    let spreadsheetToken = parsed.token;
    if (parsed.kind === 'wiki') {
      const node = await lark.resolveWikiToSpreadsheet(parsed.token, req.session.token);
      if (node.objType !== 'sheet')
        return res.status(400).json({ error: 'Link này không phải Lark Sheet (là: ' + node.objType + ')' });
      spreadsheetToken = node.objToken;
    }
    const sheets = await lark.listSheets(spreadsheetToken, req.session.token);
    req.session.spreadsheetToken = spreadsheetToken;
    res.json({ spreadsheetToken, sheets });
  } catch (e) {
    res.status(500).json({ error: e.message, larkCode: e.larkCode });
  }
});

// Preview: read the first N rows of a chosen sheet so the UI can build the
// column mapping (show header row + sample).
app.post('/api/preview', requireLogin, async (req, res) => {
  try {
    const { sheetId, lastCol = 'BZ', rows = 30 } = req.body;
    const token = req.session.token;
    const stoken = req.session.spreadsheetToken;
    const range = `${sheetId}!A1:${lastCol}${rows}`;
    const values = await lark.readValues(stoken, range, token);
    res.json({ values });
  } catch (e) {
    res.status(500).json({ error: e.message, larkCode: e.larkCode });
  }
});

// ---- compute + inject ---------------------------------------------------
// multipart: file=<xlsx>, config=<json string>
app.post('/api/generate', requireLogin, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Chưa upload file Excel mẫu' });
    const cfg = JSON.parse(req.body.config || '{}');
    const {
      sheetId,
      lastCol = 'BZ',
      maxRows = 5000,
      headerRow = 1, // Lark header row (1-based)
      skuCol, // Lark column letter for SKU
      orderCol = 'I', // SL đặt
      recvCol = 'J', // SL nhận
      statusCol = 'AP', // Tình trạng lô hàng
      processValue = 'On process',
      // Excel target:
      targetSheet = 'Tồn kho & Kế hoạch đặt hàng',
      excelKeyCol = 'A',
      excelTargetCol = 'G',
      excelFirstRow = 3,
    } = cfg;
    if (!sheetId) return res.status(400).json({ error: 'Chưa chọn sheet Lark' });
    if (!skuCol) return res.status(400).json({ error: 'Chưa chọn cột Mã SKU bên Lark' });

    const range = `${sheetId}!A1:${lastCol}${maxRows}`;
    const values = await lark.readValues(req.session.spreadsheetToken, range, req.session.token);

    const { valuesBySku, processRows } = computeValuesBySku(values, cfg);
    const result = await injectColumn(req.file.buffer, {
      sheetName: targetSheet,
      keyCol: excelKeyCol,
      targetCol: excelTargetCol,
      firstRow: excelFirstRow,
      valuesBySku,
    });
    sendResult(res, result, processRows, valuesBySku);
  } catch (e) {
    res.status(500).json({ error: e.message, larkCode: e.larkCode });
  }
});

// ---- NO-LOGIN mode: upload an exported Lark sheet (.xlsx/.csv) directly -----
// multipart: template=<xlsx mẫu>, larkfile=<file export từ Lark>, config=<json>
app.post(
  '/api/generate-local',
  upload.fields([{ name: 'template', maxCount: 1 }, { name: 'larkfile', maxCount: 1 }]),
  async (req, res) => {
    try {
      const template = req.files && req.files.template && req.files.template[0];
      const larkfile = req.files && req.files.larkfile && req.files.larkfile[0];
      if (!template) return res.status(400).json({ error: 'Chưa upload file Excel mẫu' });
      if (!larkfile) return res.status(400).json({ error: 'Chưa upload file export từ Lark' });
      const cfg = JSON.parse(req.body.config || '{}');
      const {
        larkSheetName, // tên sheet trong file export (bỏ trống = sheet đầu tiên)
        targetSheet = 'Tồn kho & Kế hoạch đặt hàng',
        excelKeyCol = 'A',
        excelTargetCol = 'G',
        excelFirstRow = 3,
      } = cfg;

      // Parse the Lark export into a 2D array keeping absolute column positions.
      const wb = XLSX.read(larkfile.buffer, { type: 'buffer' });
      const wsName = larkSheetName && wb.Sheets[larkSheetName] ? larkSheetName : wb.SheetNames[0];
      const ws = wb.Sheets[wsName];
      const values = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: '' });

      const { valuesBySku, processRows } = computeValuesBySku(values, cfg);
      const result = await injectColumn(template.buffer, {
        sheetName: targetSheet,
        keyCol: excelKeyCol,
        targetCol: excelTargetCol,
        firstRow: excelFirstRow,
        valuesBySku,
      });
      sendResult(res, result, processRows, valuesBySku);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  }
);

// List sheet names inside an uploaded Lark export (to pick the right tab).
app.post('/api/local-sheets', upload.single('larkfile'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Chưa upload file' });
    const wb = XLSX.read(req.file.buffer, { type: 'buffer', sheetRows: 5 });
    const preview = {};
    for (const name of wb.SheetNames) {
      preview[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, defval: '' });
    }
    res.json({ sheets: wb.SheetNames, preview });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Serve built client in production (optional)
const clientDist = path.join(__dirname, '..', 'client', 'dist');
app.use(express.static(clientDist));

app.listen(PORT, () => {
  console.log(`Server chạy tại http://localhost:${PORT}`);
  if (!lark.hasCreds()) console.log('⚠️  Chưa có LARK_APP_ID/LARK_APP_SECRET — xem README-SETUP.md');
});
// Lark (Larksuite international) API helpers.
// Docs: https://open.larksuite.com/document
//
// Flow used:
//   1) User OAuth 2.0 (v2) -> user_access_token  (reads with the user's own view rights)
//   2) Resolve the /wiki/<token> node -> the underlying spreadsheet token
//   3) List sheets, then read cell values (v2 values API)

const OPEN_BASE = process.env.LARK_OPEN_BASE || 'https://open.larksuite.com';
const ACCOUNTS_BASE = process.env.LARK_ACCOUNTS_BASE || 'https://accounts.larksuite.com';

const APP_ID = process.env.LARK_APP_ID;
const APP_SECRET = process.env.LARK_APP_SECRET;
const REDIRECT_URI = process.env.LARK_REDIRECT_URI || 'http://localhost:3000/auth/callback';

// Scopes must ALSO be enabled in the app console (Permissions & Scopes).
const SCOPES = (process.env.LARK_SCOPES ||
  'wiki:wiki:readonly sheets:spreadsheet:readonly drive:drive:readonly offline_access')
  .trim();

function buildAuthorizeUrl(state) {
  const p = new URLSearchParams({
    client_id: APP_ID,
    redirect_uri: REDIRECT_URI,
    response_type: 'code',
    scope: SCOPES,
    state,
  });
  return `${ACCOUNTS_BASE}/open-apis/authen/v1/authorize?${p.toString()}`;
}

async function exchangeCodeForToken(code) {
  const res = await fetch(`${OPEN_BASE}/open-apis/authen/v2/oauth/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      grant_type: 'authorization_code',
      client_id: APP_ID,
      client_secret: APP_SECRET,
      code,
      redirect_uri: REDIRECT_URI,
    }),
  });
  const data = await res.json();
  if (!res.ok || data.code !== 0 || data.error) {
    throw new Error('Lark token error: ' + JSON.stringify(data));
  }
  return data.data; // Trả về data.data chứa access_token
}

async function refreshToken(refresh_token) {
  const res = await fetch(`${OPEN_BASE}/open-apis/authen/v2/oauth/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      grant_type: 'refresh_token',
      client_id: APP_ID,
      client_secret: APP_SECRET,
      refresh_token,
    }),
  });
  const data = await res.json();
  if (!res.ok || data.code !== 0 || data.error) throw new Error('Lark refresh error: ' + JSON.stringify(data));
  return data.data; // Trả về data.data chứa access_token
}

async function larkGet(path, token, params) {
  const url = new URL(OPEN_BASE + path);
  if (params) for (const [k, v] of Object.entries(params)) url.searchParams.set(k, v);
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const data = await res.json();
  if (data.code && data.code !== 0) {
    const err = new Error('Lark API ' + path + ' -> code ' + data.code + ': ' + data.msg);
    err.larkCode = data.code;
    throw err;
  }
  return data.data;
}

// A /wiki/<token> link points to a "node" that wraps the real object (sheet/doc/base).
async function resolveWikiToSpreadsheet(wikiToken, token) {
  const data = await larkGet('/open-apis/wiki/v2/spaces/get_node', token, {
    token: wikiToken,
    obj_type: 'wiki',
  });
  const node = data.node || data;
  return { objToken: node.obj_token, objType: node.obj_type, title: node.title };
}

async function listSheets(spreadsheetToken, token) {
  const data = await larkGet(
    `/open-apis/sheets/v3/spreadsheets/${spreadsheetToken}/sheets/query`,
    token
  );
  return (data.sheets || []).map((s) => ({
    sheetId: s.sheet_id,
    title: s.title,
    rowCount: s.grid_properties && s.grid_properties.row_count,
    colCount: s.grid_properties && s.grid_properties.column_count,
  }));
}

// Read a rectangular range. `range` like "<sheetId>" or "<sheetId>!A1:Z2000".
async function readValues(spreadsheetToken, range, token) {
  const data = await larkGet(
    `/open-apis/sheets/v2/spreadsheets/${spreadsheetToken}/values/${encodeURIComponent(range)}`,
    token,
    { valueRenderOption: 'ToString', dateTimeRenderOption: 'FormattedString' }
  );
  return (data.valueRange && data.valueRange.values) || [];
}

// Extract the wiki/sheet token from a pasted Lark URL.
function parseLarkUrl(url) {
  const m = String(url).match(/\/(wiki|sheets)\/([A-Za-z0-9]+)/);
  if (!m) return null;
  return { kind: m[1], token: m[2] };
}

module.exports = {
  APP_ID,
  REDIRECT_URI,
  SCOPES,
  hasCreds: () => Boolean(APP_ID && APP_SECRET),
  buildAuthorizeUrl,
  exchangeCodeForToken,
  refreshToken,
  resolveWikiToSpreadsheet,
  listSheets,
  readValues,
  parseLarkUrl,
};

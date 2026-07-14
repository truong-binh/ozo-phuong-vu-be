// Lưu refresh_token của user (Cách B) để nút Excel chạy mà không cần đăng nhập lại.
// Lark XOAY refresh_token mỗi lần refresh -> phải ghi đè token mới sau mỗi lần dùng.
//
// Chỗ lưu: file JSON tại TOKEN_STORE_PATH. Trên Render, ổ đĩa mặc định bị xoá mỗi
// lần deploy -> hãy mount 1 persistent disk và trỏ TOKEN_STORE_PATH vào đó
// (ví dụ /data/lark-token.json). Nếu ghi file thất bại (đĩa read-only) thì vẫn
// giữ trong RAM cho tới lần restart kế tiếp.
const fs = require('fs');
const path = require('path');

const STORE_PATH = process.env.TOKEN_STORE_PATH || path.join(__dirname, '.lark-token.json');

let mem = { refresh_token: null };

function saveRefreshToken(refresh_token) {
  if (!refresh_token) return;
  mem.refresh_token = refresh_token;
  try {
    fs.writeFileSync(STORE_PATH, JSON.stringify({ refresh_token, savedAt: Date.now() }), 'utf8');
  } catch (e) {
    // Đĩa không ghi được -> vẫn còn trong RAM (mem). Sẽ mất khi restart/deploy.
    console.warn('⚠️  Không ghi được token store:', e.message, '(giữ tạm trong RAM)');
  }
}

function loadRefreshToken() {
  if (mem.refresh_token) return mem.refresh_token;
  try {
    const raw = fs.readFileSync(STORE_PATH, 'utf8');
    mem.refresh_token = JSON.parse(raw).refresh_token || null;
  } catch (e) {
    mem.refresh_token = null;
  }
  return mem.refresh_token;
}

function isSeeded() {
  return Boolean(loadRefreshToken());
}

module.exports = { saveRefreshToken, loadRefreshToken, isSeeded, STORE_PATH };

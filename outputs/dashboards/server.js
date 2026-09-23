/**
 * server.js — Beta Solutions Budget Dashboard API Server
 * Node.js + Express + SheetJS (xlsx)
 * Đọc file Excel thực tế, cấp API JSON, file-watcher tự reload khi Excel thay đổi
 * Chạy: node server.js → mở http://localhost:3000
 */

const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// ─── Đường dẫn tuyệt đối tới file Excel gốc ───────────────────────────────
const EXCEL_PATH = path.resolve(
  __dirname,
  '../../sample-data/ngan_sach_phong_ban.xlsx'
);

// ─── Cache dữ liệu trong memory ───────────────────────────────────────────
let dataCache = [];
let lastLoadTime = null;
let lastModified = null;

// ─── Hàm đọc & parse Excel ────────────────────────────────────────────────
function loadExcelData() {
  try {
    const stat = fs.statSync(EXCEL_PATH);
    const modTime = stat.mtime.toISOString();

    // Chỉ reload nếu file thực sự thay đổi
    if (modTime === lastModified && dataCache.length > 0) return;

    const workbook = XLSX.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: null });

    // Normalize dữ liệu: ép kiểu số, chuỗi ngày
    dataCache = rawRows.map(row => ({
      Ma_Giao_Dich: String(row['Ma_Giao_Dich'] || ''),
      Ngay: row['Ngay'] ? String(row['Ngay']) : '',
      Quy: String(row['Quy'] || ''),
      Phong_Ban: String(row['Phong_Ban'] || ''),
      Hang_Muc: String(row['Hang_Muc'] || ''),
      Ngan_Sach_Du_Kien: Number(row['Ngan_Sach_Du_Kien'] || 0),
      Chi_Tieu_Thuc_Te: Number(row['Chi_Tieu_Thuc_Te'] || 0),
      Chenh_Lech: Number(row['Chenh_Lech'] || 0),
      Trang_Thai: String(row['Trang_Thai'] || ''),
      Phan_Tram_Su_Dung: Number(row['Phan_Tram_Su_Dung'] || 0),
    }));

    lastModified = modTime;
    lastLoadTime = new Date().toISOString();
    console.log(`[${lastLoadTime}] ✅ Đã tải ${dataCache.length} dòng từ Excel.`);
  } catch (err) {
    console.error('❌ Lỗi đọc Excel:', err.message);
  }
}

// ─── Load lần đầu ─────────────────────────────────────────────────────────
loadExcelData();

// ─── File Watcher: tự động reload khi Excel thay đổi ─────────────────────
fs.watch(EXCEL_PATH, { persistent: true }, (eventType) => {
  if (eventType === 'change') {
    console.log(`[WATCHER] File Excel thay đổi → đang reload...`);
    // Delay nhỏ để Excel ghi xong file
    setTimeout(loadExcelData, 500);
  }
});

// ─── Middleware ───────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json());

// Serve static files (dashboard HTML) từ cùng thư mục
app.use(express.static(path.join(__dirname)));

// ─── API: GET /api/data — toàn bộ dữ liệu ────────────────────────────────
app.get('/api/data', (req, res) => {
  // Reload từ disk mỗi request để bắt thay đổi (backup cho watcher)
  loadExcelData();
  res.json({
    success: true,
    lastUpdated: lastLoadTime,
    totalRows: dataCache.length,
    data: dataCache,
  });
});

// ─── API: GET /api/meta — metadata phòng ban & quý ────────────────────────
app.get('/api/meta', (req, res) => {
  loadExcelData();
  const departments = [...new Set(dataCache.map(r => r.Phong_Ban))].sort();
  const quarters = [...new Set(dataCache.map(r => r.Quy))].sort();
  res.json({ success: true, departments, quarters });
});

// ─── API: GET /api/health — health check ─────────────────────────────────
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    server: 'Beta Solutions Budget Dashboard API',
    excelPath: EXCEL_PATH,
    lastLoaded: lastLoadTime,
    rows: dataCache.length,
  });
});

// ─── Khởi động server ─────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log('');
  console.log('╔══════════════════════════════════════════════════════╗');
  console.log('║   💜 BETA SOLUTIONS — BUDGET DASHBOARD SERVER        ║');
  console.log('╠══════════════════════════════════════════════════════╣');
  console.log(`║   🌐 Dashboard: http://localhost:${PORT}                 ║`);
  console.log(`║   📊 API Data:  http://localhost:${PORT}/api/data        ║`);
  console.log(`║   🔍 API Meta:  http://localhost:${PORT}/api/meta        ║`);
  console.log('║   ⚡ Real-time: File-watcher ON                      ║');
  console.log('╚══════════════════════════════════════════════════════╝');
  console.log('');
});

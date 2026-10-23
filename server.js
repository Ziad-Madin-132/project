const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

let XLSX;
try {
  XLSX = require('./xlsx.full.min.js');
} catch (e) {
  try {
    XLSX = require('xlsx');
  } catch (err) {
    console.error('[Athar Studio] Warning: SheetJS library not found.');
  }
}

const PORT = 3000;
const EXCEL_FILE = path.join(__dirname, 'حجوزات_استوديو_أثر.xlsx');
const JSON_FILE = path.join(__dirname, 'bookings.json');

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.webmanifest': 'application/manifest+json; charset=utf-8',
  '.png': 'image/png',
  '.webp': 'image/webp',
  '.ico': 'image/x-icon',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

function formatBookingsForExcel(bookings) {
  const labels = {
    pending: 'قيد التأكيد',
    confirmed: 'مؤكد',
    recording: 'جاري التسجيل',
    done: 'تم التسجيل',
    completed: 'مكتمل ومُسلّم',
    cancelled: 'ملغي'
  };

  return bookings.map((b, i) => {
    const code = b.id ? ('ATH-' + b.id.slice(-6).toUpperCase()) : `ATH-${String(i + 1).padStart(3, '0')}`;
    const total = Number(b.total) || 0;
    const deposit = Number(b.deposit) || 0;
    const remaining = Math.max(0, total - deposit);

    return {
      'رقم الحجز': code,
      'اسم المنشد / العميل': b.name || '',
      'رقم الهاتف': b.phone || '',
      'العمل الفني': b.workTitle || '',
      'نوع الخدمة': b.service || '',
      'المقام الموسيقي': b.maqam || 'حر',
      'مهندس الصوت': b.engineer || 'م/ أحمد فاروق',
      'تاريخ الجلسة': b.date || '',
      'وقت البداية': b.time || '',
      'مدة الجلسة (ساعات)': Number(b.duration) || 1,
      'السعر الإجمالي (ج.م)': total,
      'العربون المدفوع (ج.م)': deposit,
      'المتبقي (ج.م)': remaining,
      'حالة الحجز': labels[b.status] || b.status,
      'ملاحظات فنية': b.notes || '',
      'تاريخ التسجيل': b.createdAt || new Date().toISOString()
    };
  });
}

function saveBookingsToExcel(bookings) {
  try {
    if (XLSX) {
      const excelRows = formatBookingsForExcel(bookings);
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(excelRows);

      ws['!cols'] = [
        { wch: 14 },
        { wch: 26 },
        { wch: 16 },
        { wch: 30 },
        { wch: 22 },
        { wch: 14 },
        { wch: 18 },
        { wch: 14 },
        { wch: 12 },
        { wch: 16 },
        { wch: 18 },
        { wch: 18 },
        { wch: 16 },
        { wch: 16 },
        { wch: 45 },
        { wch: 22 }
      ];

      XLSX.utils.book_append_sheet(wb, ws, 'سجل حجوزات استوديو أثر');
      
      // Use buffer writing to guarantee Unicode/Arabic filename compatibility on Windows
      const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      
      try {
        fs.writeFileSync(EXCEL_FILE, buf);
      } catch (writeErr) {
        if (writeErr.code === 'EBUSY' || writeErr.code === 'EPERM') {
          console.warn('[Excel Warning] ملف الإكسل مفتوح حالياً في برنامج آخر، تم حفظ نسخة احتياطية إضافية.');
          const backupExcel = path.join(__dirname, 'حجوزات_استوديو_أثر_محدث.xlsx');
          try { fs.writeFileSync(backupExcel, buf); } catch (e) {}
        } else {
          throw writeErr;
        }
      }
    }

    // Always update JSON database
    fs.writeFileSync(JSON_FILE, JSON.stringify(bookings, null, 2), 'utf8');

    // Also update parent folder file if present
    const parentExcel = path.join(__dirname, '..', 'حجوزات_استوديو_أثر.xlsx');
    try {
      if (fs.existsSync(parentExcel)) {
        fs.copyFileSync(EXCEL_FILE, parentExcel);
      }
    } catch (e) {}

    console.log(`[Excel] Updated successfully: ${EXCEL_FILE} (${bookings.length} bookings)`);
    return true;
  } catch (err) {
    console.error('[Excel Error]', err);
    // Still ensure JSON backup if possible
    try {
      fs.writeFileSync(JSON_FILE, JSON.stringify(bookings, null, 2), 'utf8');
      return true;
    } catch (e) {
      return false;
    }
  }
}

function loadBookingsFromExcel() {
  // 1. Try reading from JSON database first
  if (fs.existsSync(JSON_FILE)) {
    try {
      const data = JSON.parse(fs.readFileSync(JSON_FILE, 'utf8'));
      if (Array.isArray(data) && data.length > 0) return data;
    } catch (e) {}
  }

  // 2. Fallback to reading from Excel file if JSON does not exist
  if (XLSX && fs.existsSync(EXCEL_FILE)) {
    try {
      const buf = fs.readFileSync(EXCEL_FILE);
      const wb = XLSX.read(buf, { type: 'buffer' });
      const sheetName = wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);
      if (Array.isArray(rows) && rows.length > 0) {
        const reverseStatus = {
          'قيد التأكيد': 'pending',
          'مؤكد': 'confirmed',
          'جاري التسجيل': 'recording',
          'تم التسجيل': 'done',
          'مكتمل ومُسلّم': 'completed',
          'ملغي': 'cancelled'
        };

        const bookings = rows.map((r, i) => {
          let statusKey = 'confirmed';
          const st = String(r['حالة الحجز'] || '').trim();
          for (const [k, v] of Object.entries(reverseStatus)) {
            if (st === k || st === v) { statusKey = v; break; }
          }

          let id = r['رقم الحجز'] || `ath_${i + 1}`;
          if (typeof id === 'string' && id.startsWith('ATH-')) {
            id = 'ath_' + id.replace('ATH-', '').toLowerCase();
          }

          return {
            id: String(id),
            name: r['اسم المنشد / العميل'] || '',
            phone: String(r['رقم الهاتف'] || ''),
            workTitle: r['العمل الفني'] || '',
            service: r['نوع الخدمة'] || 'تسجيل إنشاد ديني',
            maqam: r['المقام الموسيقي'] || 'حر',
            engineer: r['مهندس الصوت'] || 'م/ أحمد فاروق',
            date: r['تاريخ الجلسة'] || new Date().toISOString().slice(0, 10),
            time: r['وقت البداية'] || '17:00',
            duration: String(r['مدة الجلسة (ساعات)'] || 2),
            status: statusKey,
            total: Number(r['السعر الإجمالي (ج.م)']) || 0,
            deposit: Number(r['العربون المدفوع (ج.م)']) || 0,
            notes: r['ملاحظات فنية'] || '',
            createdAt: r['تاريخ التسجيل'] || new Date().toISOString()
          };
        });

        // Save JSON backup immediately
        try {
          fs.writeFileSync(JSON_FILE, JSON.stringify(bookings, null, 2), 'utf8');
        } catch (e) {}
        return bookings;
      }
    } catch (e) {
      console.error('[Excel Load Error]', e);
    }
  }

  return [];
}

const server = http.createServer((req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host}`);
  let pathname = parsedUrl.pathname;

  // CORS headers for local usage
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }

  // API Endpoints
  if (pathname === '/api/status') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ status: 'ok', server: 'Athar Studio Server', time: new Date() }));
    return;
  }

  if (pathname === '/api/bookings' && req.method === 'GET') {
    const data = loadBookingsFromExcel();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(data || []));
    return;
  }

  if (pathname === '/api/save' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      try {
        const bookings = JSON.parse(body);
        if (Array.isArray(bookings)) {
          const success = saveBookingsToExcel(bookings);
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success, count: bookings.length }));
          return;
        }
      } catch (e) {}
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Invalid bookings payload' }));
    });
    return;
  }

  if (pathname === '/api/open-excel' && req.method === 'POST') {
    if (fs.existsSync(EXCEL_FILE)) {
      const psCommand = `powershell -NoProfile -Command "Start-Process -FilePath '${EXCEL_FILE.replace(/'/g, "''")}'"`;
      exec(psCommand, (err) => {
        if (err) {
          exec(`start "" "${EXCEL_FILE}"`, (fallbackErr) => {
            if (fallbackErr) {
              res.writeHead(500, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ success: false, error: fallbackErr.message }));
            } else {
              res.writeHead(200, { 'Content-Type': 'application/json' });
              res.end(JSON.stringify({ success: true }));
            }
          });
        } else {
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: true }));
        }
      });
    } else {
      res.writeHead(404, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Excel file not found' }));
    }
    return;
  }

  // Static File Serving
  if (pathname === '/') pathname = '/index.html';
  let safePath = path.normalize(path.join(__dirname, pathname));
  if (!safePath.startsWith(__dirname)) {
    res.writeHead(403);
    res.end('Access Denied');
    return;
  }

  fs.stat(safePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('الملف غير موجود');
      return;
    }

    const ext = path.extname(safePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': contentType });
    fs.createReadStream(safePath).pipe(res);
  });
});

server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.log(`[Athar Studio] المنفذ ${PORT} قيد الاستخدام بالفعل. المنظومة تعمل مسبقاً على: http://localhost:${PORT}`);
  } else {
    console.error('[Athar Studio Server Error]', err);
  }
});

server.listen(PORT, () => {
  console.log(`
╔════════════════════════════════════════════════════════════════════╗
║                   استوديو أَثَـــــر للإنتاج الفني                 ║
║               ATHAR SOUND STUDIO SERVER IS RUNNING                 ║
╠════════════════════════════════════════════════════════════════════╣
║  • الرابط في المتصفح : http://localhost:${PORT}                      ║
║  • ملف الإكسل المرتبط: حجوزات_استوديو_أثر.xlsx                    ║
║  • يتم حفظ أي حجز جديد تلقائياً في الإكسل على جهازك مباشرة!        ║
╚════════════════════════════════════════════════════════════════════╝
`);
});

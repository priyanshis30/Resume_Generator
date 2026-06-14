const express  = require('express');
const cors     = require('cors');
const ExcelJS  = require('exceljs');
const nodemailer = require('nodemailer');
const rateLimit  = require('express-rate-limit');
const path     = require('path');
const fs       = require('fs');

// ── Load .env in development ───────────────────────────────────────────────────
if (process.env.NODE_ENV !== 'production') {
  try { require('dotenv').config(); } catch (_) {}
}

const app          = express();
const PORT         = process.env.PORT || 3001;
const DATA_DIR     = path.join(__dirname, 'data');
const EXCEL_FILE   = path.join(DATA_DIR, 'feedback.xlsx');
const NOTIFY_EMAIL = 'mujhekyameintohbatakhoon@gmail.com';

// ── Ensure data folder ─────────────────────────────────────────────────────────
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

// ── Middleware ─────────────────────────────────────────────────────────────────
app.set('trust proxy', 1); // needed on Render behind a proxy

app.use(cors({
  origin: '*',          // allow GitHub Pages + local dev
  methods: ['GET', 'POST'],
}));

app.use(express.json({ limit: '50kb' }));

// Rate limiter — max 10 feedback submissions per IP per 15 minutes
const feedbackLimiter = rateLimit({
  windowMs : 15 * 60 * 1000,
  max      : 10,
  message  : { error: 'Too many requests — please try again later.' },
  standardHeaders: true,
  legacyHeaders  : false,
});

// ── Sanitise helper ────────────────────────────────────────────────────────────
function sanitise(str, maxLen = 1000) {
  if (typeof str !== 'string') return '';
  return str.trim().slice(0, maxLen);
}

// ── Excel helpers ──────────────────────────────────────────────────────────────
const HEADER_COLS = [
  { header: 'Timestamp',      key: 'ts',      width: 24 },
  { header: 'Name',           key: 'name',    width: 22 },
  { header: 'Email',          key: 'email',   width: 30 },
  { header: 'Stream',         key: 'stream',  width: 24 },
  { header: 'Hired Status',   key: 'hired',   width: 24 },
  { header: 'Rating (/ 5)',   key: 'rating',  width: 14 },
  { header: 'Message',        key: 'message', width: 60 },
];

async function ensureWorkbook() {
  const wb = new ExcelJS.Workbook();

  // Try to load existing file
  if (fs.existsSync(EXCEL_FILE)) {
    try {
      await wb.xlsx.readFile(EXCEL_FILE);
      // Verify the expected worksheet exists
      if (wb.getWorksheet('Feedback')) return wb;
    } catch (err) {
      console.warn('⚠️  Corrupt workbook — recreating:', err.message);
      // Fall through to create a fresh one
    }
  }

  // Build fresh workbook
  const ws = wb.addWorksheet('Feedback');
  ws.columns = HEADER_COLS;

  const headerRow = ws.getRow(1);
  headerRow.eachCell(cell => {
    cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A2E' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
    cell.border = {
      bottom: { style: 'medium', color: { argb: 'FFC9A84C' } },
    };
  });
  headerRow.height = 28;

  await wb.xlsx.writeFile(EXCEL_FILE);
  return wb;
}

async function appendRow(data) {
  const wb = await ensureWorkbook();
  const ws = wb.getWorksheet('Feedback');

  const rowIndex = ws.rowCount + 1;
  const isEven   = rowIndex % 2 === 0;
  const bgColor  = isEven ? 'FFF7F2EA' : 'FFFFFFFF';

  const row = ws.addRow({
    ts     : new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
    name   : data.name    || '',
    email  : data.email   || '',
    stream : data.stream  || '',
    hired  : data.hired   || '',
    rating : data.rating  ? Number(data.rating) : '',
    message: data.message || '',
  });

  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    cell.alignment = { vertical: 'top', wrapText: true };
    // Highlight the rating cell with gold if present
    if (colNumber === 6 && data.rating) {
      cell.font = { color: { argb: 'FFC9A84C' }, bold: true };
    }
  });

  row.height = 20;
  await wb.xlsx.writeFile(EXCEL_FILE);
  console.log(`✅ Row ${rowIndex} written — ${data.name} (${data.email})`);
}

// ── Fetch all rows as JSON (for stats / admin) ─────────────────────────────────
async function getAllRows() {
  if (!fs.existsSync(EXCEL_FILE)) return [];
  const wb = await ensureWorkbook();
  const ws = wb.getWorksheet('Feedback');
  const rows = [];
  ws.eachRow((row, i) => {
    if (i === 1) return; // skip header
    const [ts, name, email, stream, hired, rating, message] =
      row.values.slice(1); // ExcelJS row.values[0] is undefined
    rows.push({ ts, name, email, stream, hired, rating, message });
  });
  return rows;
}

// ── Email helpers ──────────────────────────────────────────────────────────────
function buildTransporter() {
  return nodemailer.createTransport({
    service : 'gmail',
    auth    : {
      user : process.env.GMAIL_USER,
      pass : process.env.GMAIL_APP_PASSWORD,
    },
  });
}

async function sendNotification(data) {
  if (!process.env.GMAIL_USER || !process.env.GMAIL_APP_PASSWORD) {
    console.warn('⚠️  GMAIL_USER / GMAIL_APP_PASSWORD not set — skipping email.');
    return;
  }

  const stars    = '★'.repeat(Number(data.rating) || 0) + '☆'.repeat(5 - (Number(data.rating) || 0));
  const hiredBadge = data.hired
    ? `<span style="background:#22c55e22;border:1px solid #22c55e55;color:#86efac;padding:2px 12px;border-radius:999px;font-size:0.8rem;">${data.hired}</span>`
    : '—';

  const html = `
  <!DOCTYPE html>
  <html>
  <body style="margin:0;padding:0;background:#080808;font-family:Inter,Arial,sans-serif;">
    <div style="max-width:580px;margin:32px auto;background:#0f0f0f;border-radius:16px;overflow:hidden;border:1px solid rgba(201,168,76,0.2);">

      <!-- Header -->
      <div style="background:linear-gradient(135deg,#7a5e24,#c9a84c,#e4c97e);padding:28px 32px;">
        <p style="margin:0 0 4px;font-size:11px;letter-spacing:0.15em;text-transform:uppercase;color:rgba(8,8,8,0.7);">ResuméCraft · New Submission</p>
        <h2 style="margin:0;font-size:1.5rem;color:#080808;">💬 New Feedback Received</h2>
      </div>

      <!-- Body -->
      <div style="padding:28px 32px;">
        <table style="width:100%;border-collapse:collapse;font-size:0.88rem;color:#f7f2ea;">
          <tr>
            <td style="padding:10px 0;color:#777770;width:110px;vertical-align:top;">👤 Name</td>
            <td style="padding:10px 0;font-weight:700;font-size:1rem;">${sanitise(data.name)}</td>
          </tr>
          <tr style="border-top:1px solid rgba(201,168,76,0.12);">
            <td style="padding:10px 0;color:#777770;vertical-align:top;">✉️ Email</td>
            <td style="padding:10px 0;"><a href="mailto:${sanitise(data.email)}" style="color:#c9a84c;text-decoration:none;">${sanitise(data.email) || '—'}</a></td>
          </tr>
          <tr style="border-top:1px solid rgba(201,168,76,0.12);">
            <td style="padding:10px 0;color:#777770;vertical-align:top;">🎓 Stream</td>
            <td style="padding:10px 0;">${sanitise(data.stream) || '—'}</td>
          </tr>
          <tr style="border-top:1px solid rgba(201,168,76,0.12);">
            <td style="padding:10px 0;color:#777770;vertical-align:top;">🎉 Hired</td>
            <td style="padding:10px 0;">${hiredBadge}</td>
          </tr>
          <tr style="border-top:1px solid rgba(201,168,76,0.12);">
            <td style="padding:10px 0;color:#777770;vertical-align:top;">⭐ Rating</td>
            <td style="padding:10px 0;color:#c9a84c;font-size:1.2rem;letter-spacing:2px;">${stars}</td>
          </tr>
        </table>

        <!-- Message block -->
        <div style="margin-top:20px;background:#161616;border-left:3px solid #c9a84c;padding:16px 20px;border-radius:0 8px 8px 0;">
          <p style="margin:0 0 6px;font-size:10px;letter-spacing:0.12em;text-transform:uppercase;color:#7a5e24;font-weight:600;">Message</p>
          <p style="margin:0;font-size:0.88rem;color:#f7f2ea;line-height:1.75;">${sanitise(data.message, 2000).replace(/\n/g, '<br>')}</p>
        </div>
      </div>

      <!-- Footer -->
      <div style="padding:16px 32px;background:#080808;border-top:1px solid rgba(201,168,76,0.12);display:flex;justify-content:space-between;align-items:center;">
        <span style="font-size:0.72rem;color:#777770;">Received ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })} IST</span>
        <span style="font-size:0.72rem;color:#7a5e24;">ResuméCraft · IGDTUW</span>
      </div>
    </div>
  </body>
  </html>`;

  const transporter = buildTransporter();
  await transporter.sendMail({
    from    : `"ResuméCraft Feedback" <${process.env.GMAIL_USER}>`,
    to      : NOTIFY_EMAIL,
    subject : `💬 [ResuméCraft] Feedback from ${sanitise(data.name) || 'Anonymous'} — ${stars}`,
    html,
    // Plain-text fallback
    text: `New feedback from ${data.name || 'Anonymous'}\n\nEmail: ${data.email}\nStream: ${data.stream}\nHired: ${data.hired}\nRating: ${data.rating}/5\n\nMessage:\n${data.message}`,
  });

  console.log(`📧 Email sent to ${NOTIFY_EMAIL}`);
}

// ── Routes ─────────────────────────────────────────────────────────────────────

// Health check
app.get('/', (_req, res) => {
  res.json({
    status  : '✅ ResuméCraft backend is running',
    version : '2.0.0',
    time    : new Date().toISOString(),
  });
});

// POST /api/feedback — main submission endpoint
app.post('/api/feedback', feedbackLimiter, async (req, res) => {
  const name    = sanitise(req.body.name,    120);
  const email   = sanitise(req.body.email,   200);
  const stream  = sanitise(req.body.stream,  100);
  const hired   = sanitise(req.body.hired,   100);
  const message = sanitise(req.body.message, 2000);
  const rating  = Number(req.body.rating) || 0;

  // Validation
  if (!name)    return res.status(400).json({ error: 'name is required' });
  if (!message) return res.status(400).json({ error: 'message is required' });
  if (rating < 0 || rating > 5) return res.status(400).json({ error: 'rating must be 0–5' });

  const payload = { name, email, stream, hired, rating, message };

  try {
    // Save to Excel first — email failure should not block the response
    await appendRow(payload);
  } catch (excelErr) {
    console.error('❌ Excel write error:', excelErr);
    return res.status(500).json({ error: 'Failed to save feedback. Please try again.' });
  }

  // Send email in background — don't block or fail the response
  sendNotification(payload).catch(err => {
    console.error('❌ Email error (non-fatal):', err.message);
  });

  res.json({ success: true, message: 'Feedback received — thank you!' });
});

// GET /api/stats — summary stats (public, no auth needed)
app.get('/api/stats', async (req, res) => {
  try {
    const rows = await getAllRows();
    const total = rows.length;
    const avgRating = total
      ? (rows.reduce((sum, r) => sum + (Number(r.rating) || 0), 0) / total).toFixed(2)
      : 0;
    const hiredCount = rows.filter(r =>
      typeof r.hired === 'string' && r.hired.includes('Got placed')
    ).length;
    const streamCounts = rows.reduce((acc, r) => {
      if (r.stream) acc[r.stream] = (acc[r.stream] || 0) + 1;
      return acc;
    }, {});

    res.json({ total, avgRating: Number(avgRating), hiredCount, streamCounts });
  } catch (err) {
    console.error('Stats error:', err);
    res.status(500).json({ error: 'Could not compute stats' });
  }
});

// GET /api/all-feedback — all rows as JSON, protected by token
app.get('/api/all-feedback', async (req, res) => {
  if (req.query.token !== process.env.DOWNLOAD_TOKEN) {
    return res.status(403).json({ error: 'Forbidden' });
  }
  try {
    const rows = await getAllRows();
    res.json({ count: rows.length, data: rows });
  } catch (err) {
    res.status(500).json({ error: 'Failed to read data' });
  }
});

// GET /api/download-excel — download the .xlsx file, protected by token
app.get('/api/download-excel', (req, res) => {
  if (req.query.token !== process.env.DOWNLOAD_TOKEN) {
    return res.status(403).json({ error: 'Forbidden — provide ?token=YOUR_TOKEN' });
  }
  if (!fs.existsSync(EXCEL_FILE)) {
    return res.status(404).json({ error: 'No submissions yet — file will be created on first submission.' });
  }
  res.download(EXCEL_FILE, 'ResuméCraft_Feedback.xlsx', err => {
    if (err) console.error('Download error:', err);
  });
});

// GET /api/admin — HTML admin dashboard, protected by token
app.get('/api/admin', async (req, res) => {
  if (req.query.token !== process.env.DOWNLOAD_TOKEN) {
    return res.status(403).send(`
      <html><body style="font-family:sans-serif;background:#080808;color:#f7f2ea;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;">
        <div style="text-align:center;">
          <div style="font-size:2rem;margin-bottom:1rem;">🔒</div>
          <p>Access denied. Add <code style="color:#c9a84c;">?token=YOUR_TOKEN</code> to the URL.</p>
        </div>
      </body></html>`);
  }

  let rows = [];
  let stats = { total: 0, avgRating: 0, hiredCount: 0 };
  try {
    rows = await getAllRows();
    stats.total = rows.length;
    stats.hiredCount = rows.filter(r => typeof r.hired === 'string' && r.hired.includes('Got placed')).length;
    stats.avgRating = rows.length
      ? (rows.reduce((s, r) => s + (Number(r.rating) || 0), 0) / rows.length).toFixed(1)
      : 0;
  } catch (_) {}

  const rowsHTML = rows.slice().reverse().map((r, i) => `
    <tr style="border-bottom:1px solid rgba(201,168,76,0.1);${i % 2 === 0 ? 'background:rgba(255,255,255,0.02)' : ''}">
      <td style="padding:10px 12px;font-size:0.75rem;color:#777770;white-space:nowrap;">${r.ts || ''}</td>
      <td style="padding:10px 12px;font-weight:600;">${r.name || ''}</td>
      <td style="padding:10px 12px;color:#c9a84c;font-size:0.82rem;">${r.email || ''}</td>
      <td style="padding:10px 12px;font-size:0.82rem;">${r.stream || ''}</td>
      <td style="padding:10px 12px;font-size:0.78rem;">${r.hired || ''}</td>
      <td style="padding:10px 12px;color:#c9a84c;font-size:1rem;letter-spacing:1px;">${'★'.repeat(Number(r.rating)||0)}${'☆'.repeat(5-(Number(r.rating)||0))}</td>
      <td style="padding:10px 12px;font-size:0.82rem;max-width:300px;word-break:break-word;">${(r.message||'').replace(/</g,'&lt;')}</td>
    </tr>`).join('');

  res.send(`<!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8"/>
    <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
    <title>ResuméCraft Admin</title>
    <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700&family=Inter:wght@400;500;600&family=DM+Mono&display=swap" rel="stylesheet"/>
    <style>
      *{box-sizing:border-box;margin:0;padding:0;}
      body{background:#080808;color:#f7f2ea;font-family:'Inter',sans-serif;min-height:100vh;}
      .top{background:linear-gradient(135deg,#7a5e24,#c9a84c);padding:24px 36px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:1rem;}
      .top h1{font-family:'Playfair Display',serif;font-size:1.6rem;color:#080808;}
      .top small{font-size:0.75rem;color:rgba(8,8,8,0.65);}
      .stats{display:flex;gap:1.2rem;padding:24px 36px;flex-wrap:wrap;}
      .stat{background:#0f0f0f;border:0.5px solid rgba(201,168,76,0.2);border-radius:12px;padding:18px 24px;min-width:140px;}
      .stat-n{font-family:'DM Mono',monospace;font-size:2rem;color:#c9a84c;}
      .stat-l{font-size:0.72rem;color:#777770;text-transform:uppercase;letter-spacing:0.1em;margin-top:4px;}
      .actions{padding:0 36px 20px;display:flex;gap:10px;flex-wrap:wrap;}
      .btn{display:inline-flex;align-items:center;gap:6px;background:#c9a84c;color:#080808;font-family:'Inter',sans-serif;font-weight:700;font-size:0.82rem;padding:9px 20px;border-radius:999px;text-decoration:none;transition:background 0.2s;}
      .btn:hover{background:#e4c97e;}
      .btn.ghost{background:transparent;border:0.5px solid rgba(201,168,76,0.4);color:#c9a84c;}
      .btn.ghost:hover{background:rgba(201,168,76,0.08);}
      .wrap{padding:0 24px 60px;overflow-x:auto;}
      table{width:100%;border-collapse:collapse;font-size:0.85rem;}
      thead tr{background:#0f0f0f;}
      th{padding:10px 12px;text-align:left;font-size:10px;letter-spacing:0.1em;text-transform:uppercase;color:#7a5e24;font-weight:600;border-bottom:1px solid rgba(201,168,76,0.25);}
      .empty{text-align:center;padding:60px;color:#777770;}
    </style>
  </head>
  <body>
    <div class="top">
      <div>
        <h1>✦ ResuméCraft Admin</h1>
        <small>IGDTUW · Feedback Dashboard</small>
      </div>
      <small>Showing ${rows.length} submission${rows.length !== 1 ? 's' : ''}</small>
    </div>

    <div class="stats">
      <div class="stat"><div class="stat-n">${stats.total}</div><div class="stat-l">Total Submissions</div></div>
      <div class="stat"><div class="stat-n">${stats.avgRating}★</div><div class="stat-l">Avg Rating</div></div>
      <div class="stat"><div class="stat-n">${stats.hiredCount}</div><div class="stat-l">Got Placed 🎉</div></div>
    </div>

    <div class="actions">
      <a class="btn" href="/api/download-excel?token=${req.query.token}">↓ Download Excel</a>
      <a class="btn ghost" href="/api/all-feedback?token=${req.query.token}" target="_blank">View as JSON</a>
    </div>

    <div class="wrap">
      ${rows.length === 0
        ? '<div class="empty">No submissions yet. Share the site to start collecting feedback!</div>'
        : `<table>
            <thead><tr>
              <th>Timestamp</th><th>Name</th><th>Email</th>
              <th>Stream</th><th>Hired</th><th>Rating</th><th>Message</th>
            </tr></thead>
            <tbody>${rowsHTML}</tbody>
           </table>`
      }
    </div>
  </body>
  </html>`);
});

// 404 handler
app.use((_req, res) => {
  res.status(404).json({ error: 'Route not found' });
});

// Global error handler
app.use((err, _req, res, _next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({ error: 'Internal server error' });
});

// ── Start ──────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n✅  ResuméCraft backend running on port ${PORT}`);
  console.log(`   Health : http://localhost:${PORT}/`);
  console.log(`   Admin  : http://localhost:${PORT}/api/admin?token=YOUR_TOKEN`);
  console.log(`   Excel  : http://localhost:${PORT}/api/download-excel?token=YOUR_TOKEN\n`);
});

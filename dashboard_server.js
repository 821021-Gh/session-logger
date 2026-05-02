/**
 * Session Logger Dashboard — Node.js HTTP server (fixed)
 * Run: node dashboard_server.js
 * Open: http://localhost:5000
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const os = require('os');

const PORT = 5000;
const CSV_PATH = path.join(os.homedir(), 'session_log.csv');
const XLSX_PATH = path.join(os.homedir(), 'session_report.xlsx');
const EMAIL_TO = 'spunk.chang@gmail.com';

// ── HTML UI ─────────────────────────────────────────────────────────────────

const HTML = `<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Session Logger</title>
<style>
:root{--bg:#0f1117;--surface:#1a1d27;--card:#222738;--primary:#4f8ef7;--pg:rgba(79,142,247,.25);--green:#4ade80;--yellow:#fbbf24;--red:#f87171;--text:#e8eaed;--dim:#8892a4;--border:rgba(255,255,255,.07)}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI','Microsoft YaHei',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;display:flex;flex-direction:column;align-items:center;padding:28px 16px}
.header{display:flex;align-items:center;gap:14px;margin-bottom:24px;width:100%;max-width:580px}
.hi{width:48px;height:48px;background:var(--pg);border:1px solid var(--primary);border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0}
.hl h1{font-size:20px;font-weight:700}
.hl p{font-size:12px;color:var(--dim);margin-top:2px}
.dash{width:100%;max-width:580px}
.card{background:var(--card);border:1px solid var(--border);border-radius:16px;padding:18px;margin-bottom:14px}
.ch{display:flex;align-items:center;gap:10px;margin-bottom:14px}
.ci{font-size:18px}
.ct{font-size:11px;font-weight:600;color:var(--dim);letter-spacing:.8px;text-transform:uppercase}
.br{display:flex;gap:8px;flex-wrap:wrap}
.btn{flex:1;min-width:80px;padding:10px 8px;border:none;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;transition:all .15s;display:flex;align-items:center;justify-content:center;gap:5px}
.btn:active{transform:scale(.97)}
.bp{background:var(--primary);color:#fff}
.bp:hover{background:#6ba0f9;box-shadow:0 4px 16px var(--pg)}
.bs{background:#1a3a1a;color:var(--green);border:1px solid var(--green)}
.bs:hover{background:#214d21}
.by{background:#2a2200;color:var(--yellow);border:1px solid rgba(251,191,36,.5)}
.by:hover{background:#352c00}
.br2{background:#2a1010;color:var(--red);border:1px solid rgba(248,113,113,.3)}
.br2:hover{background:#351515}
.bg2{background:var(--surface);color:var(--dim);border:1px solid var(--border)}
.bg2:hover{border-color:var(--primary);color:var(--text)}
.bf{width:100%;padding:13px;font-size:14px;flex:unset}
.sr{display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-bottom:1px solid var(--border);font-size:13px}
.sr:last-child{border-bottom:none}
.sr .lbl{color:var(--dim)}
.sr .val{font-weight:600}
.sr .val.g{color:var(--green)}
.sr .val.y{color:var(--yellow)}
.sr .val.r{color:var(--red)}
.ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.75);align-items:center;justify-content:center;z-index:999}
.ov.show{display:flex}
.ob{background:var(--card);border:1px solid var(--border);border-radius:20px;padding:30px;max-width:380px;width:90%;text-align:center}
.spn{width:44px;height:44px;border:3px solid var(--border);border-top-color:var(--primary);border-radius:50%;animation:spin .7s linear infinite;margin:0 auto 18px}
@keyframes spin{to{transform:rotate(360deg)}}
.ok{color:var(--green);font-size:48px;margin-bottom:10px;display:none}
.er{color:var(--red);font-size:48px;margin-bottom:10px;display:none}
.ob h3{font-size:16px;margin-bottom:8px}
.ob p{font-size:12px;color:var(--dim);margin-bottom:0;word-break:break-all}
.ft{margin-top:16px;font-size:11px;color:var(--dim);text-align:center}
pre{font-size:12px;color:var(--text);white-space:pre-wrap;max-height:280px;overflow-y:auto;line-height:1.7;margin-top:8px;padding:4px}
canvas{width:100%!important;max-height:200px;display:block}
#chart-card{display:none}#report-card{display:none}
.ov-close{margin-top:12px;width:100%;padding:10px;border:none;border-radius:10px;background:var(--surface);color:var(--dim);cursor:pointer;font-size:13px;font-weight:600;display:none}
::-webkit-scrollbar{width:4px}::-webkit-scrollbar-thumb{background:var(--border);border-radius:4px}
</style>
</head>
<body>

<div class="header">
  <div class="hi">📊</div>
  <div class="hl"><h1>Session Logger</h1><p>開關機與登入記錄</p></div>
</div>

<div class="dash">
  <div class="card">
    <div class="ch"><span class="ci">⚡</span><span class="ct">系統狀態</span></div>
    <div class="sr"><span class="lbl">上次記錄</span><span class="val" id="s-last">—</span></div>
    <div class="sr"><span class="lbl">排程任務</span><span class="val g" id="s-sched">—</span></div>
    <div class="sr"><span class="lbl">今日筆數</span><span class="val" id="s-count">—</span></div>
    <div class="sr"><span class="lbl">CSV 檔案</span><span class="val g" id="s-csv">—</span></div>
  </div>

  <div class="card">
    <div class="ch"><span class="ci">🖥</span><span class="ct">記錄操作</span></div>
    <div class="br">
      <button class="btn bp" id="btn-log">📝 記錄</button>
      <button class="btn bs" id="btn-report">📊 報告</button>
      <button class="btn by" id="btn-today">📅 今日</button>
      <button class="btn br2" id="btn-clear">🗑 清除</button>
    </div>
  </div>

  <div class="card" id="report-card">
    <div class="ch">
      <span class="ci">📊</span><span class="ct" id="report-title">報告</span>
      <button onclick="document.getElementById('report-card').style.display='none'" style="margin-left:auto;background:none;border:none;color:var(--dim);cursor:pointer;font-size:16px;padding:0">✕</button>
    </div>
    <pre id="report-out"></pre>
  </div>

  <div class="card" id="chart-card">
    <div class="ch">
      <span class="ci">📈</span><span class="ct">使用時長圖表</span>
      <button onclick="document.getElementById('chart-card').style.display='none'" style="margin-left:auto;background:none;border:none;color:var(--dim);cursor:pointer;font-size:16px;padding:0">✕</button>
    </div>
    <canvas id="cht"></canvas>
    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap">
      <button class="btn bg2" id="p7" onclick="setP(7)">近 7 天</button>
      <button class="btn bp" id="p30" onclick="setP(30)">近 30 天</button>
    </div>
  </div>

  <div class="card">
    <div class="ch"><span class="ci">📧</span><span class="ct">寄送報告</span></div>
    <button class="btn bp bf" id="btn-email">📤 產生並寄送 Excel 圖表到 Email</button>
  </div>
</div>

<div class="ft">C:\\Users\\spunk\\session_log.csv · 排程每天 20:00</div>

<div class="ov" id="ov">
  <div class="ob">
    <div class="spn" id="spn"></div>
    <div class="ok" id="ok-icon">✅</div>
    <div class="er" id="er-icon">❌</div>
    <h3 id="ov-title">處理中...</h3>
    <p id="ov-msg">請稍候</p>
    <button class="ov-close" id="ov-close" onclick="document.getElementById('ov').classList.remove('show')">關閉</button>
  </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script>
// ── API helper ──────────────────────────────────────────────────────────────
function apiCall(action, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open('POST', '/api', true);
  xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  xhr.onreadystatechange = function() {
    if (xhr.readyState === 4) {
      if (xhr.status === 200) {
        try { callback(null, JSON.parse(xhr.responseText)); }
        catch(e) { callback('JSON parse error: ' + xhr.responseText); }
      } else {
        callback('HTTP ' + xhr.status + ': ' + xhr.statusText);
      }
    }
  };
  xhr.send('action=' + encodeURIComponent(action));
}

// ── Overlay helpers ───────────────────────────────────────────────────────────
function showOV(t, m) {
  document.getElementById('spn').style.display = 'block';
  document.getElementById('ok-icon').style.display = 'none';
  document.getElementById('er-icon').style.display = 'none';
  document.getElementById('ov-title').textContent = t;
  document.getElementById('ov-msg').textContent = m;
  document.getElementById('ov-close').style.display = 'none';
  document.getElementById('ov').classList.add('show');
}
function showOK(t, m) {
  document.getElementById('spn').style.display = 'none';
  document.getElementById('ok-icon').style.display = 'block';
  document.getElementById('er-icon').style.display = 'none';
  document.getElementById('ov-title').textContent = t;
  document.getElementById('ov-msg').textContent = m;
  document.getElementById('ov-close').style.display = 'block';
}
function showER(t, m) {
  document.getElementById('spn').style.display = 'none';
  document.getElementById('ok-icon').style.display = 'none';
  document.getElementById('er-icon').style.display = 'block';
  document.getElementById('ov-title').textContent = t;
  document.getElementById('ov-msg').textContent = m;
  document.getElementById('ov-close').style.display = 'block';
}

// ── Button handlers ──────────────────────────────────────────────────────────
document.getElementById('btn-log').onclick = function() {
  showOV('處理中', '寫入記錄...');
  apiCall('log', function(err, res) {
    if (err) { showER('錯誤', err); return; }
    showOK('已記錄 ✅', '當前 session 已寫入');
    loadStats();
  });
};

document.getElementById('btn-report').onclick = function() {
  showOV('處理中', '載入報告...');
  apiCall('report', function(err, res) {
    document.getElementById('ov').classList.remove('show');
    document.getElementById('report-card').style.display = 'block';
    document.getElementById('report-title').textContent = 'Session 報告';
    document.getElementById('report-out').textContent = res.data || '無資料';
    loadStats();
  });
};

document.getElementById('btn-today').onclick = function() {
  showOV('處理中', '載入今日...');
  apiCall('today', function(err, res) {
    document.getElementById('ov').classList.remove('show');
    document.getElementById('report-card').style.display = 'block';
    document.getElementById('report-title').textContent = '今日事件';
    document.getElementById('report-out').textContent = res.data || '今日無記錄';
  });
};

document.getElementById('btn-clear').onclick = function() {
  if (!confirm('確定清除所有記錄？此操作不可還原。')) return;
  showOV('處理中', '清除中...');
  apiCall('clear', function(err, res) {
    if (err) { showER('錯誤', err); return; }
    showOK('已清除 ✅', '所有記錄已刪除');
    loadStats();
  });
};

document.getElementById('btn-email').onclick = function() {
  showOV('產生報告', '正在生成 Excel...');
  apiCall('generate', function(err, res) {
    if (err || res.error) {
      showER('產生失敗', res.error || err);
      return;
    }
    showOV('寄送中', '正在發送 Email...');
    apiCall('email', function(err2, res2) {
      if (err2 || res2.error) {
        showER('寄送失敗', res2.error || err2);
      } else {
        showOK('寄送成功 ✅', '報告已發送到 spunk.chang@gmail.com');
      }
    });
  });
};

// ── Chart ────────────────────────────────────────────────────────────────────
var chart = null;
var period = 7;
var chartData = {};

function drawChart(data) {
  chartData = data || {};
  document.getElementById('chart-card').style.display = 'block';
  document.getElementById('report-card').style.display = 'none';

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - period);
  var cs = cutoff.toISOString().slice(0, 10);

  var fd = {};
  for (var d in data) {
    if (d >= cs) fd[d] = data[d];
  }

  var accs = [];
  var accSet = {};
  for (var d in fd) {
    for (var a in fd[d]) {
      if (!accSet[a]) { accSet[a] = true; accs.push(a); }
    }
  }
  var dates = Object.keys(fd).sort();
  var cols = ['#4f8ef7','#4ade80','#fbbf24','#f87171','#9e479e','#60a5fa','#a78bfa','#34d399'];
  var ds = accs.map(function(a, i) {
    return {
      label: a,
      data: dates.map(function(d) { return Math.round((fd[d] && fd[d][a]) || 0); }),
      backgroundColor: cols[i % cols.length] + '99',
      borderColor: cols[i % cols.length],
      borderWidth: 1,
      borderRadius: 4
    };
  });

  var ctx = document.getElementById('cht').getContext('2d');
  if (chart) { chart.destroy(); }
  chart = new Chart(ctx, {
    type: 'bar',
    data: { labels: dates.map(function(d) { return d.slice(5); }), datasets: ds },
    options: {
      responsive: true,
      plugins: { legend: { labels: { color: '#8892a4', font: { size: 11 } } } },
      scales: {
        x: { ticks: { color: '#8892a4', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,.04)' } },
        y: { ticks: { color: '#8892a4', callback: function(v) { return v + 'm'; } }, grid: { color: 'rgba(255,255,255,.04)' }, beginAtZero: true }
      }
    }
  });
}

function setP(d) {
  period = d;
  document.getElementById('p7').className = 'btn ' + (d === 7 ? 'bg2' : 'bp');
  document.getElementById('p30').className = 'btn ' + (d === 30 ? 'bg2' : 'bp');
  drawChart(chartData);
}

// ── Stats ────────────────────────────────────────────────────────────────────
function loadStats() {
  apiCall('stats', function(err, res) {
    if (err) return;
    if (res.last) document.getElementById('s-last').textContent = res.last;
    if (res.count !== undefined) document.getElementById('s-count').textContent = res.count;
    if (res.csv_exists !== undefined) document.getElementById('s-csv').textContent = res.csv_exists ? '正常' : '不存在';
    if (res.sched !== undefined) document.getElementById('s-sched').textContent = res.sched ? '已啟用' : '未設定';
  });
  // Load chart automatically
  apiCall('chart', function(err, res) {
    if (!err && res.data) { chartData = res.data; drawChart(res.data); }
  });
}

// Init
loadStats();
</script>
</body>
</html>`;

// ── CSV Helpers ──────────────────────────────────────────────────────────────

function readCSV() {
  if (!fs.existsSync(CSV_PATH)) return [];
  const content = fs.readFileSync(CSV_PATH, 'utf8');
  const lines = content.trim().split(/\r?\n/);
  if (lines.length < 2) return [];
  const headers = lines[0].split(',');
  return lines.slice(1).map(line => {
    const vals = line.split(',');
    const obj = {};
    headers.forEach((h, i) => obj[h.trim()] = (vals[i] || '').trim());
    return obj;
  });
}

function appendCSV(eventType, account, duration) {
  const exists = fs.existsSync(CSV_PATH);
  const ts = new Date().toISOString().replace('T', ' ').slice(0, 19);
  const line = [ts, eventType, account, duration].join(',');
  if (!exists) {
    fs.writeFileSync(CSV_PATH, 'timestamp,event_type,account,duration_minutes\r\n', 'utf8');
  }
  fs.appendFileSync(CSV_PATH, line + '\r\n', 'utf8');
}

function deleteCSV() {
  if (fs.existsSync(CSV_PATH)) fs.unlinkSync(CSV_PATH);
}

// ── Run Python ───────────────────────────────────────────────────────────────

function runPy(script, args) {
  return new Promise((resolve) => {
    const proc = spawn('python', [script, ...args], { shell: true });
    let out = '', err = '';
    proc.stdout.on('data', d => out += d);
    proc.stderr.on('data', d => err += d);
    proc.on('close', code => resolve({ out, err, code }));
  });
}

// ── API Handler ──────────────────────────────────────────────────────────────

async function handleAPI(action) {
  const rows = readCSV();
  const today = new Date().toISOString().slice(0, 10);

  switch (action) {
    case 'log': {
      const user = os.userInfo().username;
      appendCSV('MANUAL', user, 0);
      return { ok: true };
    }
    case 'report': {
      if (!rows.length) return { data: '尚無記錄。請先執行「記錄」功能。' };
      const accounts = {};
      for (const r of rows) {
        const acc = r.account || '';
        if (acc === 'System' || !acc) continue;
        if (!accounts[acc]) accounts[acc] = { logins: 0, logout: 0, total: 0 };
        if (r.event_type === 'LOGIN') accounts[acc].logins++;
        if (r.event_type === 'LOGOUT') {
          accounts[acc].logout++;
          accounts[acc].total += parseFloat(r.duration_minutes || 0);
        }
      }
      const lines = [];
      lines.push('帳號             登入    登出    總時長');
      lines.push('─'.repeat(42));
      for (const [acc, d] of Object.entries(accounts).sort()) {
        const h = Math.floor(d.total / 60), m = Math.floor(d.total % 60);
        lines.push(`${acc.padEnd(15)} ${String(d.logins).padStart(4)} ${String(d.logout).padStart(4)} ${String(h + 'h ' + m + 'm').padStart(8)}`);
      }
      lines.push('');
      lines.push('最近 15 筆記錄：');
      lines.push('時間                    類型       帳號             時長');
      lines.push('─'.repeat(60));
      for (const r of rows.slice(-15)) {
        lines.push(`${(r.timestamp||'').padEnd(22)} ${(r.event_type||'').padEnd(9)} ${(r.account||'').padEnd(15)} ${(r.duration_minutes||'').padStart(5)}`);
      }
      return { data: lines.join('\n') };
    }
    case 'today': {
      const tr = rows.filter(r => (r.timestamp || '').startsWith(today));
      if (!tr.length) return { data: '今日尚無記錄。' };
      const lines = ['時間                    類型       帳號             時長', '─'.repeat(60)];
      for (const r of tr) lines.push(`${(r.timestamp||'').padEnd(22)} ${(r.event_type||'').padEnd(9)} ${(r.account||'').padEnd(15)} ${(r.duration_minutes||'').padStart(5)}`);
      return { data: lines.join('\n') };
    }
    case 'clear': { deleteCSV(); return { ok: true }; }
    case 'chart': {
      const cd = {};
      for (const r of rows) {
        if (r.event_type === 'LOGOUT' && r.account && r.account !== 'System') {
          const date = (r.timestamp || '').slice(0, 10);
          const acc = r.account;
          const mins = parseFloat(r.duration_minutes || 0);
          if (date) {
            if (!cd[date]) cd[date] = {};
            cd[date][acc] = (cd[date][acc] || 0) + mins;
          }
        }
      }
      return { data: cd };
    }
    case 'generate': {
      const { out, err } = await runPy('C:\\Users\\spunk\\.qclaw-oversea\\workspace\\session_report.py', ['--generate-only']);
      if (out.includes('saved') || out.includes('Excel')) return { ok: true };
      return { error: (err || out || 'Unknown error').slice(0, 200) };
    }
    case 'email': {
      const g1 = await runPy('C:\\Users\\spunk\\.qclaw-oversea\\workspace\\session_report.py', ['--generate-only']);
      if (!g1.out.includes('saved') && !g1.out) return { error: '產生 Excel 失敗: ' + g1.err };
      const g2 = await runPy('C:\\Users\\spunk\\.qclaw-oversea\\workspace\\session_report.py', ['--email-only']);
      if (g2.out.includes('successfully') || g2.out.includes('成功') || g2.out.includes('sent')) return { ok: true };
      return { error: (g2.err || g2.out || '寄送失敗').slice(0, 200) };
    }
    case 'stats': {
      const last = rows.length ? rows[rows.length - 1].timestamp || '—' : '—';
      const count = rows.filter(r => (r.timestamp || '').startsWith(today)).length;
      const csvExists = fs.existsSync(CSV_PATH);
      let sched = false;
      try {
        const so = await runPy('schtasks', ['/query', '/tn', 'SessionReport_Daily']);
        sched = so.out.includes('Ready') || so.out.includes('Running');
      } catch (e) {}
      return { last, count, csv_exists: csvExists, sched };
    }
    default:
      return { error: 'Unknown action: ' + action };
  }
}

// ── HTTP Server ──────────────────────────────────────────────────────────────

const MIME = {
  'html': 'text/html; charset=utf-8',
  'json': 'application/json',
  'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

const server = http.createServer(async (req, res) => {
  const origin = req.headers['origin'] || '*';
  res.setHeader('Access-Control-Allow-Origin', origin);
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204, { 'Content-Type': 'text/plain' });
    res.end();
    return;
  }

  const url = req.url.split('?')[0];

  if (url === '/' || url === '/index.html') {
    res.writeHead(200, { 'Content-Type': MIME.html });
    res.end(HTML);
    return;
  }

  if (url === '/report.xlsx' && fs.existsSync(XLSX_PATH)) {
    res.writeHead(200, { 'Content-Type': MIME.xlsx, 'Content-Disposition': 'attachment; filename=session_report.xlsx' });
    res.end(fs.readFileSync(XLSX_PATH));
    return;
  }

  if (url === '/api' && req.method === 'POST') {
    let body = '';
    req.on('data', c => body += c);
    req.on('end', async () => {
      try {
        const params = new URLSearchParams(body);
        const action = params.get('action') || '';
        const result = await handleAPI(action);
        res.writeHead(200, { 'Content-Type': MIME.json });
        res.end(JSON.stringify(result));
      } catch (e) {
        res.writeHead(500, { 'Content-Type': MIME.json });
        res.end(JSON.stringify({ error: e.message }));
      }
    });
    return;
  }

  res.writeHead(404);
  res.end('Not found');
});

server.listen(PORT, '0.0.0.0', () => {
  console.log('[SessionLogger] Dashboard: http://localhost:' + PORT);
  console.log('[SessionLogger] CSV: ' + CSV_PATH);
});
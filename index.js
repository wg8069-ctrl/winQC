// v5.6 - 新增異常分類欄位
const express = require('express');
const crypto  = require('crypto');
const axios   = require('axios');
const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');
const FormData = require('form-data');

const app = express();
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  res.header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});
app.use(express.json({ limit: '10mb', verify: (req, res, buf) => { req.rawBody = buf; } }));

// ── 環境變數 ──
const LINE_CHANNEL_SECRET       = process.env.LINE_CHANNEL_SECRET;
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const NOTION_TOKEN              = process.env.NOTION_TOKEN;
const NOTION_DATABASE_ID        = process.env.NOTION_DATABASE_ID;
const NOTIFY_USERS              = (process.env.NOTIFY_USERS || '').split(',').filter(Boolean);
const CLOUDINARY_CLOUD          = 'dlpxz4qlh';
const CLOUDINARY_KEY            = '953226455671951';
const CLOUDINARY_SECRET         = process.env.CLOUDINARY_SECRET || 'Bx_qzmiTmGPtSoEPXYpJwvqLQoA';

const sessions = {};

// ── Google Sheets ──
const { google } = require('googleapis');
const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID || '1PDzqFlsjPgHJBhDsB8gwuu3_eRwJ6GS_uV2Sc6ZanXM';
const SHEET_HEADERS = ['異常單號','發生日期','狀態','需求回覆時間','發生單位','責任單位','槍型號','單號','零件名稱','異常分類','異常狀況','訂單數量','異常比例','目前處理狀態','判定','回報人','人工成本(人)','人工成本(時)','行政成本(人)','行政成本(時)','所耗人力成本','異常標註內容','品保處理回覆','異常照片','異常照片2'];

async function getSheets() {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS || '{}');
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
  });
  return google.sheets({ version: 'v4', auth });
}

let SHEET_NAME = null;
async function getSheetName() {
  if (SHEET_NAME) return SHEET_NAME;
  const sheets = await getSheets();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  SHEET_NAME = meta.data.sheets[0].properties.title;
  return SHEET_NAME;
}

async function ensureSheetHeader() {
  try {
    const sheets = await getSheets();
    const name = await getSheetName();
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `${name}!A1:Z1` });
    if (!res.data.values || !res.data.values[0] || res.data.values[0].length === 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID, range: `${name}!A1`,
        valueInputOption: 'RAW', requestBody: { values: [SHEET_HEADERS] }
      });
    }
  } catch (e) { console.error('ensureSheetHeader failed:', e.message); }
}

async function appendToSheet(dataMap, retried) {
  try {
    const sheets = await getSheets();
    const name = await getSheetName();
    const headerRes = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `${name}!1:1` });
    const headers = (headerRes.data.values && headerRes.data.values[0]) ? headerRes.data.values[0] : SHEET_HEADERS;
    const row = headers.map(h => dataMap[h] !== undefined ? dataMap[h] : '');

    // 寫入主工作表
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID, range: `${name}!A1`,
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [row] }
    });
    console.log('Sheet row appended OK');

    // 寫入「備份」工作表
    try {
      await ensureBackupSheet(sheets, headers);
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID, range: `備份!A1`,
        valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [row] }
      });
      console.log('Backup sheet row appended OK');
    } catch (backupErr) {
      console.error('Backup sheet failed (不影響主流程):', backupErr.message);
    }

  } catch (e) {
    console.error('appendToSheet failed:', e.message);
    if (!retried && e.message && (e.message.includes('502') || e.message.includes('503') || e.message.includes('500'))) {
      console.log('Sheet 502/503, retry in 2s...');
      await new Promise(r => setTimeout(r, 2000));
      return appendToSheet(dataMap, true);
    }
  }
}

// 確保「備份」工作表存在，不存在就建立並加標題列
async function ensureBackupSheet(sheets, headers) {
  try {
    await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: `備份!A1` });
  } catch (e) {
    // 工作表不存在，建立它
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: '備份' } } }] }
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID, range: `備份!A1`,
      valueInputOption: 'RAW', requestBody: { values: [headers] }
    });
    console.log('Backup sheet created with headers');
  }
}

ensureSheetHeader();

// ── Helper 函式 ──
function verifySignature(req) {
  const sig  = req.headers['x-line-signature'];
  const hash = crypto.createHmac('sha256', LINE_CHANNEL_SECRET).update(req.rawBody).digest('base64');
  return hash === sig;
}

async function getDisplayName(userId) {
  try {
    const r = await axios.get(`https://api.line.me/v2/bot/profile/${userId}`, {
      headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` }
    });
    console.log('getDisplayName OK:', userId, '->', r.data.displayName);
    return r.data.displayName || '用戶';
  } catch (e) {
    console.error('getDisplayName failed:', userId, e.response?.status, e.response?.data?.message || e.message);
    return '(未知)';
  }
}

async function replyFlex(replyToken) {
  await axios.post('https://api.line.me/v2/bot/message/reply',
    {
      replyToken,
      messages: [{
        type: 'text',
        text: 'WinGun 異常通報 👇',
        quickReply: {
          items: [
            { type: 'action', action: { type: 'uri', label: '📋 建立異常單', uri: 'https://liff.line.me/2009600334-UpN6esDu' } },
            { type: 'action', action: { type: 'uri', label: '🔍 查詢紀錄', uri: 'https://script.google.com/macros/s/AKfycbyfhN8EXhwETXwFY99WCRooWCQvusNNveeW5txuGRVnLLc9tRCBHQxc5fpYUyb03bLgzg/exec' } }
          ]
        }
      }]
    },
    { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` } }
  );
}

// ── 流水號（從 Sheet 讀取避免重啟撞號）──
const dailyCounters = {};

async function genWGNumber() {
  const now = new Date();
  const key = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;

  if (!dailyCounters[key]) {
    dailyCounters[key] = 0;
    try {
      const sheets = await getSheets();
      const name = await getSheetName();
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `${name}!A:A`
      });
      const rows = res.data.values || [];
      const prefix = `WG${key}`;
      rows.forEach(r => {
        if (r[0] && r[0].startsWith(prefix)) {
          const seq = parseInt(r[0].replace(prefix, ''));
          if (seq > dailyCounters[key]) dailyCounters[key] = seq;
        }
      });
      console.log('Daily counter for', key, ':', dailyCounters[key]);
    } catch (e) {
      console.error('Read counter failed:', e.message);
    }
  }

  dailyCounters[key] = dailyCounters[key] + 1;
  return `WG${key}${String(dailyCounters[key]).padStart(2, '0')}`;
}

// ── Cloudinary 上傳（圖片）──
async function uploadToCloudinary(base64Data) {
  try {
    const timestamp = Math.floor(Date.now() / 1000);
    const signature = crypto.createHash('sha1').update(`timestamp=${timestamp}${CLOUDINARY_SECRET}`).digest('hex');
    const formData = new URLSearchParams();
    formData.append('file', base64Data);
    formData.append('timestamp', timestamp);
    formData.append('api_key', CLOUDINARY_KEY);
    formData.append('signature', signature);
    const r = await axios.post(`https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD}/image/upload`, formData.toString());
    return r.data.secure_url;
  } catch (e) { return null; }
}

// ── Cloudinary 上傳（Excel 檔案）──
async function uploadExcelToCloudinary(buffer, filename) {
  try {
    const buf = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
    const timestamp = Math.floor(Date.now() / 1000);
    const signature = crypto.createHash('sha1').update(`timestamp=${timestamp}${CLOUDINARY_SECRET}`).digest('hex');
    const form = new FormData();
    form.append('file', buf, { filename, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    form.append('timestamp', String(timestamp));
    form.append('api_key', CLOUDINARY_KEY);
    form.append('signature', signature);
    const r = await axios.post(`https://api.cloudinary.com/v1_1/${CLOUDINARY_CLOUD}/raw/upload`, form, { headers: form.getHeaders() });
    return r.data.secure_url;
  } catch (e) { return null; }
}

// ── 產生 Excel（動態讀取 template-config.json）──
async function generateAndSendExcel(data, wgNumber, reporterName, photoUrl, photoUrl2) {
  const templatePath = path.join(__dirname, 'template.xlsx');
  if (!fs.existsSync(templatePath)) return { buffer: null, downloadUrl: null, error: 'template.xlsx 不存在' };

  let workbook, ws;
  try {
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);
    ws = workbook.worksheets[0];
  } catch (e) { return { buffer: null, downloadUrl: null, error: 'template.xlsx 讀取失敗' }; }

  // 資料對應表（Excel 匯出不含異常分類）
  const dataMap = {
    '異常單號': wgNumber,
    '發生日期': data.date || '',
    '發生單位': data.unit || '',
    '責任單位': data.resp || '',
    '客戶': data.customer || '',
    '零件名稱': data.product || '',
    '系列別': data.series || '',
    '異常狀況': data.anomaly || '',
    '訂單數量': data.qty ? String(data.qty) : '',
    '單號': data.orderNo || '',
    '異常比例': data.ratio || '',
    '判定': data.judge || '',
    '回報人': reporterName || '',
    '目前處理狀態': data.status || '',
    '人工成本人': data.laborPeople || '',
    '人工成本時': data.laborHours || '',
    '行政成本人': data.adminPeople || '',
    '行政成本時': data.adminHours || '',
    '所耗人力成本': data.laborCost || '',
    '異常標註內容': data.anomalyReply || ''
  };

  // 讀取 template-config.json（優先），否則回退掃描 {{}}
  const configPath = path.join(__dirname, 'template-config.json');
  let tmConfig = null;
  try {
    if (fs.existsSync(configPath)) {
      tmConfig = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    }
  } catch (e) { console.error('讀取 config 失敗:', e.message); }

  if (tmConfig) {
    for (const p of (tmConfig.placeholders || [])) {
      const val = dataMap[p.key];
      if (val !== undefined) {
        ws.getCell(p.cell).value = val;
      }
    }
  } else {
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        if (cell.value && typeof cell.value === 'string' && cell.value.includes('{{')) {
          let text = cell.value;
          for (const [key, val] of Object.entries(dataMap)) {
            text = text.replace(`{{${key}}}`, val);
          }
          cell.value = text;
        }
      });
    });
  }

  // 插入照片
  const fetchBuf = (url) => new Promise((resolve) => {
    try {
      const mod = url.startsWith('https') ? require('https') : require('http');
      mod.get(url, (res) => {
        const chunks = [];
        res.on('data', c => chunks.push(c));
        res.on('end', () => resolve(Buffer.concat(chunks)));
        res.on('error', () => resolve(null));
      }).on('error', () => resolve(null));
    } catch (e) { resolve(null); }
  });

  for (const [url, col, row] of [[photoUrl, 0, 4], [photoUrl2, 6, 4]]) {
    if (url) {
      const imgBuf = await fetchBuf(url);
      if (imgBuf && imgBuf.length > 100) {
        const imgId = workbook.addImage({ buffer: imgBuf, extension: 'jpeg' });
        ws.addImage(imgId, { tl: { col, row }, ext: { width: 360, height: 360 } });
      }
    }
  }

  let buffer;
  try { buffer = await workbook.xlsx.writeBuffer(); } catch (e) { return { error: e.message }; }
  const downloadUrl = await uploadExcelToCloudinary(buffer, `${wgNumber}.xlsx`);
  if (!downloadUrl) return { error: 'Cloudinary 上傳失敗' };

  return { buffer, downloadUrl, error: null };
}

// ── API Routes ──

// LIFF 表單提交
app.post('/api/anomaly', async (req, res) => {
  try {
    const d = req.body;
    let reporterName = d.userId ? await getDisplayName(d.userId) : '(未知)';
    console.log('reporterName:', reporterName);
    const wgNumber = await genWGNumber();
    const today = new Date().toISOString().split('T')[0].replace(/-/g, '/');

    const photoUrl  = d.photoData ? await uploadToCloudinary(d.photoData) : null;
    const photoUrl2 = d.photoData2 ? await uploadToCloudinary(d.photoData2) : null;

    // 1. 寫入 Notion（失敗不影響後續）
    const toText = (v) => [{ text: { content: v ? String(v) : '' } }];
    try {
      await axios.post(`https://api.notion.com/v1/pages`, {
        parent: { database_id: NOTION_DATABASE_ID },
        properties: {
          '異常單號': { title: [{ text: { content: wgNumber } }] },
          '發生日期': { date: { start: new Date().toISOString().split('T')[0] } },
          '零件名稱': { rich_text: toText(d.product) },
          '系列別':   { rich_text: toText(d.series) },
          '發生單位': { rich_text: toText(d.unit) },
          '責任單位': { rich_text: toText(d.resp) },
          '異常分類': { rich_text: toText(d.anomCat) },
          '異常狀況': { rich_text: toText(d.anomaly) },
          '判定':     { rich_text: toText(d.judge) },
          '回報人':   { rich_text: toText(reporterName) }
        }
      }, { headers: { Authorization: `Bearer ${NOTION_TOKEN}`, 'Notion-Version': '2022-06-28' } });
      console.log('Notion OK');
    } catch (notionErr) {
      console.error('Notion failed (繼續執行):', notionErr.response?.data?.message || notionErr.message);
    }

    // 2. 寫入 Google Sheets
    console.log('Writing to Sheet...');
    console.log('異常分類:', d.anomCat, '異常狀況:', d.anomaly);
    console.log('工時資料:', 'laborPeople:', d.laborPeople, 'laborHours:', d.laborHours, 'adminPeople:', d.adminPeople, 'adminHours:', d.adminHours, 'laborCost:', d.laborCost);
    console.log('其他:', 'replyDate:', d.replyDate, 'orderNo:', d.orderNo, 'status:', d.status);
    await appendToSheet({
      '異常單號': wgNumber,
      '發生日期': today,
      '狀態': '未處理',
      '需求回覆時間': d.replyDate || '',
      '發生單位': d.unit || '',
      '責任單位': d.resp || '',
      '槍型號': d.series || '',
      '單號': d.orderNo || '',
      '零件名稱': d.product || '',
      '異常分類': d.anomCat || '',
      '異常狀況': d.anomaly || '',
      '訂單數量': d.qty || '',
      '異常比例': d.ratio || '',
      '目前處理狀態': d.status || '',
      '判定': d.judge || '',
      '回報人': reporterName,
      '人工成本(人)': d.laborPeople || '0',
      '人工成本(時)': d.laborHours || '0',
      '行政成本(人)': d.adminPeople || '0',
      '行政成本(時)': d.adminHours || '0',
      '所耗人力成本': d.laborCost || '',
      '異常標註內容': '',
      '品保處理回覆': '',
      '異常照片': photoUrl || '',
      '異常照片2': photoUrl2 || '',
    });

    // 3. 推播通知
    const msg =
      `【異常通報 ${wgNumber}】\n` +
      `回報人：${reporterName}\n` +
      `📌品名：${d.product || ''}　系列：${d.series || ''}\n` +
      `📍 發生單位：${d.unit || ''}\n` +
      `🏭 責任單位：${d.resp || ''}\n` +
      `⚠️ 異常狀況：${d.anomaly || ''}\n` +
      `📃 訂單數量：${d.qty || ''}　比例：${d.ratio || ''}\n` +
      `🔥 目前處理狀態：${d.status || ''}\n` +
      `📅 日期：${today}`;

    for (const uid of NOTIFY_USERS) {
      await axios.post('https://api.line.me/v2/bot/message/push',
        { to: uid, messages: [{ type: 'text', text: msg }] },
        { headers: { Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}` } }
      ).catch(e => console.error('Push failed:', e.message));
    }

    res.json({ success: true, number: wgNumber, reporter: reporterName });
  } catch (err) {
    console.error('API Error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// 從 Sheet 資料產生 Excel（異常分類不匯出）
app.post('/api/generate-excel-from-sheet', async (req, res) => {
  try {
    const { data } = req.body;
    if (!data) return res.status(400).json({ success: false, error: 'Missing data' });

    const wgNumber     = data['異常單號'] || 'WG_UNKNOWN';
    const reporterName = data['回報人'] || '';
    const photoUrl     = data['異常照片'] || null;
    const photoUrl2    = data['異常照片2'] || null;

    const mapped = {
      date:        data['發生日期'] || '',
      replyDate:   data['需求回覆時間'] || '',
      unit:        data['發生單位'] || '',
      resp:        data['責任單位'] || '',
      customer:    data['客戶'] || '',
      product:     data['零件名稱'] || '',
      series:      data['系列別'] || data['槍型號'] || '',
      orderNo:     data['單號'] || '',
      anomaly:     data['異常狀況'] || '',
      qty:         data['訂單數量'] || '',
      ratio:       data['異常比例'] || '',
      judge:       data['判定'] || '',
      status:      data['狀態'] || data['目前處理狀態'] || '',
      laborPeople: data['人工成本(人)'] || '',
      laborHours:  data['人工成本(時)'] || '',
      adminPeople: data['行政成本(人)'] || '',
      adminHours:  data['行政成本(時)'] || '',
      laborCost:   data['所耗人力成本'] || '',
      anomalyReply: data['異常處理回覆'] || data['異常標註內容'] || '',
    };

    const result = await generateAndSendExcel(mapped, wgNumber, reporterName, photoUrl, photoUrl2);

    if (result.error) {
      res.status(500).json({ success: false, error: result.error });
    } else if (result.downloadUrl) {
      res.json({ success: true, url: result.downloadUrl });
    } else {
      res.status(500).json({ success: false, error: '上傳失敗' });
    }
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/webhook', async (req, res) => {
  if (!verifySignature(req)) return res.status(401).send('Unauthorized');
  res.status(200).send('OK');
  for (const event of (req.body.events || [])) {
    if (event.type === 'message' || event.type === 'follow') await replyFlex(event.replyToken);
  }
});

app.get('/', (req, res) => res.json({ status: 'LINE Bot running ✅' }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Bot started on port ${PORT}`));
